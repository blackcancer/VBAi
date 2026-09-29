#define WIN32_LEAN_AND_MEAN
#define NOMINMAX
#include "NativeTheme.h"
#include "ToolbarDrawing.h"
#include "ImportTable.h"
#include <commctrl.h>
#include <array>

namespace {
constexpr UINT_PTR SubclassId = 0x43565448;
HWND root = nullptr;
DWORD ownerThread = 0, textDraws = 0;
volatile LONG enabled = 0;
bool started = false;
thread_local bool drawing = false;
std::array<HWND, 128> windows{};
ToolbarDrawing renderer;
BOOL WINAPI NativePatBlt(HDC, int, int, int, int, DWORD);
int WINAPI NativeFillRect(HDC, const RECT*, HBRUSH);
int WINAPI NativeSetDIBitsToDevice(HDC, int, int, DWORD, DWORD, int, int, UINT, UINT, const void*, const BITMAPINFO*, UINT);
BOOL WINAPI NativeTextOutA(HDC, int, int, LPCSTR, int);
BOOL WINAPI NativeTextOutW(HDC, int, int, LPCWSTR, int);
BOOL WINAPI NativeExtTextOutA(HDC, int, int, UINT, const RECT*, LPCSTR, UINT, const INT*);
BOOL WINAPI NativeExtTextOutW(HDC, int, int, UINT, const RECT*, LPCWSTR, UINT, const INT*);
int WINAPI NativeDrawTextA(HDC, LPCSTR, int, LPRECT, UINT);
int WINAPI NativeDrawTextW(HDC, LPCWSTR, int, LPRECT, UINT);
#define TARGET(name) {#name, reinterpret_cast<void*>(&::name), reinterpret_cast<void*>(&Native##name)}
ImportTarget targets[]{TARGET(PatBlt), TARGET(FillRect), TARGET(SetDIBitsToDevice), TARGET(TextOutA), TARGET(TextOutW),
    TARGET(ExtTextOutA), TARGET(ExtTextOutW), TARGET(DrawTextA), TARGET(DrawTextW)};
ImportTable imports(targets, sizeof(targets) / sizeof(targets[0]));
bool OwnThread() { return GetCurrentThreadId() == ownerThread; }
bool Eligible(HDC dc) {
    if (!enabled || !OwnThread() || drawing) return false;
    HWND window = WindowFromDC(dc);
    if (!window) return false;
    for (auto item : windows) if (item == window) return true;
    return false;
}
struct DrawGuard { DrawGuard() { drawing = true; } ~DrawGuard() { drawing = false; } };
BOOL WINAPI NativePatBlt(HDC dc, int x, int y, int width, int height, DWORD operation) {
    DWORD error = GetLastError();
    if (Eligible(dc)) {
        DrawGuard guard;
        if (renderer.TryPaint(dc, x, y, width, height, operation)) { SetLastError(error); return TRUE; }
    }
    SetLastError(error); return ::PatBlt(dc, x, y, width, height, operation);
}
int WINAPI NativeFillRect(HDC dc, const RECT* bounds, HBRUSH brush) {
    DWORD error = GetLastError();
    if (Eligible(dc)) {
        DrawGuard guard; int result = 0; SetLastError(error);
        if (renderer.TryFill(dc, bounds, brush, result)) return result;
    }
    SetLastError(error); return ::FillRect(dc, bounds, brush);
}
int WINAPI NativeSetDIBitsToDevice(HDC dc, int x, int y, DWORD width, DWORD height, int sourceX, int sourceY,
    UINT startScan, UINT lines, const void* bits, const BITMAPINFO* info, UINT colors) {
    DWORD error = GetLastError();
    if (Eligible(dc)) {
        DrawGuard guard; int result = 0; SetLastError(error);
        if (renderer.TryDrawDib(dc, x, y, width, height, sourceX, sourceY, startScan, lines, bits, info, colors, result)) return result;
    }
    SetLastError(error); return ::SetDIBitsToDevice(dc, x, y, width, height, sourceX, sourceY, startScan, lines, bits, info, colors);
}
#define PREPARE_TEXT() DWORD incoming = GetLastError(); ToolbarTextPalette palette(dc, Eligible(dc)); \
    if (palette.Changed()) ++textDraws; SetLastError(incoming)
#define FINISH_TEXT() DWORD outgoing = GetLastError(); palette.Restore(); SetLastError(outgoing); return result
#define RENDER_TEXT(name, stringType) BOOL WINAPI Native##name(HDC dc, int x, int y, stringType text, int length) { \
    PREPARE_TEXT(); BOOL result = ::name(dc, x, y, text, length); FINISH_TEXT(); }
RENDER_TEXT(TextOutA, LPCSTR)
RENDER_TEXT(TextOutW, LPCWSTR)
#define EXT_TEXT(name, stringType) BOOL WINAPI Native##name(HDC dc, int x, int y, UINT flags, const RECT* bounds, stringType text, UINT length, const INT* spacing) { \
    PREPARE_TEXT(); BOOL result = ::name(dc, x, y, flags, bounds, text, length, spacing); FINISH_TEXT(); }
EXT_TEXT(ExtTextOutA, LPCSTR)
EXT_TEXT(ExtTextOutW, LPCWSTR)
#define DRAW_TEXT(name, stringType) int WINAPI Native##name(HDC dc, stringType text, int length, LPRECT bounds, UINT flags) { \
    PREPARE_TEXT(); int result = ::name(dc, text, length, bounds, flags); FINISH_TEXT(); }
DRAW_TEXT(DrawTextA, LPCSTR)
DRAW_TEXT(DrawTextW, LPCWSTR)
LRESULT CALLBACK WindowProc(HWND window, UINT message, WPARAM wParam, LPARAM lParam, UINT_PTR id, DWORD_PTR) {
    if (message == WM_NCDESTROY) {
        RemoveWindowSubclass(window, WindowProc, id);
        for (auto& item : windows) if (item == window) item = nullptr;
        if (window == root) VBAiThemeStop();
    }
    return DefSubclassProc(window, message, wParam, lParam);
}
DWORD Register(HWND window) {
    char className[128]{}; GetClassNameA(window, className, sizeof(className));
    if (strcmp(className, "MsoCommandBar")) return ERROR_NOT_SUPPORTED;
    if (!IsWindow(window)) return ERROR_INVALID_WINDOW_HANDLE;
    // Floating popups are outside this renderer's child-toolbar scope.
    if (!IsChild(root, window)) return ERROR_NOT_SUPPORTED;
    if (GetWindowThreadProcessId(window, nullptr) != ownerThread) return ERROR_INVALID_THREAD_ID;
    for (auto item : windows) if (item == window) return ERROR_SUCCESS;
    HWND* available = nullptr;
    for (auto& item : windows) if (!item) { available = &item; break; }
    if (!available) return ERROR_NOT_ENOUGH_MEMORY;
    DWORD error = imports.AddModule(reinterpret_cast<HMODULE>(GetClassLongPtrW(window, GCLP_HMODULE)));
    if (error) return error;
    if (!SetWindowSubclass(window, WindowProc, SubclassId, 0)) return ERROR_INVALID_DATA;
    *available = window; return ERROR_SUCCESS;
}
BOOL CALLBACK FindBars(HWND window, LPARAM parameter) {
    DWORD error = Register(window);
    if (error != ERROR_SUCCESS && error != ERROR_NOT_SUPPORTED) { *reinterpret_cast<DWORD*>(parameter) = error; return FALSE; }
    return TRUE;
}
}
DWORD __cdecl VBAiThemeStart(HWND editor) {
    if (started) return root == editor && OwnThread() && enabled ? ERROR_SUCCESS : ERROR_ALREADY_EXISTS;
    DWORD process = 0;
    DWORD thread = GetWindowThreadProcessId(editor, &process);
    if (!thread || process != GetCurrentProcessId()) return ERROR_INVALID_WINDOW_HANDLE;
    if (thread != GetCurrentThreadId()) return ERROR_INVALID_THREAD_ID;
    char className[128]{}; GetClassNameA(editor, className, sizeof(className));
    if (strcmp(className, "wndclass_desked_gsk")) return ERROR_INVALID_WINDOW_HANDLE;
    HMODULE self;
    if (!GetModuleHandleExW(GET_MODULE_HANDLE_EX_FLAG_FROM_ADDRESS | GET_MODULE_HANDLE_EX_FLAG_PIN,
        reinterpret_cast<LPCWSTR>(&VBAiThemeStart), &self)) return GetLastError();
    root = editor; ownerThread = thread; textDraws = 0; started = true;
    DWORD error = ERROR_SUCCESS;
    try {
        if (!renderer.Initialize()) error = ERROR_NOT_ENOUGH_MEMORY;
        if (!error && !SetWindowSubclass(root, WindowProc, SubclassId, 0)) error = ERROR_INVALID_DATA;
        if (!error) error = imports.AddModule(GetModuleHandleW(L"VBE7.DLL"));
        if (!error) EnumChildWindows(root, FindBars, reinterpret_cast<LPARAM>(&error));
        if (!error) error = imports.Refresh();
    } catch (...) { error = ERROR_NOT_ENOUGH_MEMORY; }
    if (error) { VBAiThemeStop(); return error; }
    InterlockedExchange(&enabled, 1); return ERROR_SUCCESS;
}
DWORD __cdecl VBAiThemeRegister(HWND window) {
    if (!started) return ERROR_INVALID_STATE;
    if (!OwnThread()) return ERROR_INVALID_THREAD_ID;
    try { DWORD error = Register(window); return error ? error : imports.Refresh(); }
    catch (...) { return ERROR_NOT_ENOUGH_MEMORY; }
}
DWORD __cdecl VBAiThemeRefresh() {
    if (!started) return ERROR_INVALID_STATE;
    if (!OwnThread()) return ERROR_INVALID_THREAD_ID;
    try { return imports.Refresh(); } catch (...) { return ERROR_NOT_ENOUGH_MEMORY; }
}
DWORD __cdecl VBAiThemeStop() {
    if (!started) return ERROR_SUCCESS;
    if (!OwnThread()) return ERROR_INVALID_THREAD_ID;
    InterlockedExchange(&enabled, 0);
    DWORD error = imports.Stop();
    for (auto& window : windows) { if (window && IsWindow(window)) RemoveWindowSubclass(window, WindowProc, SubclassId); window = nullptr; }
    if (root && IsWindow(root)) RemoveWindowSubclass(root, WindowProc, SubclassId);
    renderer.Dispose();
    if (!error) { root = nullptr; started = false; }
    return error;
}
DWORD __cdecl VBAiThemeStatus(NativeThemeStatus* status) {
    if (!status || status->size != sizeof(NativeThemeStatus)) return ERROR_INVALID_PARAMETER;
    if (started && ownerThread && !OwnThread()) return ERROR_INVALID_THREAD_ID;
    DWORD count = 0; for (auto window : windows) if (window) ++count;
    *status = {sizeof(NativeThemeStatus), 1, static_cast<DWORD>(enabled), count, imports.Count(), imports.restoredCount,
        renderer.painted, renderer.dibPainted, textDraws, renderer.fillsPainted,
        renderer.unsupported + renderer.dibUnsupported, renderer.failures};
    return ERROR_SUCCESS;
}

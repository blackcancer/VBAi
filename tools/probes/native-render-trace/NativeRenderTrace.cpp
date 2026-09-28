#define WIN32_LEAN_AND_MEAN
#define NOMINMAX
#include "NativeRenderTrace.h"
#include "ToolbarPatternPilot.h"
#include <commctrl.h>
#include <delayimp.h>
#include <intrin.h>
#include <cstdint>
#include <cstdio>
#include <cstring>
#include <memory>
#include <new>
#include <sstream>
#include <string>
#include <vector>

// This diagnostic changes only one selected module's normal IAT pointers.
// Trace-only mode preserves calls. An explicit toolbar-pattern pilot redirects
// PATCOPY through a preallocated DIB before presenting its recolored result.
// No file I/O, heap allocation or COM call occurs in a hook.
namespace {
enum class Api : unsigned {
    SetTextColor, SetBkColor, TextOutA, TextOutW, ExtTextOutA, ExtTextOutW,
    CreateCompatibleDC, CreateCompatibleBitmap, BitBlt, SelectObject, DeleteDC,
    GetDC, GetDCEx, ReleaseDC, BeginPaint, EndPaint,
    FillRect, PatBlt, DrawTextA, DrawTextW, DrawFrameControl, DrawEdge, SetDIBitsToDevice, Count
};
COLORREF WINAPI TraceSetTextColor(HDC, COLORREF);
COLORREF WINAPI TraceSetBkColor(HDC, COLORREF);
BOOL WINAPI TraceTextOutA(HDC, int, int, LPCSTR, int);
BOOL WINAPI TraceTextOutW(HDC, int, int, LPCWSTR, int);
BOOL WINAPI TraceExtTextOutA(HDC, int, int, UINT, const RECT*, LPCSTR, UINT, const INT*);
BOOL WINAPI TraceExtTextOutW(HDC, int, int, UINT, const RECT*, LPCWSTR, UINT, const INT*);
HDC WINAPI TraceCreateCompatibleDC(HDC);
HBITMAP WINAPI TraceCreateCompatibleBitmap(HDC, int, int);
BOOL WINAPI TraceBitBlt(HDC, int, int, int, int, HDC, int, int, DWORD);
HGDIOBJ WINAPI TraceSelectObject(HDC, HGDIOBJ);
BOOL WINAPI TraceDeleteDC(HDC);
HDC WINAPI TraceGetDC(HWND);
HDC WINAPI TraceGetDCEx(HWND, HRGN, DWORD);
int WINAPI TraceReleaseDC(HWND, HDC);
HDC WINAPI TraceBeginPaint(HWND, LPPAINTSTRUCT);
BOOL WINAPI TraceEndPaint(HWND, const PAINTSTRUCT*);
int WINAPI TraceFillRect(HDC, const RECT*, HBRUSH);
BOOL WINAPI TracePatBlt(HDC, int, int, int, int, DWORD);
int WINAPI TraceDrawTextA(HDC, LPCSTR, int, LPRECT, UINT);
int WINAPI TraceDrawTextW(HDC, LPCWSTR, int, LPRECT, UINT);
BOOL WINAPI TraceDrawFrameControl(HDC, LPRECT, UINT, UINT);
BOOL WINAPI TraceDrawEdge(HDC, LPRECT, UINT, UINT);
int WINAPI TraceSetDIBitsToDevice(HDC, int, int, DWORD, DWORD, int, int, UINT, UINT, const void*, const BITMAPINFO*, UINT);

struct Target { const char* name; void* replacement; void* original; };
#define TARGET(name) { #name, reinterpret_cast<void*>(&Trace##name), reinterpret_cast<void*>(&::name) }
Target targets[] = {
    TARGET(SetTextColor), TARGET(SetBkColor), TARGET(TextOutA), TARGET(TextOutW),
    TARGET(ExtTextOutA), TARGET(ExtTextOutW), TARGET(CreateCompatibleDC),
    TARGET(CreateCompatibleBitmap), TARGET(BitBlt), TARGET(SelectObject),
    TARGET(DeleteDC), TARGET(GetDC), TARGET(GetDCEx), TARGET(ReleaseDC),
    TARGET(BeginPaint), TARGET(EndPaint), TARGET(FillRect), TARGET(PatBlt),
    TARGET(DrawTextA), TARGET(DrawTextW), TARGET(DrawFrameControl), TARGET(DrawEdge), TARGET(SetDIBitsToDevice)
};
#define ORIGINAL(name) reinterpret_cast<decltype(&::name)>(targets[static_cast<unsigned>(Api::name)].original)

constexpr UINT_PTR SubclassId = 0x43565452;
constexpr DWORD MaxWindowCount = 512;
constexpr DWORD MaxDcCount = 1024;
struct Patch {
    void** slot; void* original; void* replacement; const char* name; bool restored;
    DWORD initialProtection; bool protectionSaved; bool protectionRestored; HMODULE module;
};
struct ImportEntry { std::string name; bool delayed; void* target; bool selected; HMODULE module; };
struct Frame { HWND window; UINT message; Frame* previous; };
struct DcEntry { HDC dc; HWND owner; HDC source; };
struct WindowEntry { HWND window; DWORD thread; char name[128]; bool removed; };
struct Event {
    unsigned api;
    DWORD thread;
    LONGLONG ticks;
    UINT message;
    HWND context, directWindow, lineageWindow;
    HDC dc, source;
    HGDIOBJ bitmap, sourceBitmap, font, brush;
    void* caller;
    DWORD dcType, foreground, background, brushStyle, brushColor, flags, count;
    DWORD dibDepth, dibCompression, dibHeaderSize;
    int dibWidth, dibHeight;
    bool brushKnown;
    int x, y, width, height, sourceX, sourceY;
    std::uint64_t argument, result;
};
std::unique_ptr<Event[]> events;
DWORD capacity = 0, count = 0, dropped = 0;
DWORD startError = 0;
DWORD ownerThread = 0;
HWND rootWindow = nullptr;
HWND moduleSourceWindow = nullptr;
bool patternMode = false;
ToolbarPatternPilot patternPilot;
HMODULE targetModule = nullptr;
HMODULE toolbarTextModule = nullptr;
std::uintptr_t toolbarTextBase = 0;
DWORD toolbarTextSize = 0, toolbarTextDraws = 0;
std::uintptr_t targetBase = 0;
DWORD targetSize = 0;
std::wstring outputPath, modulePath;
std::vector<Patch> patches;
std::vector<ImportEntry> importInventory;
WindowEntry windows[MaxWindowCount]{};
DWORD windowCount = 0;
DcEntry dcs[MaxDcCount]{};
DWORD dcOverflow = 0;
volatile LONG enabled = 0;
bool sessionExists = false;
thread_local Frame* frame = nullptr;
thread_local bool recording = false;
LARGE_INTEGER frequency{}, startTicks{};

bool IsOwnerThread() { return GetCurrentThreadId() == ownerThread; }
bool IsVbeWindow(HWND window) {
    return window && (window == rootWindow || IsChild(rootWindow, window));
}
HWND Lineage(HDC dc) {
    if (!dc) return nullptr;
    HWND direct = WindowFromDC(dc);
    if (direct) return direct;
    for (const auto& item : dcs) if (item.dc == dc) return item.owner;
    return nullptr;
}
void TrackDc(HDC dc, HWND window, HDC source = nullptr) {
    if (!dc || !IsOwnerThread() || !enabled) return;
    if (!window && frame) window = frame->window;
    if (!window) window = Lineage(source);
    for (auto& item : dcs) if (item.dc == dc) { item.owner = window; item.source = source; return; }
    for (auto& item : dcs) if (!item.dc) { item = {dc, window, source}; return; }
    ++dcOverflow;
}
void ForgetDc(HDC dc) {
    if (!IsOwnerThread()) return;
    for (auto& item : dcs) if (item.dc == dc) item = {};
}
Event* AddEvent(Api api, HDC dc, HDC source, void* caller) {
    if (!enabled || !IsOwnerThread() || recording) return nullptr;
    if (count >= capacity) { ++dropped; return nullptr; }
    recording = true;
    Event* item = &events[count++];
    *item = {};
    item->api = static_cast<unsigned>(api);
    item->thread = ownerThread;
    LARGE_INTEGER tick;
    QueryPerformanceCounter(&tick);
    item->ticks = tick.QuadPart - startTicks.QuadPart;
    item->dc = dc;
    item->source = source;
    item->caller = caller;
    if (frame) { item->context = frame->window; item->message = frame->message; }
    if (dc) {
        item->directWindow = WindowFromDC(dc);
        item->lineageWindow = Lineage(dc);
        item->dcType = GetObjectType(dc);
        item->foreground = GetTextColor(dc);
        item->background = GetBkColor(dc);
        item->bitmap = GetCurrentObject(dc, OBJ_BITMAP);
        item->font = GetCurrentObject(dc, OBJ_FONT);
        item->brush = GetCurrentObject(dc, OBJ_BRUSH);
        LOGBRUSH brush{};
        item->brushKnown = item->brush && GetObjectW(item->brush, sizeof(brush), &brush) == sizeof(brush);
        if (item->brushKnown) {
            item->brushStyle = brush.lbStyle;
            item->brushColor = item->brush == GetStockObject(DC_BRUSH) ? GetDCBrushColor(dc) : brush.lbColor;
        }
    }
    if (source) item->sourceBitmap = GetCurrentObject(source, OBJ_BITMAP);
    recording = false;
    return item;
}
// The hooks capture metadata only after the original operation and restore its
// last-error value after the diagnostic's own harmless metadata queries.
#define RECORD(name, dc, source) DWORD savedError = GetLastError(); Event* item = AddEvent(Api::name, dc, source, _ReturnAddress())
#define FINISH(value) SetLastError(savedError); return value
#define APPLY_TOOLBAR_COLORS(dc) DWORD inputError = GetLastError(); \
    ToolbarTextPalette textPalette(dc, enabled && patternMode && IsOwnerThread() && moduleSourceWindow && WindowFromDC(dc) == moduleSourceWindow); \
    if (textPalette.Changed()) ++toolbarTextDraws; SetLastError(inputError)
#define RESTORE_TOOLBAR_COLORS() DWORD outputError = GetLastError(); textPalette.Restore(); SetLastError(outputError)

COLORREF WINAPI TraceSetTextColor(HDC dc, COLORREF color) {
    COLORREF result = ORIGINAL(SetTextColor)(dc, color);
    RECORD(SetTextColor, dc, nullptr);
    if (item) { item->argument = color; item->result = result; }
    FINISH(result);
}
COLORREF WINAPI TraceSetBkColor(HDC dc, COLORREF color) {
    COLORREF result = ORIGINAL(SetBkColor)(dc, color);
    RECORD(SetBkColor, dc, nullptr);
    if (item) { item->argument = color; item->result = result; }
    FINISH(result);
}
#define TEXT_HOOK(name, charType) \
BOOL WINAPI Trace##name(HDC dc, int x, int y, charType text, int length) { \
    APPLY_TOOLBAR_COLORS(dc); \
    BOOL result = ORIGINAL(name)(dc, x, y, text, length); \
    RESTORE_TOOLBAR_COLORS(); \
    RECORD(name, dc, nullptr); \
    if (item) { item->x = x; item->y = y; item->count = static_cast<DWORD>(length); item->result = result; } \
    FINISH(result); \
}
TEXT_HOOK(TextOutA, LPCSTR)
TEXT_HOOK(TextOutW, LPCWSTR)
#define EXT_TEXT_HOOK(name, charType) \
BOOL WINAPI Trace##name(HDC dc, int x, int y, UINT options, const RECT* rect, charType text, UINT length, const INT* dx) { \
    APPLY_TOOLBAR_COLORS(dc); \
    BOOL result = ORIGINAL(name)(dc, x, y, options, rect, text, length, dx); \
    RESTORE_TOOLBAR_COLORS(); \
    RECORD(name, dc, nullptr); \
    if (item) { item->x = x; item->y = y; item->count = length; item->flags = options; item->result = result; \
        if (rect) { item->width = rect->right - rect->left; item->height = rect->bottom - rect->top; } } \
    FINISH(result); \
}
EXT_TEXT_HOOK(ExtTextOutA, LPCSTR)
EXT_TEXT_HOOK(ExtTextOutW, LPCWSTR)
HDC WINAPI TraceCreateCompatibleDC(HDC source) {
    HDC result = ORIGINAL(CreateCompatibleDC)(source);
    DWORD originalError = GetLastError();
    TrackDc(result, nullptr, source);
    SetLastError(originalError);
    RECORD(CreateCompatibleDC, result, source);
    if (item) item->result = reinterpret_cast<std::uintptr_t>(result);
    FINISH(result);
}
HBITMAP WINAPI TraceCreateCompatibleBitmap(HDC dc, int width, int height) {
    HBITMAP result = ORIGINAL(CreateCompatibleBitmap)(dc, width, height);
    RECORD(CreateCompatibleBitmap, dc, nullptr);
    if (item) { item->width = width; item->height = height; item->result = reinterpret_cast<std::uintptr_t>(result); }
    FINISH(result);
}
BOOL WINAPI TraceBitBlt(HDC dc, int x, int y, int width, int height, HDC source, int sx, int sy, DWORD rop) {
    BOOL result = ORIGINAL(BitBlt)(dc, x, y, width, height, source, sx, sy, rop);
    RECORD(BitBlt, dc, source);
    if (item) { item->x = x; item->y = y; item->width = width; item->height = height;
        item->sourceX = sx; item->sourceY = sy; item->flags = rop; item->result = result; }
    FINISH(result);
}
HGDIOBJ WINAPI TraceSelectObject(HDC dc, HGDIOBJ object) {
    HGDIOBJ result = ORIGINAL(SelectObject)(dc, object);
    RECORD(SelectObject, dc, nullptr);
    if (item) { item->argument = reinterpret_cast<std::uintptr_t>(object);
        item->flags = GetObjectType(object); item->result = reinterpret_cast<std::uintptr_t>(result); }
    FINISH(result);
}
BOOL WINAPI TraceDeleteDC(HDC dc) {
    // Metadata must be read before deletion, while the HDC remains valid.
    DWORD incoming = GetLastError();
    Event* item = AddEvent(Api::DeleteDC, dc, nullptr, _ReturnAddress());
    SetLastError(incoming);
    BOOL result = ORIGINAL(DeleteDC)(dc);
    DWORD savedError = GetLastError();
    if (item) item->result = result;
    if (result) ForgetDc(dc);
    FINISH(result);
}
HDC WINAPI TraceGetDC(HWND window) {
    HDC result = ORIGINAL(GetDC)(window);
    DWORD originalError = GetLastError();
    TrackDc(result, window);
    SetLastError(originalError);
    RECORD(GetDC, result, nullptr);
    if (item) { item->argument = reinterpret_cast<std::uintptr_t>(window); item->result = reinterpret_cast<std::uintptr_t>(result); }
    FINISH(result);
}
HDC WINAPI TraceGetDCEx(HWND window, HRGN region, DWORD flags) {
    HDC result = ORIGINAL(GetDCEx)(window, region, flags);
    DWORD originalError = GetLastError();
    TrackDc(result, window);
    SetLastError(originalError);
    RECORD(GetDCEx, result, nullptr);
    if (item) { item->argument = reinterpret_cast<std::uintptr_t>(window); item->flags = flags; item->result = reinterpret_cast<std::uintptr_t>(result); }
    FINISH(result);
}
int WINAPI TraceReleaseDC(HWND window, HDC dc) {
    DWORD incoming = GetLastError();
    Event* item = AddEvent(Api::ReleaseDC, dc, nullptr, _ReturnAddress());
    SetLastError(incoming);
    int result = ORIGINAL(ReleaseDC)(window, dc);
    DWORD savedError = GetLastError();
    if (item) { item->argument = reinterpret_cast<std::uintptr_t>(window); item->result = result; }
    if (result) ForgetDc(dc);
    FINISH(result);
}
HDC WINAPI TraceBeginPaint(HWND window, LPPAINTSTRUCT paint) {
    HDC result = ORIGINAL(BeginPaint)(window, paint);
    DWORD originalError = GetLastError();
    TrackDc(result, window);
    SetLastError(originalError);
    RECORD(BeginPaint, result, nullptr);
    if (item) { item->argument = reinterpret_cast<std::uintptr_t>(window); item->result = reinterpret_cast<std::uintptr_t>(result);
        if (result && paint) { item->x = paint->rcPaint.left; item->y = paint->rcPaint.top;
            item->width = paint->rcPaint.right - paint->rcPaint.left; item->height = paint->rcPaint.bottom - paint->rcPaint.top; } }
    FINISH(result);
}
BOOL WINAPI TraceEndPaint(HWND window, const PAINTSTRUCT* paint) {
    DWORD incoming = GetLastError();
    HDC dc = paint ? paint->hdc : nullptr;
    Event* item = AddEvent(Api::EndPaint, dc, nullptr, _ReturnAddress());
    SetLastError(incoming);
    BOOL result = ORIGINAL(EndPaint)(window, paint);
    DWORD savedError = GetLastError();
    if (item) { item->argument = reinterpret_cast<std::uintptr_t>(window); item->result = result; }
    if (result) ForgetDc(dc);
    FINISH(result);
}

void RecordRect(Event* item, const RECT& bounds) {
    if (!item) return;
    item->x = bounds.left; item->y = bounds.top;
    item->width = bounds.right - bounds.left; item->height = bounds.bottom - bounds.top;
}
int WINAPI TraceFillRect(HDC dc, const RECT* rect, HBRUSH brush) {
    DWORD incoming = GetLastError();
    int result = 0;
    bool eligible = enabled && patternMode && IsOwnerThread() && moduleSourceWindow && WindowFromDC(dc) == moduleSourceWindow;
    SetLastError(incoming);
    bool painted = eligible && patternPilot.TryFill(dc, rect, brush, result);
    if (!painted) { SetLastError(incoming); result = ORIGINAL(FillRect)(dc, rect, brush); }
    RECORD(FillRect, dc, nullptr);
    if (item) { item->argument = reinterpret_cast<std::uintptr_t>(brush); item->result = result; if (rect) RecordRect(item, *rect); }
    FINISH(result);
}
BOOL WINAPI TracePatBlt(HDC dc, int x, int y, int width, int height, DWORD operation) {
    DWORD incoming = GetLastError();
    bool painted = enabled && patternMode && IsOwnerThread() && moduleSourceWindow &&
        WindowFromDC(dc) == moduleSourceWindow && patternPilot.TryPaint(dc, x, y, width, height, operation);
    SetLastError(incoming);
    BOOL result = painted ? TRUE : ORIGINAL(PatBlt)(dc, x, y, width, height, operation);
    RECORD(PatBlt, dc, nullptr);
    if (item) { item->x = x; item->y = y; item->width = width; item->height = height; item->flags = operation; item->result = result; }
    FINISH(result);
}
int WINAPI TraceSetDIBitsToDevice(HDC dc, int x, int y, DWORD width, DWORD height, int sourceX, int sourceY,
    UINT startScan, UINT lines, const void* bits, const BITMAPINFO* info, UINT colorUse) {
    DWORD incoming = GetLastError();
    int result = 0;
    bool eligible = enabled && patternMode && IsOwnerThread() && moduleSourceWindow && WindowFromDC(dc) == moduleSourceWindow;
    SetLastError(incoming);
    bool painted = eligible && patternPilot.TryDrawDib(dc, x, y, width, height, sourceX, sourceY, startScan, lines, bits, info, colorUse, result);
    if (!painted) { SetLastError(incoming); result = ORIGINAL(SetDIBitsToDevice)(dc, x, y, width, height, sourceX, sourceY, startScan, lines, bits, info, colorUse); }
    RECORD(SetDIBitsToDevice, dc, nullptr);
    if (item) {
        item->x = x; item->y = y; item->width = width; item->height = height;
        item->sourceX = sourceX; item->sourceY = sourceY; item->count = lines; item->argument = startScan; item->flags = colorUse; item->result = result;
        if (info && info->bmiHeader.biSize >= sizeof(BITMAPINFOHEADER)) {
            item->dibDepth = info->bmiHeader.biBitCount; item->dibCompression = info->bmiHeader.biCompression;
            item->dibHeaderSize = info->bmiHeader.biSize; item->dibWidth = info->bmiHeader.biWidth; item->dibHeight = info->bmiHeader.biHeight;
        }
    }
    FINISH(result);
}
// Keep the incoming rectangle: DT_CALCRECT/BF_ADJUST may modify it. Do not
// inspect text, including when a caller passes -1 for a terminated string.
#define DRAW_TEXT_HOOK(name, charType) \
int WINAPI Trace##name(HDC dc, charType text, int length, LPRECT rect, UINT flags) { \
    RECT bounds = rect ? *rect : RECT{}; \
    APPLY_TOOLBAR_COLORS(dc); \
    int result = ORIGINAL(name)(dc, text, length, rect, flags); \
    RESTORE_TOOLBAR_COLORS(); \
    RECORD(name, dc, nullptr); \
    if (item) { RecordRect(item, bounds); item->count = static_cast<DWORD>(length); item->flags = flags; item->result = result; } \
    FINISH(result); \
}
DRAW_TEXT_HOOK(DrawTextA, LPCSTR)
DRAW_TEXT_HOOK(DrawTextW, LPCWSTR)
#define DRAW_CONTROL_HOOK(name) \
BOOL WINAPI Trace##name(HDC dc, LPRECT rect, UINT kind, UINT flags) { \
    RECT bounds = rect ? *rect : RECT{}; \
    BOOL result = ORIGINAL(name)(dc, rect, kind, flags); \
    RECORD(name, dc, nullptr); \
    if (item) { RecordRect(item, bounds); item->argument = kind; item->flags = flags; item->result = result; } \
    FINISH(result); \
}
DRAW_CONTROL_HOOK(DrawFrameControl)
DRAW_CONTROL_HOOK(DrawEdge)

LRESULT CALLBACK WindowProc(HWND window, UINT message, WPARAM wParam, LPARAM lParam, UINT_PTR id, DWORD_PTR) {
    DWORD incomingError = GetLastError();
    Frame current{ window, message, frame };
    if (message == WM_DRAWITEM && lParam) {
        auto item = reinterpret_cast<const DRAWITEMSTRUCT*>(lParam);
        if (item->hwndItem && IsVbeWindow(item->hwndItem)) current.window = item->hwndItem;
    }
    frame = &current;
    SetLastError(incomingError);
    LRESULT result = DefSubclassProc(window, message, wParam, lParam);
    DWORD resultError = GetLastError();
    frame = current.previous;
    if (message == WM_NCDESTROY) {
        if (window == moduleSourceWindow) moduleSourceWindow = nullptr;
        RemoveWindowSubclass(window, WindowProc, id);
        for (DWORD i = 0; i < windowCount; ++i) if (windows[i].window == window) windows[i].removed = true;
    }
    SetLastError(resultError);
    return result;
}
BOOL CALLBACK AddWindow(HWND window, LPARAM) {
    DWORD process;
    DWORD thread = GetWindowThreadProcessId(window, &process);
    if (process != GetCurrentProcessId() || thread != ownerThread) return TRUE;
    if (windowCount >= MaxWindowCount) return FALSE;
    if (!SetWindowSubclass(window, WindowProc, SubclassId, 0)) return FALSE;
    WindowEntry& item = windows[windowCount++];
    item = {};
    item.window = window;
    item.thread = thread;
    GetClassNameA(window, item.name, sizeof(item.name));
    return TRUE;
}

bool ContainsRva(DWORD rva, std::size_t size) {
    return rva < targetSize && size <= targetSize - rva;
}
DWORD DiscoverTable(DWORD lookupBase, DWORD addressBase, bool delayed) {
    auto base = reinterpret_cast<unsigned char*>(targetModule);
    for (DWORD index = 0; index < 65536; ++index) {
        auto offset = static_cast<std::uint64_t>(index) * sizeof(IMAGE_THUNK_DATA64);
        if (lookupBase + offset > MAXDWORD || addressBase + offset > MAXDWORD) return ERROR_BAD_EXE_FORMAT;
        DWORD lookup = lookupBase + static_cast<DWORD>(offset), address = addressBase + static_cast<DWORD>(offset);
        if (!ContainsRva(lookup, sizeof(IMAGE_THUNK_DATA64)) || !ContainsRva(address, sizeof(IMAGE_THUNK_DATA64))) return ERROR_BAD_EXE_FORMAT;
        auto thunk = reinterpret_cast<IMAGE_THUNK_DATA64*>(base + lookup);
        if (!thunk->u1.AddressOfData) return ERROR_SUCCESS;
        if (IMAGE_SNAP_BY_ORDINAL64(thunk->u1.Ordinal)) continue;
        if (thunk->u1.AddressOfData > MAXDWORD - sizeof(WORD)) return ERROR_BAD_EXE_FORMAT;
        DWORD nameRva = static_cast<DWORD>(thunk->u1.AddressOfData) + sizeof(WORD);
        if (!ContainsRva(nameRva, 1)) return ERROR_BAD_EXE_FORMAT;
        const char* name = reinterpret_cast<const char*>(base + nameRva);
        if (strnlen_s(name, targetSize - nameRva) == targetSize - nameRva) return ERROR_BAD_EXE_FORMAT;
        void** slot = reinterpret_cast<void**>(base + address);
        importInventory.push_back({name, delayed, *slot, false, targetModule});
        for (auto& target : targets) if (!strcmp(name, target.name)) {
            // A delay thunk still pointing into the importing image is unresolved.
            // Never force its helper or alter the unload/bound tables.
            auto current = reinterpret_cast<std::uintptr_t>(*slot);
            if (delayed && current >= targetBase && current < targetBase + targetSize) break;
            if (*slot != target.original) return ERROR_ALREADY_EXISTS;
            patches.push_back({slot, *slot, target.replacement, target.name, false, 0, false, false, targetModule});
            importInventory.back().selected = true;
        }
    }
    return ERROR_BAD_EXE_FORMAT;
}
DWORD DiscoverImports() {
    auto base = reinterpret_cast<unsigned char*>(targetModule);
    auto dos = reinterpret_cast<IMAGE_DOS_HEADER*>(base);
    if (dos->e_magic != IMAGE_DOS_SIGNATURE || dos->e_lfanew <= 0 || dos->e_lfanew > 1024 * 1024) return ERROR_BAD_EXE_FORMAT;
    auto nt = reinterpret_cast<IMAGE_NT_HEADERS64*>(base + dos->e_lfanew);
    if (nt->Signature != IMAGE_NT_SIGNATURE || nt->FileHeader.Machine != IMAGE_FILE_MACHINE_AMD64 ||
        nt->OptionalHeader.Magic != IMAGE_NT_OPTIONAL_HDR64_MAGIC) return ERROR_BAD_EXE_FORMAT;
    targetSize = nt->OptionalHeader.SizeOfImage;
    auto directory = nt->OptionalHeader.DataDirectory[IMAGE_DIRECTORY_ENTRY_IMPORT];
    if (!ContainsRva(directory.VirtualAddress, directory.Size) || !directory.Size) return ERROR_BAD_EXE_FORMAT;
    auto descriptors = reinterpret_cast<IMAGE_IMPORT_DESCRIPTOR*>(base + directory.VirtualAddress);
    for (DWORD d = 0; (d + 1) * sizeof(IMAGE_IMPORT_DESCRIPTOR) <= directory.Size; ++d) {
        auto& desc = descriptors[d];
        if (!desc.Name) break;
        if (!desc.OriginalFirstThunk || !desc.FirstThunk) return ERROR_BAD_EXE_FORMAT;
        DWORD error = DiscoverTable(desc.OriginalFirstThunk, desc.FirstThunk, false);
        if (error) return error;
    }
    auto delayDirectory = nt->OptionalHeader.DataDirectory[IMAGE_DIRECTORY_ENTRY_DELAY_IMPORT];
    if (delayDirectory.Size) {
        if (!ContainsRva(delayDirectory.VirtualAddress, delayDirectory.Size)) return ERROR_BAD_EXE_FORMAT;
        auto delays = reinterpret_cast<ImgDelayDescr*>(base + delayDirectory.VirtualAddress);
        for (DWORD d = 0; (d + 1) * sizeof(ImgDelayDescr) <= delayDirectory.Size; ++d) {
            auto& desc = delays[d];
            if (!desc.rvaDLLName) break;
            if (desc.grAttrs != dlattrRva || !desc.rvaINT || !desc.rvaIAT) return ERROR_BAD_EXE_FORMAT;
            DWORD error = DiscoverTable(desc.rvaINT, desc.rvaIAT, true);
            if (error) return error;
        }
    }
    return patches.empty() ? ERROR_PROC_NOT_FOUND : ERROR_SUCCESS;
}
DWORD DiscoverToolbarTextModule() {
    if (!GetModuleHandleExW(0, L"VBE7.DLL", &toolbarTextModule)) return GetLastError();
    if (toolbarTextModule == targetModule) { FreeLibrary(toolbarTextModule); toolbarTextModule = nullptr; return ERROR_SUCCESS; }
    HMODULE primary = targetModule;
    std::uintptr_t primaryBase = targetBase;
    DWORD primarySize = targetSize;
    targetModule = toolbarTextModule; targetBase = reinterpret_cast<std::uintptr_t>(toolbarTextModule);
    DWORD error;
    try { error = DiscoverImports(); }
    catch (...) { targetModule = primary; targetBase = primaryBase; targetSize = primarySize; throw; }
    toolbarTextBase = targetBase; toolbarTextSize = targetSize;
    targetModule = primary; targetBase = primaryBase; targetSize = primarySize;
    return error;
}
void ReleaseToolbarTextModule() {
    if (toolbarTextModule) FreeLibrary(toolbarTextModule);
    toolbarTextModule = nullptr; toolbarTextBase = 0; toolbarTextSize = 0;
}
DWORD ExchangeSlot(Patch& patch, bool restore) {
    DWORD oldProtection;
    if (!VirtualProtect(patch.slot, sizeof(void*), PAGE_READWRITE, &oldProtection)) return GetLastError();
    // A failed protection restore can leave the page writable. A later retry
    // must still restore the protection observed before our first mutation.
    if (!patch.protectionSaved) { patch.initialProtection = oldProtection; patch.protectionSaved = true; }
    void* expected = restore ? patch.replacement : patch.original;
    void* desired = restore ? patch.original : patch.replacement;
    void* previous = InterlockedCompareExchangePointer(reinterpret_cast<void* volatile*>(patch.slot), desired, expected);
    DWORD ignored;
    BOOL protectionRestored = VirtualProtect(patch.slot, sizeof(void*), patch.initialProtection, &ignored);
    DWORD protectionError = protectionRestored ? ERROR_SUCCESS : GetLastError();
    MEMORY_BASIC_INFORMATION information{};
    patch.protectionRestored = protectionRestored && VirtualQuery(patch.slot, &information, sizeof(information)) == sizeof(information) &&
        information.Protect == patch.initialProtection;
    if (restore) patch.restored = (previous == expected || previous == desired) && *patch.slot == desired;
    if (!patch.protectionRestored) return protectionError ? protectionError : ERROR_INVALID_DATA;
    if (previous != expected && !(restore && previous == desired)) return ERROR_INVALID_DATA;
    if (*patch.slot != desired) return ERROR_INVALID_DATA;
    return ERROR_SUCCESS;
}
std::string Utf8(const std::wstring& value) {
    if (value.empty()) return {};
    int size = WideCharToMultiByte(CP_UTF8, 0, value.data(), static_cast<int>(value.size()), nullptr, 0, nullptr, nullptr);
    std::string result(size, '\0');
    WideCharToMultiByte(CP_UTF8, 0, value.data(), static_cast<int>(value.size()), &result[0], size, nullptr, nullptr);
    return result;
}
std::string Quote(const std::string& value) {
    std::ostringstream result;
    result << '"';
    for (unsigned char c : value) {
        if (c == '"' || c == '\\') result << '\\' << c;
        else if (c < 32) { char buffer[7]; sprintf_s(buffer, "\\u%04x", c); result << buffer; }
        else result << c;
    }
    result << '"';
    return result.str();
}
std::string Pointer(const void* value) {
    std::ostringstream result;
    result << '"' << "0x" << std::hex << reinterpret_cast<std::uintptr_t>(value) << '"';
    return result.str();
}
DWORD WriteReport(DWORD stopError) {
    // All serialization and I/O is outside hooks, after they have been disabled.
    std::ostringstream data;
    data << "{\"type\":\"session\",\"schema\":1,\"process\":" << GetCurrentProcessId()
        << ",\"thread\":" << ownerThread << ",\"root\":" << Pointer(rootWindow)
        << ",\"module\":" << Quote(Utf8(modulePath)) << ",\"moduleBase\":" << Pointer(targetModule)
        << ",\"frequency\":" << frequency.QuadPart << ",\"maxEvents\":" << capacity
        << ",\"startError\":" << startError
        << ",\"moduleSourceWindow\":" << Pointer(moduleSourceWindow)
        << ",\"toolbarTextModule\":" << Pointer(toolbarTextModule)
        << ",\"moduleSize\":" << targetSize
        << ",\"colorsChanged\":" << (patternMode ? "true" : "false")
        << ",\"scope\":\"Selected module(s) normal and resolved delay IAT; owner thread only\"}\n";
    for (const auto& entry : importInventory) data << "{\"type\":\"discoveredImport\",\"api\":" << Quote(entry.name)
        << ",\"delayed\":" << (entry.delayed ? "true" : "false") << ",\"target\":" << Pointer(entry.target)
        << ",\"selected\":" << (entry.selected ? "true" : "false") << ",\"module\":" << Pointer(entry.module) << "}\n";
    for (const auto& patch : patches) data << "{\"type\":\"import\",\"api\":" << Quote(patch.name)
        << ",\"slot\":" << Pointer(patch.slot) << ",\"original\":" << Pointer(patch.original)
        << ",\"module\":" << Pointer(patch.module)
        << ",\"restored\":" << (patch.restored ? "true" : "false")
        << ",\"initialProtection\":" << patch.initialProtection
        << ",\"protectionRestored\":" << (patch.protectionRestored ? "true" : "false") << "}\n";
    for (DWORD i = 0; i < windowCount; ++i) data << "{\"type\":\"window\",\"hwnd\":" << Pointer(windows[i].window)
        << ",\"class\":" << Quote(windows[i].name) << ",\"subclassRemoved\":" << (windows[i].removed ? "true" : "false") << "}\n";
    for (DWORD i = 0; i < count; ++i) {
        const Event& item = events[i];
        auto caller = reinterpret_cast<std::uintptr_t>(item.caller);
        data << "{\"type\":\"call\",\"n\":" << i << ",\"api\":" << Quote(targets[item.api].name)
            << ",\"ticks\":" << item.ticks << ",\"thread\":" << item.thread << ",\"message\":" << item.message
            << ",\"contextWindow\":" << Pointer(item.context) << ",\"directWindow\":" << Pointer(item.directWindow)
            << ",\"lineageWindow\":" << Pointer(item.lineageWindow) << ",\"dc\":" << Pointer(item.dc)
            << ",\"sourceDc\":" << Pointer(item.source) << ",\"dcType\":" << item.dcType
            << ",\"bitmap\":" << Pointer(item.bitmap) << ",\"sourceBitmap\":" << Pointer(item.sourceBitmap)
            << ",\"font\":" << Pointer(item.font) << ",\"caller\":" << Pointer(item.caller)
            << ",\"brush\":" << Pointer(item.brush) << ",\"brushKnown\":" << (item.brushKnown ? "true" : "false")
            << ",\"brushStyle\":" << item.brushStyle << ",\"brushColorRef\":" << item.brushColor
            << ",\"callerModule\":" << Pointer(caller >= targetBase && caller < targetBase + targetSize ? targetModule :
                caller >= toolbarTextBase && caller < toolbarTextBase + toolbarTextSize ? toolbarTextModule : nullptr)
            << ",\"callerRva\":" << (caller >= targetBase && caller < targetBase + targetSize ? caller - targetBase :
                caller >= toolbarTextBase && caller < toolbarTextBase + toolbarTextSize ? caller - toolbarTextBase : 0)
            << ",\"foregroundColorRef\":" << item.foreground << ",\"backgroundColorRef\":" << item.background
            << ",\"argument\":" << item.argument << ",\"result\":" << item.result << ",\"flags\":" << item.flags
            << ",\"count\":" << item.count << ",\"x\":" << item.x << ",\"y\":" << item.y
            << ",\"width\":" << item.width << ",\"height\":" << item.height
            << ",\"sourceX\":" << item.sourceX << ",\"sourceY\":" << item.sourceY
            << ",\"dibDepth\":" << item.dibDepth << ",\"dibCompression\":" << item.dibCompression
            << ",\"dibHeaderSize\":" << item.dibHeaderSize << ",\"dibWidth\":" << item.dibWidth << ",\"dibHeight\":" << item.dibHeight << "}\n";
    }
    data << "{\"type\":\"summary\",\"events\":" << count << ",\"dropped\":" << dropped
        << ",\"patternPainted\":" << patternPilot.painted << ",\"patternUnsupported\":" << patternPilot.unsupported
        << ",\"patternFailures\":" << patternPilot.failures
        << ",\"dibPainted\":" << patternPilot.dibPainted << ",\"dibUnsupported\":" << patternPilot.dibUnsupported
        << ",\"toolbarTextDraws\":" << toolbarTextDraws
        << ",\"toolbarFills\":" << patternPilot.fillsPainted
        << ",\"dcTrackingOverflow\":" << dcOverflow << ",\"stopError\":" << stopError << "}\n";
    std::string bytes = data.str();
    HANDLE file = CreateFileW(outputPath.c_str(), GENERIC_WRITE, FILE_SHARE_READ, nullptr, OPEN_EXISTING, FILE_ATTRIBUTE_NORMAL, nullptr);
    if (file == INVALID_HANDLE_VALUE) return GetLastError();
    DWORD written = 0;
    BOOL ok = WriteFile(file, bytes.data(), static_cast<DWORD>(bytes.size()), &written, nullptr);
    DWORD error = ok && written == bytes.size() ? ERROR_SUCCESS : (ok ? ERROR_WRITE_FAULT : GetLastError());
    if (error == ERROR_SUCCESS && !SetEndOfFile(file)) error = GetLastError();
    if (error == ERROR_SUCCESS && !FlushFileBuffers(file)) error = GetLastError();
    CloseHandle(file);
    return error;
}
DWORD RemoveInstrumentation() {
    InterlockedExchange(&enabled, 0);
    DWORD error = ERROR_SUCCESS;
    for (auto& patch : patches) {
        DWORD result = ExchangeSlot(patch, true);
        if (result && !error) error = result;
    }
    for (DWORD i = 0; i < windowCount; ++i) {
        auto& item = windows[i];
        if (item.removed || !IsWindow(item.window)) { item.removed = true; continue; }
        if (GetWindowThreadProcessId(item.window, nullptr) != ownerThread) { if (!error) error = ERROR_INVALID_THREAD_ID; continue; }
        item.removed = RemoveWindowSubclass(item.window, WindowProc, SubclassId) != FALSE;
        if (!item.removed && !error) error = ERROR_INVALID_DATA;
    }
    return error;
}
}

static DWORD StartTrace(HWND editorRoot, HWND sourceWindow, const wchar_t* outputJsonl, DWORD maxEvents, bool renderPatterns = false) {
    if (sessionExists) return ERROR_ALREADY_EXISTS;
    if (!editorRoot || !IsWindow(editorRoot) || !outputJsonl || !*outputJsonl || maxEvents < 64 || maxEvents > 100000) return ERROR_INVALID_PARAMETER;
    DWORD process;
    DWORD thread = GetWindowThreadProcessId(editorRoot, &process);
    if (process != GetCurrentProcessId()) return ERROR_ACCESS_DENIED;
    if (thread != GetCurrentThreadId()) return ERROR_INVALID_THREAD_ID;
    if (renderPatterns) {
        char className[128]{};
        if (!sourceWindow || !GetClassNameA(sourceWindow, className, sizeof(className)) || strcmp(className, "MsoCommandBar"))
            return ERROR_INVALID_WINDOW_HANDLE;
    }
    HMODULE sourceModule = nullptr;
    if (sourceWindow) {
        if (!IsChild(editorRoot, sourceWindow) || GetWindowThreadProcessId(sourceWindow, nullptr) != thread)
            return ERROR_INVALID_WINDOW_HANDLE;
        sourceModule = reinterpret_cast<HMODULE>(GetClassLongPtrW(sourceWindow, GCLP_HMODULE));
        if (!sourceModule) return ERROR_MOD_NOT_FOUND;
    }
    HMODULE self;
    // Keep the pass-through hook code alive even if an already-fetched IAT target
    // is invoked after Stop. No hook can call into an unloaded diagnostic DLL.
    if (!GetModuleHandleExW(GET_MODULE_HANDLE_EX_FLAG_FROM_ADDRESS | GET_MODULE_HANDLE_EX_FLAG_PIN,
        reinterpret_cast<LPCWSTR>(&CodexVbeTraceStart), &self)) return GetLastError();
    if (!GetModuleHandleExW(sourceModule ? GET_MODULE_HANDLE_EX_FLAG_FROM_ADDRESS : 0,
        sourceModule ? reinterpret_cast<LPCWSTR>(sourceModule) : L"VBE7.DLL", &targetModule)) return GetLastError();
    try {
        events.reset(new (std::nothrow) Event[maxEvents]);
        if (!events) { FreeLibrary(targetModule); targetModule = nullptr; return ERROR_NOT_ENOUGH_MEMORY; }
        outputPath = outputJsonl;
        wchar_t path[32768];
        DWORD length = GetModuleFileNameW(targetModule, path, 32768);
        if (!length || length == 32768) { FreeLibrary(targetModule); targetModule = nullptr; return ERROR_BAD_PATHNAME; }
        modulePath.assign(path, length);
        targetBase = reinterpret_cast<std::uintptr_t>(targetModule);
        capacity = maxEvents; count = 0; dropped = 0; dcOverflow = 0; windowCount = 0; startError = 0; toolbarTextDraws = 0;
        ZeroMemory(dcs, sizeof(dcs));
        patches.clear(); patches.reserve(32); importInventory.clear();
        rootWindow = editorRoot; moduleSourceWindow = sourceWindow; ownerThread = thread; patternMode = renderPatterns;
        patternPilot.Dispose(); patternPilot.painted = patternPilot.unsupported = patternPilot.failures = patternPilot.dibPainted = patternPilot.dibUnsupported = patternPilot.fillsPainted = 0;
        // Reserve a new report path. Refuse to overwrite an earlier trace.
        HANDLE file = CreateFileW(outputPath.c_str(), GENERIC_WRITE, FILE_SHARE_READ, nullptr, CREATE_NEW, FILE_ATTRIBUTE_NORMAL, nullptr);
        if (file == INVALID_HANDLE_VALUE) { DWORD error = GetLastError(); FreeLibrary(targetModule); targetModule = nullptr; return error; }
        CloseHandle(file);
        DWORD error = DiscoverImports();
        if (!error && patternMode) error = DiscoverToolbarTextModule();
        if (error) {
            // Report the selected module even when no graphics import exists.
            // No instrumentation has been installed at this stage.
            startError = error;
            WriteReport(0);
            ReleaseToolbarTextModule();
            FreeLibrary(targetModule); targetModule = nullptr; return error;
        }
        sessionExists = true;
        if (patternMode && !patternPilot.Initialize()) error = ERROR_NOT_ENOUGH_MEMORY;
        QueryPerformanceFrequency(&frequency); QueryPerformanceCounter(&startTicks);
        if (!error && (!AddWindow(rootWindow, 0) || !EnumChildWindows(rootWindow, AddWindow, 0))) error = ERROR_INVALID_DATA;
        if (!error) for (auto& patch : patches) { error = ExchangeSlot(patch, false); if (error) break; }
        if (error) {
            DWORD restoreError = RemoveInstrumentation();
            patternPilot.Dispose();
            WriteReport(restoreError ? restoreError : error);
            if (!restoreError) { sessionExists = false; ReleaseToolbarTextModule(); FreeLibrary(targetModule); targetModule = nullptr; }
            return error;
        }
        InterlockedExchange(&enabled, 1);
        return ERROR_SUCCESS;
    } catch (...) {
        if (sessionExists) { RemoveInstrumentation(); patternPilot.Dispose(); return ERROR_NOT_ENOUGH_MEMORY; }
        ReleaseToolbarTextModule();
        if (targetModule) { FreeLibrary(targetModule); targetModule = nullptr; }
        return ERROR_NOT_ENOUGH_MEMORY;
    }
}
DWORD __cdecl CodexVbeTraceStart(HWND editorRoot, const wchar_t* outputJsonl, DWORD maxEvents) {
    return StartTrace(editorRoot, nullptr, outputJsonl, maxEvents);
}
DWORD __cdecl CodexVbeTraceStartForWindow(HWND editorRoot, HWND sourceWindow, const wchar_t* outputJsonl, DWORD maxEvents) {
    if (!sourceWindow) return ERROR_INVALID_WINDOW_HANDLE;
    return StartTrace(editorRoot, sourceWindow, outputJsonl, maxEvents);
}
DWORD __cdecl CodexVbeToolbarPatternStart(HWND editorRoot, HWND sourceWindow, const wchar_t* outputJsonl, DWORD maxEvents) {
    return StartTrace(editorRoot, sourceWindow, outputJsonl, maxEvents, true);
}
DWORD __cdecl CodexVbeTraceStop() {
    if (!sessionExists) return ERROR_INVALID_STATE;
    if (!IsOwnerThread()) return ERROR_INVALID_THREAD_ID;
    DWORD error = RemoveInstrumentation();
    DWORD reportError = ERROR_SUCCESS;
    try {
        reportError = WriteReport(error);
        if (patternMode && !patternPilot.WriteSamples(outputPath) && !reportError) reportError = ERROR_WRITE_FAULT;
    } catch (...) { reportError = ERROR_NOT_ENOUGH_MEMORY; }
    patternPilot.Dispose();
    if (!error) { sessionExists = false; ReleaseToolbarTextModule(); FreeLibrary(targetModule); targetModule = nullptr; }
    return error ? error : reportError;
}
DWORD __cdecl CodexVbeTraceCount() { return IsOwnerThread() ? count : 0; }
DWORD __cdecl CodexVbeTraceDropped() { return IsOwnerThread() ? dropped : 0; }

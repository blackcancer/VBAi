#define WIN32_LEAN_AND_MEAN
#define NOMINMAX
#include <windows.h>
#include <commctrl.h>
#include <cstdio>
#include <cstdint>
#include <cstring>
#include <string>
#include <vector>

using DrawFn = DWORD(__cdecl*)(HWND, HDC);
using PaintFn = void(__cdecl*)(HWND);
using ErrorFn = DWORD(__cdecl*)(HDC);
using StartFn = DWORD(__cdecl*)(HWND, const wchar_t*, DWORD);
using StartWindowFn = DWORD(__cdecl*)(HWND, HWND, const wchar_t*, DWORD);
using StopFn = DWORD(__cdecl*)();
DrawFn draw = nullptr;
PaintFn paint = nullptr;
LRESULT CALLBACK WindowProc(HWND window, UINT message, WPARAM wParam, LPARAM lParam) {
    if (message == WM_PRINTCLIENT) return draw(window, reinterpret_cast<HDC>(wParam));
    if (message == WM_PAINT) { paint(window); return 0; }
    return DefWindowProc(window, message, wParam, lParam);
}
struct ForeignCall { HWND window; StartFn start; StopFn stop; const wchar_t* path; DWORD startResult, stopResult; DrawFn draw; HDC dc; DWORD drawResult; };
DWORD WINAPI ForeignThread(void* argument) {
    auto call = static_cast<ForeignCall*>(argument);
    call->startResult = call->start(call->window, call->path, 64);
    call->stopResult = call->stop();
    // No window dispatch: exercise IAT pass-through on an unrelated thread.
    call->drawResult = call->draw(nullptr, call->dc);
    return 0;
}
std::string ReadFileText(const std::wstring& path) {
    HANDLE file = CreateFileW(path.c_str(), GENERIC_READ, FILE_SHARE_READ, nullptr, OPEN_EXISTING, FILE_ATTRIBUTE_NORMAL, nullptr);
    if (file == INVALID_HANDLE_VALUE) return {};
    DWORD size = GetFileSize(file, nullptr), read = 0;
    std::string text(size, '\0');
    BOOL ok = ReadFile(file, text.empty() ? nullptr : &text[0], size, &read, nullptr);
    CloseHandle(file);
    return ok && read == size ? text : std::string();
}
int Fail(const char* reason, DWORD code = 0) { fprintf(stderr, "FAIL: %s (code=%lu)\n", reason, code); return 1; }

int wmain(int argc, wchar_t** argv) {
    if (argc != 3) return Fail("Pass absolute trace DLL path and fresh output-directory path.");
    std::wstring directory = argv[2];
    if (!CreateDirectoryW(directory.c_str(), nullptr)) return Fail("Output directory must not exist", GetLastError());
    wchar_t exe[32768];
    GetModuleFileNameW(nullptr, exe, 32768);
    std::wstring fixturePath(exe);
    fixturePath.resize(fixturePath.find_last_of(L"\\/"));
    fixturePath += L"\\VBE7.dll";
    HMODULE fixture = LoadLibraryExW(fixturePath.c_str(), nullptr, LOAD_LIBRARY_SEARCH_DLL_LOAD_DIR | LOAD_LIBRARY_SEARCH_DEFAULT_DIRS);
    HMODULE diagnostic = LoadLibraryExW(argv[1], nullptr, LOAD_LIBRARY_SEARCH_DLL_LOAD_DIR | LOAD_LIBRARY_SEARCH_DEFAULT_DIRS);
    if (!fixture || !diagnostic) return Fail("Loading isolated fixtures", GetLastError());
    draw = reinterpret_cast<DrawFn>(GetProcAddress(fixture, "FixtureDraw"));
    paint = reinterpret_cast<PaintFn>(GetProcAddress(fixture, "FixturePaint"));
    auto colorError = reinterpret_cast<ErrorFn>(GetProcAddress(fixture, "FixtureColorLastError"));
    auto start = reinterpret_cast<StartFn>(GetProcAddress(diagnostic, "VBAiTraceStart"));
    auto startWindow = reinterpret_cast<StartWindowFn>(GetProcAddress(diagnostic, "VBAiTraceStartForWindow"));
    auto stop = reinterpret_cast<StopFn>(GetProcAddress(diagnostic, "VBAiTraceStop"));
    auto count = reinterpret_cast<StopFn>(GetProcAddress(diagnostic, "VBAiTraceCount"));
    auto dropped = reinterpret_cast<StopFn>(GetProcAddress(diagnostic, "VBAiTraceDropped"));
    if (!draw || !paint || !colorError || !start || !startWindow || !stop || !count || !dropped) return Fail("Export lookup");
    INITCOMMONCONTROLSEX controls{sizeof(controls), ICC_STANDARD_CLASSES};
    InitCommonControlsEx(&controls);
    WNDCLASSW wc{};
    wc.lpfnWndProc = WindowProc; wc.hInstance = fixture; wc.lpszClassName = L"VBAiTraceSyntheticWindow";
    if (!RegisterClassW(&wc)) return Fail("RegisterClass", GetLastError());
    HWND window = CreateWindowW(wc.lpszClassName, L"Native trace diagnostic", WS_POPUP, -12000, -12000, 320, 100, nullptr, nullptr, wc.hInstance, nullptr);
    HWND child = CreateWindowW(wc.lpszClassName, L"Synthetic child", WS_CHILD, 0, 0, 320, 100, window, nullptr, wc.hInstance, nullptr);
    if (!window || !child) return Fail("Window creation", GetLastError());
    // Resolve BeginPaint/EndPaint before discovery; untouched delay thunks are
    // deliberately not forced by the diagnostic.
    paint(child);
    HDC dc = CreateCompatibleDC(nullptr);
    BITMAPINFO info{};
    info.bmiHeader.biSize = sizeof(BITMAPINFOHEADER); info.bmiHeader.biWidth = 320; info.bmiHeader.biHeight = -100;
    info.bmiHeader.biPlanes = 1; info.bmiHeader.biBitCount = 32; info.bmiHeader.biCompression = BI_RGB;
    void* pixels = nullptr;
    HBITMAP bitmap = CreateDIBSection(dc, &info, DIB_RGB_COLORS, &pixels, nullptr, 0);
    HGDIOBJ oldBitmap = SelectObject(dc, bitmap);
    if (!bitmap || !pixels) return Fail("DIB creation");
    constexpr std::size_t bytes = 320 * 100 * 4;
    memset(pixels, 0, bytes);
    if (SendMessageW(child, WM_PRINTCLIENT, reinterpret_cast<WPARAM>(dc), PRF_CLIENT)) return Fail("Baseline rendering");
    GdiFlush();
    std::vector<unsigned char> baseline(static_cast<unsigned char*>(pixels), static_cast<unsigned char*>(pixels) + bytes);
    DWORD baselineError = colorError(dc);
    std::wstring firstPath = directory + L"\\identity.jsonl";
    if (startWindow(window, nullptr, firstPath.c_str(), 4096) != ERROR_INVALID_WINDOW_HANDLE ||
        startWindow(child, window, firstPath.c_str(), 4096) != ERROR_INVALID_WINDOW_HANDLE)
        return Fail("Source-window ancestry guard");
    DWORD code = startWindow(window, child, firstPath.c_str(), 4096);
    if (code) return Fail("Starting trace", code);
    memset(pixels, 0, bytes);
    if (SendMessageW(child, WM_PRINTCLIENT, reinterpret_cast<WPARAM>(dc), PRF_CLIENT)) { stop(); return Fail("Traced rendering"); }
    GdiFlush();
    if (memcmp(pixels, baseline.data(), bytes)) { stop(); return Fail("Trace changed pixels"); }
    if (colorError(dc) != baselineError) { stop(); return Fail("Trace changed GetLastError"); }
    DWORD beforeForeign = count();
    ForeignCall foreign{window, start, stop, firstPath.c_str(), 0, 0, draw, dc, 0};
    HANDLE worker = CreateThread(nullptr, 0, ForeignThread, &foreign, 0, nullptr);
    if (!worker || WaitForSingleObject(worker, 5000) != WAIT_OBJECT_0) { stop(); return Fail("Foreign thread timed out"); }
    CloseHandle(worker);
    if (foreign.stopResult != ERROR_INVALID_THREAD_ID || foreign.startResult != ERROR_ALREADY_EXISTS || foreign.drawResult) {
        stop(); return Fail("Foreign thread guard/pass-through", foreign.stopResult);
    }
    if (count() != beforeForeign) { stop(); return Fail("Foreign thread was logged"); }
    ShowWindow(window, SW_SHOWNOACTIVATE);
    InvalidateRect(window, nullptr, FALSE);
    UpdateWindow(window);
    code = stop();
    if (code) return Fail("Stop/restoration", code);
    DWORD afterStop = count();
    memset(pixels, 0, bytes);
    if (SendMessageW(child, WM_PRINTCLIENT, reinterpret_cast<WPARAM>(dc), PRF_CLIENT)) return Fail("Post-stop rendering");
    GdiFlush();
    if (memcmp(pixels, baseline.data(), bytes)) return Fail("Restored renderer changed pixels");
    if (count() != afterStop) return Fail("IAT remained active after stop");
    std::string report = ReadFileText(firstPath);
    if (report.empty() || report.find("SYNTHETIC_TRACE_SENTINEL") != std::string::npos) return Fail("Report missing or leaked source text");
    auto unresolved = report.find("\"type\":\"discoveredImport\",\"api\":\"GetDCEx\"");
    if (unresolved == std::string::npos) return Fail("Missing unresolved import inventory");
    std::string unresolvedRow = report.substr(unresolved, report.find('\n', unresolved) - unresolved);
    if (unresolvedRow.find("\"delayed\":true") == std::string::npos || unresolvedRow.find("\"selected\":false") == std::string::npos)
        return Fail("Unresolved delay import was selected");
    for (const char* api : {"SetTextColor", "SetBkColor", "TextOutA", "TextOutW", "ExtTextOutA", "ExtTextOutW", "CreateCompatibleDC", "BitBlt", "BeginPaint", "EndPaint", "FillRect", "PatBlt", "DrawTextA", "DrawTextW", "DrawFrameControl", "DrawEdge"}) {
        std::string needle = std::string("\"type\":\"call\"");
        bool found = false;
        std::size_t p = 0;
        while ((p = report.find(needle, p)) != std::string::npos) {
            auto end = report.find('\n', p);
            if (report.substr(p, end - p).find(std::string("\"api\":\"") + api + "\"") != std::string::npos) { found = true; break; }
            p += needle.size();
        }
        if (!found) return Fail(api);
    }
    if (report.find("\"restored\":false") != std::string::npos || report.find("\"subclassRemoved\":false") != std::string::npos ||
        report.find("\"protectionRestored\":false") != std::string::npos) return Fail("Unrestored instrumentation");
    std::wstring boundedPath = directory + L"\\bounded.jsonl";
    code = start(window, boundedPath.c_str(), 64);
    if (code) return Fail("Restarting trace", code);
    for (int i = 0; i < 100; ++i) SendMessageW(child, WM_PRINTCLIENT, reinterpret_cast<WPARAM>(dc), PRF_CLIENT);
    if (count() != 64 || !dropped()) { stop(); return Fail("Trace was not bounded"); }
    code = stop();
    if (code) return Fail("Bounded trace stop", code);
    // Wrong-thread Start must also fail while no session is active.
    foreign.startResult = 0;
    worker = CreateThread(nullptr, 0, ForeignThread, &foreign, 0, nullptr);
    if (!worker || WaitForSingleObject(worker, 5000) != WAIT_OBJECT_0) return Fail("Foreign start timed out");
    CloseHandle(worker);
    if (foreign.startResult != ERROR_INVALID_THREAD_ID) return Fail("Wrong-thread Start accepted", foreign.startResult);
    SelectObject(dc, oldBitmap); DeleteObject(bitmap); DeleteDC(dc);
    DestroyWindow(window);
    FreeLibrary(fixture);
    // The diagnostic deliberately remains pinned until this process exits.
    printf("PASS: identical pixels before/during/after; preserved error state; actual GDI calls; thread guard; bounded buffer; verified IAT and subclass restoration.\n");
    return 0;
}

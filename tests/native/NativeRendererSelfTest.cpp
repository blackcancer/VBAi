#define WIN32_LEAN_AND_MEAN
#define NOMINMAX
#include <windows.h>
#include <cstdio>
#include <vector>
#include <stdexcept>
#include <cstring>
#include <delayimp.h>
struct Status { DWORD size, abi, active, windows, imports, restored, patterns, images, text, fills, unsupported, failures; };
using WindowCall = DWORD(__cdecl*)(HWND);
using SimpleCall = DWORD(__cdecl*)();
using Query = DWORD(__cdecl*)(Status*);
struct SlotSnapshot { void** address; void* value; DWORD protection; };
std::vector<SlotSnapshot> ReadSlots(HMODULE module) {
    std::vector<SlotSnapshot> result;
    auto base = reinterpret_cast<BYTE*>(module);
    auto dos = reinterpret_cast<IMAGE_DOS_HEADER*>(base);
    auto nt = reinterpret_cast<IMAGE_NT_HEADERS64*>(base + dos->e_lfanew);
    auto read = [&](DWORD lookup, DWORD address) {
        auto names = reinterpret_cast<IMAGE_THUNK_DATA64*>(base + lookup);
        auto slots = reinterpret_cast<IMAGE_THUNK_DATA64*>(base + address);
        for (size_t i = 0; names[i].u1.AddressOfData; ++i) {
            if (IMAGE_SNAP_BY_ORDINAL64(names[i].u1.Ordinal)) continue;
            auto name = reinterpret_cast<IMAGE_IMPORT_BY_NAME*>(base + names[i].u1.AddressOfData)->Name;
            const char* targets[] = {"FillRect", "PatBlt", "SetDIBitsToDevice", "TextOutA", "TextOutW",
                "ExtTextOutA", "ExtTextOutW", "DrawTextA", "DrawTextW"};
            for (auto target : targets) if (!strcmp(name, target)) {
                auto pointer = reinterpret_cast<void**>(&slots[i].u1.Function);
                auto value = reinterpret_cast<BYTE*>(*pointer);
                // Unresolved delay thunks may legitimately resolve later.
                if (value >= base && value < base + nt->OptionalHeader.SizeOfImage) break;
                MEMORY_BASIC_INFORMATION info{};
                if (!VirtualQuery(pointer, &info, sizeof(info))) throw std::runtime_error("VirtualQuery failed");
                result.push_back({pointer, *pointer, info.Protect}); break;
            }
        }
    };
    auto normal = nt->OptionalHeader.DataDirectory[IMAGE_DIRECTORY_ENTRY_IMPORT];
    if (normal.VirtualAddress) {
        auto row = reinterpret_cast<IMAGE_IMPORT_DESCRIPTOR*>(base + normal.VirtualAddress);
        for (; row->Name; ++row) read(row->OriginalFirstThunk, row->FirstThunk);
    }
    auto delayed = nt->OptionalHeader.DataDirectory[IMAGE_DIRECTORY_ENTRY_DELAY_IMPORT];
    if (delayed.VirtualAddress) {
        auto row = reinterpret_cast<ImgDelayDescr*>(base + delayed.VirtualAddress);
        for (; row->rvaDLLName; ++row) read(row->rvaINT, row->rvaIAT);
    }
    return result;
}
bool Restored(const std::vector<SlotSnapshot>& snapshots) {
    for (auto slot : snapshots) {
        MEMORY_BASIC_INFORMATION info{};
        if (*slot.address != slot.value || !VirtualQuery(slot.address, &info, sizeof(info)) ||
            info.Protect != slot.protection) return false;
    }
    return true;
}
SimpleCall stop;
DWORD WINAPI ForeignThread(void* argument) { *static_cast<DWORD*>(argument) = stop(); return 0; }
int Fail(const char* message) { fprintf(stderr, "FAIL: %s (win32=%lu)\n", message, GetLastError()); return 1; }
int wmain(int argc, wchar_t** argv) {
    if (argc != 3) return Fail("Expected native renderer and synthetic VBE7 fixture paths.");
    HMODULE fixture = LoadLibraryExW(argv[2], nullptr, 0x1100); if (!fixture) return Fail("Fixture load");
    HMODULE module = LoadLibraryExW(argv[1], nullptr, 0x1100);
    if (!module) return Fail("Renderer load");
    auto start = reinterpret_cast<WindowCall>(GetProcAddress(module, "VBAiThemeStart"));
    auto add = reinterpret_cast<WindowCall>(GetProcAddress(module, "VBAiThemeRegister"));
    auto refresh = reinterpret_cast<SimpleCall>(GetProcAddress(module, "VBAiThemeRefresh"));
    stop = reinterpret_cast<SimpleCall>(GetProcAddress(module, "VBAiThemeStop"));
    auto query = reinterpret_cast<Query>(GetProcAddress(module, "VBAiThemeStatus"));
    if (!start || !add || !refresh || !stop || !query) return Fail("Exports");
    Status status{sizeof(Status)};
    if (query(&status) || status.abi != 1 || status.active || start(nullptr) == 0) return Fail("Initial ABI/state");
    WNDCLASSW cls{}; cls.lpfnWndProc = DefWindowProcW; cls.hInstance = GetModuleHandleW(nullptr);
    cls.lpszClassName = L"wndclass_desked_gsk"; if (!RegisterClassW(&cls)) return Fail("Root class");
    cls.lpszClassName = L"MsoCommandBar"; if (!RegisterClassW(&cls)) return Fail("Toolbar class");
    HWND root = CreateWindowW(L"wndclass_desked_gsk", L"Isolated renderer test", WS_POPUP, 0, 0, 400, 200, nullptr, nullptr, cls.hInstance, nullptr);
    HWND bar = CreateWindowW(L"MsoCommandBar", L"", WS_CHILD, 0, 0, 320, 30, root, nullptr, cls.hInstance, nullptr);
    if (!root || !bar) return Fail("Windows");
    RECT bounds{0, 0, 20, 20};
    HBRUSH face = CreateSolidBrush(RGB(240, 240, 240));
    HDC dc = GetDC(bar);
    // Resolve calls before hooking; repeat below to verify direct HWND filtering.
    FillRect(dc, &bounds, face);
    auto ownSlots = ReadSlots(GetModuleHandleW(nullptr));
    auto fixtureSlots = ReadSlots(fixture);
    if (ownSlots.empty() || fixtureSlots.empty()) return Fail("Independent IAT snapshot");
    DWORD baseline = GetGuiResources(GetCurrentProcess(), GR_GDIOBJECTS);
    for (int cycle = 0; cycle < 20; ++cycle) {
        if (start(root) || start(root) || query(&status) || !status.active || status.windows != 1 || !status.imports) return Fail("Start");
        HWND popup = CreateWindowW(L"MsoCommandBar", L"", WS_POPUP, 0, 0, 50, 50, root, nullptr, cls.hInstance, nullptr);
        if (!popup || add(popup) != ERROR_NOT_SUPPORTED) return Fail("Floating popup scope");
        DestroyWindow(popup);
        DWORD imports = status.imports;
        DWORD foreign = 0;
        HANDLE thread = CreateThread(nullptr, 0, ForeignThread, &foreign, 0, nullptr);
        if (!thread) return Fail("Thread");
        WaitForSingleObject(thread, INFINITE); CloseHandle(thread);
        if (foreign != ERROR_INVALID_THREAD_ID) return Fail("Foreign thread guard");
        DWORD before = status.fills;
        FillRect(dc, &bounds, face);
        if (query(&status) || status.fills != before + 1) return Fail("Toolbar drawing interception");
        HDC other = GetDC(root); FillRect(other, &bounds, face); ReleaseDC(root, other);
        if (query(&status) || status.fills != before + 1) return Fail("Other HWND pass-through");
        HWND late = CreateWindowW(L"MsoCommandBar", L"", WS_CHILD, 0, 30, 320, 30, root, nullptr, cls.hInstance, nullptr);
        if (!late || add(late) || refresh() || query(&status) || status.windows != 2) return Fail("Late registration");
        DestroyWindow(late);
        if (query(&status) || status.windows != 1) return Fail("Toolbar destruction");
        if (stop() || stop() || query(&status) || status.active || status.windows || status.imports || status.restored < imports) return Fail("Stop/restore");
        if (!Restored(ownSlots) || !Restored(fixtureSlots)) return Fail("IAT addresses or page protections not restored");
        before = status.fills; FillRect(dc, &bounds, face);
        if (query(&status) || status.fills != before) return Fail("Original drawing restored");
        if (GetGuiResources(GetCurrentProcess(), GR_GDIOBJECTS) != baseline) return Fail("GDI resource balance");
    }
    ReleaseDC(bar, dc);
    if (start(root)) return Fail("Final start");
    DestroyWindow(root);
    if (query(&status) || status.active || status.windows || status.imports) return Fail("Root destruction cleanup");
    DeleteObject(face);
    puts("PASS: 20 start/stop cycles; native drawing scope; late toolbar; thread guard; destruction; restored imports; balanced GDI resources.");
    return 0;
}

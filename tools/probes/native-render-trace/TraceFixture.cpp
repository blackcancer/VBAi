#define WIN32_LEAN_AND_MEAN
#define NOMINMAX
#include <windows.h>

// Synthetic VBE7.dll, loaded only by TraceSelfTest.exe in a separate process.
// It never replaces or loads an installed Office/VBA DLL.
extern "C" __declspec(dllexport) DWORD __cdecl FixtureDraw(HWND window, HDC destination) {
    HDC client = GetDC(window);
    if (client) ReleaseDC(window, client);
    HDC buffer = CreateCompatibleDC(destination);
    HBITMAP bitmap = CreateCompatibleBitmap(destination, 320, 100);
    if (!buffer || !bitmap) return 1;
    HGDIOBJ oldBitmap = SelectObject(buffer, bitmap);
    RECT bounds{0, 0, 320, 100};
    HBRUSH brush = CreateSolidBrush(RGB(40, 45, 53));
    FillRect(buffer, &bounds, brush);
    DeleteObject(brush);
    HFONT font = CreateFontW(-15, 0, 0, 0, FW_NORMAL, FALSE, FALSE, FALSE,
        DEFAULT_CHARSET, OUT_DEFAULT_PRECIS, CLIP_DEFAULT_PRECIS, CLEARTYPE_QUALITY, FIXED_PITCH, L"Consolas");
    HGDIOBJ oldFont = SelectObject(buffer, font);
    SetBkColor(buffer, RGB(40, 45, 53));
    SetTextColor(buffer, RGB(220, 220, 220));
    SetBkMode(buffer, OPAQUE);
    BOOL a = TextOutA(buffer, 4, 3, "SYNTHETIC_TRACE_SENTINEL", 24);
    BOOL w = TextOutW(buffer, 4, 23, L"Wide diagnostic fixture", 23);
    RECT clip{0, 43, 310, 63};
    SetTextColor(buffer, RGB(86, 156, 214));
    BOOL ea = ExtTextOutA(buffer, 4, 43, ETO_OPAQUE | ETO_CLIPPED, &clip, "Public Sub Fixture()", 20, nullptr);
    clip.top = 63; clip.bottom = 83;
    SetTextColor(buffer, RGB(87, 166, 74));
    BOOL ew = ExtTextOutW(buffer, 4, 63, ETO_OPAQUE | ETO_CLIPPED, &clip, L"' Comment", 9, nullptr);
    RECT button{260, 3, 310, 30};
    BOOL frame = DrawFrameControl(buffer, &button, DFC_BUTTON, DFCS_BUTTONPUSH);
    BOOL edge = DrawEdge(buffer, &button, EDGE_RAISED, BF_RECT | BF_ADJUST);
    int ta = DrawTextA(buffer, "A", 1, &button, DT_CENTER | DT_SINGLELINE);
    int tw = DrawTextW(buffer, L"W", 1, &button, DT_RIGHT | DT_SINGLELINE);
    BOOL pattern = PatBlt(buffer, 310, 0, 10, 100, BLACKNESS);
    BOOL blit = BitBlt(destination, 0, 0, 320, 100, buffer, 0, 0, SRCCOPY);
    SelectObject(buffer, oldFont);
    DeleteObject(font);
    SelectObject(buffer, oldBitmap);
    DeleteObject(bitmap);
    DeleteDC(buffer);
    return a && w && ea && ew && frame && edge && ta && tw && pattern && blit ? 0 : 2;
}
// Keep one matching delayed API unresolved when Start discovers the IAT.
extern "C" __declspec(dllexport) void __cdecl FixtureLateDc(HWND window) {
    HDC extended = GetDCEx(window, nullptr, DCX_CACHE | DCX_CLIPSIBLINGS);
    if (extended) ReleaseDC(window, extended);
}
extern "C" __declspec(dllexport) DWORD __cdecl FixtureColorLastError(HDC dc) {
    SetLastError(0x13572468);
    SetTextColor(dc, RGB(71, 89, 103));
    return GetLastError();
}
extern "C" __declspec(dllexport) void __cdecl FixturePaint(HWND window) {
    PAINTSTRUCT paint{};
    HDC dc = BeginPaint(window, &paint);
    if (dc) FixtureDraw(window, dc);
    EndPaint(window, &paint);
}

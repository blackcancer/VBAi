#define WIN32_LEAN_AND_MEAN
#define NOMINMAX
#include "ToolbarPatternPilot.h"
#include <cstdio>
#include <vector>

struct Surface {
    HDC dc = CreateCompatibleDC(nullptr);
    HBITMAP bitmap = nullptr;
    HGDIOBJ previous = nullptr;
    DWORD* pixels = nullptr;
    Surface(int width = 64, int height = 32) {
        BITMAPINFO info{};
        info.bmiHeader.biSize = sizeof(BITMAPINFOHEADER); info.bmiHeader.biWidth = width; info.bmiHeader.biHeight = -height;
        info.bmiHeader.biPlanes = 1; info.bmiHeader.biBitCount = 32; info.bmiHeader.biCompression = BI_RGB;
        bitmap = CreateDIBSection(dc, &info, DIB_RGB_COLORS, reinterpret_cast<void**>(&pixels), nullptr, 0);
        if (bitmap) previous = SelectObject(dc, bitmap);
    }
    ~Surface() { if (previous) SelectObject(dc, previous); if (bitmap) DeleteObject(bitmap); if (dc) DeleteDC(dc); }
};
int Fail(const char* reason) { fprintf(stderr, "FAIL: %s\n", reason); return 1; }
int main() {
    if (ToolbarPatternPilot::Map(0x0020242b) != 0x0020242b) return Fail("Existing dark background became light");
    for (DWORD neutral = 0; neutral < 256; ++neutral) {
        DWORD mapped = ToolbarPatternPilot::Map(0x7f000000 | neutral * 0x010101);
        if (ToolbarPatternPilot::Map(mapped) != mapped || (mapped & 0xff000000) != 0x7f000000)
            return Fail("Repeated palette conversion changes output or alpha");
    }
    Surface native, themed, tile;
    if (!native.pixels || !themed.pixels || !tile.pixels) return Fail("DIB allocation");
    // The pattern includes neutral glyphs and saturated colors. Its source
    // bitmap is discarded to exercise the brush's retained pattern copy.
    for (int y = 0; y < 32; ++y) for (int x = 0; x < 64; ++x)
        tile.pixels[y * 64 + x] = x % 4 == 0 ? 0x00ffffff : x % 4 == 1 ? 0x00000000 : x % 4 == 2 ? 0x00808080 : 0x00cc4020;
    HBRUSH brush = CreatePatternBrush(tile.bitmap);
    if (!brush) return Fail("Pattern allocation");
    SelectObject(tile.dc, tile.previous); DeleteObject(tile.bitmap);
    tile.previous = nullptr; tile.bitmap = nullptr; tile.pixels = nullptr;
    HGDIOBJ oldNative = SelectObject(native.dc, brush), oldThemed = SelectObject(themed.dc, brush);
    // Establish destination-owned clip copies before counting pilot resources.
    HRGN initialClip = CreateRectRgn(8, 4, 38, 22);
    SelectClipRgn(native.dc, initialClip); SelectClipRgn(themed.dc, initialClip); DeleteObject(initialClip);
    HBRUSH face = CreateSolidBrush(RGB(240, 240, 240));
    if (!face) return Fail("Face fixture allocation");
    DWORD resourcesBefore = GetGuiResources(GetCurrentProcess(), GR_GDIOBJECTS);
    fprintf(stderr, "GDI baseline=%lu\n", resourcesBefore);
    static ToolbarPatternPilot pilot;
    if (!pilot.Initialize()) return Fail("Pilot allocation");
    constexpr DWORD sentinel = 0x0066bbdd;
    for (int pass = 0; pass < 100; ++pass) {
        GdiFlush();
        std::fill(native.pixels, native.pixels + 64 * 32, sentinel);
        std::fill(themed.pixels, themed.pixels + 64 * 32, sentinel);
        for (HDC dc : {native.dc, themed.dc}) {
            SetViewportOrgEx(dc, 3, 2, nullptr); SetWindowOrgEx(dc, 1, -2, nullptr);
            SetBrushOrgEx(dc, (pass % 7) - 4, (pass % 5) - 2, nullptr);
            HRGN clip = CreateRectRgn(8, 4, 38, 22); SelectClipRgn(dc, clip); DeleteObject(clip);
        }
        POINT brushBefore{}, brushAfter{}, viewportBefore{}, viewportAfter{};
        GetBrushOrgEx(themed.dc, &brushBefore); GetViewportOrgEx(themed.dc, &viewportBefore);
        if (!PatBlt(native.dc, 4, 3, 40, 24, PATCOPY) || !pilot.TryPaint(themed.dc, 4, 3, 40, 24, PATCOPY)) return Fail("Pattern rendering");
        GdiFlush();
        for (int index = 0; index < 64 * 32; ++index) {
            DWORD expected = native.pixels[index] == sentinel ? sentinel : ToolbarPatternPilot::Map(native.pixels[index]);
            if ((themed.pixels[index] & 0xffffff) != (expected & 0xffffff)) return Fail("Pattern phase, clip or color mismatch");
            if ((native.pixels[index] & 0xffffff) == 0xcc4020 && (themed.pixels[index] & 0xffffff) != 0xcc4020) return Fail("Chromatic icon changed");
        }
        GetBrushOrgEx(themed.dc, &brushAfter); GetViewportOrgEx(themed.dc, &viewportAfter);
        if (memcmp(&brushBefore, &brushAfter, sizeof(POINT)) || memcmp(&viewportBefore, &viewportAfter, sizeof(POINT)) ||
            GetCurrentObject(themed.dc, OBJ_BRUSH) != brush) return Fail("Destination DC was modified");
    }
    {
        constexpr int width = 2200, height = 270;
        Surface wideNative(width, height), wideThemed(width, height);
        if (!wideNative.pixels || !wideThemed.pixels) return Fail("Wide surface allocation");
        GdiFlush();
        std::fill(wideNative.pixels, wideNative.pixels + width * height, sentinel);
        std::fill(wideThemed.pixels, wideThemed.pixels + width * height, sentinel);
        HGDIOBJ selections[2]{}; int index = 0;
        for (HDC target : {wideNative.dc, wideThemed.dc}) {
            selections[index++] = SelectObject(target, brush);
            SetViewportOrgEx(target, 5, 3, nullptr); SetWindowOrgEx(target, 2, 1, nullptr);
            SetBrushOrgEx(target, -3, 7, nullptr);
            HRGN clip = CreateRectRgn(7, 5, width - 8, height - 6);
            SelectClipRgn(target, clip); DeleteObject(clip);
        }
        if (!PatBlt(wideNative.dc, 1, 1, width - 5, height - 4, PATCOPY) ||
            !pilot.TryPaint(wideThemed.dc, 1, 1, width - 5, height - 4, PATCOPY)) return Fail("Tiled pattern rendering");
        GdiFlush();
        for (int pixel = 0; pixel < width * height; ++pixel) {
            DWORD expected = wideNative.pixels[pixel] == sentinel ? sentinel : ToolbarPatternPilot::Map(wideNative.pixels[pixel]);
            if ((wideThemed.pixels[pixel] & 0xffffff) != (expected & 0xffffff)) return Fail("Tiled pattern seam/clip/color mismatch");
        }
        SelectObject(wideNative.dc, selections[0]); SelectObject(wideThemed.dc, selections[1]);
    }
    std::vector<DWORD> snapshot(themed.pixels, themed.pixels + 64 * 32);
    fprintf(stderr, "GDI after patterns=%lu\n", GetGuiResources(GetCurrentProcess(), GR_GDIOBJECTS));
    if (pilot.TryPaint(themed.dc, 0, 0, 64, 32, PATINVERT) || pilot.TryPaint(themed.dc, 0, 0, 16385, 32, PATCOPY)) return Fail("Unsupported operation accepted");
    if (memcmp(snapshot.data(), themed.pixels, snapshot.size() * sizeof(DWORD))) return Fail("Unsupported operation changed pixels");
    for (HDC dc : {native.dc, themed.dc}) {
        SelectClipRgn(dc, nullptr); SetViewportOrgEx(dc, 0, 0, nullptr); SetWindowOrgEx(dc, 0, 0, nullptr);
    }
    for (WORD depth : {WORD(8), WORD(24), WORD(32)}) {
        struct DibInfo { BITMAPINFOHEADER header; RGBQUAD palette[256]; } info{};
        info.header.biSize = sizeof(BITMAPINFOHEADER); info.header.biWidth = 4; info.header.biHeight = -2;
        info.header.biPlanes = 1; info.header.biBitCount = depth; info.header.biCompression = BI_RGB;
        BYTE bits[32]{};
        DWORD colors[4]{0x00ffffff, 0x00000000, 0x00cc4020, 0x00808080};
        for (int index = 0; index < 4; ++index)
            info.palette[index] = {static_cast<BYTE>(colors[index]), static_cast<BYTE>(colors[index] >> 8), static_cast<BYTE>(colors[index] >> 16), 0};
        if (depth == 8) { info.header.biClrUsed = 4; for (int i = 0; i < 8; ++i) bits[i] = static_cast<BYTE>(i % 4); }
        else for (int i = 0; i < 8; ++i) {
            BYTE* p = bits + i * (depth / 8); DWORD color = colors[i % 4];
            p[0] = static_cast<BYTE>(color); p[1] = static_cast<BYTE>(color >> 8); p[2] = static_cast<BYTE>(color >> 16);
        }
        BYTE sourceCopy[32]; memcpy(sourceCopy, bits, sizeof(bits));
        GdiFlush();
        std::fill(native.pixels, native.pixels + 64 * 32, sentinel); std::fill(themed.pixels, themed.pixels + 64 * 32, sentinel);
        int expectedLines = SetDIBitsToDevice(native.dc, 0, 0, 4, 2, 0, 0, 0, 2, bits, reinterpret_cast<BITMAPINFO*>(&info), DIB_RGB_COLORS);
        int actualLines = 0;
        if (!pilot.TryDrawDib(themed.dc, 0, 0, 4, 2, 0, 0, 0, 2, bits, reinterpret_cast<BITMAPINFO*>(&info), DIB_RGB_COLORS, actualLines) || actualLines != expectedLines || expectedLines != 2)
            return Fail("DIB result differs");
        GdiFlush();
        for (int index = 0; index < 64 * 32; ++index) {
            DWORD expected = native.pixels[index] == sentinel ? sentinel : ToolbarPatternPilot::Map(native.pixels[index]);
            if ((themed.pixels[index] & 0xffffff) != (expected & 0xffffff)) return Fail("DIB palette/pixels mismatch");
        }
        if (memcmp(sourceCopy, bits, sizeof(bits))) return Fail("Source icon bitmap changed");
    }
    COLORREF originalFg = RGB(0, 0, 0), originalBg = RGB(240, 240, 240);
    SetTextColor(themed.dc, originalFg); SetBkColor(themed.dc, originalBg);
    { ToolbarTextPalette palette(themed.dc, true);
      if (!palette.Changed() || GetTextColor(themed.dc) == originalFg || GetBkColor(themed.dc) == originalBg) return Fail("Text palette not applied"); }
    if (GetTextColor(themed.dc) != originalFg || GetBkColor(themed.dc) != originalBg) return Fail("Text palette not restored");
    fprintf(stderr, "GDI with face=%lu\n", GetGuiResources(GetCurrentProcess(), GR_GDIOBJECTS));
    RECT field{2, 3, 20, 18}; int fillResult = 0;
    COLORREF brushColorBefore = GetDCBrushColor(themed.dc);
    for (int i = 0; i < 100; ++i)
        if (!pilot.TryFill(themed.dc, &field, face, fillResult) || !fillResult) return Fail("Toolbar field fill failed");
    GdiFlush();
    fprintf(stderr, "GDI after field=%lu\n", GetGuiResources(GetCurrentProcess(), GR_GDIOBJECTS));
    if ((themed.pixels[4 * 64 + 4] & 0xffffff) != ToolbarPatternPilot::Map(0xf0f0f0) || GetDCBrushColor(themed.dc) != brushColorBefore)
        return Fail("Toolbar field color/DC restoration failed");
    // Reinstate destination-owned clip copies for the resource comparison.
    initialClip = CreateRectRgn(8, 4, 38, 22);
    SelectClipRgn(native.dc, initialClip); SelectClipRgn(themed.dc, initialClip); DeleteObject(initialClip);
    pilot.Dispose();
    DWORD resourcesAfter = GetGuiResources(GetCurrentProcess(), GR_GDIOBJECTS);
    if (resourcesAfter != resourcesBefore) {
        fprintf(stderr, "GDI before=%lu after=%lu\n", resourcesBefore, resourcesAfter);
        return Fail("Pilot leaked GDI objects");
    }
    SelectObject(native.dc, oldNative); SelectObject(themed.dc, oldThemed); DeleteObject(brush);
    if (!DeleteObject(face)) return Fail("Face fixture deletion failed");
    printf("PASS: 100 pattern paints; phase/clip/DC preserved; 8/24/32-bit icon DIBs preserve source/colors; text colors restored; unsupported operations untouched; no GDI leak.\n");
    return 0;
}

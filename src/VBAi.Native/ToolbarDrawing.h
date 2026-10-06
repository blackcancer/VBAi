#pragma once
#include <windows.h>
#include <algorithm>
#include <array>
#include <cmath>
#include <climits>
#include <cstdint>
#include <string>

// Experimental PATCOPY backend. The caller must verify the exact live toolbar
// HWND and its owner thread before entry. No source window pixels are captured.
/// <summary>Owns a fixed top-down DIB for supported command-bar recoloring without reading screen pixels.</summary>
class ToolbarDrawing {
public:
    /// <summary>Reusable buffer dimensions in pixels, independent of destination toolbar size.</summary>
    static constexpr int MaxWidth = 1024, MaxHeight = 128;
    /// <summary>Counters for handled pattern tiles, rejected pattern operations and renderer failures.</summary>
    DWORD painted = 0, unsupported = 0, failures = 0;
    /// <summary>Counters for handled and rejected source DIB draws.</summary>
    DWORD dibPainted = 0, dibUnsupported = 0;
    /// <summary>Number of solid-brush fills dispatched through the mapped palette.</summary>
    DWORD fillsPainted = 0;
    /// <summary>Dimensions of the last handled pattern tile, used for optional bitmap samples.</summary>
    int lastWidth = 0, lastHeight = 0;
    /// <summary>Releases the owned DIB and memory DC after restoring the prior bitmap selection.</summary>
    ~ToolbarDrawing() { Dispose(); }
    /// <summary>Resets counters and creates the reusable DIB/DC; cleans partial GDI allocation on failure.</summary>
    /// <returns>True when the bitmap is selected into the DC, false after incomplete resources are released.</returns>
    bool Initialize() {
        Dispose(); painted = unsupported = failures = dibPainted = dibUnsupported = fillsPainted = 0; lastWidth = lastHeight = 0;
        dc = CreateCompatibleDC(nullptr);
        BITMAPINFO info{};
        info.bmiHeader.biSize = sizeof(BITMAPINFOHEADER);
        info.bmiHeader.biWidth = MaxWidth; info.bmiHeader.biHeight = -MaxHeight;
        info.bmiHeader.biPlanes = 1; info.bmiHeader.biBitCount = 32; info.bmiHeader.biCompression = BI_RGB;
        bitmap = CreateDIBSection(dc, &info, DIB_RGB_COLORS, reinterpret_cast<void**>(&pixels), nullptr, 0);
        if (dc && bitmap) oldBitmap = SelectObject(dc, bitmap);
        if (!dc || !bitmap || !oldBitmap || oldBitmap == HGDI_ERROR) { Dispose(); return false; }
        return true;
    }
    /// <summary>Restores the previous bitmap, releases owned GDI objects and clears handles; repeated calls are harmless.</summary>
    void Dispose() {
        if (dc && oldBitmap && oldBitmap != HGDI_ERROR) SelectObject(dc, oldBitmap);
        if (bitmap) DeleteObject(bitmap);
        if (dc) DeleteDC(dc);
        dc = nullptr; bitmap = nullptr; oldBitmap = nullptr; pixels = nullptr;
    }
    /// <summary>Maps neutral RGB pixels to the dark ramp while preserving chromatic/already mapped pixels and the high byte.</summary>
    /// <param name="pixel">Packed high-byte/R/G/B value; RGB channels occupy bits 16, 8 and 0.</param>
    /// <returns>The mapped pixel in the same representation.</returns>
    static DWORD Map(DWORD pixel) {
        // VBE can reuse colors supplied by the themed WM_CTLCOLOR path.
        // Preserve this ramp exactly, like the managed recovery renderer.
        static const std::array<DWORD, 256> output = [] {
            std::array<DWORD, 256> colors{};
            for (size_t i = 0; i < colors.size(); ++i) {
                double t = i / 255.0;
                DWORD r = static_cast<DWORD>(std::nearbyint(226 - 194 * t));
                DWORD g = static_cast<DWORD>(std::nearbyint(232 - 196 * t));
                DWORD b = static_cast<DWORD>(std::nearbyint(240 - 197 * t));
                colors[i] = (r << 16) | (g << 8) | b;
            }
            return colors;
        }();
        DWORD rgb = pixel & 0xffffff;
        if (std::binary_search(output.rbegin(), output.rend(), rgb)) return pixel;
        int r = (pixel >> 16) & 255, g = (pixel >> 8) & 255, b = pixel & 255;
        // Preserve chromatic icon pixels. Translate only the neutral ramp, using
        // the current product palette so a legacy recovery pass is idempotent.
        if (std::max(r, std::max(g, b)) - std::min(r, std::min(g, b)) > 12) return pixel;
        double t = ((r + g + b) / 3) / 255.0;
        DWORD rr = static_cast<DWORD>(std::nearbyint(226 - 194 * t));
        DWORD gg = static_cast<DWORD>(std::nearbyint(232 - 196 * t));
        DWORD bb = static_cast<DWORD>(std::nearbyint(240 - 197 * t));
        return (pixel & 0xff000000) | (rr << 16) | (gg << 8) | bb;
    }
    /// <summary>Tiles a bounded pattern operation through the fixed buffer, preserving brush phase and destination clipping.</summary>
    /// <returns>True if every tile was drawn; false on unsupported bounds or a failed tile. Earlier tiles can already be painted.</returns>
    bool TryPaint(HDC destination, int x, int y, int width, int height, DWORD operation) {
        // Reuse the fixed buffer for wide or vertically docked bars. Each tile
        // derives its phase from the destination DC, never from screen pixels.
        if (width <= 0 || height <= 0 || width > 16384 || height > 2048 ||
            static_cast<int64_t>(x) + width > INT_MAX || static_cast<int64_t>(y) + height > INT_MAX) {
            ++unsupported; return false;
        }
        for (int row = 0; row < height; row += MaxHeight)
            for (int col = 0; col < width; col += MaxWidth)
                if (!TryPaintTile(destination, x + col, y + row,
                    std::min(MaxWidth, width - col), std::min(MaxHeight, height - row), operation)) return false;
        return true;
    }
    /// <summary>Reconstructs one PATCOPY tile from the destination brush, recolors it and copies it through native clipping.</summary>
    /// <returns>True when copied; false for unsupported DC/brush or GDI failure. Original screen pixels are not captured.</returns>
    bool TryPaintTile(HDC destination, int x, int y, int width, int height, DWORD operation) {
        LOGBRUSH brush{};
        HGDIOBJ sourceBrush = GetCurrentObject(destination, OBJ_BRUSH);
        if (!dc || operation != PATCOPY || width <= 0 || height <= 0 || width > MaxWidth || height > MaxHeight ||
            GetMapMode(destination) != MM_TEXT || GetGraphicsMode(destination) != GM_COMPATIBLE || GetLayout(destination) != 0 ||
            GetObjectW(sourceBrush, sizeof(brush), &brush) != sizeof(brush) || brush.lbStyle != BS_PATTERN) {
            ++unsupported; return false;
        }
        POINT origin{}, device{x, y};
        if (!GetBrushOrgEx(destination, &origin) || !LPtoDP(destination, &device, 1)) { ++failures; return false; }
        // Keep the pattern phase when a draw starts inside the button or when
        // the destination has nonzero viewport/window origins. Its clip remains
        // untouched and applies to the final BitBlt in native coordinates.
        int saved = SaveDC(dc);
        if (!saved) { ++failures; return false; }
        bool ready = SetBrushOrgEx(dc, origin.x - device.x, origin.y - device.y, nullptr) != FALSE;
        HGDIOBJ prior = SelectObject(dc, sourceBrush);
        ready = ready && prior && prior != HGDI_ERROR;
        SetTextColor(dc, GetTextColor(destination)); SetBkColor(dc, GetBkColor(destination));
        bool drawn = ready && PatBlt(dc, 0, 0, width, height, PATCOPY) && GdiFlush();
        // Restore the shared brush selection before inspecting the CPU pixels.
        bool restored = RestoreDC(dc, saved) != FALSE;
        if (!drawn || !restored) { ++failures; return false; }
        for (int row = 0; row < height; ++row) for (int col = 0; col < width; ++col) {
            auto index = row * MaxWidth + col;
            before[index] = pixels[index];
            pixels[index] = Map(pixels[index]);
            after[index] = pixels[index];
        }
        if (!BitBlt(destination, x, y, width, height, dc, 0, 0, SRCCOPY)) { ++failures; return false; }
        // Complete access to the reusable source before the next native draw.
        GdiFlush();
        lastWidth = width; lastHeight = height; ++painted;
        return true;
    }
    /// <summary>Writes before/after samples of the last pattern tile to fresh bitmap paths without overwriting files.</summary>
    /// <param name="prefix">Prefix for .pattern-before.bmp and .pattern-after.bmp.</param>
    /// <returns>True after both writes, or when no tile was painted. A first successful sample remains after a second-write failure.</returns>
    bool WriteSamples(const std::wstring& prefix) const {
        if (!painted) return true;
        return WriteBitmap(prefix + L".pattern-before.bmp", before.data()) &&
            WriteBitmap(prefix + L".pattern-after.bmp", after.data());
    }
    /// <summary>Maps supported indexed or full-scan 24/32-bit BI_RGB source DIBs and dispatches the image draw.</summary>
    /// <param name="result">Receives the native SetDIBitsToDevice result only when handled.</param>
    /// <returns>True when handled, even if the native result is zero; false when the caller must use its original fallback.</returns>
    /// <remarks>Caller source pixels stay unchanged. The draw's last-error value survives GdiFlush.</remarks>
    bool TryDrawDib(HDC destination, int x, int y, DWORD width, DWORD height, int sourceX, int sourceY,
        UINT startScan, UINT lines, const void* bits, const BITMAPINFO* info, UINT colorUse, int& result) {
        DWORD incomingError = GetLastError();
        if (!dc || !bits || !info || info->bmiHeader.biSize != sizeof(BITMAPINFOHEADER) ||
            info->bmiHeader.biCompression != BI_RGB || info->bmiHeader.biPlanes != 1 || colorUse != DIB_RGB_COLORS ||
            info->bmiHeader.biWidth <= 0 || info->bmiHeader.biWidth > MaxWidth ||
            info->bmiHeader.biHeight == 0 || info->bmiHeader.biHeight < -MaxHeight || info->bmiHeader.biHeight > MaxHeight) {
            ++dibUnsupported; return false;
        }
        unsigned depth = info->bmiHeader.biBitCount;
        struct PaletteInfo { BITMAPINFOHEADER header; RGBQUAD colors[256]; } mapped{};
        mapped.header = info->bmiHeader;
        const void* mappedBits = bits;
        if (depth == 1 || depth == 4 || depth == 8) {
            unsigned colors = info->bmiHeader.biClrUsed ? info->bmiHeader.biClrUsed : 1u << depth;
            if (colors > (1u << depth)) { ++dibUnsupported; return false; }
            for (unsigned i = 0; i < colors; ++i) {
                RGBQUAD entry = info->bmiColors[i];
                DWORD rgb = Map((static_cast<DWORD>(entry.rgbRed) << 16) | (entry.rgbGreen << 8) | entry.rgbBlue);
                mapped.colors[i] = {static_cast<BYTE>(rgb), static_cast<BYTE>(rgb >> 8), static_cast<BYTE>(rgb >> 16), entry.rgbReserved};
            }
        } else if ((depth == 24 || depth == 32) && info->bmiHeader.biClrUsed == 0 && startScan == 0 && lines == static_cast<UINT>(std::abs(info->bmiHeader.biHeight))) {
            unsigned stride = ((info->bmiHeader.biWidth * depth + 31) / 32) * 4;
            size_t bytes = stride * lines;
            memcpy(dibPixels.data(), bits, bytes);
            unsigned pixelBytes = depth / 8;
            for (unsigned row = 0; row < lines; ++row) for (int col = 0; col < info->bmiHeader.biWidth; ++col) {
                BYTE* p = dibPixels.data() + row * stride + col * pixelBytes;
                DWORD rgb = Map((static_cast<DWORD>(p[2]) << 16) | (p[1] << 8) | p[0]);
                p[0] = static_cast<BYTE>(rgb); p[1] = static_cast<BYTE>(rgb >> 8); p[2] = static_cast<BYTE>(rgb >> 16);
            }
            mappedBits = dibPixels.data();
        } else { ++dibUnsupported; return false; }
        SetLastError(incomingError);
        result = SetDIBitsToDevice(destination, x, y, width, height, sourceX, sourceY, startScan, lines,
            mappedBits, reinterpret_cast<const BITMAPINFO*>(&mapped), colorUse);
        DWORD resultError = GetLastError();
        GdiFlush(); ++dibPainted; SetLastError(resultError);
        return true;
    }
    /// <summary>Maps a solid/system brush through DC_BRUSH and restores the previous destination DC-brush color.</summary>
    /// <param name="result">Receives the native FillRect result when handled.</param>
    /// <returns>True when dispatched, not a guarantee of drawing success; false for unsupported brush/color state.</returns>
    bool TryFill(HDC destination, const RECT* rect, HBRUSH sourceBrush, int& result) {
        DWORD incomingError = GetLastError();
        if (!dc || !rect) return false;
        COLORREF color;
        auto pseudo = reinterpret_cast<ULONG_PTR>(sourceBrush);
        LOGBRUSH brush{};
        if (pseudo >= 1 && pseudo <= COLOR_MENUBAR + 1) color = GetSysColor(static_cast<int>(pseudo - 1));
        else if (GetObjectW(sourceBrush, sizeof(brush), &brush) == sizeof(brush) && brush.lbStyle == BS_SOLID)
            color = sourceBrush == GetStockObject(DC_BRUSH) ? GetDCBrushColor(destination) : brush.lbColor;
        else return false;
        DWORD rgb = Map((GetRValue(color) << 16) | (GetGValue(color) << 8) | GetBValue(color));
        COLORREF previousColor = GetDCBrushColor(destination);
        if (previousColor == CLR_INVALID) return false;
        SetDCBrushColor(destination, RGB((rgb >> 16) & 255, (rgb >> 8) & 255, rgb & 255));
        SetLastError(incomingError);
        result = FillRect(destination, rect, static_cast<HBRUSH>(GetStockObject(DC_BRUSH)));
        DWORD resultError = GetLastError();
        SetDCBrushColor(destination, previousColor); ++fillsPainted; SetLastError(resultError);
        return true;
    }
private:
    /// <summary>Owned memory DC holding the reusable DIB.</summary>
    HDC dc = nullptr;
    /// <summary>Owned top-down 32-bit DIB selected into dc.</summary>
    HBITMAP bitmap = nullptr;
    /// <summary>Borrowed prior DC selection restored before deletion of the owned bitmap.</summary>
    HGDIOBJ oldBitmap = nullptr;
    /// <summary>Borrowed bitmap-storage pointer, valid only while the owned DIB exists.</summary>
    DWORD* pixels = nullptr;
    /// <summary>CPU copies of the last handled pattern tile before and after palette mapping.</summary>
    std::array<DWORD, MaxWidth * MaxHeight> before{}, after{};
    /// <summary>Bounded scratch copy for mapped 24/32-bit source DIB bytes.</summary>
    std::array<BYTE, MaxWidth * MaxHeight * 4> dibPixels{};
    /// <summary>Creates a top-down 32-bit sample using lastWidth/lastHeight and the fixed row stride.</summary>
    /// <returns>True after all headers/rows are written; partial files remain and are never overwritten.</returns>
    bool WriteBitmap(const std::wstring& path, const DWORD* data) const {
        HANDLE file = CreateFileW(path.c_str(), GENERIC_WRITE, FILE_SHARE_READ, nullptr, CREATE_NEW, FILE_ATTRIBUTE_NORMAL, nullptr);
        if (file == INVALID_HANDLE_VALUE) return false;
        BITMAPFILEHEADER header{};
        BITMAPINFOHEADER info{};
        info.biSize = sizeof(info); info.biWidth = lastWidth; info.biHeight = -lastHeight;
        info.biPlanes = 1; info.biBitCount = 32; info.biCompression = BI_RGB;
        info.biSizeImage = lastWidth * lastHeight * sizeof(DWORD);
        header.bfType = 0x4d42; header.bfOffBits = sizeof(header) + sizeof(info); header.bfSize = header.bfOffBits + info.biSizeImage;
        DWORD written;
        bool ok = WriteFile(file, &header, sizeof(header), &written, nullptr) && written == sizeof(header) &&
            WriteFile(file, &info, sizeof(info), &written, nullptr) && written == sizeof(info);
        for (int row = 0; ok && row < lastHeight; ++row)
            ok = WriteFile(file, data + row * MaxWidth, lastWidth * sizeof(DWORD), &written, nullptr) && written == lastWidth * sizeof(DWORD);
        CloseHandle(file); return ok;
    }
};

/// <summary>Temporarily maps a borrowed DC's text/background colors and restores their exact captured values.</summary>
class ToolbarTextPalette {
public:
    /// <summary>Captures current DC colors and maps them only when enabled and readable.</summary>
    ToolbarTextPalette(HDC target, bool enabled) : dc(target) {
        if (!enabled) return;
        foreground = GetTextColor(dc); background = GetBkColor(dc);
        if (foreground == CLR_INVALID || background == CLR_INVALID) return;
        COLORREF fg = MapColor(foreground), bg = MapColor(background);
        changed = fg != foreground || bg != background;
        if (changed) { SetTextColor(dc, fg); SetBkColor(dc, bg); }
    }
    /// <summary>Restores captured colors unless they were already restored explicitly.</summary>
    ~ToolbarTextPalette() { Restore(); }
    /// <summary>Reports whether mapped colors are still applied and awaiting restoration.</summary>
    bool Changed() const { return changed; }
    /// <summary>Restores captured foreground/background once and clears the pending-change flag.</summary>
    void Restore() { if (changed) { SetTextColor(dc, foreground); SetBkColor(dc, background); changed = false; } }
private:
    /// <summary>Borrowed destination DC, never released by this scope.</summary>
    HDC dc;
    /// <summary>Captured original text and background colors.</summary>
    COLORREF foreground = 0, background = 0;
    /// <summary>Whether this scope applied colors that still need restoration.</summary>
    bool changed = false;
    /// <summary>Converts COLORREF to renderer RGB and returns the mapped COLORREF.</summary>
    static COLORREF MapColor(COLORREF color) {
        DWORD rgb = (GetRValue(color) << 16) | (GetGValue(color) << 8) | GetBValue(color);
        DWORD mapped = ToolbarDrawing::Map(rgb);
        return RGB((mapped >> 16) & 255, (mapped >> 8) & 255, mapped & 255);
    }
};

# Chrome bitmap and mapping matrix

Prepared before execution. Raster readback uses an in-memory GDI DC and never a host window, cursor action, keyboard input or VBE setting.

| Area | Cases | Evidence |
| --- | --- | --- |
| Chrome pixel mapping | output identity, red, both greens, yellow, both blues, neutral, inactive caption, colored edge, alpha | exact expected RGB/alpha |
| Code mapping | editor face, border, text/comment/keyword ramps, independent ClearType channels, native markers, arbitrary bright color | exact constants, ramp endpoints, alpha |
| Background detection | all four dark, one light sample, each channel threshold, bright pixels outside samples | classification from synthetic bitmap |
| Code margin | empty scan, width/4 bound, 48 pixel bound, gaps/nonmatching RGB | last qualifying column within bounded scan |
| Property row | actual conversion and repeated conversion, region clipping, zero/negative/oversized bounds, invalid DC | bitmap readback and unchanged outside region |
| Invalid windows | Paint/PaintBorder/PaintComboButton missing HWND | harmless rejection; no log/error or bitmap mutation |

Full native chrome/window painting remains separate from these memory-raster contracts.

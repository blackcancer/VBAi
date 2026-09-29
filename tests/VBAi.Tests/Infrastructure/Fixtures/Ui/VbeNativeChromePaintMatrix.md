# Chrome window painting matrix

Prepared before execution. Window ownership and bitmap readback are real; API geometry/DC/cursor results can be substituted from deterministic memory fixtures. The cursor is never moved and no user-window action occurs.

| Method | Complete cases |
| --- | --- |
| Paint | reentrance, bounds/client/origin failure, client/nonclient, hosted child absent/invalid/valid, every dimension bound, supplied/acquired/missing DC, failed BitBlt, normal/code/light/dark/preserve paths, margin, unchanged, allocation exception, release ownership |
| Border | geometry failures, zero dimensions/no edges, independent left/right/bottom, clamps, missing DC, readback and release |
| Combo | info/bounds/client failures, collapsed button dimensions, missing DC, enabled/disabled, cursor absent/inside/outside and conversion failure, pressed state, exact arrow/background pixels and release |

Production defaults call the same native APIs. The scoped fixture restores all delegates and the thread's painting flag. Tests assert raster/state changes and ownership counts; they never click or reposition a host window.

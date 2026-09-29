# Native VBE rendering trace

Optional x64 diagnostic, isolated from the product build. It records actual GDI
calls made through the **import address tables of one selected loaded module**.
The default is VBE7.dll; the toolbar probe selects the module that registered
its MsoCommandBar class (VBEUI.dll on the measured Excel installation).
Default trace mode does not recolor text. An explicit toolbar rendering pilot is
also available below; neither mode modifies any Microsoft binary on disk.

## Build and synthetic verification

Requires the installed Visual Studio C++ x64 toolchain and Windows SDK. No hook
library or package download is used.

```powershell
powershell.exe -NoProfile -ExecutionPolicy Bypass -File .\tools\probes\native-render-trace\Build-NativeRenderTrace.ps1 -BuildSelfTest
$trace = Join-Path $PWD 'artifacts\native-renderer-pilot\trace'
& "$trace\selftest\TraceSelfTest.exe" "$trace\VBAiNativeTrace.dll" "$trace\selftest-result-fresh"
```

The last directory must not already exist. The self-test runs in its own EXE,
using a synthetic DLL named VBE7.dll next to that EXE. It does not open Excel or
SOLIDWORKS, and never replaces the installed VBA DLL. It checks identical pixels
before/during/after tracing, last-error preservation, actual GDI call records,
pass-through on another thread, a bounded buffer, and restoration of IAT slots,
page protection, and subclasses. It is not evidence of VBE compatibility.

## In-process ABI

```cpp
DWORD __cdecl VBAiTraceStart(HWND editorRoot, const wchar_t* outputJsonl, DWORD maxEvents);
DWORD __cdecl VBAiTraceStartForWindow(HWND editorRoot, HWND sourceWindow, const wchar_t* outputJsonl, DWORD maxEvents);
DWORD __cdecl VBAiTraceStop();
DWORD __cdecl VBAiTraceCount();
DWORD __cdecl VBAiTraceDropped();
```

The DLL must already be loaded inside the target host. Start and Stop must run
on the owning UI thread of `editorRoot`, in that window's process. They return a
Win32 error code; zero means success. The root should be the VBE main window.
This library provides no remote injection, host launcher, or product installer.
Load the absolute diagnostic path only for a coordinated disposable-host probe.

StartForWindow requires an existing descendant on the editor's thread. It reads
GCLP_HMODULE, retains that loaded module and observes its imports. This identifies
the class registration module, not necessarily every module involved in drawing.
No library is loaded by name to guess a toolbar implementation. Use the main
probe's `-TraceToolbarModule -NativeRenderTraceDll <path>` options for this mode.

Both normal imports and already-resolved delay imports are considered. Delay
thunks still pointing into the selected image are left untouched; their helper,
bound and unload tables are not changed. A slot pointing somewhere other than
the exact expected API target is rejected. Delay imports resolved after Start
are outside this session's coverage. No process-wide GDI detour is installed.

`outputJsonl` must be a new absolute path in an existing output directory.
`maxEvents` is 64 through 100000. Start reserves an empty report file. Stop disables
capture, restores instrumentation, verifies restoration, then writes UTF-8 JSONL.
No file I/O or heap allocation occurs inside the graphics hooks. A bounded
in-memory array drops records after its limit; the summary reports this count.
Always call Stop in probe cleanup, even when another probe action fails.

The diagnostic DLL pins itself until process termination: **never try to unload
it** after Stop. A thread that fetched an old IAT function pointer can still enter
its pass-through code safely. The actual IAT slots are restored with compare and
exchange; another component's unexpected replacement is not overwritten. The
initial page protection is retained and checked with VirtualQuery. Nonzero Stop
means restoration or report writing failed and must be investigated.

## Captured evidence

- API, timestamp, caller RVA inside VBE7, thread, message and window context.
- DC and memory-DC lineage where discoverable, source DC and selected bitmap for
  BitBlt, current selected font handle, foreground/background COLORREF values.
- Selected brush handle, known/style/color metadata. For BS_PATTERN, the color
  field is not a solid fill color and must not be interpreted as one.
- Text **length only**, positions, options, dimensions and native return value.
- Import inventory and verified restoration of pointers/protection/subclasses.
- FillRect, PatBlt, DrawTextA/W, DrawFrameControl and DrawEdge are also recorded.
  DrawText length -1 remains the unsigned sentinel 4294967295; text is not read.
  Rectangles are captured before APIs that may modify them (DT_CALCRECT/BF_ADJUST).

`discoveredImport` records identify normal/delayed imports, their observed target
and selection. A failed discovery also writes the module and `startError`, before
any instrumentation; an empty trace with nonzero startError is not a successful
measurement. The synthetic fixture includes delayed USER32 imports.

No text contents, source code, captions, or file document contents are recorded.
Window class names and the installed VBE module path appear in the report.
COLORREF is the native `0x00BBGGRR` encoding, not web RGB.

Normal imports inspected on this workstation: VBE7 7.01.1039 and Office VBE7
7.01.1158 import TextOutA and ExtTextOutA, SetTextColor/SetBkColor, BitBlt,
CreateCompatibleDC/Bitmap, SelectObject/DeleteDC, GetDC/GetDCEx/ReleaseDC and
BeginPaint/EndPaint. W variants are supported by the trace if actually imported.

## Exact limits

- It patches only the selected module's normal/resolved delay IAT entries, not
  all GDI calls in the process. Calls from other modules, dynamic GetProcAddress
  targets, unresolved delay imports, Uniscribe, DrawTextEx and other unlisted
  functions may be absent. An absent event
  does not prove that a rendering operation did not happen.
- Calls on other threads pass through without logging. Start subclasses the
  root and up to 511 existing descendants on its thread. Windows created later,
  owned popups outside that child tree and independent threads have incomplete
  message context. The GDI call can still be recorded with null context.
- Memory DCs created before capture have incomplete lineage. WindowFromDC alone
  cannot identify them. A 1024-entry lineage table has its own overflow counter.
- DC lineage is observational and can be incomplete; do not use it as the sole
  authorization to recolor an arbitrary HDC in production.
- It refuses pre-existing IAT redirection rather than guessing how to chain it.
- It does not prove that a native dark renderer, ClearType preservation, or full
  VBE/SOLIDWORKS coverage is solved. It supplies evidence for the next design.

## Microsoft references

- [PE imports and import-address tables](https://learn.microsoft.com/en-us/windows/win32/debug/pe-format)
- [GetClassLongPtrW and GCLP_HMODULE](https://learn.microsoft.com/en-us/windows/win32/api/winuser/nf-winuser-getclasslongptrw)
- [Delay-load helper](https://learn.microsoft.com/en-us/cpp/build/reference/understanding-the-helper-function?view=msvc-170)
- [GDI drawing functions](https://learn.microsoft.com/en-us/windows/win32/api/_gdi/)
- [VirtualProtect](https://learn.microsoft.com/en-us/windows/win32/api/memoryapi/nf-memoryapi-virtualprotect)
- [InterlockedCompareExchangePointer](https://learn.microsoft.com/en-us/windows/win32/api/winnt/nf-winnt-interlockedcompareexchangepointer)
- [SetWindowSubclass thread restriction](https://learn.microsoft.com/en-us/windows/win32/api/commctrl/nf-commctrl-setwindowsubclass)
- [WindowFromDC limitation](https://learn.microsoft.com/en-us/windows/win32/api/winuser/nf-winuser-windowfromdc)
- [GetModuleHandleEx module pinning](https://learn.microsoft.com/en-us/windows/win32/api/libloaderapi/nf-libloaderapi-getmodulehandleexw)

These document the underlying APIs and format. They do not establish a supported
VBE extension contract for changing an Office module's live import table.

## Experimental toolbar rendering pilot

`VBAiToolbarPatternStart(root, toolbar, path, limit)` uses the same lifecycle
but deliberately changes the drawing of exactly one live `MsoCommandBar` HWND.
It also observes VBE7 imports because that module paints the position field.
The report sets `colorsChanged: true` and identifies the module for each import
and call. It is not a product dependency or a default trace behavior.

`ToolbarPatternPilot.h` implements these paths:

- PATCOPY with a BS_PATTERN brush: render the requested tile to a preallocated
  DIB, preserving brush phase and destination clipping, map neutral pixels, then
  present the completed tile. Never capture or remap the previous screen image.
- SetDIBitsToDevice: copy neutral palette/pixel colors before drawing icon DIBs.
  Supports BI_RGB 1/4/8-bit indexed and complete 24/32-bit scan arrays, with a
  40-byte header and bounded dimensions. Source pixels and palettes are unchanged.
- FillRect: substitute the neutral color for a solid/system brush just for the
  draw, restoring the DC brush color afterward.
- TextOut/ExtTextOut/DrawText: apply the neutral foreground/background palette
  before rasterization and restore the DC colors after the call.

Color-bearing pixels are kept unchanged by this pilot. Unknown formats, memory
DCs without the exact toolbar identity, unrelated HWNDs/threads and unsupported
operations use the original API. Dimensions are bounded to 1024 × 128 pixels.
The target HWND is invalidated on destruction. The existing managed theme still
runs alongside this experiment; this is not complete replacement coverage.

Run the isolated pattern test after building with `-BuildSelfTest`:

```powershell
& "$trace\selftest\PatternPilotSelfTest.exe"
```

For a disposable Excel observation, use the main probe with
`-InProcessAddInExperiment -OpenCodeWindow -CheckPropertyRows -CheckToolbarBlink
-TraceToolbarModule -ToolbarPatternPilot -NativeRenderTraceDll <dll>`.
`-ToolbarOnly` ends the campaign after the toolbar observation, with normal
trace/Excel cleanup. `-KeepToolbarProbeVisible` temporarily makes only the test
VBE topmost without activation during sampling, then restores its topmost flag.
This test-only option prevents another app from masking the sampled pixels; it
does not change how the product runs in the background.

The last pattern tile before/after conversion is saved as BMP at Stop. These
samples are toolbar graphics, not document captures. Resource/call counters are
included in the completion record. The pilot is not yet installed or qualified
for all toolbars, popups, SOLIDWORKS, DPI, and theme/lifecycle transitions.

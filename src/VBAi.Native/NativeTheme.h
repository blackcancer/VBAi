#pragma once
#include <windows.h>
/// <summary>Versioned native-renderer counters; the caller supplies the structure size before querying status.</summary>
struct NativeThemeStatus {
    /// <summary>Structure size in bytes, ABI version, enabled flag, tracked toolbar count, discovered imports and last Stop restoration count.</summary>
    DWORD size, abi, active, windows, imports, restoredImports;
    /// <summary>Handled pattern/image/text/fill draws, unsupported pattern/image draws and renderer failures.</summary>
    DWORD patterns, images, text, fills, unsupported, failures;
};
/// <summary>Initializes the experimental renderer for the current-process VBE root on its owning UI thread.</summary>
/// <param name="editor">Live wndclass_desked_gsk HWND owned by the calling process and thread.</param>
/// <returns>ERROR_SUCCESS on activation or an already enabled identical root; a Win32 error on refusal/setup failure.</returns>
/// <remarks>Pins this module, subclasses the root and eligible child bars, and patches retained imports. Failed setup attempts Stop; cleanup can remain incomplete.</remarks>
extern "C" __declspec(dllexport) DWORD __cdecl VBAiThemeStart(HWND editor);
/// <summary>Registers one eligible child MsoCommandBar and refreshes drawing imports on the original UI thread.</summary>
/// <param name="window">Live child command-bar HWND under the original VBE root.</param>
/// <returns>ERROR_SUCCESS or a Win32 error; a registered HWND is not added twice.</returns>
/// <remarks>Requires an active Start. Floating popups are outside scope.</remarks>
extern "C" __declspec(dllexport) DWORD __cdecl VBAiThemeRegister(HWND window);
/// <summary>Rediscovers drawing imports in retained modules and installs replacement pointers.</summary>
/// <returns>ERROR_SUCCESS or a Win32 error; refuses entry before Start or on another thread.</returns>
extern "C" __declspec(dllexport) DWORD __cdecl VBAiThemeRefresh();
/// <summary>Disables drawing, attempts import restoration, removes subclasses and releases buffers on the owning thread.</summary>
/// <returns>ERROR_SUCCESS if stopped/already stopped; the first import-restoration error or a thread error otherwise.</returns>
/// <remarks>A restoration error retains module/import tracking and started state for diagnosis. The pinned native module is not unloaded.</remarks>
extern "C" __declspec(dllexport) DWORD __cdecl VBAiThemeStop();
/// <summary>Copies the current ABI and counters into a caller-owned status structure.</summary>
/// <param name="status">Non-null writable buffer whose size equals sizeof(NativeThemeStatus).</param>
/// <returns>ERROR_SUCCESS, ERROR_INVALID_PARAMETER for a missing/wrong-size buffer, or ERROR_INVALID_THREAD_ID while started on another thread.</returns>
extern "C" __declspec(dllexport) DWORD __cdecl VBAiThemeStatus(NativeThemeStatus* status);

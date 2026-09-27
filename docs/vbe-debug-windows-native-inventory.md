# VbeDebugWindows native coverage inventory

Status: **219 passed, 2 skipped in the global VSTest run on 27 September 2026**. The tests in `VbeDebugWindowsNativeTests` use an internal
`INativeProbe` fake. Public entry points still use the same Win32, UIA and MSAA
operations through `NativeProbe`.

## Deterministic cases added

- Snapshot: absent VBE root, resolved or missing pane handles, optional call
  stack, and the explicit missing-pane limit.
- Dialog preflight: open compile and Options dialogs block another action.
- Diagnostic read: absent dialog, invisible controls, one or multiple visible
  messages, and visible buttons.
- Diagnostic response: absent dialog, changed or unsupported message, missing
  or ambiguous button, failed click, closed dialog, and pending verification.
- Compile wait: diagnostic and own OK click, absent or failed OK, completed
  compilation without diagnostic, and timeout without completion.

## Native branches still needing a real VBE

- `ReadList`, `ReadImmediate`, `ImmediateDocument`, and `ImmediateText` depend on
  a visible native debugger pane with UI Automation patterns. The fake proves
  routing and handling of missing handles, not the actual text or row parser.
- `ExecuteImmediate` posts keystrokes into the native pane and verifies output
  from UIA. Cursor focus, timing, and actual VBA evaluation need live VBE.
- `ReadCallStack` opens a VBE dialog and reads its MSAA accessible tree. The
  live call stack and dialog ownership need host integration coverage.
- `ChangeDebugItem`, `CompleteAddWatch`, `CompleteEditWatch`,
  `CompleteQuickWatch`, `SelectWatch`, and `VerifyWatchRemoved` use native watch
  dialogs or UIA rows. The watch mutation, selection, and removal results need
  a paused VBA project in an actual host.
- `ReadDebugOptions` and `ReadVbeOptions` enumerate native UIA tabs and
  controls. The host-specific tree and localized labels need live inspection.
- `ReadSignatureDialog` and `CompleteProjectSignature` depend on MSAA and the
  native certificate picker. Certificate selection and trust are not proved by
  fake controls.
- Win32 discovery (`FindVbeRoot`, `FindPane`, `FindDialog`) and timing of a
  dialog after a posted command require a host window and message loop. The
  fake tests only verify the decisions made from returned handles and states.

These limits remain after the fake suite passes. No coverage exclusion was
added. The grouped VSTest/Cobertura report is at
`artifacts/coverage/batch-comprehensive-final/9bf3adb1-a9d7-4eac-bc7c-3913122c6063/coverage.cobertura.xml`.

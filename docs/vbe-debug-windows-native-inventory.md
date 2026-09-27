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

## Follow-up deterministic batch (pending grouped VSTest)

The new `IWatchProbe` cases cover Add/Edit Watch dialog absence, missing
controls, changed selection/context, type selection failure, edit echo failure,
refused OK, native validation error, close timeout, hidden Watches pane,
pending readback, and verified edit readback. `ISignatureProbe` cases cover
dialog timeout, inaccessible or missing Cancel control, exact labels and
Cancel action, and failure to close. `IOptionsProbe` cases cover dialog
timeout, tab bounds and names, control visibility, disabled and blank text
filtering, control bound, and close timeout. These cases were compiled but
were **not run** in this batch; the earlier 219/2 result above does not
include them.

The Win32 message delivery, VBE watch mutation, live UIA values, MSAA
certificate tree, native dialog ownership, and actual close timing still
require a VBE host. The probe tests verify decisions from simulated native
observations only. `ReadDebugOptions`, `CompleteProjectSignature`,
`ExecuteImmediate`, `ChangeDebugItem`, `CompleteQuickWatch`, and the
low-level UIA/MSAA extraction branches are not fully reachable by this
deterministic batch. They remain explicit obstacles to a claim of 100% line
and branch coverage.

## Additional native decision scenarios (pending grouped VSTest)

The watch probe now exercises Quick Watch dialog controls, expression and
context readback, selection of a unique row, and watch removal readback,
including hidden, ambiguous, unsupported, pending, and observed states.
The Options probe also exercises the three error trapping radio choices:
invalid count/name/pattern, zero or multiple selections, successful readback,
and a dialog that remains open. `VbeDebugTests` adds host-independent
branches for absent editor selection, Watches/Immediate command routing,
pending break/reset transitions, and watch-dialog preflight.

The probe does not simulate UIA provider behavior itself. `ReadList`,
`ImmediateDocument`, `ReadCallStack`, `MatchingWatchRows`, low-level
`NativeOptionsProbe` extraction, and actual command effects still need a
visible native VBE in an appropriate host. `CompleteProjectSignature` also
needs the protected Windows certificate picker and live certificate readback.

`IImmediateProbe` also covers command validation, missing host/pane, rejected
character messages, exact echo timeout, rejected Enter, pending output, and
changed text. The native `TextPattern` document, focus, and posted-message
effects remain a live-host acceptance boundary.

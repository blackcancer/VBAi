# Native-host qualification

[Testing overview](README.md) · [Compatibility](../docs/compatibility.md)

These checks exercise Windows COM and native application state. They require a
prepared disposable fixture, the selected host and the exact registered build.
Host-free or protocol-only tests cannot substitute for their acceptance oracles.

## Prepare before dispatch

1. Freeze the complete ordered scenarios, acceptance oracles and no-retry policy.
2. Identify product/test/helper bytes by source manifest, MVID and SHA-256.
3. Record executable version/hash, process/COM/window identity and desktop.
4. Declare owned documents, mutations, persistence/reopen and cleanup obligations.
5. Check pre-existing hosts, trust prerequisites, provider access and registration.
6. Dispatch the reviewed complete bank once; retain original handles and receipts.

Normal exit must be observed through the original owned process handle. PID
disappearance, a successful Quit call or a released desktop alone is insufficient.
Blocked, failed, retained and not-run scenarios are reported separately.

## Host opt-ins

| Family | Basic opt-in | Additional requirements |
| --- | --- | --- |
| Excel | `VBAi_RUN_EXCEL_TESTS=1` | Explicit disposable workbook, installed candidate and fixture-specific flags |
| Word/PowerPoint/Access/Publisher | `VBAi_RUN_OFFICE_TESTS=1` | Selected host executable, document adapter prerequisites and case filter |
| Classic Outlook | `VBAi_RUN_OUTLOOK_TESTS=1` | Configured profile; preserve mail and personal OTM; use the case's synthetic fixture |
| SOLIDWORKS | Selected preloaded host under the fixture's opt-in | Explicit version/PID/ROT; no autonomous launch/termination under normal repository rules |

Basic flags do not authorize every scenario. Read the integration fixture's
environment gate for the selected test, then remove its opt-ins after completion.
Avoid running native UI banks concurrently against a shared host or desktop.

## Desktop and coordinator contracts

`tools/tests/Invoke-IsolatedTests.ps1` defaults to the original-handle PowerShell
launcher. `-DirectGuiLauncher` uses the Q026 GUI helper `--run-plan` entry point.
Both retain the guarded worker, sentinel, desktop lease and failed-worker
observation. Their helper/task receipt contracts differ; preserve the actual one.
The helper is a WinExe, so launchers must retain and explicitly wait its original
process handle instead of treating early script return as completion.

An inactive desktop can isolate qualification from a working desktop, but it
does not establish real-desktop rendering. When a campaign specifically requires
the real desktop, use its reviewed same-user Limited x64 STA coordinator. Never
switch the user's desktop or select windows with keyboard/mouse coordinates.
Private-desktop descriptors cannot be silently adopted by main-desktop workers.

## Campaign entry points

| Scope | Reviewed runner |
| --- | --- |
| General isolated managed/native plan | `tools/tests/Invoke-IsolatedTests.ps1` |
| Q006 Excel declared values and persistence | `tools/tests/Invoke-Q006Qualification.ps1` |
| Q012 existing Office adapter contract | `tools/tests/Invoke-Q012Qualification.ps1` |
| Q014 selected SOLIDWORKS workflows | `tools/tests/Invoke-Q014Qualification.ps1` |
| Q020/Q030 native macro creation/publication/reopen | `tools/tests/q020-main/Invoke-Q020MainNativeMacroQualification.ps1` |
| Q028 provider and embedded assistant | [Provider campaign](providers.md#q028-complete-office-assistant-campaign) |

Runner parameter declarations and source are authoritative. Preparation freezes
inputs; it does not qualify them. A claimed evidence root is not reusable for
another execution. Do not bypass a campaign's authorization gate to start a host.

## Registration, state and shutdown

Temporary candidate registration uses
`tools/testing-explorer/Set-TestExplorerCandidate.ps1` with the reviewed DLL/MVID.
Back up the exact scoped registry state, verify loaded identity after connection,
then restore and verify the same baseline. A shell's virtualized registry view
alone is not real-host loading proof. Preserve concurrent settings changes.

COM/UI work stays on its owning STA and the observer on its original thread.
Freeze exact process, native thread, HWND/parent/owner and control identities
before each action; a replaced or foreign identity refuses the action.
The settled Word collection diagnostic is an explicit fixture option, not native
RCW-release proof. Original unchanged exit deadlines still apply.

After uncertain native dispatch, do not repeat Save, import, Reset, Send, Stop or
Quit to obtain a passing result. Retain owned processes/references, pending
requests, recovery files and failure receipts for a separate diagnosis. Forced
termination is cleanup evidence only and is never normal-exit acceptance.

## Special diagnostic scopes

Path visibility/effective-token inspection, paired export tracing, native scalar
setters, owner-import fonts/resources, signatures and authenticated Git each have
additional explicit gates in their integration fixture. They are diagnostic
operations, not default test behavior. Record authorization and exact synthetic
inputs; do not change permissions, Office security or an unrelated account.

UserForm persistence compares raw FRM/FRX, typed native properties/fonts,
references and fresh-process saved-file readback. Initialize the actual designer
where the host requires it. Pixel inspection, raw resources and runtime form
events are different evidence scopes.

Signature trials use only the exact disposable certificate/key and workbook,
record selected certificate/dialog ownership and signed-file readback, then
remove only owned test material after terminal cleanup. Certificate dialogs and
trust prerequisites can block acceptance; preparation is not a signing pass.

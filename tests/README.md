# Testing VBAi

## Isolated launcher choices

`tools/tests/Invoke-IsolatedTests.ps1` keeps the original-handle PowerShell
launcher by default. `-DirectGuiLauncher` selects the Q026 GUI entry point
(`--run-plan`) instead, avoiding an intermediate console process. Both routes
use the same guarded worker, hidden inventory sentinel, checked desktop lease
and retained failed-worker observation. The helper is a WinExe in either route.
The default route records the original outer helper handle and observed exit;
the direct route records its own helper identity and terminal result, with the
scheduled task's actual exit checked separately. Do not substitute those
different receipt contracts or turn resource release into a passing campaign.
Scripts that launch a WinExe subprocess must explicitly retain and wait for its
original process handle. An early script return and an empty window inventory
do not prove that a GUI child has finished. The script must observe that exit
before allowing its parent worker and desktop to close.

The test suite covers the shared VBE core, provider protocols, storage, editor
services and selected host integrations. A passing local suite is not a claim
that every feature works in every application that embeds the VBE.

## Q-014 inactive-desktop SOLIDWORKS qualification

`tests/VBAi.Q014.Tests` is a scoped frozen-product harness for SOLIDWORKS
persistence guards and private-desktop guards. It references an explicit
`FrozenProductDirectory`, never rebuilds the installed product and does not
replace the full managed suite. Build it and `VBAi.Desktop.Helper` using the
same isolated output with `BuildProjectReferences=false` for the helper and an
exact candidate copy at `build/VBAi/Debug/net48`. Use the preparation script
only after those outputs and the installed candidate hashes agree.

```powershell
& "$env:WINDIR/System32/WindowsPowerShell/v1.0/powershell.exe" -NoProfile -File tools/tests/Invoke-Q014Qualification.ps1 -Prepare -EvidenceRoot "$PWD/artifacts/q014-owned-campaign" -InstalledDirectory '<absolute installed candidate directory>'
```

Preparation writes a frozen plan before execution. Run its saved
`Invoke-FrozenQ014.ps1` only through `Invoke-IsolatedTests.ps1` and the prepared
helper. The private-desktop canary and actual testhost desktop must pass first.
SOLIDWORKS launch additionally requires explicit maintainer authorization: the
default repository rule remains preloaded, selected hosts only. The current
profile defaults to 2019 SP5, revision `27.5.0`. For explicitly authorized 2025
qualification, pass `-SolidWorksYear 2025 -ExpectedNativeRevision 33.1.1` and
the exact 2025 `-SolidWorksExecutable`; preparation validates the major version.

The generated Visual Studio utility solution launches through the exact owned
DTE's `Debug.Start`. Do not pass quoted `/Command` arguments to the desktop
launcher's generic argument formatter. Capture internal `NativeChild` receipts
through reflection with instance/nonpublic flags. SOLIDWORKS identity is read
through its installed typed interop inside C#, since the observed 2019 ROT
object refuses PowerShell's dynamic type-info scan. Validate the actual native
Afx frame's PID/UI thread/desktop, rather than assuming a `SWFrame` class.

The plan orders managed guards, launch, fresh native macro creation, add-in
loading, module/class/form operations, stale-state refusals, export/import,
compile, debugger, product Save, and independent saved-copy Edit Macro. Native
macro file commands use typed `ISldWorks.RunCommand` with a retained STA worker;
SOLIDWORKS API command IDs are not a qualified `WM_COMMAND` transport. Dialog
input uses only addressed filename readback and the exact owned button. Each
request has a durable intent; a timeout stops dependencies without replay or
cleanup mutation. Original exports and partial files remain available.

Embedded Monaco/designer rendering, resize/restore, stale-project UI, assistant
dispatch/cancellation/recovery, and normal cleanup remain explicit acceptance
requirements. The runner does **not** implement the embedded UI/assistant bank
and must retain them as `NOT_RUN`, even if its bridge scenarios pass. Normal
SOLIDWORKS/IDE closure requires maintainer authorization. Automatic normal
closure has been explicitly authorized for the current Q-014 campaign and does
not require a repeated approval for its exact owned hosts. This permission does
not authorize force termination or uncertain mutation replay. The initial
campaign and any narrowly guarded continuation retain separate ledgers; do not
erase harness failures or rerun completed tests to manufacture a clean result.
The typed file-command path received independent 2019 and 2025 acceptance in the
campaign. Its external observers also record actual embedded rendering and
normal owned-process closure; they do not turn the runner's UI/assistant
placeholders into executed results. Read host-held SWP files with explicit
ReadWrite/Delete sharing, retain a verified fresh copy, and never repeat Save
merely because an ordinary file-hash reader conflicts with a host writer.

`-RetainedManagedRoot` can preserve an earlier scoped TRX only when the installed
candidate, scoped assembly and all frozen runtime dependencies match. Observe
the exact loaded bridge MVID/path before assigning native results to that
candidate. A different loaded candidate needs its own frozen harness/results;
do not waive the identity oracle or rewrite registration to conceal a mismatch.
Wait for exact DTE project/document readiness before the first Debug.Start and
for a newly started process path before rejecting its identity. Duplicate menu
and toolbar presentations with the same native ID/caption share a command
identity. Read-only observations may settle; uncertain mutations may not replay.

The desktop helper retains the original worker handle after campaign failure.
It continues read-only exit and desktop observations; after that exact worker
exits and the inactive desktop is empty, it releases its own handles and writes
`retained-terminal.json`. It still returns failure and preserves `failure.json`.
Observation errors never prove exit, an empty desktop or successful release.
The timer-driven owned-form regression and focused guard results are recorded
in [retention recovery evidence](../docs/test-coverage.md#q-014-private-desktop-retention-recovery-2026-10-04).
This source correction does not change earlier already-loaded helper copies.

Before normally closing a host from an additional assistant bank, call
`tools/tests/Assert-Q014AssistantTerminal.ps1 -EvidenceRoot <absolute-bank-root>`.
Every sent scenario must have its independent native success or failure receipt;
a local HTTP completion alone is insufficient. This read-only barrier throws
before teardown when a native result readback is still pending.

See [the current Q-014 gate](../docs/release-qualification.md#q-014-private-desktop-campaign-2026-10-04)
and [its final recorded results](../docs/test-coverage.md#q-014-final-decision-and-cleanup-2026-10-04).

## Q012 campaign on an inactive Windows desktop

`tools/tests/Invoke-Q012Qualification.ps1` prepares the complete fixed Access and
Publisher adapter matrix before execution. Preparation is read-only and freezes
source, binaries, installed executables and exact test selectors. `-Execute`
claims that plan once; it runs the focused managed gate, then each native case
at most once, retaining its original TRX and all blocked/not-run rows. Existing
or retained same-host processes block dependent cases. A reviewed `-HostScope`
and frozen `-BlockedHostReason` can preserve a proven prerequisite failure without
repeating that unchanged host bootstrap; every excluded row stays BLOCKED. No native mutation or
cleanup is replayed to obtain a pass.

`-ScenarioNumbers` can freeze a smaller diagnostic selection from the numbered
matrix. Each number must be distinct; unselected rows remain NOT_RUN with zero
invocations. A selected-case pass does not qualify the full matrix. Use a fresh
plan/evidence directory for a changed hypothesis or instrumentation, retaining
the original failures and avoiding a replay of already accepted cases.
Its terminal state is `DIAGNOSTIC_PASS` only when all selected cases pass;
`FullMatrixQualified` remains false unless the complete matrix passes and the
exact temporary COM registration is restored and verified at the terminal gate.

`-StopOnNativeFailure` blocks later selected cases after the first native failure.
`-NativeScenarioTimeoutSeconds` applies a VSTest hang deadline to each selected
native case, without creating dumps. `-ManagedGateTimeoutSeconds` independently
bounds the focused managed gate's hang wait. A deadline termination does not prove normal
Office shutdown; inspect the retained ownership and cleanup receipts. Defaults
preserve the original behavior. Rows 16 and 17 exercise the explicit native
General HelpFile and HelpContextID paths; rows 13 and 15 retain their distinct
legacy COM results. Row 18 separately exercises accented HelpFile text
representable by the observed ANSI control. Row 19 verifies refusal before
writing unsupported Unicode. A known encoding refusal cancels the unchanged
original dialog once, with the same ownership, authorization and revision checks
as a read. Acceptance requires observed closure and original command return,
unchanged source/references/file, and normal host shutdown; no Save or fresh
reopen follows. Unknown failures and unsettled cancellation retain uncertainty
and require separately guarded cleanup. Row 16 retains the original Unicode persistence
test; it is not replaced by the narrower accented-path scenario.
Rows 20 and 21 apply the shared native General persistence scenario to an owned
Access database, for compatible HelpFile text and HelpContextID respectively.
They observe host trust without enabling content or changing trust policy.
Rows 22–26 separately use the audited serialized Publisher seed for source/class/
UserForm-state Save, reference addition by GUID/file, reference removal and
Description. Rows 27–30 verify Access/Publisher legacy HelpFile/HelpContextID
refusal before any setter, with unchanged project/source/reference/Saved state
and normal close. They perform no Save or fresh reopen after the refusal.
The catalogue's `HostLegacyWriteUnsupported` status requires an explicit,
approved native General call; it is not a fallback or a ReadOnly reclassification.

Row 19 keeps its audited saved seed unchanged: it does not replace source markers,
perform a preparation Save or read live publication bytes. The original seed's
closed-file hash is bound to the verified pre-open copy; byte comparison of the
owned publication occurs only after successful normal fixture disposal. A
known refusal requires Field/OK attempts absent, one Cancel, closure and original
Execute return, unchanged revision/source/references/Saved state, and exact closed
bytes. Unknown outcomes remain quarantined. Successful compatible-text or refusal
cases do not replace the original Unicode persistence failure.

Current Q-012 acceptance is the required existing-document contract recorded in
[recorded validation](../docs/test-coverage.md#q-012-completed-adapter-contract-2026-10-04),
not every diagnostic row or a complete managed suite. Original NewDocument and
legacy scalar failures remain preserved; form-state persistence is not rendering
or event execution. Keep focused/native plans sequential so an unrelated smoke
test cannot add a screenshot to a frozen binary inventory during qualification.
The runner hashes the full source/binary inventory before
each actual invocation and at the terminal gate, avoiding repeated hashing for
unselected rows that cannot dispatch.

The desktop helper checks the creator lease's native close result and records
the attempt before entering `CloseDesktop`. Failed closes retain the owned
handle and their original error; disposal never repeats that native call.
Successful terminal publication follows sentinel shutdown and verified creator
lease release. Remaining host windows or an uncertain close retain ownership
and require separately guarded cleanup. This does not make forced host exits
normal shutdown acceptance.

Publisher fixture Open requires verified process-local
`AutomationSecurity=ForceDisable`. It never enables macro content, changes a
global trust policy or retries an uncertain Open. An unfamiliar modal is refused.

`-MetadataGetterProbe` explicitly freezes the existing read-only Access metadata
diagnostic into a selected campaign. It compares descriptor, CLR and raw getters
at the planned mutation/save/reopen phases, including observation after a failed
setter. That observation cannot promote a failed or uncertain mutation to success.
The getter opt-in is enabled only after the focused managed gate; setter probes
remain disabled and no additional metadata write or Save is introduced.

Build the solution with an isolated `BuildOutputRoot`. Prepare with absolute
`-CandidateAssembly`, `-BuildOutputRoot`, `-EvidenceDirectory` and
`-DesktopHelperAssembly` paths using Windows PowerShell 5.1 x64. The helper is
`VBAi.Desktop.Helper.exe` in its project output. Preparation does not change COM
registration or start Office. Review `plan.json` and the registration preview.

Create a reviewed script containing the same arguments plus `-Execute`, then
pass its absolute path to `tools/tests/Invoke-IsolatedTests.ps1` with a fresh
absolute evidence directory and that helper. The one-shot limited-user task
creates a generated inactive Windows desktop; it never switches the input
desktop. Its canary and worker receipts are environment proofs, not Office
acceptance. Missing desktop evidence refuses execution without fallback.
The outer limited launcher uses `UseShellExecute=false` and `CreateNoWindow=true`
for the unchanged console helper, drains both redirected streams asynchronously,
and retains its original process handle/identity until observed exit. No WinExe
conversion, control-handler ignore, console attachment or keyboard input is used.
Parent and worker console observations do not establish a helper-local console
inventory or explain older `0xC000013A` interruptions. Exceptional launcher faults
need their separate checked evidence; normal native acceptance proves the settled
path only. Run `tools/tests/Test-IsolatedHelperRetention.ps1` under PowerShell 5.1
for pure receipt-failure and uncertain-Start regressions; it starts no helper or
Office application. A helper/worker that disappears without terminal evidence is not
treated as normal cleanup.

Only Access and Publisher have the private host bootstrap. The campaign passes
`VBAi_TEST_ACCESS_EXE` and `VBAi_TEST_PUBLISHER_EXE` from the frozen installed
paths. Original and reopened processes are explicitly started on the private
desktop, attached without COM activation, and verified by original handle, image,
PID and exact HWND membership in successful private/input desktop inventories
before mutation. The live hidden sentinel must be observed on the private desktop;
no owned host window may appear on the input desktop. Foreign thread-desktop
queries are not used as proof. Ordinary Office tests without
`VBAi_TEST_DESKTOP_NAME` retain their existing behavior.

Private Access bootstrap explicitly observes the registered `VBAi.AddIn` VBE
entry and, when disconnected, requests its host-managed connection once.
Exact ProgID/GUID, original COM/process/window identities, fixture STA and
private/input inventories must remain valid. Read-only LoadBehavior observations
must remain unchanged, and unknown visible modals or a disabled VBE refuse the
subsequent bridge/scenario work. A true Connect readback is not candidate
acceptance: the existing bridge PID and MVID checks remain mandatory. Failed or
uncertain connection is retained without reconnecting or changing trust policy.

Access starts as an explicitly owned normal application and attaches through ROT;
Publisher also uses an explicitly owned normal GUI launch with no arguments.
Its registered server command/arguments are recorded only and never executed;
manual server arguments do not perform COM class activation. Before the first
disposable NewDocument/Open, the application association is explicitly provisional:
empty prelaunch inventory, sole retained original Publisher process, exact native
handle/image/session, application interface/canonical identity, fixture STA and
private/input/sentinel proofs must agree. A failed or uncertain first creation/open
has no retry, fallback or adoption. The returned publication must then be the exact
retained, active and sole canonical document with an owned native window before
any baseline SaveAs, bridge or scenario work. Final ownership is not inferred
from ROT attachment alone.
Access records Boolean `UserControl` and `Visible` once outside ROT discovery.
An already visible instance needs no visibility setter; hidden automation permits
one setter, while a hidden user-controlled instance refuses. The decision is bound
to the original native handle/application/PID and cannot be reused. Ownership and
private desktop proofs precede database creation or open.
Publisher's observed French recovery prompt is handled before ROT attachment:
only its exact owned dialog and native No button can receive one bounded click
requesting normal startup. Unknown prompts/locales, changed identities or uncertain
message delivery retain the host and refuse further action. No safe-mode,
macro-trust, licensing or global input setting is changed.

Every positive case requires adapter-only Save, exact source/project/reference
readback, normal original exit, a fresh process reopening the saved file and
normal final exit. Existing ACCDB/PUB is the accepted scope; first SaveAs remains
refused. The campaign preserves the uncertainty and single-save guards. It
restores the exact temporary registration without replacing installed binaries
or changing trust policies. A retained host or helper stays separately recorded;
forced cleanup, if separately performed, cannot qualify normal exit or persistence.

## Q-006 inactive-desktop Excel qualification

`tools/tests/Invoke-Q006Qualification.ps1 -Prepare` builds a deliberately scoped
`tests/VBAi.Q006.Tests` harness against an explicit installed candidate directory.
It does not rebuild/register the product or overwrite the installed DLL. The
plan retains the source revision/status, candidate and harness hashes, exact
scenario filters, acceptance oracles and fresh evidence paths before execution.
This separate friend assembly is not a replacement for the full solution suite.

```powershell
$q006Root = Join-Path $PWD 'artifacts/q006-new-campaign'
& tools/tests/Invoke-Q006Qualification.ps1 -Prepare `
  -EvidenceRoot $q006Root -InstalledDirectory '<absolute installed net48 directory>'
& tools/tests/Invoke-IsolatedTests.ps1 `
  -ScriptPath (Join-Path $q006Root 'Invoke-FrozenQ006.ps1') `
  -HelperAssembly (Join-Path $q006Root 'build/VBAi.Desktop.Helper/Debug/net48/VBAi.Desktop.Helper.exe') `
  -EvidenceDirectory (Join-Path $q006Root 'isolation')
```

The wrapper creates a one-shot task at the current user's limited privilege level.
Its helper owns a generated `WinSta0/VBAiTests_<guid>` desktop and starts the
worker and Excel there with original native process handles. It never switches
the input desktop. The UI canary reads UI Automation identity and posts `BM_CLICK`
only to its own verified button; it does not use UIA Invoke or global input.
The fixture checks the actual application window's desktop and loaded assembly
MVID before native work. This isolates windows and focus, not user/profile data.

Execution refuses an existing Excel process and a previously used campaign
ledger. It records each scenario once and stops subsequent work when a host
remains, no terminal report exists, or managed preflight fails. A known assertion
failure with verified host exit can be followed by the next independent planned
scenario. `-ScenarioIds` prepares a specific subset after a diagnosed harness
change; there is no automatic retry of a native command or failed campaign.

The scalar cases require exact synthetic values, declared-candidate refusals,
identity/source/selection/mode preservation and terminal native phases. The
extended type matrix uses bounded pages to fit the installed diagnostic trace
budget. Its single full-page diagnostic remains strict and available separately;
missing terminal evidence cannot pass. The persistence case uses product Save,
discard-close and normal exit before read-only disk reopening in a fresh owned
process with macros disabled. Expected module/class/form sources and the Label
are fixed before Save. Neither reopen nor cleanup invokes helper Save.

Successful detailed reports live in the retained host directories under
`VBAi_EXCEL_RESULTS`, independently of VSTest deployment cleanup. Keep plans,
TRX files, phase traces, request ledgers, startup and shutdown receipts together.
The helper closes its handles/desktop only after original child exit and verified
empty window inventory. An invisible inventory sentinel distinguishes a valid
empty desktop from an enumeration failure. Uncertain ownership retains the
launcher without forced termination or cleanup replay. Any later diagnostic
intervention must retain the failed lifecycle result.

See the [Q-006 results](../docs/test-coverage.md#q-006-inactive-desktop-excel-qualification-2026-10-02)
for candidate identity, observed passes, failed preparations and remaining gates.
SOLIDWORKS is outside this autonomous campaign.

## Settled Word scope collection (qualification only)

`VBAi_TEST_WORD_SETTLED_SCOPE_GC=1` enables a disabled-by-default testhost
experiment in `RegisteredOfficeRunsBatchAndSingleWithVerifiedResults("Word")`.
After its successful non-inlined qualification scope returns, the fixture verifies
terminal bridge work, settled native execution and the original owned process
before collecting unreachable testhost objects. `word-settled-scope-gc.json`
records timestamps and collection counts. This does not prove an RCW leak or run
collection in the product. Failed or uncertain scopes skip collection; the single
Close/Quit sequence, original process handle and 5-second exit bound remain intact.

`VBAi_TEST_WORD_EXIT_WAIT_BOUND_MS=15000` separately opts Word qualification
into a 15-second exit observation after its single confirmed Quit and COM release.
The fixture captures the bound before Close/Quit and records it with elapsed time,
pump attempts and the original process handle. Other hosts and an unset variable
retain the 5-second bound; any other configured Word value is refused before
Close/Quit. This observation diagnostic neither enables scope collection nor
replays native cleanup. An exit requires an observed code from the retained handle;
a timeout preserves ownership and remains a failure.

## Owned Excel teardown trace (qualification only)

`VBAi_TEST_EXCEL_TEARDOWN_TRACE_GATE=1` enables a disabled-by-default diagnostic
only in the two `ExcelProcedureValuesTests` native array/ParamArray cases.
`VBAi_RUN_EXCEL_TESTS=1` and an absolute `VBAi_EXCEL_RESULTS` are required. Run one
fresh disposable case, never the complete batch with this gate. Pending or
uncertain bridge work, identity mismatch or the bounded arming deadline preserves
the exact host without Close/Quit. The ordinary fixture is unchanged when the
diagnostic variable is absent. The gate never replays a procedure or cleanup.

Compile `tools/probes/OwnedTeardownTrace.Helper.cs` as a separate x64 executable
named `VBAi.OwnedTeardown.Helper.exe`. The helper is not an Office host. First
prepare, then explicitly execute `Test-OwnedExcelTeardownPreflight.ps1` with its
absolute `-HelperPath` and `-OutputRoot`. Its four owned helper trials must prove
normal detach/STOP/exit, synthetic terminal `c0000409` and `c0000005` collectors,
and a software first-chance AV forwarded unhandled to the helper's local handler.
These synthetic exceptions do not reproduce or explain an Office fault.
The earlier filesystem-trace preflight does not prove these handlers.
Default script invocation prepares a plan; `-Execute` is required to run it.

The opted-in native test writes `teardown.pending.json` in its owned GUID root
immediately before its existing single Close/Quit sequence. Prepare
`Trace-OwnedExcelTeardown.ps1 -PendingReport <absolute path> -ExpectedMvid <guid>
-ExpectedAssemblySha256 <hash>`; then add `-DebuggerPreflightReport <absolute
passing preflight.json> -Execute` for the exact same candidate/PID/start/nonce.
Only exception record, live event-thread registers and stack are captured, without full
memory dumps, global WER policy or additional bridge/COM calls. A controller
timeout permits bounded debugger detachment only after proving the exact owned
stop-breakpoint event; pending faults are retained without `qd`. A safe trial
without a verified terminal fault is `NOT_REPRODUCED`, including a deadline;
it never terminates Excel. Primary scenario and diagnostic cleanup errors remain
separate failures. A captured crash does not turn the original case into a pass.
A raw VSTest pass after detaching at a pending fault is invalid for crash
qualification; preserve that result and its intervention evidence separately.

See [development setup](../docs/development.md), [recorded results](../docs/test-coverage.md)
and the [compatibility matrix](../docs/compatibility.md). Run commands from the
repository root in a Windows development environment.

## Build and run

### Grouped execution without interrupting the working desktop

Prepare the complete qualification scenario set, build once into an isolated
output and freeze the source revision and assembly hashes before execution.
Avoid repeated full-suite runs between individual scenario changes. The default
suite includes real WinForms `Show`/focus tests even with Office opt-ins disabled;
running it directly can display windows on the working desktop.

`VBAi.Desktop.Helper` and `tools/tests/Invoke-IsolatedTests.ps1` launch a reviewed
campaign script on a generated, inactive Windows desktop under the same user's
limited interactive token. The helper uses explicit
[CreateProcess desktop selection](https://learn.microsoft.com/en-us/windows/win32/api/processthreadsapi/ns-processthreadsapi-startupinfow),
never switches the input desktop and has no process-termination path. A disposable
WinForms canary must prove the exact private desktop, native/UIA control
identity, an addressed `BM_CLICK` on its owned button and normal exit before the campaign script starts.
The legacy UIA `Button.Invoke` proxy uses input injection; the canary reads UIA
identity without invoking that proxy. This does not qualify other UIA actions. A failed
canary refuses the campaign without falling back to the working desktop.

```powershell
& "$env:WINDIR/System32/WindowsPowerShell/v1.0/powershell.exe" -NoProfile -File tools/tests/Invoke-IsolatedTests.ps1 `
  -ScriptPath 'E:\Qualification\Invoke-FrozenCampaign.ps1' `
  -HelperAssembly "$PWD/artifacts/build/VBAi.Desktop.Helper/Debug/net48/VBAi.Desktop.Helper.exe" `
  -EvidenceDirectory "$PWD/artifacts/qualification-private-desktop"
```

The script and helper paths must be absolute and existing; the evidence directory
must be new. The one-shot limited task writes its terminal receipt separately.
The helper retains original process handles across pending or uncertain outcomes;
it does not replay a host action or terminate a host. Remaining private windows
after a script exit refuse completed cleanup and retain desktop ownership. Record
desktop/canary, candidate, TRX and original host-exit evidence separately: desktop
isolation is not a security sandbox, a coverage result or Office acceptance.

The candidate script must explicitly prepare its runtime dependencies, including
the Node executable required by JavaScript scenarios. A limited scheduled task
does not inherit the development shell's extra PATH entries. Freeze the reviewed
executable path/hash and add only its directory to the worker's process PATH.
Office tests without their exact opt-in become inactive before native desktop
checks; enabled private Word tests validate their desktop before COM activation.

`VBAi.Desktop.Helper --probe-actions <generated desktop> <fresh absolute receipt>`
prepares a complete synthetic action matrix for text, tabs, scope selection,
native/custom buttons, a virtual ToolStrip item and modal closure. Invoke it only
inside the isolated worker. It records each action once with independent readback;
failed actions are not replayed. An action gap is not native acceptance and must
block the corresponding Word UI phase. The basic launcher canary separately
checks both populated and empty inventories on the owned desktop before starting
any campaign. On the observed Windows 11 host an empty valid desktop returns
false/error 0 without callbacks; errored or partial inventories remain refusals.

The private Word fixture persists its returned CreateProcess identity and captures
the original process before window discovery. The primary thread's desktop query
is a recorded diagnostic: an unavailable name provides no placement proof, and
an observed different desktop refuses. Attachment requires the unique `_WwG`
document child under an owned visible `OpusApp` root, the exact launched PID and
actual UI thread's named inactive desktop, before the first NativeOM/COM call.
The receipt also records the requested raw command line. The single Word `/a`
switch stays literal and only the macro-free seed path is quoted, following
[Microsoft's Office command-line syntax](https://support.microsoft.com/en-us/office/lifecycle/command-line-switches-for-microsoft-office-products).
That requested string is not evidence of startup behavior or cleared recovery state.
During startup, only an unambiguous window whose process, root, classes, UI thread
and exact private desktop have already been verified may be observed again while
hidden. A foreign or unreadable identity refuses immediately. The bounded wait
records native metadata before the readiness check; both document and root must
be visible and pass the final identity check before any NativeOM call.
Discovery uses a bounded inventory of the exact named desktop; duplicate document
windows, partial enumeration, foreign identities or an unavailable actual UI
desktop refuse without activation or native replay. The only open document must
still be the exact macro-free seed. A recovered older disposable document does
not satisfy that condition and is retained for a separate diagnosis.
The seed fingerprint is captured before launch. Once Word owns the file, the
fixture reads it with read/write sharing and requires the same fingerprint; it
never writes through that stream. Sharing violations, missing files and changed
bytes remain failures without retries. Executable fingerprinting retains its
ordinary read-only sharing behavior.

Discover an expanded ComboBox's list through its exact `GetComboBoxInfo` list
HWND, then require the unique enabled UIA ListItem under that HWND, the same
process/thread and private desktop. A dropdown need not expose its items as
descendants of the ComboBox itself. A WinForms Options popup may expose Menu or
the legacy ToolBar UIA role. The latter is accepted only with the exact native
WinForms dropdown class, retaining the unique newly visible popup, frozen owner
shape, localized virtual MenuItem and ancestor checks. Its MSAA default action
requires a unique exact-label child with MenuItem role 12; discovery does not
authorize a different toolbar or a repeated action.
After private scope selection, the Options popup may be owned by that exact
native ComboBox child. Only the previously frozen scope-picker HWND/class and
complete native owner metadata are admitted, with its chat ancestry rechecked;
other children remain refusals. This does not change the Git modal's separately
verified VBE-root owner.

Native Word isolation requires `VBAi_TEST_DESKTOP_NAME` and the reviewed installed
`VBAi_TEST_WORD_EXE`. Its fixture must explicitly launch that binary on the private
desktop and attach only its verified owned PID through NativeOM. Ordinary COM
activation is not evidence of private desktop placement. Keep native opt-ins
disabled for the default managed phase; run the prepared Word identity, owner Git
and Chat-to-Git scenarios together in the subsequent native phase. SOLIDWORKS
continues to require a preloaded, explicitly selected instance.

### Owned Word embedded Git qualification

`WordEmbeddedGitWindowTests` requires `VBAi_RUN_WORD_EMBEDDED_GIT_TESTS=1`,
`VBAi_RUN_WORD_GIT_TESTS=1`, `VBAi_RUN_OFFICE_TESTS=1`, an absolute
`VBAi_OFFICE_RESULTS`, and the exact retained synthetic repository manifest in
`VBAi_TEST_GITHUB_MANIFEST`. Set `VBAi_TEST_EMBEDDED_GIT_MVID` and
`VBAi_TEST_EMBEDDED_GIT_SHA256` to the frozen candidate. Run this case as an independent owned-host case in the grouped campaign,
with no existing Word process. It creates and saves an inert owned DOCM, exports
an independent baseline through the installed bridge, then invokes the tagged
production Git menu once. The MTA UI worker verifies the native PID, VBE thread
and modal owner before fetching the pinned synthetic branch, creating an exact
local checkpoint and comparing. No remote push, import or macro execution occurs.
An uncertain menu/UI outcome retains the original host; known terminal closure
permits normal teardown on its owning STA. Saved bytes, source, references and
normal exit remain required. This scenario does not qualify the chat Git button
or provider conversation; native acceptance must be recorded separately.

`WordChatGitWindowTests` additionally requires `VBAi_RUN_WORD_CHAT_GIT_TESTS=1`.
Run it as an independent owned-host case in the grouped frozen campaign. It maps the owned saved
Word project through canonical bridge fields, selects its unique chat UI label,
then invokes the actual chat GitHub menu item once. Discovery uses a bounded native
HWND inventory and the unique chat caption/control shape, with a content-free
receipt before action. Native
`EnumChildWindows` inventory uses an explicit callback limit, preserving partial
inventory refusal while ignoring its
[unused return value](https://learn.microsoft.com/en-us/windows/win32/api/winuser/nf-winuser-enumchildwindows).
The Git modal must belong to the observed chat's verified
top-level root and VBE thread. An anchored child chat resolves its native modal
owner to the VBE root, following the
[Win32 owner contract](https://learn.microsoft.com/en-us/windows/win32/winmsg/window-features#owned-windows).
Git ToolStrip items may expose no AutomationId or leaf HWND. The bank selects the
exact localized product label only inside the unique newly visible native Options
popup, verifies its PID/thread/owner and rechecks the virtual item's native
ancestor before invocation. Inventory receipts contain identities and match
counts, without menu/conversation contents.
Top-window enumeration retains a global callback bound, then counts only visible
windows on the exact owned VBE thread against its target limit. A refusal records
the API return, global/owned/target counts and failure branch; API failure remains
a failure. Hidden and foreign-thread windows do not authorize actions or consume
the target limit.
The exact root relationship is frozen before invocation and rechecked before
closure. Exact closure, unchanged Word
source/references/saved bytes and normal exit remain required; opening must not
create a repository binding. No provider prompt, Git connection or import occurs.
Both Word scenarios read reference identities through the owning-thread bridge;
external testhost reference property getters are not part of this preparation.
The existing `VBAi_TEST_WORD_EXIT_WAIT_BOUND_MS=15000` opt-in changes only the
single read-only original-handle exit observation. Record that bound with the
native result; it does not replay Close/Quit or make an earlier timeout successful.

`WordGitProjectIsolationTests.TwoSameNameWordDocumentsUseCanonicalGitScopesAndRefuseStaleSaveAs`
qualifies the owned two-document Word path, duplicate-name refusal, native
selection, stale SaveAs binding refusal before CodePane access, and preservation
of the other document's source without calling `VbaGitProject.Capture` or
exporting a component. The historical external-STA Capture and distinct export
routes are a separate diagnostic method requiring
`VBAi_RUN_WORD_GIT_CAPTURE_DIAGNOSTIC=1` as well as the Word Git/Office opt-ins.
Its failures remain failures and do not supply canonical-path acceptance.
For grouped runs on an inactive test desktop, the campaign runner supplies the
generated `VBAi_TEST_DESKTOP_NAME`, an absolute `VBAi_TEST_WORD_EXE`, and its
exact `VBAi_TEST_WORD_EXE_SHA256`. The Word-only fixture then launches that
binary through the desktop launcher with Microsoft's documented
[`/a` switch](https://support.microsoft.com/en-us/office/lifecycle/command-line-switches-for-microsoft-office-products)
and one disposable macro-free DOCX seed. It attaches through the launched PID's
unique Word NativeOM window, verifies the application, sole seed document and
actual UI-thread desktop before showing Word or creating a DOCM. Attachment
failure retains the original launched process without a COM activation fallback.
The single `/a` switch prevents automatic loading of Normal and global templates
before attachment. This is an isolated startup qualification, not evidence for
ordinary Word startup with user templates. A missing VBAi connection remains a
failure; the fixture does not retry with ordinary startup. Before creating the
owned DOCM, the fixture keeps `AutomationSecurity=3` and calls
[`WordBasic.DisableAutoMacros`](https://learn.microsoft.com/en-us/office/vba/word/concepts/customizing-word/auto-macros)
on its verified owning STA. Suppression is not reenabled before Quit. An uncertain
command retains the original host and WordBasic lease without Close/Quit replay.
Without this opt-in, the existing Office fixture startup path applies.

The separate `WordChatGitWindowTests` observes the Options popup after the
owned Word chat selects its canonical saved project. A native Menu popup may be
owned by the observed VBE root or by a hidden, standalone WinForms drop-down
owner on the same process and UI thread. The hidden owner must share the popup's
WinForms application-domain class suffix, have no native parent or owner, and
carry the observed tool-window and window-edge styles. Its exact native shape is
checked again before the Git item is invoked. The Git modal retains its separate
VBE-root ownership rule. WinForms defines a distinct
[drop-down owner window](https://referencesource.microsoft.com/System.Windows.Forms/winforms/Managed/System/WinForms/ToolStrip.cs.html)
and [uses it for top-level drop-downs](https://github.com/dotnet/winforms/blob/main/src/System.Windows.Forms/System/Windows/Forms/Controls/ToolStrips/ToolStripDropDown.cs).

### Disposable owner-import root-font observation

The existing owner-import scenario can be instrumented with
`VBAi_RUN_ROOT_FONT_OBSERVATION_TESTS=1`, an absolute fresh
`VBAi_TEST_ROOT_FONT_OBSERVATION_MANIFEST` path inherited before Excel starts,
and `VBAi_TEST_ROOT_FONT_OBSERVATION_MODE=ObserveWrites`. It also requires the
ordinary owner-import opt-ins, pinned synthetic remote and exact candidate
identity. The fixture publishes the manifest after preparing the saved baseline
and before invoking the UI; the product validates it immediately before import.
The manifest binds the disposable project, target form source and root font
descriptor, loaded MVID and fresh nonce-named evidence directory.

The diagnostic records exported root bindings before font getters, child-value
readback and attached font descriptors using `IPersistStream.Save(false)`.
It preserves the final exact snapshot check. The separate `DistinctChildName`
mode predeclares one temporary Arial child-name write before the ordinary target
writes; it refuses an Arial target and is never a fallback after a failed trial.
The separate `AfterInitialCapture` mode defers the selected form's font delivery
until the first post-import snapshot. It requires that only that form's FRX
differs, then transfers one fresh exact StdFont to `Designer.Font` on the owning
thread. It performs no child-value writes or diagnostic exports before that
first capture, retains the imported component identity and preserves the final
strict snapshot check. This is a predeclared experiment, not an automatic retry
after an uncertain mutation. Native acceptance of this mode is not established.
For a separately declared synthetic baseline, set
`VBAi_TEST_ROOT_FONT_OBSERVATION_SEED_PROFILE=SyntheticExplicitArial9` with
`AfterInitialCapture`. This tests-only profile delivers one fresh Arial 9.00
root font before the initial SaveAs, records its intent and requires the exact
descriptor in the saved/reopened baseline before publishing the manifest. An
absent profile keeps the default preparation unchanged. A successful explicit
profile would not qualify default Tahoma serialization or Frame 8.27 fidelity.
The separate `RetainedSyntheticTahoma825` profile is restricted to LabelButton
and `AfterInitialCapture`. Its
`VBAi_TEST_ROOT_FONT_OBSERVATION_SOURCE_WORKBOOK` must identify the retained
synthetic `EmbeddedGit.xlsm` from the declared observation artifact, with
SHA-256 `B5264F941E0FD398A9DE03B203DB7A31A6DB9FF939B894E7829A0203C76F4B73`.
The fixture normally closes its empty owned workbook, creates one fresh copy,
and opens it with events and macros disabled. It does not save, replace code
markers or seed a font. Before scalar font readback or menu emission, it must
capture the pinned component/source/resource hashes and explicit Tahoma 8.25
root descriptor. Independent bridge export must then remain exact. A new
UI correlation nonce is separate from the retained synthetic source marker.
Wrong paths/hashes, occupied destinations, reparse points or an implicit-default
baseline are refused; an uncertain native Close/Open preserves the host.
This is a prepared discriminator, not proof of root restoration or Frame fidelity.
The disabled-by-default `VBAi_TEST_RETAINED_VBE_LIFETIME=BeforeAfterCopy`
observation is restricted to this retained profile. After confirming events
disabled and AutomationSecurity 3, it reads `Application.VBE` once while the
initial inert workbook remains open, releases that acquired reference, then
allows the single planned replacement. Getter or receipt/cleanup failure stops
before replacement. A second named getter is observed after verified copy-open;
neither access is retried. This compares two lifecycle stages under the same
security settings, but the initial getter itself can initialize VBE state.
Success cannot establish a causal repair or qualify cached VBE references as
authority for import. Unknown configuration values are refused before mutation.
Capture-only and persistence scenarios cannot enable this diagnostic. Without
the opt-in, there are no additional font getters, exports or writes. Recorded
instrumented trials reach terminal exact snapshot refusals; this diagnostic
does not qualify native import or persistence. See
[recorded validation](../docs/test-coverage.md).

### Disposable owner-import UserForm persistence qualification

`EmbeddedGitWindowTests.InstalledOwnerImportedFormSurvivesOneSaveAndFreshReadOnlyProcess`
is a separate prepared scenario for the declared native layouts. It requires
`VBAi_RUN_USERFORM_OWNER_PERSISTENCE_TESTS=1` in addition to the existing Excel,
UserForm GitHub, embedded Git UI and owner-restore opt-ins, the pinned synthetic
GitHub manifest, absolute evidence directory and exact candidate MVID/hash.
It must first verify the actual owner-UI checkpoint import, then attest the
selected bare repository and absence of recovery before one Save. Only normal
original exit permits a new explicit owned process to reopen the saved file
read-only with events/macros disabled. Exact source/resources/native fonts,
loaded candidate, unchanged post-Save disk bytes and fresh normal exit are
required. The existing unchanged-on-disk import scenario remains separate.
These persistence scenarios are currently NOT_RUN because the preceding
owner-import canary fails. See [recorded validation](../docs/test-coverage.md).

### Disposable native Excel signature qualification

`ExcelSignatureQualificationTests` requires `VBAi_RUN_EXCEL_TESTS=1`,
`VBAi_RUN_EXCEL_SIGNATURE_TESTS=1` and an absolute `VBAi_EXCEL_RESULTS`.
The first scenario saves synthetic nonexecuted code and checks native unsigned
state plus the closed-file Windows Office SIP result. The signing scenario also
requires `VBAi_TEST_SIGNING_CERTIFICATE_MANIFEST`, pointing to a JSON manifest
with `FormatVersion: 1`, a 32-character uppercase hexadecimal `Nonce`,
`Subject: "CN=VBAi Disposable Signature Qualification <Nonce>"`, an uppercase
40-character `Thumbprint`, `Store: "CurrentUser/My"` and
`TrustedRootInstalled: false`. Use an explicitly owned, short-lived synthetic
code-signing certificate only; never borrow a personal signing certificate or
change Root/TrustedPublisher stores or Office trust policies for a test.

The protected Windows certificate selector requires the user to confirm the
exact named certificate within the product deadline. The fixture sends one
signing request and retains its owned host if the response is nonterminal;
it never retries signing. Persistence requires a distinct fresh Excel process,
read-only reopen with macros disabled, unchanged source/file hashes, and normal
original-process exits without helper saving. Closed-file SIP verification
distinguishes absence, untrusted-root and verifier-unavailable outcomes; a
self-signed trust failure is never promoted to a valid/trusted signature.
Remove only the exact owned certificate and its key after terminal cleanup.
The recorded trial is failed; preparation of these scenarios is not acceptance.

Use a separate output directory while an application has the installed DLL loaded:

```powershell
dotnet build VBAi.sln -c Debug -p:BuildOutputRoot="$PWD/artifacts/build"
dotnet test tests/VBAi.Tests/VBAi.Tests.csproj -c Debug --no-build -p:BuildOutputRoot="$PWD/artifacts/build"
```

An isolated build does not replace the registered add-in. Native tests must verify
which assembly the target process actually loaded. Do not overwrite a loaded DLL
or close unrelated applications to make a build succeed.

For a category-specific pass, append `--filter "TestCategory=Unit"`. Some local tests
create real Windows controls, WebView2 instances, temporary Git repositories and
simulated provider processes. They require the relevant runtimes and an interactive
Windows desktop; they are not all platform-independent pure unit tests.

## Test organization

| Location | Purpose |
| --- | --- |
| `tests/VBAi.Tests/Unit/` | Contracts and branch behavior, with source-file mirrors. |
| `tests/VBAi.Tests/Integration/` | Storage, processes, runtime controls and host boundaries. |
| `tests/VBAi.Tests/Scenarios/` | Cross-component workflows and recovery scenarios. |
| `tests/VBAi.Tests/Infrastructure/` | Shared doubles, fixtures and host helpers. |
| `tests/VBAi.Git.Smoke/` | Standalone Git diagnostics; shared scenarios also run in VSTest. |
| `tests/VBAi.Providers.Smoke/` | Provider diagnostics and a simulated CLI process. |
| `tests/native/` | Native-renderer lifecycle and managed-loader checks. |

Tests reference the production assembly rather than recompiling its sources.
Mirrors follow production paths with a `.Tests.cs` suffix, including partial
classes. Cross-cutting scenarios should complement those mirrors rather than
create a second implementation of production logic.

```powershell
powershell.exe -NoProfile -File tools/tests/Test-TestLayout.ps1 -ReportPath artifacts/test-layout/mirror-inventory.json
```

Prompt files under `Infrastructure/Fixtures/Prompts/` and the test-environment skill
under `tests/Infrastructure/.agents/` are test inputs, not user documentation.
Preserve their contents unless deliberately changing the corresponding fixture.

## Native host tests are opt-in

The Q-027 path-visibility diagnostic is a temporary, explicit host opt-in. Run
only `ExcelPathVisibilityTests.OwnedOwnerStaObservesTwoTesthostCreatedGuidDirectoriesAndEffectiveToken`
with `VBAi_RUN_EXCEL_PATH_VISIBILITY=1`, `VBAi_RUN_EXCEL_TESTS=1` and an absolute
durable `VBAi_EXCEL_RESULTS` directory. The owned `/x /automation` bootstrap
passes `VBAi_TEST_PATH_VISIBILITY_MANIFEST` explicitly to the child. The manifest
contains only two direct GUID children of LocalAppData and TEMP. Fixed synthetic
files and the manifest remain under read leases during observation and are
retained as evidence. No personal file is read, no macro/export/source edit runs,
and no ACL, attribute, trust policy or token is modified.

The internal, parameter-free `diagnostic_path_visibility` bridge command is
disabled unless the host received this manifest before add-in connection. It is
absent from the LLM catalogue. It checks captured owner PID/native TID/STA and
records managed existence/attribute results, explicit exception types/HRESULTs,
native attributes and immediately captured LastError (meaningful only on native
failure). Effective token evidence queries the thread with OpenAsSelf=true and
falls back to the primary token only for ERROR_NO_TOKEN (1008); metadata is
limited to SID, integrity, restricted/AppContainer state, type/impersonation and
AuthenticationId/TokenId. Compare testhost before/after with the actual owner STA;
do not label the testhost's STA as Excel's. Visibility differences are recorded,
not assumed equal or attributed to a cause. Pending delivery/startup retains the
host and read leases without replay or native cleanup. This diagnostic changes
the product candidate; qualification of an earlier binary does not cover it.

After deploying and verifying the matching candidate, run exactly one native
observation (the command below is a recipe, not recorded execution):

```powershell
$env:VBAi_RUN_EXCEL_TESTS = '1'
$env:VBAi_RUN_EXCEL_PATH_VISIBILITY = '1'
$env:VBAi_EXCEL_RESULTS = "$PWD/artifacts/path-visibility-native"
dotnet test tests/VBAi.Tests/VBAi.Tests.csproj -c Debug --no-build -p:BuildOutputRoot="$PWD/artifacts/build" --filter 'FullyQualifiedName=VBAi.Tests.Integration.ExcelPathVisibilityTests.OwnedOwnerStaObservesTwoTesthostCreatedGuidDirectoriesAndEffectiveToken' --logger 'trx;LogFileName=path-visibility-native.trx' --results-directory "$PWD/artifacts/path-visibility-native"
```

The paired export diagnostic is separate from that read-only observation. Its
two data rows create fresh GUID children of LocalAppData and TEMP, then bootstrap
one fresh owned Excel per row with the same fixed-allowlist manifest. Production
commands create only a disposable UserForm and Label. The parameter-free owner
STA diagnostic immediately precedes one `export_component` request to
`ASCIIForm.frm` in the selected GUID child. Both commands use the same exact
PID pipe and production dispatcher; no COM call or host command intervenes.
Their sequencing is not an atomic proof that a thread token cannot change.

Evidence preserves actual owner native TID/STA, candidate MVID and SHA-256,
project/component/design identity, source SHA-256, native Label properties,
effective tokens, all four attribute outcomes and raw export responses/files.
Visibility differences remain observations. A received `Ok=false` is a terminal
failed export, with independent after-readback and normal-exit checks; it is not
a transport timeout. Pending startup/delivery/export retains the exact host and
manifest/synthetic read leases without Close/Quit, retry or target cleanup. No
ACL, EFS, permission, trust policy or token changes are made. The GUID directories
and original partial/successful FRM/FRX remain evidence. This does not qualify
Git capture normalization/import/recovery or equate explicit bootstrap with
historical COM-activation trials.

Prepare a **separate frozen-candidate test output** as described below, verify
its referenced product hash against the installed candidate, then run the whole
two-row matrix with this exact filter. This recipe is not recorded execution;
do not rebuild/deploy the product to run a tests-only diagnostic:

```powershell
$env:VBAi_RUN_EXCEL_TESTS = '1'
$env:VBAi_RUN_EXCEL_PAIRED_EXPORT_PATH_VISIBILITY = '1'
$env:VBAi_EXCEL_RESULTS = "$PWD/artifacts/paired-export-native"
dotnet test tests/VBAi.Tests/VBAi.Tests.csproj -c Debug --no-build -p:BuildProjectReferences=false -p:BuildOutputRoot="$PWD/artifacts/build-frozen-paired-export" --filter 'FullyQualifiedName=VBAi.Tests.Integration.ExcelPairedExportPathVisibilityTests.OwnerStaVisibilityImmediatelyPrecedesOneExportInOwnedGuidDirectory' --logger 'trx;LogFileName=paired-export-native.trx' --results-directory "$PWD/artifacts/paired-export-native"
```

For an owned-host scalar-inspection investigation, set `VBAi_VBE_INSPECTION_TRACE`
to an absolute local JSONL file path in the host's environment before launching it.
The parent directory must already exist. Tracing is disabled when absent or invalid;
unavailable logging never changes command execution. Evidence is bounded to 128
events per inspection and 1 MiB per file. Rows contain correlation, fixed phase,
PID, thread/apartment, elapsed time and exception type; they contain no project
path, code, expression or inspected value. A client timeout does not cancel or
authorize replay of a pending native command. Retain the host and phase evidence
when the bridge stops responding.

The `ExcelScalarDiagnostics` scenarios use a dedicated explicit bootstrap,
not the ordinary Excel COM-activation fixture. Set `VBAi_RUN_EXCEL_TESTS=1`,
`VBAi_VBE_INSPECTION_TRACE` to an absolute local JSONL path with an existing parent
directory, and `VBAi_EXCEL_RESULTS` to a durable local evidence directory. The
bootstrap resolves the installed x64 Excel executable from HKLM App Paths and
launches it once with `/x /automation` and an owned, empty, macro-free `.xlsx`
seed. Microsoft documents that
[/automation suppresses automatically opened files and auto-run macros](https://learn.microsoft.com/en-us/troubleshoot/microsoft-365-apps/excel/files-open-automatically).
The trace variable is supplied explicitly in the child process environment;
setting it only in the test process does not prove that COM activation inherited
it. No machine/user environment setting is changed.

NativeOM attachment accepts only the launched PID's `EXCEL7` document. Before
any workbook/VBE mutation, the fixture checks the retained process handle,
executable/start identity, application HWND/PID and an inventory containing only
the exact seed workbook. It never falls back to COM activation or an active
application. Startup, loaded add-in MVID and shutdown evidence remain in the
fixture directory even on success. If attachment or startup is uncertain, the
process is retained without Close, Quit, termination or another launch. Execute
`InstalledBridgeSkipsUnsupportedScalarPageWithoutQuickWatch` first, then
`InstalledBridgeReadsOneLongScalarWithNativePhaseEvidence`, and then
`InstalledBridgeReadsFullScalarPageWithThreeNativeObservers`, using separate exact
method filters; a category batch does not guarantee their order. Each requires
nonempty correlated host phase evidence and the installed product MVID must
match the assembly referenced by the tests.

The full page issues one `inspect_local_scalars` request with Offset=0/Limit=6.
It checks the exact Long/String/Boolean probe values, declaration order/type/
position, and array/Variant/object rows skipped without a value or error. Native
phase evidence must contain one correlation and terminal inspection, with a
complete Command229/observer lifecycle for each eligible scalar. Selection,
source and break mode are checked before and after; normal owned Close/Quit and
exit are required. These synthetic runtime assertions do not establish complete
Locals enumeration or explain a prior crash. Recorded acceptance and loaded
candidate identity are in [validation results](../docs/test-coverage.md).

For a tests-only follow-up against a frozen installed candidate, use a separate
output containing hash-verified copies of its product/dependency binaries and
build the test project with `BuildProjectReferences=false`. Verify the product
hash again in the final test output; do not rebuild the product implicitly or
overwrite a test assembly whose native campaign is still active. The exact full
page filter is:

```text
FullyQualifiedName=VBAi.Tests.Integration.ExcelLocalScalarInspectionTests.InstalledBridgeReadsFullScalarPageWithThreeNativeObservers
```

Office adapter-only diagnostics distinguish transient Access CurrentProject
wrapper identity from the database path and selected VBProject identity. Keep
sampled objects alive across identity comparisons and balance their temporary
IUnknown references; do not weaken native dispatch guards or invoke Save to
diagnose an identity getter. A read-only identity trial with normal exit does
not qualify adapter persistence. Adapter/property/reference acceptance requires
its own verified save and fresh-disk reopen without post-save helper saving.

For Access save/reopen, preserve the exact live and reopened metadata strings,
not only a successful Save response or Saved flag. Reference-removal acceptance
requires the initial owned process to exit normally before a fresh process
reopens the database. A returned Quit is insufficient; preserve the disposable
database and report disk readback as NOT_RUN if exit is unverified. Do not replay
Quit or native mutations to obtain a passing result.

Office fixtures retain the original process handle, PID/start/image identity and
shutdown failure when Quit returns but exit is not observed within the fixture
deadline. `shutdown-lifecycle.json` records the single Quit outcome, bounded exit
observation and handle disposition; the final qualification report retains this
ledger too. An unverified shutdown refuses further native requests, saving,
Close/Quit attempts and reopening. Later PID absence, manual termination or an
eventual exit does not overwrite the original failed gate or establish an exit
code that was not read. Releasing a client COM reference is not proof that every
server reference disappeared. These retention checks do not diagnose a host's
failure to exit.

### Read-only Access metadata getter probe

Set both `VBAi_RUN_OFFICE_TESTS=1` and
`VBAi_RUN_OFFICE_METADATA_GETTER_PROBE=1` to enable this diagnostic. The
`FreshAccessMetadataGetterContractsReadOnly` test starts a new owned disposable
database through the existing Access fixture. Database creation and its initial
fixture save remain prerequisites; the probe adds no setter, save, macro or help
invocation. The baseline requires three successful, exactly equal getters and
runtime getter/setter contracts for both DISPIDs. It does not qualify metadata
persistence.

The diagnostic reads `HelpFile` and `HelpContextID` on the exact mapped, selected
VBProject through `PropertyDescriptor.GetValue`, CLR `InvokeMember` with
`GetProperty`, and raw `IDispatch.Invoke` with `DISPATCH_PROPERTYGET`. The raw
result uses the x64 24-byte VARIANT ABI, a checked boundary canary and OLE cleanup.
Runtime `GetTypeInfo` records the getter/setter VARTYPEs for DISPIDs 116 and 117
and the containing library identity. Reads run on the external fixture STA;
COM marshaling dispatches them to the Office object's apartment. These are not
in-process bridge getter observations.

Each `metadata-getters-<phase>.json` retains candidate MVID and matching assembly
file hashes, PID/start identity,
project path/IUnknown identity, mode/protection/Saved state, getter outcomes and
bounded exact BSTR bytes. No ANSI repair or value normalization is performed.
A pending-read marker identifies an unfinished call. Normal owned process exit
must be verified separately through the fixture's lifecycle evidence.

With the same diagnostic opt-in, the existing Access HelpFile/HelpContextID
adapter scenarios also record reads before and after their existing mutation,
after their existing adapter save, and immediately after fresh-disk reopen.
These hooks add no mutation or save. Getter disagreements remain observations;
the original exact metadata assertions still decide adapter acceptance. An
`OBSERVED` report does not mean its getters agree or persistence passed.

Prepare a new durable evidence directory and use a test output referencing the
exact installed candidate as described above. The following filter selects the
complete diagnostic batch; no ordering between tests is assumed:

```powershell
$env:VBAi_RUN_OFFICE_TESTS = "1"
$env:VBAi_RUN_OFFICE_METADATA_GETTER_PROBE = "1"
$env:VBAi_OFFICE_RESULTS = "$PWD/artifacts/metadata-getter-evidence"
dotnet test tests/VBAi.Tests/VBAi.Tests.csproj -c Debug --no-build -p:BuildOutputRoot="$PWD/artifacts/build" --filter "FullyQualifiedName=VBAi.Tests.Integration.OfficeMetadataGetterProbeTests.FreshAccessMetadataGetterContractsReadOnly|FullyQualifiedName=VBAi.Tests.Integration.OfficeAdapterOnlyMetadataQualificationTests.Access16HelpFilePathAdapterSaveReopen|FullyQualifiedName=VBAi.Tests.Integration.OfficeAdapterOnlyMetadataQualificationTests.Access16HelpContextIdAdapterSaveReopen" --results-directory "$PWD/artifacts/metadata-getter-evidence" --logger "trx;LogFileName=metadata-getters.trx"
```

### One-shot Access metadata setter comparison

The additional `VBAi_RUN_OFFICE_METADATA_SETTER_PROBE=1` opt-in enables four
`OfficeMetadataSetterProbe` cases: HelpFile/HelpContextID, each through the
existing production CLR `SetNative` implementation or raw `IDispatch` PROPERTYPUT.
Every case creates a separate owned Access database, prepares the existing
synthetic module/class baseline and attempts one metadata setter. A returned
failure is followed only by read-only getter observations before the original
error is rethrown. Independent read/cleanup failures are aggregated; uncertain
bridge delivery refuses further native reads or mutations. There is no setter
replay, rollback claim or post-failure save.

Both new setter paths run from the same external fixture STA, use the exact
selected project identity and current production revision, and require design
mode, no project protection and the verified candidate MVID/assembly-file hashes.
Their context differs from the in-process bridge path; include the existing
adapter HelpFile/HelpContextID scenarios as separate controls. A difference
between a bridge case and an external case alone does not establish a binder bug.

The raw setter uses exact BSTR/I4 inputs without coercion, architecture-correct
24-byte argument/result buffers, checked canaries and OLE cleanup. It passes one
named `DISPID_PROPERTYPUT` argument, `DISPATCH_PROPERTYPUT` only and the invariant
locale used by the CLR comparison. Microsoft documents that
[PROPERTYPUT requires the named argument and ignores the result](https://learn.microsoft.com/en-us/windows/win32/api/oaidl/nf-oaidl-idispatch-invoke).
`metadata-setter.json` retains the original native HRESULT and input/result
buffer observations. Successful calls proceed through the unchanged adapter
save, normal process exit and fresh-disk exact metadata assertions. Byte-packing
is retained as observed; these diagnostics add no ANSI recovery heuristic.

After preparing a matching frozen-candidate test output and a new evidence
directory, run the four cases and both bridge controls together:

```powershell
$env:VBAi_RUN_OFFICE_TESTS = "1"
$env:VBAi_RUN_OFFICE_METADATA_GETTER_PROBE = "1"
$env:VBAi_RUN_OFFICE_METADATA_SETTER_PROBE = "1"
$env:VBAi_OFFICE_RESULTS = "$PWD/artifacts/metadata-setter-evidence"
dotnet test tests/VBAi.Tests/VBAi.Tests.csproj -c Debug --no-build -p:BuildOutputRoot="$PWD/artifacts/build" --filter "TestCategory=OfficeMetadataSetterProbe|FullyQualifiedName=VBAi.Tests.Integration.OfficeAdapterOnlyMetadataQualificationTests.Access16HelpFilePathAdapterSaveReopen|FullyQualifiedName=VBAi.Tests.Integration.OfficeAdapterOnlyMetadataQualificationTests.Access16HelpContextIdAdapterSaveReopen" --results-directory "$PWD/artifacts/metadata-setter-evidence" --logger "trx;LogFileName=metadata-setters.trx"
```

The native project-properties dialog exposes these Help fields on its
[General tab](https://learn.microsoft.com/en-us/office/vba/language/reference/user-interface-help/project-properties-dialog-box).
The existing guarded `queue_project_properties_dialog` command opens that
dialog; it does not edit its controls. UI field discovery, a separately owned
single-change trial and normal save/reopen proof are prerequisites for evaluating
that alternative. No native dialog edit is included in this setter batch.

### Controlled native export tracing

`tools/probes/Trace-NativeUserFormExport.ps1` defaults to a prepare-only plan.
Its explicit `-Execute` path requires the absolute pending report, expected
MVID/SHA-256 and a measured `-DebuggerPreflightReport` from
`tools/probes/Test-NativeTracePreflight.ps1`. That preflight uses only a new
disposable non-Office helper and must prove paired native tracing, attachment,
target survival, normal shutdown and debugger detach with the exact CDB,
JsProvider and trace-script hashes. It does not qualify native Office export.

Trace execution validates the owned Excel PID/start identity and existing
absolute GUID export child, refuses a preexisting trace output or attached
debugger, and arms the export permission marker only after attachment is proven.
The bounded capture records paired syscall arguments/statuses and its detach
lifecycle; it does not issue exports or change ACLs, EFS, tokens or trust. The
fixture issues one export only. If setup fails before arming, preserve zero-export
evidence; if a native export fails, retain its original response and do not replay
it. Record detach and owned-host shutdown independently from export acceptance.

Pending/preflight JSON is read explicitly as UTF-8, including accented repository
paths. CDB command files intentionally use the active Windows ANSI code page
without a BOM; debugger logs use Unicode. These are separate encoding contracts.
A native path/name-not-found status despite a successful synthetic root write
localizes the observed failure but does not prove an EFS, ACL or token cause.
Exact candidates and measured results belong in recorded validation.

If a native campaign aborts, retain the final TRX and reconcile individual
outcomes with its raw counters and the planned scenario inventory. An in-flight
scenario absent from the TRX and unreached cases are NOT_RUN; NotExecuted is not
a successful fixture skip. Explicitly authorized forced cleanup is never normal
host shutdown acceptance. Record native debugger ownership, any termination
timeout and whether an exit code was actually observed; transient PID absence
does not establish completed cleanup.

Use disposable documents and identify the intended process/project before any
write. Record application version, architecture, language, DPI and loaded VBAi
build. Restore temporary settings and verify that unrelated documents remain
unchanged. Skip unavailable hosts honestly instead of treating a skip as a pass.

| Opt-in variable | Scope |
| --- | --- |
| `VBAi_RUN_EXCEL_TESTS=1` | Excel integration tests (`TestCategory=Excel`). |
| `VBAi_RUN_OFFICE_TESTS=1` | The additional Office host qualification fixtures. |
| `VBAi_RUN_OUTLOOK_TESTS=1` | Outlook tests; a usable profile is also required. |
| `VBAI_EDITOR_EXCEL_TEST=1` | Monaco/Excel roundtrip (`TestCategory=MonacoExcel`). |
| `VBAI_EDITOR_LANGUAGE_EXCEL_TEST=1` | Language-service qualification (`FullyQualifiedName~MonacoLanguageExcelTests`). |
| `VBAI_NATIVE_PALETTE_EXCEL_TEST=1` | Native palette (`TestCategory=NativePaletteExcel`); observe the fixture's other-host exclusions. |
| `VBAi_SOLIDWORKS_PID` | PID of a user-preloaded SOLIDWORKS/VBE instance (`TestCategory=SolidWorks`). |

Set only the variables needed for the intended run and remove them afterward.
The formatting-options scenario can additionally use an absolute
`VBAi_TEST_FORMAT_OPTIONS_OUTPUT` directory to retain request/response and
before/after evidence independently of VSTest attachment retention, including
successful runs. It creates a unique subdirectory and never retries a failed
mutation to gather evidence.

For Q-026 investigation, `tools/probes/inspect_q026_options.py <before.json>
<after.json> --output <fresh-report.json>` compares recorded complete snapshots
offline. It preserves catalogue order, palette fields and value types, and refuses
truncated, ambiguous or incomplete receipts. It neither invokes a host nor proves
native restoration. The [prepared Q-026 scenario matrix](../tools/tests/q026-scenarios.json)
records the native cases and stop rules; its planned status is not execution proof.

### Q-026 inactive-desktop campaign

`tools/tests/Invoke-Q026Qualification.ps1 -Prepare` copies the reviewed
candidate into a fresh evidence directory and builds only the focused harness
and desktop helper. It does not replace the installed DLL or modify registration.
Use Windows PowerShell 5.1 for this .NET Framework preparation:

```powershell
$q026Evidence = Join-Path $PWD ('artifacts/q026-' + [Guid]::NewGuid().ToString('N'))
powershell.exe -NoProfile -File tools/tests/Invoke-Q026Qualification.ps1 -Prepare `
    -EvidenceRoot $q026Evidence -InstalledDirectory 'E:\path\to\installed\net48'
powershell.exe -NoProfile -File tools/tests/Invoke-IsolatedTests.ps1 `
    -ScriptPath (Join-Path $q026Evidence 'Invoke-FrozenQ026.ps1') `
    -HelperAssembly (Join-Path $q026Evidence 'build/VBAi.Desktop.Helper/Debug/net48/VBAi.Desktop.Helper.exe') `
    -EvidenceDirectory (Join-Path $q026Evidence 'isolation')
```

The target path must identify the reviewed candidate; it can be frozen before
temporary registration, but must match the registered CodeBase before execution. Preparation
freezes hashes, MVID, source revision/status and the scenario matrix. Execution
requires a verified private-desktop worker, passing managed guards and a continuous
30-second interval without another Excel, Word, PowerPoint, Access, Publisher or
SOLIDWORKS process. Preferences can be shared across desktops: a host retained by
another qualification must be recovered by its owner before this campaign starts.
A quiet process inventory can also be a gap between cases in another native
batch. Obtain that batch owner's explicit release before scheduling Q-026;
the automated 30-second interval does not establish that the batch has ended.
The campaign waits without terminating another process or switching desktops.
It checks the per-user x64 CodeBase before launching its owned disposable Excel.

`campaign.json`, the native TRX and durable `native/phases` receipts are separate
evidence. A passing TRX requires independent review of the complete baseline and
restored tab structures, refusal closure, actual loaded assembly, original-handle
normal exit and the private-desktop terminal inventory. An empty size catalogue
qualifies only the documented refusal; it does not qualify a font-size mutation.
The evidence root is one-shot. Failed, pending or uncertain work must not be
replayed; preserve its receipts and exact host ownership for recovery.

Preparation with `-Scenario Margin` selects the separate checkbox-only native
method and a reduced matrix (read stability, real margin transition, complete
restoration and normal exit). A pass qualifies that diagnostic, not full Format.
`-Scenario FontSizeCatalogue` selects only font/size inspection and mutation,
with full-snapshot restoration and original-handle normal exit. It records the
native catalogue before and after the font change. An empty list still exercises
the explicit refusal and does not establish size-mutation acceptance; palette,
margin and historical-causality qualification remain separate.
`-Scenario SizeFocus` opens the existing native Options menu once on the fixture's
owning STA. A worker on the same inactive desktop selects the observed Format
tab, posts one `WM_NEXTDLGCTL` to the owned Size control, observes focus and the
native catalogue, then cancels once. No preference is written. Acceptance requires
complete unchanged Tabs/revision and normal exit through the original launch
handle; unknown focus, modal or dispatch outcomes retain the owned host.
`-Scenario HistoricalFullFormat` uses the unchanged historical product and the
preflighted exact-guard collector through the nondefault category, margin and
deliberate stale-revision request. Initial palette writes retain their original
null Query. Positive compensation entries bind the category observed before each
write, so the later category selection cannot redirect recovery. Independent
closed readbacks and terminal recovery proof are stronger than the original
test; this is a causal diagnostic, not an unchanged historical test harness.
The deliberate stale request is classified separately from any unexpected guard
hit. The earlier prefix intentionally omits later stages; its passing result
does not identify the historical failure's position.
`-Scenario HistoricalCatalogueDrift` uses the same frozen guard collector for a
controlled Size-focus diagnostic. `-Scenario CatalogueDrift` compares the current
candidate. The request's property is first verified absent, so neither path can
reach a preference writer. An actor posts one owned dialog-local focus message
while observing the product's own tab navigation. The old candidate must refuse
the fresh revision after catalogue population; the prepared current candidate
must reach the missing-property refusal with a stable revision. Both require an
unchanged complete closed readback and original-handle normal exit. A captured
old guard must still be decoded and compared offline. This deliberately induced
metadata drift is not a natural reproduction or proof of the original failure;
it does not replace the complete Format matrix or close historical causality.
`tools/tests/Set-Q026Candidate.ps1` temporarily changes only the existing per-user
x64 add-in CodeBase and matching version entry, with an explicit fresh backup.
It refuses loaded VBE hosts or an unexpected assembly/original CodeBase. Restore
the original registration after all test-owned hosts have exited normally; the
script refuses to overwrite a registration changed by another task.
Apply and restore registration in the same user/token context as the limited
private-desktop worker. An elevated shell's successful readback does not prove
the worker sees that CodeBase. The campaign records the worker's actual value
and refuses a mismatch before launching Excel. Reserve the entire activation,
native execution and restoration interval: an empty host inventory alone is
insufficient if another qualification is about to launch. The isolated worker
also records script exceptions in `worker-error.json` without replaying cleanup.
The helper uses the installer's unescaped `file:///` CodeBase representation;
URI percent encoding is not substituted for the repository's COM registration
format. Check actual class activation before opening a host after a path-format
change; a matching registry string alone is insufficient loading evidence.
Execution also requires the existing VBAi `Addins64` LoadBehavior to be 3.
If necessary, explicit `Set-Q026Candidate.ps1 -EnableAutoLoad` temporarily
enables that installed add-in only from an existing DWORD 0, 2 or 3,
records the exact previous DWORD and restores
it with CodeBase after the owned hosts exit. It creates no registration and
changes no Office trust policy. A changed setting from another task prevents
restoration; it must not be overwritten silently.
The disconnected startup value 2 is accepted only by this explicit opt-in;
execution still requires 3 before activation. Microsoft's
[COM add-in flag documentation](https://learn.microsoft.com/en-us/office/vba/outlook/concepts/getting-started/customizing-outlook-using-com-add-ins)
describes startup and connection bits. This reference does not establish that
the current VBE installation successfully loads; its actual MVID must be verified.

For a retained Q-026 host, `tools/probes/Observe-Q026Host.ps1` defaults to an
independent PID/start/desktop window inventory, without bridge calls or input.
Its explicit `-CancelObservationPath` recovery targets only the captured owned
Options dialog and sends Cancel once on that desktop. It must not be used to
replay an uncertain cancellation. `Recover-Q026Preferences.ps1` compensates only
the captured positively committed entries with fresh revisions and exact
readbacks; any failed guard or unknown reply stops further dispatch. A complete
baseline hash and all tab structures must match before
`Close-Q026RecoveredHost.ps1` can close the sole saved disposable seed and Quit
once. Recovery receipts do not convert a failed native TRX into acceptance, and
their newly captured process handle does not prove the lost campaign handle's
shutdown lifecycle. These helpers do not recover arbitrary user documents.
`Recover-Q026Preferences.ps1 -RecoveryHostStartup` permits a separately owned
recovery host only when the original process is absent and its failed request
was a read with no terminal response. The exact frozen product must be loaded;
the complete snapshot may differ from the original baseline only in a positively
committed supported font. This path does not replay the lost read or transfer
an unknown mutation. The original failed qualification remains failed.
An empty commit ledger and null failed request can be released only for the
recorded pre-dispatch isolation stop, with its immediately preceding competitor
observation and a fresh complete baseline comparison. No compensation is sent.
Unknown null requests and incomplete states are refused.
If the same isolation stop occurs after the verified nonmutating guard warmup
but before the scenario baseline, `Resolve-Q026RecoveryBaseline.ps1` accepts
only the complete unchanged warmup readback for that exact PID/start/MVID.
Both warmup receipts and the immediate competitor observation are required;
intervening native dispatch, uncertain responses or positive commits refuse
this path. The recovery still requires a fresh full comparison after competing
hosts have exited. Offline evidence gates can be run without Office:

```powershell
powershell.exe -NoProfile -File tools/probes/tests/Test-Q026RecoveryBaseline.ps1
```

The existing `VBAi_VBE_INSPECTION_TRACE` opt-in also records bounded Options
combo metadata on the public read/write completion paths. `OptionsComboInspection`
distinguishes native ComboBox and UI Automation fallback, with a fixed Font,
Size, Palette or Other role. Native records retain handle/parent/control identity,
style, list counts before/after the existing single expansion, selection index
and observed dropdown cleanup. Missing observations remain null. No labels,
choice text, values, request content or exception messages are logged. These
records share the existing event/file limits and can be truncated; unavailable
logging cannot change the native result. They do not populate an empty catalogue,
relax revision checks or qualify font-size writes without native acceptance.

Optional `VBAi_TEST_Q026_CLR_TRACE_PLAN` enables test-only capture of the
historical revision guard. Prefer a separately verified plan with
`TraceMode: ExactGuardILBreakpoint` and a `GuardILOffset` derived from the frozen
binary's IL, not a guessed source line. Its own
`Test-Q026ClrTracePreflight.ps1 -GuardILOffset` receipt must verify the complete
synthetic graph, request, exact branch and normal detach/exit with matching
collector and product hashes. The Format fixture first compiles the native
guard through one intentional nonmutating refusal and verifies complete
unchanged state; only then can the exact breakpoint bind. A missing or ambiguous
breakpoint retains ownership and prevents the matrix. Ordinary CLR exceptions
are not stopped in this mode. The legacy first-chance mode remains available
for its separately verified plans, but changes timing much more substantially.
Neither mode establishes causality without a natural guard rejection and its
complete matching before/observed snapshots. Format dispatches independently
check owned PID/start and absence of competing VBE hosts throughout the run;
this supplements the reservation requirement rather than guaranteeing that
another chat cannot start a host between observations.

`Observe-Q026Host.ps1 -ReadDialogText` optionally reads a bounded inventory of
owned dialog controls without input. A terminal bootstrap failure before any
preference dispatch may use `Close-Q026RecoveredHost.ps1
-BootstrapClosureObservation <receipt>` instead of a preference-recovery
receipt. This separate path requires the recorded failed bootstrap, worker exit,
no Format phase receipts, independently observed dialog absence and the unchanged
saved disposable seed. It closes that seed and quits once; its receipt remains
unqualified and does not certify the failed campaign's original process handle.

The SOLIDWORKS workflow must not create or kill an application instance on the
user's behalf. A host fixture can use its own native save helper; that result does
not automatically qualify VBAi's `save_host_document` adapter.

For example, after confirming that Excel tests are safe to run:

```powershell
$env:VBAi_RUN_EXCEL_TESTS = '1'
try {
    dotnet test tests/VBAi.Tests/VBAi.Tests.csproj -c Debug --no-build -p:BuildOutputRoot="$PWD/artifacts/build" --filter "TestCategory=Excel"
} finally {
    Remove-Item Env:VBAi_RUN_EXCEL_TESTS -ErrorAction SilentlyContinue
}
```

### VBA test explorer native execution

The `VbaTestExcel` category uses the Excel opt-in above and creates its own
disposable macro-enabled workbook in a newly owned Excel process. Run this
category alone to qualify the generated VBA wrapper, assertions, fixtures,
results, stale-source refusal and the fixture-owned Excel coverage-copy path:

```powershell
$env:VBAi_RUN_EXCEL_TESTS = '1'
try {
    dotnet test tests/VBAi.Tests/VBAi.Tests.csproj -c Debug --no-build -p:BuildOutputRoot="$PWD/artifacts/build" --filter "TestCategory=VbaTestExcel"
} finally {
    Remove-Item Env:VBAi_RUN_EXCEL_TESTS -ErrorAction SilentlyContinue
}
```

The compiled service runs in the external STA test host with a fixture-owned
Excel adapter. Its synthetic production module contains two eligible procedures,
only one entered by the selected test bodies: the expected hand-calculated
procedure coverage is 1/2 = 50%. This is an assertion in the fixture, not a claim
that its current native run passed. It verifies original-source preservation and
mapped measurement when execution completes. It does not measure statements or
branches.

This external fixture does not qualify installed in-process callback activation,
native tool-window docking or other VBE hosts. The fixture records the assembly MVID
and Excel version, verifies normal owned-process exit and preserves report
examples under the test output's `test-explorer-native` directory. It does not
change Trust Center settings or execute an existing user macro.

For installed Excel qualification,
[Test-RegisteredVbaTestExplorer.ps1](../tools/testing-explorer/Test-RegisteredVbaTestExplorer.ps1)
checks the candidate assembly/MVID, registered add-in bytes and x64 callback
registration before creating a host. Its required arguments are
`CandidateAssemblyPath`, `ExpectedMvid` and `OutputDirectory`; use Windows
PowerShell 5.1 x64 in STA mode. Without `ExecuteOwnedFixture`, it only performs
preflight and reports that execution was not started. The explicit switch creates
an owned disposable Excel fixture. The script neither installs/registers the
candidate nor changes Office trust settings; deployment requires a separate
reviewed decision.

Use `-TestSubsystemOnly -ExecuteOwnedFixture` to qualify the VBA testing
subsystem: discovery/support installation, the registered native batch, human and
JSON reports, measured procedure coverage on a disposable copy, exact native
explorer actions, report consistency and stale-revision refusal. This option
skips the general Excel `Workbook_BeforeSave` positive save control and explicitly
records `NOT_TESTED_BY_SCOPE`. The synthetic event remains an eligible coverage
procedure; its counters and entry probe are retained as observations only, without
claiming that an independently qualified handler was suppressed. Original source,
probe mappings, coverage totals/percentage and exact `EnableEvents` restoration
remain required assertions. The fixture is saved, then the event-counter baseline
is read after the completed batch with owned PID/project/path/source checks.

The legacy default additionally requires the positive save control.
`-ExecutionAndUiOnly` skips that control **and measured coverage**; it cannot
qualify coverage. These two scope switches are mutually exclusive and are rejected
before any host activation if combined. None of these scenarios qualifies unrelated
VBAi features or the LLM permission/approval boundary; run the focused testing-tool
boundary regressions separately.

Shared callback code recognizes known VBE host names, but recognition is not
native qualification. The requested production scope includes Excel, Word,
PowerPoint, Access, Publisher, Outlook and a user-preloaded, explicitly selected
SOLIDWORKS instance; Visio and Project are excluded. Identify each loaded assembly,
actual callback transport, host version and tested operation. Do not infer that a
returned-value Excel fixture qualified the shared COM callback in another host.
Record failures and skipped scenarios in [validation](../docs/test-coverage.md),
not as successful host coverage.

Excel, Word and PowerPoint returned-value and coverage-copy adapters exist in the
source; that fact does not qualify those adapters or the shared callback natively.
The fixture and installed qualification script above remain Excel-specific.
PowerPoint qualification must independently verify its PIA argument-array
invocation, macro-enabled presentation copy and application-level event behavior.
Word qualification must verify saved DOCM/DOTM/DOC/DOT copying, refusal of unsaved
VBA source/reference mismatches, exact owned activation and document-qualified
positional invocation without retries after uncertain completion.

`VbaTestOffice` uses `VBAi_RUN_OFFICE_TESTS=1` for registered in-process Office
execution. Its Word/PowerPoint rows include a separate synthetic production module
with two eligible functions, only one called by the selected test: the expected
procedure measurement is 1/2 = 50%. This is a fixture assertion, not an observed
passing result. It checks canonical IDs/revision/probes, human/compact reports,
unchanged original source, file and counters, closed copies and retained artifacts.
The fixture may coexist with existing hosts only after identifying exactly one
new PID against its initial process inventory; it never reuses a user host.

`VbaTestExcelLarge` requires `VBAi_RUN_EXCEL_LARGE_TESTS=1`. It builds a fresh
catalogue spanning dispatcher leaf and route boundaries, verifies compilation of
the complete support module through a returned VBA verdict, and executes only
explicit boundary and failure selections. It uses the external STA Excel adapter;
it does not qualify the registered callback or explorer UI. Sources, references,
revisions and both reports are retained in its private temporary evidence folder.

`VbaTestOutlook` requires `VBAi_RUN_OUTLOOK_TESTS=1`, a usable existing profile,
no running Outlook process and an initially absent `VbaProject.OTM`. It creates
only an unsaved disposable inspector. It never sends or saves a mail item,
changes a profile, replaces an existing OTM or changes trust settings. An opaque
FileName on a new unpersisted project is selected by its unique project name;
the fixed per-user OTM path is used only for recovery. The fixture verifies its
initial blank source/reference baseline, removes only its unchanged owned modules,
and deletes a newly created OTM only after verified normal exit and retained backup.
Unknown native completion prevents both inspector closure and Quit. All failures
and retained recovery files are reported.

`VbaTestSolidWorks` requires `VBAi_RUN_SOLIDWORKS_TEST_EXPLORER=1`,
`VBAi_SOLIDWORKS_PID` and the absolute `VBAi_SOLIDWORKS_TEST_MACRO` path. The user
must preload that process and open the selected disposable `.swp` manually through
**Tools > Macro > Edit**. Save a blank macro containing only empty standard modules
or `Option Explicit`; remove the default generated macro body before qualification.
The fixture refuses other source, a different loaded candidate or an unsaved macro.
It never opens or closes a macro, activates application COM, saves, launches or
terminates SOLIDWORKS. Its own test/support modules are removed only after verified
settled completion and unchanged ownership hashes. The macro remains open; its
disk bytes remain unchanged, and the in-memory project may remain marked unsaved.
Other projects' sources are not inspected. This scenario covers ordinary tests,
reports, stale refusal and native window identity; it does not measure coverage.

Coverage qualification must verify Excel/Word/PowerPoint clone identity, unchanged
original live/disk source, original-source probes, explicit exclusions,
unsafe-syntax refusal and incomplete/unknown measurements. Verify restored Excel
event settings, Word AutoOpen/document/application/Normal-template effects and
PowerPoint open/before-close handler effects, including
cancelled closure and retained copies. These copies are not external-system
sandboxes. Retained document copies and coverage plans can contain full private source and data. Keep evidence
outside maintained documentation and follow [privacy](../docs/privacy.md).

## JavaScript and native renderer

The editor's JavaScript tests use Node's test runner:

```powershell
node --test tools/tests/Test-MonacoLanguage.mjs tests/VBAi.Tests/Infrastructure/Fixtures/Editor/MonacoEditing.Scenarios.mjs
```

See [native renderer testing](native/README.md) for the C++ self-test and loader.
These checks are distinct from real VBE rendering, host integration and measured
native-code coverage.

## Provider qualification

### Q028 Office assistant campaign

`tools/tests/Invoke-Q028Qualification.ps1 -Prepare -EvidenceRoot <fresh-absolute-root>
-ModelRoot <existing-absolute-model-store>` prepares the ordered Q028 matrix. It
freezes the product, test and private-desktop helper bytes, model manifest/blobs,
local Ollama executable, Office executable versions, selectors and acceptance
oracles before any provider request or native launch. Preparation never downloads
or substitutes a model. The current explicit profile is CPU, context 8192,
parallelism 1, temperature 0 and top-p 0.8; other profiles remain separate.

The worker freezes the exact Node, .NET and Git executable paths because scheduled
GUI tasks do not inherit a shell's temporary runtime path. Its managed gate selects
transport, chat, privacy, desktop ownership and Office lifecycle regressions;
it does not replace a previously failed full aggregate. Each bank uses a frozen
x64 runsettings file with a 15-minute session deadline. A deadline is a failed or
uncertain outcome, never a normal Office exit. When retaining product bytes while
rebuilding only the test harness, pass their original `-ProductSourceCommit`;
the plan records that identity separately from the harness revision.

Run the generated `Invoke-FrozenQ028.ps1` through
`tools/tests/Invoke-IsolatedTests.ps1` with the frozen helper and a fresh launcher
directory. It runs the managed gate, strict synthetic tool roundtrip,
cancellation/recovery and detached UI first. Only successful prerequisites allow
the six installed-assistant banks: Excel, Word, PowerPoint, Access, Publisher and
classic Outlook. A failed bank retains its result and leaves later banks NOT_RUN;
no native action is replayed. Existing Office processes are preserved.

The native banks require `VBAi_RUN_OLLAMA_OFFICE_TESTS=1`, the private desktop pair,
`VBAi_Q028_RESULTS`, and the existing host-specific opt-ins configured by the
worker. Each bank checks the installed MVID, native parent and exact project
scope, visible streaming while busy, Stop, interruption acknowledgement, next
reply, a native module marker and unchanged project source/references. No VBA
procedure or Outlook send is requested. Outlook requires an absent personal OTM
and an existing configured profile; it launches only the selected executable on
the inactive desktop and attaches through the ROT after sole-process/native-window
identity checks. It never falls back to COM activation on another desktop.

Excel's macro-free `/automation` bootstrap can leave the registered VBE add-in
disconnected. Q028 explicitly enables `VBAi_Q028_CONNECT_OWNED_ADDIN=1`: after
original process, STA, VBE window, private-desktop, ProgID/GUID and registered
candidate-byte checks, the shared single-use connection gate observes `Connect`
and permits one setter only for an exact false Boolean. It verifies true readback
and unchanged LoadBehavior. Failed or uncertain delivery retains the host; no
trust policy or registry repair is attempted. Other Excel banks keep their
existing startup behavior. `tools/probes/Inspect-Q028ExcelBootstrap.ps1` supports
read-only diagnosis and a separately guarded close of the unchanged saved seed
after an independent disconnected-add-in observation. Recovery never qualifies
the failed bank.

For installed-host requests, an explicit loopback observer captures only synthetic
JSON bodies and response bytes, without headers or credentials. Its capture is
bounded and preserves early-disposal outcomes. The final wire review must verify
that the marker was absent before the actual native read and appeared only in
its tool result; a visible marker alone is insufficient. The functional matrix
remains pending until that independent review and all lifecycle receipts pass.

The worker temporarily changes only the six documented non-secret provider
settings, and registers the exact candidate for future COM activation. It
restores those fields only if their current values still match its own applied
values; concurrent changes refuse restoration. Existing loaded add-ins are not
replaced. Unsettled assistant work retains its host and forbids teardown. Office
normal-exit receipts remain distinct from stopping the owned headless model
server. Historical Q028 failures and unknown historical causes are not erased
by a successful later profile.

`VBAi_RUN_OLLAMA_TESTS=1` enables `TestCategory=Ollama` against the loopback
server through the production HTTP client. `VBAi_TEST_OLLAMA_MODEL` selects an
already-installed model (default `qwen2.5:7b-instruct`). The shared selector applies
to the headless, detached UI and native Excel scenarios. An absent override uses
the default; outer whitespace is trimmed, while an empty explicit override,
internal whitespace or control characters are refused. It never loads or rewrites
personal provider configuration, downloads a model or silently substitutes another
model. These synthetic scenarios cover
streamed text, a harmless tool roundtrip, cancellation and a subsequent request;
they do not execute native VBE tools or read saved provider settings. Remove the
opt-in variables after the run.

The headless HTTP, detached chat UI and native Excel cases accept
`VBAi_TEST_OLLAMA_ENDPOINT` when an
owned local server uses another port. Its default remains
`http://127.0.0.1:11434/v1/chat/completions`. An override must be a canonical
`http://127.0.0.1:<port>/v1/chat/completions` URL without credentials, query or
fragment. Optional synthetic wire capture is restricted to that exact server's
chat and `/api/tags` routes and records the selected port. Retain the backend
version and model digest; another port does not prove the default port is usable.

The qualification backend profile is an already-installed `qwen2.5:7b-instruct`
model with `OLLAMA_CONTEXT_LENGTH=8192`, `OLLAMA_NUM_PARALLEL=1`,
`OLLAMA_NO_CLOUD=1` and `OLLAMA_NOPRUNE=1`, passed only to an owned local backend.
Attest the selected manifest/blob hashes and the backend's cloud-disabled startup
state before running the enabled scenarios. Do not change personal/global
environment. The shared test profile explicitly requests temperature `0` and
`top_p=0.8` through isolated `LlmSettings`, selected by
`VBAi_TEST_OLLAMA_TEMPERATURE` and `VBAi_TEST_OLLAMA_TOP_P` overrides when needed.
The overrides use finite invariant-culture numbers: temperature is in `[0,2]`,
and top-p is in `(0,1]`. The scenarios report the actual model, endpoint and
sampling before requests; they do not rewrite user settings or change prompts or
assertions. Production Ollama sampling properties remain nullable by default;
old settings therefore retain the backend's previous request behavior, and other
providers ignore these Ollama-only fields. This preparation is distinct from a passing
provider qualification: the original scalar, tool, UI and native readback
assertions remain required. The qualification requests do not add `logprobs` or
`top_logprobs`; controls using those diagnostic fields remain separate protocol
experiments and cannot replace the unmodified qualification request.

Passive synthetic body capture requires a separate explicit flag. For the two
headless scenarios, set `VBAi_OLLAMA_HEADLESS_CAPTURE_WIRE=1` and optionally
`VBAi_OLLAMA_HEADLESS_RESULTS` to an absolute local evidence directory (default:
`ollama-headless-diagnostics` under the test output). For the detached UI scenario,
the existing `VBAi_OLLAMA_UI_CAPTURE_WIRE=1` and `VBAi_OLLAMA_UI_RESULTS` variables
retain their behavior and `ollama-ui-diagnostics` default. The scenario opt-in is
still required; enabling capture alone does not run a model.

The shared `OllamaSyntheticWireCapture` helper requires the real production
`HttpClientHandler` with redirects disabled and the exact selected IPv4 loopback
port, POST chat or GET catalogue route. It observes buffered synthetic requests
and response bytes only as the production reader consumes them. Each body capture
is bounded to 1 MiB; read summaries distinguish observed EOF, early disposal,
in-flight reads, read errors and truncation. SSE `[DONE]` can stop the production
reader before physical EOF, so a false EOF flag alone is not a truncated-response
claim. Headers, credentials and personal history are not recorded. Diagnostic
write/serialization/wrapping failures do not replace the HTTP outcome.

The headless tool scenario records raw argument JSON, parsed argument type and
the marker's type/value before its existing scalar assertion. A nested object is
preserved as an object; it is never flattened or accepted as the expected string.
Evidence includes request intents, completion/error phases and the actually loaded
product MVID. No native tool is dispatched, and later evidence does not establish
the cause of an earlier response without its own captured wire.

For a detached helper-only regression batch, use the filter
`FullyQualifiedName~OllamaSyntheticWireCaptureTests|FullyQualifiedName~OllamaQualificationEndpointTests|FullyQualifiedName~OllamaQualificationModelTests|FullyQualifiedName~OllamaQualificationProfileTests`.
These tests use memory streams and explicit fake production-client responses;
they do not open sockets, run models or launch a native host. Remove the capture
and scenario opt-ins after real qualification.

`VBAi_RUN_OLLAMA_UI_TESTS=1` enables `TestCategory=OllamaUi` with the same
model selector. It shows the real chat controls and checks send, rendered
streaming, Stop and a subsequent completed response through the production
loopback HTTP client. Settings and history are isolated; the VBE project is
simulated and native tools are refused. This qualifies a detached chat workflow,
not Office or SOLIDWORKS integration. Run it separately from native host UI tests
to avoid competing for focus, then remove its opt-in variable.

`VBAi_RUN_OLLAMA_EXCEL_TESTS=1` enables `TestCategory=OllamaExcel` using the
same guarded loopback endpoint and already-installed model selectors. It creates a disposable Excel module containing a
random marker absent from the prompt, dispatches the model's `read_module` call
through the real project-bound VBE tools, and verifies the final answer and
unchanged source and normal exit of its owned Excel process. No macro runs.
This covers production provider/tools/session code dispatching through native
Excel COM from the test process, not the installed bridge or embedded assistant UI.
Run it separately from other native host tests
and remove the opt-in afterwards.

Existing-account checks require both `VBAi_CONNECTED_PROVIDER_TESTS=1` and
`VBAi_CONNECTED_SOURCE_TESTS=1`. The Git read check additionally requires
`VBAi_TEST_GITHUB_MANIFEST` identifying an explicitly authorized synthetic private
repository (URL, ID, verified ownership/private state and expected main commit).
The Codex check requires `VBAi_TEST_SOLIDWORKS_MANIFEST` with `OwnedDisposable`,
absolute `Path`, `FileSha256`, `Module`, `ModuleSha256`, `Marker`, `Pid` and `Mvid`.
It accepts only the manifested disposable SWP under the qualification artifacts,
uses an already-connected account and never signs in or copies authentication state.

Native UserForm GitHub qualification is a separate explicit scope:
`VBAi_RUN_USERFORM_GITHUB_TESTS=1`, `VBAi_RUN_EXCEL_TESTS=1`, the Git manifest and
an absolute `VBAi_TEST_USERFORM_GIT_OUTPUT`. The fixture is constrained to the
maintainer-authorized retained qualification repository and creates a dedicated
branch containing synthetic exports. It verifies controls, code and FRX bytes,
retains backups and captures the owned designer windows for visual review.
`VBAi_RUN_USERFORM_CORRUPTION_TESTS=1` separately enables the local-only malformed
FRX diagnostic; it must not publish corrupt content. Run either scenario only
while it owns the desktop. Normal host shutdown is part of acceptance; a verified
transfer does not excuse a subsequent crash. These fixtures are not enabled by
ordinary connected-account opt-ins.

### Installed GitWindow capture and local checkpoint

`NativeEmbeddedGitUi` is a separate, disabled-by-default scenario. It opens the
installed add-in through exactly one existing CommandBarButton with tag
`VBAi.GitHub`. A dedicated owner STA retains the owned Excel fixture and the modal
COM `Execute` call. An MTA worker acknowledges readiness before that call, discovers
only the exact owned window and native leaf controls, and requires the actual
UIA Value, Invoke, SelectionItem and Window patterns before each action. It uses
no coordinates, keyboard input, new product hook or external-STA Git capture.

Enable all three explicit flags: `VBAi_RUN_EMBEDDED_GIT_UI_TESTS=1`,
`VBAi_RUN_USERFORM_GITHUB_TESTS=1` and `VBAi_RUN_EXCEL_TESTS=1`. Set absolute
`VBAi_EXCEL_RESULTS` and `VBAi_TEST_GITHUB_MANIFEST` paths. Set
`VBAi_TEST_EMBEDDED_GIT_MVID` and `VBAi_TEST_EMBEDDED_GIT_SHA256` to the exact frozen
product referenced by the tests and installed in Excel. Existing Excel processes
refuse launch. The explicit `/x /automation` disposable seed bootstrap must attest
the same PID, executable, start time and loaded candidate before preparation.
Run only with the desktop available and an already-connected authorized account;
the test neither signs in nor chooses a fallback repository.

The manifest is constrained to the retained synthetic repository ID `1396566119`
at `https://github.com/blackcancer/vbai-qualification-20260929203712-7267b1e6`.
It additionally requires `embeddedBranch`, `embeddedBranchCommit` and the actual
observed `checkpointTabName`. The retained branch
`qualification-userform-20260929225354-93ed53dc` is accepted only at
`f5fb1a004dcb287c4c820b4bf308c673ce3b64e6`. Alternatively, the maintainer can
explicitly authorize an existing `qualification-embedded-ui-...` synthetic branch
and exact forty-character lowercase commit. No branch is created by this test;
`main`, `qualification-change` and arbitrary branches are refused. Preserve the
original manifest and make a separately reviewed manifest for this scope.

```powershell
dotnet test tests/VBAi.Tests/VBAi.Tests.csproj -c Debug --no-build -p:BuildProjectReferences=false -p:BuildOutputRoot="$PWD/artifacts/build-frozen-embedded-git" --filter 'FullyQualifiedName=VBAi.Tests.Integration.EmbeddedGitWindowTests.InstalledOwnerGitWindowCapturesSyntheticProjectAndCreatesLocalCheckpoint' --logger 'trx;LogFileName=embedded-git-ui.trx' --results-directory "$PWD/artifacts/embedded-git-ui"
```

Before opening the menu, independent `export_component` requests run through the
installed owner-VBE bridge with current component versions. `inspect_code_file`
attests the host ANSI code page; strict decoding produces the comparison snapshot.
Original native exports, resource bytes, raw hashes and parsed FRM resource offsets
remain durable. The local checkpoint must have the exact unique label/ref and
match every baseline manifest/source file and logical FRX resource using the
product's comparison rules. Those rules do not rewrite raw resources. Native
readbacks also preserve project COM identity, active project/module, mode,
protection, selection, references, source and form/control metadata. Idle UI alone
cannot prove success. Linking an existing remote revision does not prove equality
between the remote snapshot and the independent local baseline.

Every action has a durable intent and is emitted once. Known terminal errors remain
failures and permit a single normal modal close when no action is pending. Unknown
ownership, a deadline, delivery failure or unclassified status retains the original
workers/RCWs/owned host without another native action. Primary and cleanup failures
remain separate. Successful acceptance requires observed normal Excel exit; a
local checkpoint is not complete GitHub, import or recovery qualification.

The remaining embedded UI matrix is deliberately unexecuted:

| Operation | Required missing UI evidence and terminal oracle |
| --- | --- |
| Commit selection | The detached probe exposed `commitMessage` as Document/Text, without ValuePattern. A supported accessible editing route and checked-item selection/toggle patterns must be observed before any emission; do not invent a Value setter. Prove the exact selected source/ref and local commit contents. |
| Push/fetch | Observe interactive button patterns, unique delivery and terminal status in the owned host, then verify only the authorized branch's remote commit/fetched ref and unchanged native source. Idle/enabled buttons do not prove transfer. |
| Pull/import | Observe the actual import preview/report controls, exact selected revision, preserved backup refs and native mode/protection/revision guards; verify full designer/source state and save/reopen separately. |
| Checkpoint restore | Observe the actual list SelectionItem pattern and exact checkpoint identity, then the guarded import preview and full native/backup readbacks. Creating a checkpoint does not qualify restoration. |
| Recovery | Use separately authorized synthetic retained markers/backups and verify refusal plus explicit measured recovery. Do not manufacture a production failure or replay an uncertain mutation. |

The targeted detached discovery artifact proves available patterns only. This
prepared scenario and its pure helper checks are not evidence that the native
embedded workflow has run, and do not close the full Q-024/Q-027 release gates.

Native **local Git** UserForm layout qualification uses
`VBAi_RUN_USERFORM_LOCAL_GIT_TESTS=1` and `VBAi_RUN_EXCEL_TESTS=1`, with optional
absolute `VBAi_TEST_USERFORM_LOCAL_GIT_OUTPUT` for durable evidence. Filter on
`FullyQualifiedName~NativeUserFormLocalGitTests`. Each layout owns a separate
disposable Excel process and tests unchanged captures, a native property change,
exact local Git transport, checkpoint import, measured explicit rollback and
helper save/reopen with macros disabled. It covers Label/Button, TextBox,
ComboBox, ListBox, CheckBox, OptionButton, ToggleButton, ScrollBar, SpinButton,
TabStrip, an Image with a picture, and nested Frame/MultiPage. Unsupported
comparison grammars are retained unchanged and may fail these scenarios; no
normalizer exception or coverage claim is added. No remote Git operation occurs.
Run only while owning the desktop; retain failure artifacts and review the
source/reopened designer captures separately. Normal host exit is required.

The separate `ExplicitOwnedLayoutCaptureImportRecoveryAndReopenPreserveNativeState`
matrix adds `VBAi_RUN_USERFORM_EXPLICIT_BOOTSTRAP=1` and requires an absolute
`VBAi_EXCEL_RESULTS` directory. Filter on that exact method name to run its complete
prepared layout set without the COM-activation cases. It reuses the existing
`/x /automation` seed/bootstrap with exact PID, image, start-time and loaded-MVID
checks; any existing Excel process refuses launch. No path-visibility/token
manifest is passed. All native layout, capture, local Git, recovery and reopen
assertions remain the same. Launch context and final shutdown observations are
retained separately. A pass in this context does not explain a failed COM launch
or qualify remote GitHub transfer. Timeout or I/O uncertainty preserves the host
without a cleanup mutation or retry.

The Image layout reads its actual content digest through the production
`form_tree` descriptor inside Excel. It does not marshal a process-local OLE
picture to the test process. That readback records the descriptor type and PNG
content SHA-256, rather than an external picture handle or HIMETRIC dimensions.

To diagnose native UserForm export failures separately, enable
`VBAi_RUN_USERFORM_EXPORT_PROBES=1` and `VBAi_RUN_EXCEL_TESTS=1`, then filter on
`FullyQualifiedName~NativeUserFormExportProbeTests`. The six disposable trials
compare external STA and production host-bridge dispatch to fixture temporary,
production `GitTemporary` and evidence directories. Each trial uses a new
directory and the same form filename, attempts at most one export, records exact
PID, loaded assembly, document/component identity and raw files, and requires
normal owned-host exit. Failed exports are not retried or redirected. Partial
files remain available for diagnosis. Use the same optional absolute
`VBAi_TEST_USERFORM_LOCAL_GIT_OUTPUT` to retain reports. These probes do not
qualify Git capture comparison, imports or recovery.

The separate `ControlledSiblingEfsExportPreservesIdentityAndRetainsRawEvidence`
method runs two synthetic host-bridge trials with the same opt-ins. Each creates
a fresh unencrypted parent under the evidence root and two ASCII-named sibling
directories, encrypts only the empty `efs001` sibling, verifies both attributes,
then exports once to the selected sibling. Setup failures remain failures before
export. It records parent/temp attributes and retains both siblings; it never
decrypts a directory or modifies the real VBAi profile. This isolates EFS from
the production cache path without relocating any user's macro data.

`ControlledVolumeAndAncestorEfsExportRetainsOwnedIdentity` adds four synthetic
host-bridge trials: C/E volumes crossed with encryption on a fresh parent before
creating its children, or only on the empty export leaf. C trials use a unique
owned directory under the existing C TEMP root; E trials require an explicit
evidence root on E. The new parent must initially be unencrypted, and inherited
versus directly applied EFS is verified before the single export. All siblings,
raw files and reports remain available. This changes no existing directory's
encryption, decrypts nothing and does not qualify the actual GitTemporary cache.

`InheritedStorageAncestorExportRetainsExactOwnedIdentity` runs four host-bridge
trials in fresh children of the existing LocalAppData, LocalAppData/VBAi,
GitTemporary and TEMP directories. It changes no parent's ACL or encryption.
Each child records inherited attributes and ACL, verifies a new synthetic text
file can be written/read, and queries public EFS certificate hashes while keeping
metadata access denial explicit. Each process then attempts exactly one native
form export. Reports, raw and partial files are retained; no private key is read.
The same four trials record the testhost and exact owned Excel process tokens:
public user SID, session, integrity SID, elevation type, restricted-token flag and
AppContainer flag. These observations use query-only handles, change no token or
permission, and retain access failures as unverified metadata.

Designer screenshots reobserve the COM window after bounded UI settlement. A
zero designer HWND permits capture of the owned VBE root only after verifying
the exact active project and designer COM identities, captions, type, visibility
and root PID. Each screenshot has a `.png.json` sidecar recording observations,
capture scope and failures. A root capture covers the entire VBE with the exact
designer active and must be reviewed with that scope; it is not a form-only crop.

Local provider tests use simulated HTTP or CLI transports. Do not use personal
credentials, paid API calls or private project data without explicit permission.
A successful model catalog lookup alone does not qualify streamed responses,
tool execution, cancellation or recovery.

For a live qualification, use synthetic content and a harmless tool first. Record
the provider, model, CLI/protocol version and precisely what was exercised. Keep
secrets out of fixtures, logs, screenshots and reports. See the
[provider guide](../docs/providers.md).

## Coverage and evidence

A coverage run can use the installed collector:

```powershell
dotnet test tests/VBAi.Tests/VBAi.Tests.csproj -c Debug --no-build -p:BuildOutputRoot="$PWD/artifacts/build" --collect:"XPlat Code Coverage" --results-directory artifacts/coverage
```

Inspect the resulting report's assembly scope and filters. State numerator,
denominator, tested commit, exclusions and skipped host tests. Coverlet results
for the managed add-in do not measure C++, JavaScript or every native COM path.
Do not exclude production code or weaken assertions to manufacture a target.

### Managed VBA testing coverage gate

The managed testing subsystem has a strict coverage gate in
[check_managed_coverage.py](../tools/testing-explorer/check_managed_coverage.py).
Its scope is every `src/VBAi/Testing/*.cs` file, including Designers and native
adapters, plus `src/VBAi/Llm/Chat/LlmVbeTools.Testing.cs`. The collector still
instruments the complete assembly; the gate selects this subsystem from the
result without excluding production code.

Use a fresh isolated build directory for each collection. Run the focused test
classes with both collector formats, leaving exclusion settings unset:

```powershell
$managedCoverageBuild = "$PWD/artifacts/build-vba-testing-managed-coverage"
$managedTestFilter = @(
    "FullyQualifiedName~VBAi.Tests.Unit.VbaTest",
    "FullyQualifiedName~VBAi.Tests.Unit.VbaCoverage",
    "FullyQualifiedName~VBAi.Tests.Unit.VbaNativeTest",
    "FullyQualifiedName~VBAi.Tests.Unit.VbeTestExplorerService",
    "FullyQualifiedName~VBAi.Tests.Unit.TestExplorerWindow",
    "FullyQualifiedName~VBAi.Tests.Unit.TestSupportReviewDialog",
    "FullyQualifiedName~VBAi.Tests.Unit.LlmVbaTestingBoundary"
) -join "|"
dotnet build tests/VBAi.Tests/VBAi.Tests.csproj -c Debug -p:BuildOutputRoot=$managedCoverageBuild
dotnet test tests/VBAi.Tests/VBAi.Tests.csproj -c Debug --no-build -p:BuildOutputRoot=$managedCoverageBuild --filter $managedTestFilter --collect:"XPlat Code Coverage" --results-directory artifacts/coverage/vba-testing-managed -- DataCollectionRunSettings.DataCollectors.DataCollector.Configuration.Format=cobertura,json
```

Use `coverage.cobertura.xml` and `coverage.json` from the **same collector UUID
folder**. Replace `<collection-id>` below with the identifier reported by the
test command:

```powershell
python tools/testing-explorer/check_managed_coverage.py artifacts/coverage/vba-testing-managed/<collection-id>/coverage.cobertura.xml --json artifacts/coverage/vba-testing-managed/gate.json
python -m unittest discover -s tools/testing-explorer -p test_check_managed_coverage.py
```

The first command prints a readable per-file summary, saves the complete JSON
result, and exits nonzero for any uncovered line or branch, missing expected
source file, malformed report or disagreement between formats. The sole
sequence-point exception is `VbaTestExplorerService.cs`, and only while its
contents remain interface declarations without executable bodies. New source
files automatically enter the gate's expected scope.

Line totals retain each method's collector records, including generated classes
and methods that share a physical source line. Branch totals retain every raw IL
branch record, including repeated offsets and branches without an emitted source
line. Cobertura line condition counters alone can omit such branches; the
[Coverlet reporter implementation](https://github.com/coverlet-coverage/coverlet/blob/v6.0.4/src/coverlet.core/Reporters/CoberturaReporter.cs)
explains why its companion JSON is required. The gate reconciles root counters,
class and method rates, source-line records and visible condition counters across
both formats before accepting the result.

Reports do not encode the complete exclusion configuration or establish the
compiled source revision. Preserve the collector command/settings, tested source
revision and assembly identity with the validation evidence. This gate measures
C# implementation coverage; it does not qualify native applications or measure
line/branch coverage within a user's VBA macros.

Keep machine-local artifacts outside the maintained guide tree. Publish a concise,
versioned summary in [recorded validation](../docs/test-coverage.md), separating
unit/runtime tests, native host observations, Designer checks and live-provider
runs. An old 100% result does not describe a later build.

Publisher serialized-publication qualification uses the explicit
`StartPublisherSerializedQualificationSeed` fixture entry and
`VBAi_TEST_PUBLISHER_SERIALIZED_SEED` opt-in. This accepts only the frozen,
owned qualification file with its original normal-exit and closed-file receipts;
it does not accept arbitrary macro files. The fixture copies once and opens once
with macros disabled on the guarded private desktop. Before source or metadata
writes, it audits the exact inert document/module/class sources, reference
manifest and native General baseline. Dedicated `PublisherSerialized*` adapter
tests qualify existing publications; the original NewDocument tests remain
separate. These results do not qualify empty-publication VBA initialization or
first SaveAs, and no VBA procedures are executed.

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

`VBAi_RUN_MAIN_DESKTOP_TESTS=1` explicitly selects the real `WinSta0\Default`
desktop for the native Format diagnostics and the owned Office fixtures. Leave
`VBAi_TEST_DESKTOP_NAME` and `VBAi_QUALIFICATION_DESKTOP` empty. The fixtures check
the current and input desktop before native work; the serialized Publisher seed
keeps its original-process, COM identity, macro safety and shutdown checks.
The existing Word-specific main-desktop opt-in remains supported. Execute the
prepared batch through a same-user Limited STA coordinator.

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

Capture, apply and restore COM snapshots inside the campaign's intended Windows
actor. Matching user SIDs do not prove that a tool terminal and a scheduled
limited worker observe the same registry trees. Compare complete scoped trees
and fingerprints in the actual actor before attributing a mismatch to a
concurrent writer. An exact guarded restoration in that actor may recover the
registration; it does not change the original host or test verdict.

COM/UI work stays on its owning STA and the observer on its original thread.
Freeze exact process, native thread, HWND/parent/owner and control identities
before each action; a replaced or foreign identity refuses the action.
The settled Word collection diagnostic is an explicit fixture option, not native
RCW-release proof. Original unchanged exit deadlines still apply.
The separate `VBAi_TEST_WORD_POST_QUIT_GC=1` diagnostic collects once in the
testhost after the sole Quit has returned and the fixture's external document
and application references have been released. It does not wait for finalizers,
dispatch COM, or collect inside Word. Its elapsed time is included in the same
original exit-observation deadline. Capture the opt-in before shutdown and keep
the lifecycle receipt; collection alone never proves reference release or normal
process exit. Unset the variable for the original cleanup sequence.

After uncertain native dispatch, do not repeat Save, import, Reset, Send, Stop or
Quit to obtain a passing result. Retain owned processes/references, pending
requests, recovery files and failure receipts for a separate diagnosis. Forced
termination is cleanup evidence only and is never normal-exit acceptance.

## Saved SOLIDWORKS Test Explorer fixture

Use an owned disposable macro and an explicitly selected SOLIDWORKS PID and path
(`VBAi_SOLIDWORKS_PID` and `VBAi_SOLIDWORKS_TEST_MACRO`). The preparation and
execution tests have separate opt-ins so a consumer cannot install source as a
side effect of running a saved fixture:

1. With a blank saved macro open, enable
   `VBAi_PREPARE_SOLIDWORKS_TEST_EXPLORER=1` and select only
   `PrepareSavedTestExplorerFixture`. This deliberately adds the exact test and
   support modules, warms their batch/single dispatcher, and explicitly saves
   through the guarded host adapter. Preserve its source and lifecycle receipts.
2. Close the original owned host normally. Seal the resulting closed file hash,
   year, candidate identity, sources and references. Creating this fixture does
   not qualify preservation of an existing file's bytes.
3. Open a copy of that closed fixture in a fresh selected process. Enable
   `VBAi_RUN_SOLIDWORKS_SAVED_TEST_EXPLORER=1` and select only
   `PreloadedSavedMacroRunsVerifiedReportsWithoutSourceMutation`. Admission
   rejects any other code or altered generated support. The consumer runs batch,
   single and the real explorer without source edits, installation, removal or
   saving. Check the original process's normal exit and unchanged closed bytes
   separately; an in-process test pass alone does not satisfy this oracle.

The existing blank-project edit/restoration scenario remains separate, including
its dirty-project persistence failures. Do not apply `VBProjects.Remove` to these
native host projects or set their `Saved` flag to manufacture a clean baseline.

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

### Excel audit completion

`ExcelAuditCompletionTests` requires both the basic Excel opt-in and
`VBAi_RUN_EXCEL_AUDIT_COMPLETION=1`. Its prepared bank runs a newly generated
procedure once from an inactive target pane, then checks its cell marker and
unchanged source. Separate workbooks contain a generated `Workbook_BeforeSave`
handler that counts and accepts or cancels the single product Save. Cancellation
must preserve the exact disk baseline; acceptance must persist the code and
marker through a fresh, read-only process. Reopening disables macros and events.
These scenarios execute synthetic fixture VBA; they never execute a supplied
user macro or change host trust settings.

Combine the selected persistence, options, signature and embedded Monaco cases
into the prepared bank before dispatch. The Monaco receipt records elapsed
bridge/native command time as an observation, not a renderer frame-time threshold.
A signature case additionally requires its owned certificate manifest and
`VBAi_RUN_EXCEL_SIGNATURE_TESTS=1`. An untrusted synthetic certificate tests
signature persistence without granting it trust or signing the installer.
For a first signature, arrange certificate selection with the operator before
dispatch. VBAi validates the requested thumbprint and opens the Windows Security
picker, but that protected picker requires the operator to select and confirm
the certificate. A prepared manifest does not prove that selection completed.
Leave this case explicitly not run in an unattended bank; do not repeatedly
launch it, weaken trust settings or infer signing from an unsigned-file check.

The current complete Format matrix prepares font, size, every editable palette
category and the margin indicator together. `set_vbe_option` accepts a bounded
`OptionChanges` array with one `ExpectedOptionsVersion`; it validates the complete
batch before preference writes and requests OK once after final readback. A
separate complete inspection verifies the committed values. Restoration uses
one guarded batch and then checks the entire original revision. Read-only
inspections and the stale-revision refusal still open Options independently;
they do not apply a preference per window. Historical partial diagnostic
sequences remain explicit options and cannot be combined with grouped mode.

### Word and PowerPoint coverage-copy diagnostics

`VBAi_RUN_OFFICE_COVERAGE_EVENTS=1` attaches lifecycle observers to the exact owned
Word or PowerPoint process during the synthetic copy run. Event delivery and
processing threads are recorded separately. Borrowed document arguments are read
synchronously on the fixture's original STA while its transport waits pump
messages; callbacks do not change cancellation, save files or release arguments.
The oracle requires one opening and one closing event for the same copy. Missing
events remain failures, independently of the coverage result.

The separate `VBAi_RUN_OFFICE_COPY_CONSENT=1` opt-in permits one addressed action
on the observed French macro-security warning for that reviewed synthetic copy.
It requires the real main desktop, the original process handle and birth time,
an unchanged saved source, a unique new coverage folder, stable native/UIA
identities and exactly one eligible button. It never enables macros globally or
changes Trust Center settings. Other warning layouts are refused. Intent and
return receipts are durable; uncertain invocation is not retried.

Keep the complete negative consent matrix in the managed gate before dispatching
the native bank. The original baseline and verified preservation receipts are
stored separately from event acceptance: a later event failure cannot erase
partial evidence or turn it into a successful qualification.

`VBAi_RUN_WORD_LIFECYCLE_CALIBRATION=1` adds four diagnostic controls before the
Word copy scenario: visible and hidden macro-free documents, each with automatic
macro suppression enabled or disabled. The owned Word process starts with `/a`;
the source and controls use a generated, macro-free DOTX. Disabling suppression
requires an exact reviewed template inventory and document relationship. A
foreign template refuses the control before any policy change. In the exact
owned `/a` generation, Word's own Normal object is admitted only after a separate
metadata review proves design mode, no protection, and exactly one document
class with zero code lines. This reads no code text or template file; an arbitrary
template, extra component or any code is refused. The inventory is checked again
after the suppression flag changes and before opening. Every control
opens read-only under `AutomationSecurity = ForceDisable`, restores suppression
and the previous application policy, and verifies unchanged source and control
bytes. The original STA pumps event reads while a separate STA owns the control's
COM references. Durable intent, terminal and uncertain-outcome receipts prevent
replay. These controls diagnose missing events; they do not replace the copy
scenario's opening/closing oracle or authorize personal-template code/file reads.
Generated template and document structures are parsed from bounded immutable
snapshots, read with sharing compatible with Word and checked against their
frozen SHA-256 before inspection. Partial metadata reviews are recorded even if
a later package check refuses admission.

Microsoft documents `/a` as excluding the existing Normal template from startup
in its [Word troubleshooting guide](https://learn.microsoft.com/en-us/office/troubleshoot/word/damaged-documents-in-word).
The Normal object's reported path is not by itself proof of loaded personal
content; the fixture still independently refuses any code or foreign identity.

`VBAi_RUN_WORD_COVERAGE_EVENT_SCOPE=1` uses the same owned `/a` template review for
the actual synthetic Word coverage copy. It requires confirmed startup suppression
and ForceDisable before changing the temporary application flag, then uses the
existing reviewed ByUI/copy-consent scenario. It restores suppression once the
copy settles, before restoring the original application security policy. Every
flag write has its own intent/return receipt; failed restoration or uncertain
copy execution retains the original lease and prevents replay/automatic cleanup.
The option also creates the source from the generated macro-free template. It
does not change Trust Center settings or the exact copy Open/BeforeClose oracle.

### Word assistant observer lifetime

With the explicit real-desktop Ollama opt-in, the Word assistant fixture runs
its UI Automation observer in a separate owned helper process. The helper
validates the frozen test/product assemblies, user, original parent and Word
process identities, profile and desktop before observing the genuine chat.
It exposes only the fixed Word scenario; it cannot activate or close a host.

The fixture requires the observer's original creation-handle exit and a matching
settled terminal receipt before closing Word. Streaming, cancellation, the next
reply and the native-module marker remain separate acceptance oracles. A missing,
partial or contradictory receipt retains the host without a second Send, Stop
or Quit. Observer exit does not establish Word shutdown: the existing fixture
must still verify Word's own original-handle normal exit independently.
The observer pins its original parent process generation and refuses new UI
actions after that parent exits, including a button delivery after a slow
preflight. Synchronous observations returning beyond their configured deadline
are rejected; a late successful property read cannot silently extend the window.

Each control inspection fetches identity, caption and state together in a fresh,
element-scoped UI Automation snapshot. Bound control peers keep their original
HWND; refreshing properties never admits a replacement control. Ownership is
checked before and after a snapshot read and immediately before button delivery.
Snapshot batching reduces cross-process calls; its native performance benefit
must be measured and does not itself establish a responsive hosted chat.

Fixed controls are discovered among owned native WinForms HWNDs, with UIA identity
checks on each matching peer. A cached handle is never replaced during a turn.
Drop-down selection readback uses size-bounded WM_GETTEXT and an independently
read CB_GETCURSEL index, with native role/ownership checks and stable index/length
checks around the read. It rejects editable controls, truncated labels and
concurrent selection changes without a write or read retry. UIA expansion and
selection remain one-shot actions on the real controls. Composer discovery is
limited to the fixed native prompt host so earlier transcript content is not
traversed on each Send.

Transcript RichEdit children are transient read-only observation targets, rather
than fixed action targets. Missing or detached rows contribute no oracle text;
subsequent reads remain bounded by the current observation deadline. A changed
fixed panel or foreign process/thread is rejected. This distinction never allows
a second Send, Stop, selector mutation or Quit after an uncertain action.

### Opt-in Word add-in cleanup observation

For an explicitly launched Main Word fixture, `VBAi_TEST_ADDIN_SHUTDOWN_ROOT`
can enable the existing bounded add-in cleanup trace. Use a fresh GUID directory
under the current user's temporary `VBAi.AddInShutdownDiagnostic` directory and
pin the candidate with `VBAi_TEST_ADDIN_SHUTDOWN_PRODUCT`,
`VBAi_TEST_EMBEDDED_GIT_MVID` and `VBAi_TEST_EMBEDDED_GIT_SHA256`.

The fixture publishes a single request from its verified ready-state inventory,
using the original process generation, candidate identity and VBE owner thread.
It acquires no additional COM reference. After the original Close/Quit attempt
and exit wait, it reads the cleanup receipt once and preserves any independent
shutdown failure. A missing or invalid trace fails observation; a returned cleanup
trace never establishes process exit or complete native reference release. The
default five-second Word exit deadline is unchanged. The diagnostic admits exact
Excel, Word and SOLIDWORKS process names, with the same process-generation,
loaded-candidate and owner-thread requirements. The prepared SOLIDWORKS worker
can publish the canonical request after verifying its VBE inventory. Admission
tests do not qualify native SOLIDWORKS cleanup; that observation is measured
separately. Enabling the diagnostic does not activate an add-in or relax host trust.

While an admitted chat shutdown is active, the trace can also record
`ExceptionObserved` on its exact owning thread, including framework callbacks
made during handle destruction. This phase records a thrown exception before
its handler is selected; it does not assert a failed cleanup. Exceptions from
the trace's own formatting and publication are excluded. The subscription is
removed when chat shutdown returns or propagates an error.

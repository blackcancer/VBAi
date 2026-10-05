# VBE compatibility and host adapters

## Product scope

VBAi targets the **Visual Basic Editor embedded in Windows applications**. It is
not an Excel-only or SOLIDWORKS-only add-in. These applications are examples of
VBE hosts and test environments, not the product's architectural boundary.

The current build targets **x64** and **.NET Framework 4.8**. The application must
expose a compatible VBE/COM add-in environment. Monaco uses WebView2; chat storage
uses the Windows SQLite runtime. This build does not claim x86, macOS, the Visual
Basic 6 IDE or Visual Studio's VB.NET editor as supported targets.

## Three different compatibility questions

| Layer | Responsibility | What must be established |
| --- | --- | --- |
| Shared VBE/VBIDE integration | Projects, modules, references, forms, code panes, native editor operations and the assistant workspace. | The host exposes the required interfaces and accepts the add-in. |
| Host-specific compatibility | Document identity, persistence, execution and application-specific behavior not provided by VBIDE alone. | An adapter exists for the requested operation and handles this host's behavior. |
| Qualification | Observed results in a specific application/runtime/environment. | A reproducible test passed on the identified build. |

A missing save adapter does not negate working shared editor capabilities. Equally,
a host that loads the add-in is not thereby qualified for every mutation, debugger
operation, language feature or UI surface.

## Recorded host evidence

### Current operation-specific refresh

Q-012 is CLOSED for the identified Access/Publisher **existing-document adapter
contract**. Active-module and module/class/UserForm-state Save, reference addition
by GUID/file, reference removal and Description are verified through the loaded
candidate, one adapter Save, exact independent fresh-process readback and normal
exits. Explicit native General HelpContextID and HelpFile text representable by
the owning ANSI control have the same persistence acceptance. Candidate identities,
test scopes and lifecycle evidence are in
[recorded validation](test-coverage.md#q-012-completed-adapter-contract-2026-10-04).

The Access French multi-object confirmation retains approved object/source,
selection, canonical project, mode, revision, runtime authorization and deadline
checks before one queued confirmation. It does not compile implicitly, accept
unknown prompts/locales or fabricate a document Saved flag. Publisher opens an
audited owned serialized publication with process-local macros disabled; exact
retained/active/sole document and native-window identities precede adapter work.
This does not qualify empty-publication VBA initialization, NewDocument or first
SaveAs. Stored UserForm state does not establish rendering or event execution.

For real Access/Publisher Type100 projects, legacy `set_project_property`
HelpFile/HelpContextID writes are unsupported and refused before setter entry.
The catalogue's additive `SetterStatus=HostLegacyWriteUnsupported` preserves raw
Value/Type/ReadOnly/digest/error fields. Refresh the opaque project version and
explicitly use `read_project_general` then approved asynchronous
`set_project_general`; no automatic fallback or Save follows. Other hosts,
project types and properties retain their existing routes.

Unsupported Unicode is accepted only as a known pre-write refusal with one
guarded Cancel, observed command return, unchanged state and exact saved bytes
after normal close. It is not Unicode persistence acceptance. Earlier malformed
HelpFile COM getters, failed scalar HRESULTs, partial live mutations, startup
failures, forced exits and interrupted-helper receipts remain recorded failures
or uncertainty on their original candidates. The explicit General contract does
not normalize or repair those legacy results; `open_project_help` is not qualified
by path-storage acceptance.

Final inactive-desktop shutdown and temporary COM/startup registration restoration
are verified. The launcher variation separates the helper from the parent console
without changing its binary, injecting input or suppressing control signals;
earlier interruption causes remain unattributed. Earlier managed full-suite
acceptance retains its own candidate, while the final gate is focused. No
whole-host, full release or code-coverage acceptance follows. Historical
[D2 matrix](test-coverage.md#q-012-access-confirmation-and-native-matrix-2026-10-03)
and [D1 diagnosis](test-coverage.md#q-012-access-save-confirmation-and-publisher-ownership-diagnostic-2026-10-03)
remain distinct from the current operation scope.

Latest gated source `8f5e16f` passes the installed Word owner-menu capture, exact
local checkpoint and compare, preserving source, full references and saved bytes
with normal original exit code 0 and a stated 15-second read-only exit bound.
Its actual chat trial identifies the unique panel, selects and reauthorizes the
saved canonical document and verifies the VBE-root modal owner. The bank's
top-window inventory now finds the unique Options popup and exact localized Git
item. Its owner differs from the VBE root, and the bank refuses before Git intent.
A native-parented default-Designer probe reproduces a separate hidden WinForms
menu owner, but does not qualify Word or relax the actual Git modal owner guard.
Normal exit and registry restoration are verified. Actual Chat-to-Git opening
remains unqualified pending a fresh owner-aware bank trial. See
[recorded validation](test-coverage.md#word-chat-popup-owner-refusal-and-owner-menu-control-2026-10-02).

The gated Q-024 source `70f6cc0` reaches the installed Word VBE Git menu and
production owner-thread capture, exact local checkpoint and compare with unchanged
source, references and document bytes. Default 5-second exit observation fails;
a fresh control with the existing 15-second read-only bound observes normal exit
code 0 after 5,558 ms. This is the accepted operation scope with its explicit bound,
not a repair or green verdict for the earlier external-STA/default-bound cases.
The separate chat trial fails before any UI action at discovery of `ChatWindow`
AutomationId and exits normally. Follow-up gated source `4b68a77` also stops before
UI action: its bank incorrectly uses the unused `EnumChildWindows` return as a
success status. Original normal exit code 0 and registry restoration are verified;
the callback-bound correction passes its full managed gate on `c283c6e`.
Isolated Designer preflight finds the real Git item has no AutomationId, so
native Word is NOT_RUN on that candidate. Corrected native-popup selection still
needs Word observation. Q-024 stays open for Chat-to-Git; Q-027 is deferred. See
[recorded validation](test-coverage.md#word-chat-menu-identity-preflight-2026-10-02).

The gated source `b6f13fb` observes `Application.VBE` succeeding before the
retained workbook replacement in owned Excel `16.0.20430.20092`, then failing
with `0x800A03EC` after exact copy-open verification. Events/macros are disabled
for opening; no security settings are relaxed. The lifecycle cause is unproven,
and no form capture, import or transfer occurs. Normal original exit and restored
registration are verified; Q-027 remains open. See
[recorded validation](test-coverage.md#retained-vbe-lifecycle-observation-and-managed-gate-2026-10-02).

The gated source `9ba099e` narrows the retained-copy Excel refusal to
`Application.VBE`, with `0x800A03EC` recorded on the fixture owner thread.
Excel `16.0.20430.20092` loads the expected candidate and verifies the copied
workbook, but no form capture, import or transfer is reached. The cause is still
unknown. Normal original process exit, unchanged file bytes and restored
registration are verified; Q-027 remains open. See
[recorded validation](test-coverage.md#named-vbe-getter-observation-after-pr-19-integration-2026-10-02).

The gated PR integration source `3ff51bd` now verifies a retained synthetic
workbook copy's Open/path/count/hash checks in Excel `16.0.20430.20092`.
Before capture, COM `0x800A03EC` interrupts a source line containing both
`Application.VBE` and `VBE.MainWindow`; the failing getter and cause are unknown.
The owned host exits normally, source/copy bytes stay unchanged and registration
is restored. The installed DLL is unchanged. No UserForm import or font transfer
is exercised; Q-027 stays open. See
[recorded validation](test-coverage.md#retained-userform-copy-vbe-access-refusal-2026-10-02).

The retained-workbook Q-027 diagnostic source `6d3d802` passes its complete
managed gate. A native Excel copy-open succeeds, then a fixture file-sharing
error stops preparation before form capture or import. Normal owned exit,
unchanged source/copy bytes and registry restoration are verified. The reader
correction has a later scoped native result above; that trial still stops before capture and Q-027 remains open.
See [recorded validation](test-coverage.md#retained-userform-copy-preparation-failure-2026-10-02).

The latest Q-027 diagnostic source `a5649c3` passes its complete managed gate.
Its explicit Arial 9.00 baseline and actual owner-dispatched LabelButton import
have exact initial and independent snapshots and matching native readback.
The diagnostic requires an inexact FRX for a transfer experiment and therefore
refuses this already exact state before any font put; the scenario remains
failed. Ordinary workflow, default Tahoma explicit-font fidelity, Frame 8.27,
recovery and persistence are not qualified. Normal owned exit, unchanged disk
and registry restoration are verified. See
[recorded validation](test-coverage.md#explicit-root-font-import-and-diagnostic-gate-refusal-2026-10-02).

The preceding Q-027 diagnostic source `269b187` passes its complete managed gate.
Its fresh Excel baseline lacks the root font binding required to arm the
post-capture transfer experiment. The trial stops before manifest publication,
Git menu emission or import; deferred transfer is NOT_RUN. The owned host exits
normally and registration is restored. The cause of the baseline variation,
exact import, recovery and post-import persistence remain unqualified. See
[recorded validation](test-coverage.md#post-capture-font-diagnostic-baseline-refusal-2026-10-02).

The preceding Q-027 observation source `be9dcd8` passes its complete managed gate.
Actual owner-UI LabelButton observations find the root font binding absent
before diagnostic font getters. A separate temporary Arial child-name write
reads back Arial while `Font.Object` and `Designer.Font` still serialize Tahoma.
The child collection route therefore does not update the attached descriptor
at that observed boundary. Both strict imports fail with only the FRX differing;
owned hosts exit normally, saved files stay unchanged and registration is
restored. Exact import, Frame font-size fidelity, recovery, transfer and
post-import reopen remain unqualified. See
[recorded validation](test-coverage.md#owner-thread-root-font-observation-and-distinct-child-name-2026-10-02).

The earlier Q-027 scalar-child restoration source `7745e1c` passes its complete
managed gate. The first actual owner-UI LabelButton import reaches final
comparison, which still refuses the missing persisted root StdFont descriptor.
Source files match, but the FRX differs. The original owned Excel exits normally,
saved disk bytes stay unchanged and temporary registration is restored. Scalar
setter readback, exact import, recovery and post-import reopen remain unqualified.
See [recorded validation](test-coverage.md#root-font-child-value-import-canary-2026-10-02).

The earlier Q-027 diagnostic source `e3d94bc` identifies the failing root font
operation as `VBIDE.Property.Object.set`. A separate read-only observation
qualifies its Font.Value child getters and declaration metadata with exact
state preservation and normal owned Excel exit. It does not qualify native
setters, import or recovery. The following scalar-child restoration change
requires its own matching managed and native gates. See
[recorded validation](test-coverage.md#root-font-operation-and-read-only-child-observation-2026-10-02).

The latest complete Q-027 matrix source `de5c619` uses the root component's Font
`Property.Object` route. Its complete default managed gate passes, but the
declared native checkpoint-import matrix fails with terminal property/method
errors. All owned Excel processes exit normally and temporary registration is
restored. The first case still loses its root StdFont descriptor after import;
the retained exception message does not locate the exact COM member. No native
setter, import or recovery acceptance is established. Installed files remain
unchanged. See
[recorded validation](test-coverage.md#root-userform-font-property-route-2026-10-02)
for candidate identity and exact scope.

The earlier Q-027 isolated source `420b3da` adds guarded owner-thread exact-font
restoration, bounded owner names and a native accessibility ListBox selection
cache correction. Managed regressions verify those paths. After correcting the
selection observation, that owned Excel trial reaches actual LabelButton
import and normal original exit. Source files match, but the root FRX StdFont
descriptor is omitted despite a decoded restoration binding and a nonthrowing
assignment. Exact snapshot comparison refuses success. The following TextBox
trial fails seed attachment before form construction; remaining layouts cannot
launch. Earlier UI failures remain separate observations. Strict FRX comparison
and the earlier Frame font-size difference remain unresolved; no complete
UserForm import acceptance is claimed.
The installed DLL has a separate read-only identity; see
[recorded validation](test-coverage.md#q-027-userform-follow-up-2026-10-02).

#### Earlier 2026-10-01 operation evidence

The isolated source `4a323196` / MVID
`25facee2-6b10-40fa-880c-cb9328333713` now has observed Excel core execution,
declared-scalar inspection, Format restoration, embedded Monaco and actual Git
capture/checkpoint/compare results. Word/PowerPoint adapter-only source/class/form
save and fresh-process reopen pass. Access/Publisher reference workflows and
selected metadata saves have passing scopes, while Access help metadata exit and
Publisher HelpFile readback remain failed. The retained Publisher file stores
the correct help path; fresh native readback is altered. A later metadata batch
also fails Access Description exit and Publisher HelpContextID mutation, leaving
its remaining scenarios unexecuted. A longer exit bound does not establish a
lifecycle correction. Initial UserForm property differences also occur
without Git import. Saved/reopened baseline properties remain exact in the
accepted scopes, while import still fails strict FRX comparison. Expanded
diagnostics identify lost root/Frame font descriptors and a Frame size change
from 8.27 to 8.25. The tested member-setter, owner-assignment and persisted-font
restoration paths do not establish a successful import;
signature remains unqualified. The selected Ollama CPU profile passes its
headless, shown detached-chat and real-Excel read-only scopes after commit-memory
availability improves; this does not qualify the embedded-host assistant.
SOLIDWORKS is not exercised in this earlier refresh. At that checkpoint the
installed DLL was unchanged and temporary registration selections were restored.

See the [current qualification checkpoint](release-qualification.md#pr22-integration-checkpoint-2026-10-04)
and [recorded validation](test-coverage.md#native-qualification-refresh-2026-10-01)
for the exact tested scopes and identities. Historical observations below retain
their original candidate boundaries.

### Earlier candidate observations

The observations below come from the **2026-09-29 through 2026-10-01** qualification artifacts for
Microsoft 365 **16.0.20326.20158 x64**. The `office-final`, `monaco-save`,
`monaco-privacy`, `ui-native-startup`, `ui-native-placement` and `outlook` runs used
the candidate with DLL SHA-256 beginning `482864942BAC`; `excel-final` used the
candidate beginning `F3C48EAEA590`. These are separate build observations, not a
blanket qualification of v1.0.0. Full provenance and test totals belong in
[recorded validation](test-coverage.md).

The previously installed product at that checkpoint, source `8f2315d`, was MVID
`d8f31d57-8612-465e-871c-93a62f2b3eae`, SHA-256
`C900BA09D92DA7CF50CC09033C63F5226DD04C18426CC386B30AF86EB0BA0941`.
`followup-20260930/deployment-v6.json` retains deployment and backup evidence.
Its complete instrumented suite passes with unchanged product hash. This is
managed acceptance for this exact binary; native/provider skips remain separate
scopes. Preceding Monaco-status `6a74af33` and v5 `f9a36c85` results retain their
own candidate identities. The current
native owner-STA diagnostic sees both fixed synthetic LocalAppData/TEMP GUID
directories and files with attributes matching the testhost, effective primary
token evidence and normal owned Excel exit. This read-only result does not
explain the earlier export path-not-found observation or qualify export; Q-027
remains open. Paired actual owner-STA observation and one synthetic export then
pass under LocalAppData and TEMP with explicit `/x /automation` bootstrap and
normal exits. Independent fresh COM-activation trials on this same product
still fail below LocalAppData/VBAi/GitTemporary. These different contexts do
not prove a cause or complete Git capture/import/recovery acceptance.

Current-v6 Access HelpContextID bridge and external CLR setters return 321 through
all three getters after one verified product Save. Both hosts fail normal exit,
so fresh-disk reopen is NOT_RUN; authorized force cleanup and retained stable
database copies do not qualify persistence. External HelpFile production CLR and
raw IDispatch PUT setters both return normally, but both fresh-disk reopen values
are altered through descriptor, CLR binder and raw IDispatch getters. Those hosts
exit normally, yet exact metadata persistence fails. The later raw HelpContextID
PUT returns HRESULT 0 and live value 321 after verified Save but fails bounded
exit; disk reopen is NOT_RUN and authorized forced cleanup preserves the failure;
the earlier HRESULTs and altered BSTR cause remain unexplained, without a
conversion heuristic or causal product fix. The initial Publisher trial verifies the exact sole
publication/persistence identity but fails the active-project startup guard with
null SelectedProject/ActiveModule, before baseline edits or product Save. Its
later guarded, authorized window close exits normally without force termination;
the startup failure and unproven selection/guard cause remain. External Publisher
VBE inventory observations are UNVERIFIED rather than evidence of an empty
inventory. These failed or unexecuted scopes do not qualify whole-host mutation or
lifecycle behavior; see recorded validation for terminal outcomes and additive
corrections preserving the original driver evidence.

Current-v6 owned Excel also passes ordered unsupported, one-Long and complete
declared scalar pages with terminal owner-STA traces, exact pagination and
identity/source/selection/mode preservation, followed by normal exits. This is
declared-page acceptance, not complete runtime Locals coverage or qualification
in another host; historical scalar stall/crash causes remain open. Two native
Monaco preparation attempts remain failed despite normal cleanup: a malformed
selection request and a bare-module/decorated-tab caption mismatch prevent the
closed-project/live-status scenario from running. A corrected fixture subsequently
passes that scope against the actual installed embedded editor: the selected
live tab keeps healthy status while the closed tab is retained, with native
source/identity preservation, normal exit and a reviewed real capture.
Exact product/test identities and terminal evidence are in recorded validation.

The broader current-v6 Excel core batch remains failed despite scoped rename,
form-fitting and options passes. Array/ParamArray operations return before
abnormal cleanup exits; protection cleanup, uncertain extended-option delivery
and unready breakpoint UI remain failures. The separate editor wrapper records
a combase access violation and a distinct abnormal final exit, with subsequent
cases NOT_RUN. Publisher's later module/class adapter Save/reopen passes with
normal initial/fresh exits; metadata/reference cases reach fresh reopen but fail
a common original-name/canonical-path selector assertion before final readback.
Those complete tests remain failed, without evidence of native value loss or
successful metadata/reference persistence. The corrected semantic reopen guard
then qualifies current Publisher module/class/form, Description and reference
addition by GUID/file and removal with exact native disk readback and normal
exits. HelpFile remains altered after reopen; HelpContextID changes the live
value despite an error response. Its preserved unsaved state and authorized
discard/normal cleanup do not qualify persistence. See recorded validation for terminal
and event scopes; whole-host compatibility is not qualified.

On preceding v5, source `2e75161`, MVID `f9a36c85`, the stable Access
application/PID, database-path and mapped/selected-project guard passes scoped
module/class, Description and reference-addition save/reopen with exact readback
and normal exits. HelpFile is altered after fresh reopen through descriptor,
CLR binder and raw IDispatch despite intact raw VARIANT canary. Native
HelpContextID getter/setter type metadata confirms I4, but the setter fails.
No heuristic getter repair or product correction is inferred. Reference removal
cannot reach disk readback because initial Quit fails to exit; its preserved
database and authorized forced cleanup do not qualify persistence. Publisher's
campaign stops during preparation because a recovered publication is mistaken
for the disposable document. HelpFile, HelpContextID, reference removal and
Publisher save scopes remain unqualified; no whole-host qualification follows.

The earlier v4 installed product is MVID
`d5e25e25-e3b4-4be0-ad45-a2dceb7be6b1`, with deployment and backup records in
`followup-20260930/deployment-v4.json`. Its Excel declared-scalar page passes
native Long/String/Boolean readback, unsupported array/Variant/object refusal,
correlated QuickWatch observer completion, unchanged source/selection/mode and
normal owned-process exit. This is a declared-page result, not complete runtime
Locals enumeration or acceptance in another host; the historical scalar stall
and crash cause remain open in Q-006.

On this preceding product, the Office adapter/property/reference campaign is terminal
after authorized forced cleanup. Publisher Description passed, while Access
save identity checks and Access/Publisher HelpContextID setters failed. Word
stalled during form preparation, with no terminal scenario result; unexecuted
cases remain NOT_RUN. The subsequent read-only Access diagnostic observes
changing CurrentProject wrapper IUnknown identities with a stable database path
and mapped/selected VBProject, followed by normal exit. It qualifies diagnosis
only; the subsequent stable guard and scoped v5 acceptance are described above.

Preceding-product SOLIDWORKS 2019 SP5 load, copied Edit Macro/source/form readback, Monaco
return-to-code and designer/code resize passed. Later class-source drift remains
unproven and unmodified. ExitApp stalled at native heap corruption `0xc0000374`;
authorized force/debugger cleanup ended the retained target without an observed
normal exit code. That cleanup does not qualify lifecycle behavior or establish
the crash cause. Preceding v5 `f9a36c85` separately passes scoped 2025 load,
disposable source/form preparation, compile, verified Save, content reload and
normal exit. Its strict post-execution SWP byte-preservation trial remains failed.
A genuine reviewed designer capture is distinct from earlier mislabeled code
captures, and the resize helper measures Monaco rather than designer geometry.
Current-v6 SOLIDWORKS 2019/2025 acceptance remains NOT_RUN.
Exact provenance and terminal outcomes belong in recorded validation.

A controlled UserForm export on preceding `d5e25e25` still fails under the
LocalAppData parent. Paired native tracing observes VBE `CreateFileA` and managed
parent-open path/name-not-found statuses for the exact existing GUID child,
despite a root synthetic write/read succeeding. A disposable non-Office CDB
preflight and verified debugger detach qualify the diagnostic tool only;
normal Excel exit is not export acceptance. The trace does not establish an
EFS, ACL or token cause and adds no current-v6 export qualification. Q-027
remains open.

The later `office-prerequisites-ready` and `outlook-ready` runs used loaded MVID
`3553ced4-f24c-4982-8a33-a593681d867e`, with DLL SHA-256 beginning
`026640435307`. Word and PowerPoint project-access prerequisites were enabled
and a classic Outlook profile was present for these runs.

The historical `office-accepted` and `outlook-accepted` runs used MVID
`ce19a20c-9708-4c17-b998-f3415b8e6303`, DLL SHA-256 beginning `581E9889F99E`.
Every owned Office process in that batch exited normally with code 0; none was
forcibly terminated. These results supersede the earlier Word save and Outlook
fixture failures within the tested scope.

Later Excel fitting and scalar acceptance used MVID
`353ddf2a-e065-4e41-8312-7a52c9b16cdd`, DLL SHA-256 beginning `629141C4B388`.
`scalar-excel-native`, `excel-form-close/353d-fit-scroll-only` and
`native-scalar/native-353d` establish the corrected paths below, with normal
owned-process exits. They do not replace the separate Office results on ce19.

Earlier SOLIDWORKS observations include 2019 process `37308` and 2025 process
`25984` on MVID `15416749-225a-4141-8b4f-2091bcddf418`. Those trials independently
retained saved-copy content despite premature uncertain save responses.
The corrected asynchronous-save candidate has MVID
`aaf3a555-76d4-4b18-ae09-1e7b3e085934`, DLL SHA-256 beginning `9A036C779999`.
It was exercised separately in 2019 SP5, PID `47384`, revision `27.5.0`, and
2025, PID `1236`, revision `33.1.1`. The maintainer explicitly authorized host
close/relaunch for these trials. Each result retains its build and host identity;
none establishes behavior in another version.

The same aaf3 binary also passed the later `sw-async-office` regression: Excel
save isolation across two owned instances, and Word/PowerPoint adapter-only
save/reopen with module/class/form readback and normal exits. The source tree
changed concurrently after compilation; this acceptance belongs to the identified
binary, not the later edits listed in the coverage summary.

| Host/environment | Observed scope | Important boundaries |
| --- | --- | --- |
| Excel x64 | Earlier builds covered load/bridge, save, protection, debugger inspection, Monaco, palette and Ollama tool read. On candidate 353d, `scalar-excel-native` passes complete UserForm fitting, arrays and persistence with normal exits. Independent scroll-only and generic scalar trials also pass: project Description, module Name with unchanged code, and Label BackColor verified through native getters and a reviewed designer capture. | Q-025's tested fitting/scalar paths are corrected; earlier crashes and abnormal exits remain recorded. This does not qualify every control, all Git workflows or other hosts. Format-options revision drift remains Q-026 despite a later successful rerun and verified preference restoration. |
| Word, Office 16 x64 | `office-accepted` verifies disposable DOCM editing, references, module/class/form content and adapter-only save/reopen; PID 35984 exited normally. On 353d, two-document native evidence (local artifact: `artifacts/qualification-v1/word-git-native/353-native-03/hosts/Word/2cb46193741e415ab5773c1539745096/qualification.json`) verifies same-name project resolution by canonical document paths, native selection, refusal of the old Git binding after owned SaveAs, and unchanged source in the other document; PID 50144 exited normally. | The path contract preserves raw `VBProject.FileName` and uses PID/IUnknown-matched `Document.FullName`. Complete Git remains blocked (**Q-024**): production Capture receives 0x800AC35C during export into GitTemporary, despite successful byte-identical bridge/external exports to artifact paths. The suite remains failed. Native chat/Git UI opening, every SaveAs/cancellation path and template support are not qualified; `Normal` was not modified. |
| PowerPoint, Office 16 x64 | `office-accepted`, disposable `.pptm`: shared VBE scenarios and VBAi adapter save passed, with `Verified=true`, `Uncertain=false`, followed by close/reopen without a helper save and module/class/form readback. Owned PID `50432` exited with code 0. | This supersedes the earlier trust-blocked save observation. It qualifies the tested existing-document save path, not every SaveAs, event cancellation, execution or Git workflow. |
| Access, Office 16 x64 | Earlier `office-accepted` shared VBE inspection/editing, references, compilation and MSForms/native-helper persistence retain their candidate-specific scope. Current Q-012 accepts existing ACCDB adapter module/class-source persistence, reference workflows, Description and explicit native General HelpContextID/representable HelpFile, with exact fresh readback and normal exits. | Legacy HelpFile/HelpContextID setters are refused pre-write with an explicit native General alternative. Access binds CurrentProject.FullName to its canonical mapped/selected injected-VBE project. French multi-object confirmation requires approved Type1/Type2 objects, complete selection and final authorization/deadline checks. No implicit compile, invented document Saved flag, first SaveAs, unsupported Unicode persistence or UserForm execution/rendering claim. |
| Publisher, Office 16 x64 | Earlier `office-accepted` shared VBE and native-helper scopes remain historical. Current Q-012 accepts existing serialized PUB adapter source/class/UserForm-state persistence, reference workflows, Description and explicit native General HelpContextID/representable HelpFile, with exact fresh readback and normal exits. | Bootstrap uses the frozen owned serialized seed with macros disabled and exact application/document/window identity. Empty NewDocument VBA initialization and first SaveAs remain unqualified. Document.Save is attempted once after final path/format/writable/revision checks. Legacy help metadata writes and unsupported Unicode are refused before mutation through their explicit guarded contracts; form rendering/execution and whole-host acceptance are not inferred. |
| Classic Outlook, Office 16 x64 | `outlook-accepted`: startup, exact loaded MVID/PID, project inventory, explicit-project debug state and VBE environment passed using the existing profile. Owned PID `37044` exited with code 0. | Read-only metadata qualification only. No profile was configured, mail read/sent, or user VBA code modified. Persistence, macro execution and full Outlook UI workflows remain unqualified. |
| New Outlook for Windows | Installed; **NOT_APPLICABLE** to the VBE add-in. | Microsoft documents that [new Outlook does not support VBA/macros](https://learn.microsoft.com/en-us/microsoft-365-apps/outlook/get-started/vba-alternatives). |
| OneNote desktop | Installed; **NOT_APPLICABLE** to the VBE add-in. | Its [documented Application interface](https://learn.microsoft.com/en-us/office/client-developer/onenote/application-interface-onenote) exposes content/window automation, not a VBE host. COM registration alone does not establish VBE compatibility. |
| Visio desktop | **NOT_RUN**: not installed on the qualification machine. | Microsoft documents a [VBE property](https://learn.microsoft.com/en-us/office/vba/api/visio.application.vbe); runtime and adapter behavior still require qualification. |
| Project desktop | **NOT_RUN**: not installed on the qualification machine. | Microsoft documents a [VBE property](https://learn.microsoft.com/en-us/office/vba/api/project.application.vbe); runtime and adapter behavior still require qualification. |
| SOLIDWORKS 2019 SP5 | On aaf3, the existing Type100 SWP save returned Verified=true/Uncertain=false after deferred owner-thread verification. An exact saved copy opened through Edit Macro retained module/class hashes and the form label. The reviewed designer capture shows the synthetic form without clipping and preserves native panes. PID 47384 exited normally. Earlier 1541 trials verified safe standalone-Open refusal and a specific VBE close/reopen layout. | **Q-021** is closed for this existing-SWP save path. Type100 SaveAs, signatures and full-restart original-file reload remain unqualified. Resize after opening the form failed with `Editor did not adapt`; original placement was restored. A separate synthetic marker macro ran successfully, but its post-run SWP hash changed and the unload scenario failed. **Q-030:** unsafe `VBProjects.Open` remains disabled; standalone Open functionality is not restored. Full assistant and debugger workflows remain open. |
| SOLIDWORKS 2025, revision 33.1.1 | On aaf3, PID 1236: native metadata/add-in connection, existing Type100 save with Verified=true/Uncertain=false, and exact saved-copy module/class/form readback through Edit Macro passed independently. Resize/restoration passed before opening the form; the reviewed designer capture shows the synthetic form and label without clipping, with native panes preserved. The earlier 1541 baseline also verified safe standalone-Open refusal and normal host exit. | **Q-021** is closed only for the tested existing-SWP save path. Type100 SaveAs, signatures, full-restart original-file reload, macro execution and complete debugger/assistant workflows remain unqualified. The assistant is absent from the aaf3 designer frame, so that capture does not qualify its behavior. No result is inherited from 2019 or Excel. |
| Other applications with a compatible VBE | Within the intended product scope. | Require load testing and operation-specific qualification; do not infer a result from another host. |

**Saving through a test helper's application API does not validate VBAi's
`save_host_document` tool.** Native VBE Save, an application adapter and the host's
own save API are distinct paths.

The earlier follow-up campaigns do not inherit the historical Office successes
in the table. Candidate `096b2e2b` passed the disposable PowerPoint adapter-only
save/reopen path with normal exit. Word's form creation timed out after its
project-access prerequisite was enabled; the later user closure produced a
Windows-recorded managed-exception crash, whose cause is not established. The
Excel Monaco and format-options trials also failed; their form-fitting success
does not clear those failures. The later installed `95576771` Access trial stopped
its bounded dialog worker and verified the application/project association. It
invoked native Save once, but `ProjectSaved` remained false and the adapter reported
an uncertain, unverified outcome. The fixture's shutdown deadline failed and reopen
was refused. A later revalidated owned-instance Quit exited normally; it does not
clear the persistence failure.

Publisher candidate `de3c5a79` passed module/class/form editing and helper-assisted
reopen with normal exits, but still refused the adapter trial because no current
application PID was verified through the ROT. The new pathless-project association
is limited to exactly one native document and one selected VBIDE project with
matching COM/VBE identities; it was not reached by that native trial. Microsoft
documents [Office ROT registration and instance-selection limitations](https://learn.microsoft.com/en-us/previous-versions/office/troubleshoot/office-developer/use-visual-c-automate-run-program-instance).
This is a discovery prerequisite, not permission to change focus, launch another
instance or accept an unverified application. Exact artifacts and source versions
are recorded in [validation results](test-coverage.md).

The Office batch did not qualify macro execution, stepping, signatures, all
controls, the entire theme/Monaco UI or LLM providers in every application.

## Native UI evidence and remaining scope

In Excel, `ui-native-startup/monaco-startup.json` records automatic Monaco opening
without invoking its menu, attachment to the VBE `MDIClient`, filling the workspace
after resizing/restoring, and preserving the native Object Browser layout while
Monaco is hidden. `ui-native-placement/states.json` records the assistant pane
retaining its bounds and parent when the VBE closes and reopens, with the add-in
still connected. These observations cover that Excel session and layout only.

On candidate 353d, `native-scalar/native-353d/owned-vbe.png` was reviewed against
its exact-PID capture record: the native UserForm designer displays the synthetic
yellow Label and Project Explorer displays the renamed module. This verifies
that visual result; it does not qualify assistant startup or every designer layout.
`ollama-ui-wire-353d` separately passes visible streaming, Stop and a subsequent
completed response in a detached assistant with synthetic VBE context. The earlier
empty-response failures remain unexplained (**Q-028**); this is not an embedded
Excel or SOLIDWORKS assistant result.

For SOLIDWORKS 2019, the reviewed host-only capture is
`solidworks-ui/sw2019-initial/vbe-47126866.png`. It shows Monaco in the VBE
workspace and an assistant welcome card with its actions visible.
`solidworks-2019-ui-resize.json` records matching Monaco/client bounds before and
after resize, followed by verification of the original window placement.
`solidworks-2019-ui-reopen.json` records native title-bar `SC_CLOSE`, then
`ShowWindow` on the same existing VBE handle. This proves that specific
close/reopen path; reopening through a SOLIDWORKS menu was not exercised.

The later `solidworks-ui/sw2019-newmacro/vbe-47126866.png` shows a new native
module selected while Monaco retains a closed project's draft. Current source
also follows native code-pane activation on the owning STA, with project,
component, active-window, mode and protection guards and revalidation after
renderer capture. It preserves an explicitly selected Monaco tab while the
native pane is unchanged. The `d5e25e25` 2019 return-to-code captures verify the
activated module/class in Monaco; a mere project-tree selection without native
code activation is outside that following path. Source `d7a1c75` corrects a
closed background document overwriting the selected Monaco tab's status, with
separate detached/managed acceptance. Current-v6 native UI acceptance passes
for the actual embedded Excel closed-project/live-status scenario, including
native preservation, reviewed capture and normal exit. Other UI/host scopes
remain open. Chat
project scope deliberately remains independent of editor selection; when its
project closes, the scope should clear and sending should become unavailable.

The initial 2019 lifecycle failure in `solidworks/sw2019-first/qualification.json`
is historical evidence. Later 353d trials completed standalone creation, initial
SaveAs, code/form preparation, compilation, save and close. A subsequent
`VBProjects.Open` terminated the host (**Q-030**); that API is now refused in
SOLIDWORKS. Native Edit Macro is the separately tested path for opening saved
Type100 copies, and does not restore standalone Open support.

The current source also implements two explicit SOLIDWORKS creation routes.
`create_solidworks_macro` requests a fresh absolute `.swp` path before native
creation. `publish_solidworks_macro` copies an unprotected Type101 draft into a
separate native Type100 macro and reports the original and destination identities;
the original project is retained. It verifies transported code, attributes,
UserForms and pictures, references and permitted General metadata around one
native save. Nondefault help or conditional-compilation metadata, project
references and uncertain native outcomes are refused. Signatures are not
transported. Generic Type101 Save/SaveAs remains unavailable in SOLIDWORKS.
These implemented routes do not inherit the earlier candidates' host acceptance:
their first-save and independent original-file reopening gates remain
candidate-specific. See [the recorded candidate validation](test-coverage.md#q-020-and-q-030-native-macro-candidate-2026-10-05).

The current frozen native-macro candidate qualifies native creation and explicit
publication on selected SOLIDWORKS 2019 SP5 and 2025 SP1.1. Both versions retain
complete synthetic code/class/form/picture content through verified save and
same-original-file reopening in a fresh host, with reviewed native designer
images, normal owned-host exits and restored temporary COM registration.
Cold published-form inspection requires an explicit designer opening before
reading component properties; the observed closed-designer E_FAIL remains a
documented limit. Generic standalone Open and Type101 SaveAs remain unavailable.
These are operation-specific results, not complete UI or release acceptance.
See the [current qualification checkpoint](release-qualification.md#q-020-and-q-030-native-macro-checkpoint-2026-10-05)
for candidate identity, retained failures and scope.

On aaf3, `solidworks/type100-2019/async-save-02/` and
`solidworks/type100-2025/async-save-01/` independently record a single product
Save followed by verified deferred completion. Their
`saved-copy-content-verification.json` files establish retained module/class
hashes and form labels after opening exact disk copies through Edit Macro.
This does not prove signature preservation or original-file reload after a full
application restart. In the separate 2019 `async-save-01` trial, the disposable
marker macro ran successfully, but the subsequent SWP hash changed; the unload
scenario remains failed and provides no reload acceptance.

The reviewed aaf3 frames in `solidworks-ui/sw2019-aaf3-form/` and
`solidworks-ui/sw2025-aaf3-form/` show the synthetic forms and labels without
clipping, with native panes preserved. Neither frame contains the assistant.
`solidworks-2025-aaf3-resize.json` passes the layout check before form opening.
`solidworks-2019-aaf3-resize.json` instead reports `Editor did not adapt` after
form opening and verifies restoration of the original placement. This remains
historical failed evidence with no established cause. The later `d5e25e25` native
2019 designer/code resize and restoration pass for their recorded layout; they
do not qualify another layout or current-candidate 2025 behavior.
Both versions still require fuller code/designer/Object Browser switching,
focus and command routing, visible debugger feedback, embedded assistant
approval, streaming/cancellation and project privacy qualification. A bridge
response or detached WinForms screenshot cannot establish these behaviors.

## Cross-host constraints

Native dialogs and accessibility depend on VBE version, UI language, DPI and window
layout. COM references, ActiveX availability, protected projects and trust settings
also affect individual operations. Third-party controls require their own evidence.
The native dark theme remains experimental and uses additional OS-dependent
behavior; it is not part of a universal host-compatibility guarantee.

Word's observed saved document initially had no accessible `VBProject.FileName`;
after native Save, that property identified `~WRL0002.tmp` while the document's
`FullName` remained the `.docm`. Candidate 353d now resolves the canonical document
path through COM project identity. The native two-document scenario verified
selection, disambiguation, refusal of the stale Git binding after SaveAs and an
unchanged second document. Full Git qualification remains blocked by native export
failure in the production temporary directory; successful diagnostic exports to
the artifact directory do not qualify that capture path (**Q-024/Q-027**).

## Report another environment

Use the host-compatibility issue template. Record application/version, VBE version,
Windows build, process architecture, UI language, DPI, VBAi commit and the exact
operation. Distinguish **pass**, **failed**, **blocked by a prerequisite** and
**not run**. A minimal disposable file and steps are more useful than a blanket
“works” or “does not work” report. Never submit confidential production documents.

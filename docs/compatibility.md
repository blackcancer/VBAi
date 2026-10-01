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

The observations below come from the **2026-09-29 through 2026-10-01** qualification artifacts for
Microsoft 365 **16.0.20326.20158 x64**. The `office-final`, `monaco-save`,
`monaco-privacy`, `ui-native-startup`, `ui-native-placement` and `outlook` runs used
the candidate with DLL SHA-256 beginning `482864942BAC`; `excel-final` used the
candidate beginning `F3C48EAEA590`. These are separate build observations, not a
blanket qualification of v1.0.0. Full provenance and test totals belong in
[recorded validation](test-coverage.md).

The currently installed product, source `8f2315d`, is MVID
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
| Access, Office 16 x64 | `office-accepted`, disposable `.accdb`: shared VBE inspection/editing, references, compilation, an MSForms UserForm and native-helper save/reopen passed. Owned PIDs `36512` and `57496` exited with code 0. | The existing-document adapter now has scoped v5 module/class, Description and reference-addition save/reopen acceptance above; HelpFile, HelpContextID and reference-removal acceptance remain incomplete. Access matches CurrentProject.FullName to a unique injected-VBE project and uses stable Application/PID and mapped/selected-project guards before the built-in VBE Save command for the exact active project. It does not compile implicitly or invent a document Saved flag; first SaveAs is unavailable. |
| Publisher, Office 16 x64 | `office-accepted`, disposable `.pub`: shared VBE inspection/editing, references, compilation, UserForm and native-helper save/reopen passed. Owned PIDs `30152` and `27104` exited with code 0. | The follow-up source includes an existing-document adapter; native adapter save/reopen remains `NOT_QUALIFIED`. Publisher matches Document.FullName to a unique injected-VBE project and invokes Document.Save once after final identity/path/format/writable checks. First SaveAs is unavailable. |
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

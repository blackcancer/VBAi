# Recorded validation

## Current qualification checkpoint

Qualification on **2026-09-29–30** uses branch `qualification/v1-20260929`, based on
**`9e43a9df3bd5b64b66c47490d7c4cf7ac284ab3a`** with uncommitted corrections.
This is not a released source revision. The latest completed whole-suite managed
coverage run is `sw-async-global/qualification.trx`, tied to
`source-manifest-sw-async.json` and candidate `aaf3a555`. Concurrent source edits
were detected after compilation while this run was active; `ChangedSourceFiles`
in its summary records the drift. Results qualify the compiled candidate, not
the subsequently edited working tree. The pre-run content-manifest SHA-256 is
`e92bf2a345e299eaefb3a2b66d839cd81c2ea268e9c6751258f771d0a2d6da7a`.
This includes the native scalar, Word path, FRX preflight, SOLIDWORKS Open refusal,
empty-scope chat and deferred SOLIDWORKS save corrections. Native operation
acceptance remains independent and identifies the loaded binary.
The original activation ledger maps historical conditional scenarios to their
explicit attempts; it does not override later native failures or imply that
every opt-in passed on this candidate.

The earlier combined source checkpoint is captured in `source-manifest-final.json`
(content-manifest SHA-256 `bd0f3e5b64bf77fc6dd43341bbc2758485a7b8775a06373796c572ef1aa9a1c4`).
`qualification-final-build/VBAi/Debug/net48/VBAi.dll` compiled with zero warnings
and errors: MVID `5cc513d1-5569-4835-bf6c-cf70a18274fb`, SHA-256
`A12378FB826CAAD6C0BBE79BF09C52332760E6C32FDB4F65138876CB90F5F167`.
This candidate's deployment is retained in `deployment-5cc-before-word-path.json`.
Its full managed
collector run passed; loading and operation acceptance in each native host remain
separate gates. AssemblyVersion is still `0.1.0.0`; no version bump
or release was performed. The final Release compilation passed with zero warnings/errors:
MVID `58c74b92-264e-493b-a591-9f618d33ef28`, SHA-256
`1B35D45F0AB44BAAC2F2EB1CAD765183C176CA0DDB87988C5B83848CE796DEAF`
(`qualification-final-release/`, `final-build-qualification.json`).

The earlier native scalar candidate is `native-scalar-final-build/`, MVID
`353ddf2a-e065-4e41-8312-7a52c9b16cdd`, SHA-256
`629141C4B388CD0E8146B1744762CDF80AEFBE4EA8747A3EF74D8D8DC05A0C9A`.
It centralizes Word document paths, avoids an undersized native COM setter buffer,
and rejects empty or out-of-range FRX resource references. The previous ce19
candidate remains identified in `deployment-ce19-before-native-scalar.json`. Deployment and
original binary backups are recorded in `final-native-deployment.json`.
The historical whole default suite (`scalar-global/qualification.trx`) and coverage
passed on the identical Debug binary. Its source-manifest SHA-256 is
`5bcc3071b7e5d5c1903bfa6f8af6778b58a7a15cda3a554215c88c705ce50e05`.
The Release build also passed with zero warnings/errors: MVID
`713d7ac5-039a-41c7-a600-2670604ede6a`, SHA-256
`014EA65CEAB013726BE3342376E5D8BB4FCFC99F5E6A66A64A4A20AD47BC7C29`.
`scalar-global-build-Debug.json` and `scalar-global-build-Release.json` record
assembly identities and verified third-party notice delivery. The default run
leaves native/provider scenarios conditional; their explicit executions below
remain the authority for host acceptance and failures.

The preceding native candidate used `sw-protection-build/`, MVID
`15416749-225a-4141-8b4f-2091bcddf418`, SHA-256
`BFA41E659C2463ABFB0CADF8C807DF5CA0E99B76D3E1CD73134A2DEB32C9A398`.
It refuses unsafe standalone Open in SOLIDWORKS, preserves Stop while disabling
Send without a project, and identifies individual native-save verification failures.
`sw-protection-build/candidate.json` and `deployment-353d-before-sw-protection.json`
record provenance. That intermediate candidate had focused validation only.

The latest tested binary candidate is `sw-async-final-build/`, MVID
`aaf3a555-76d4-4b18-ae09-1e7b3e085934`, SHA-256
`9A036C779999A82A6104A1815A2354BF5A5FBDACEFC217B896A1DF338F55005C`. It completed Debug and Release builds without
warnings or errors. Release MVID is `f2aef6d0-6203-48a5-bb1b-79ea62eb9b4f`, SHA-256
`1934EFB212DE10FA37C70821330D441DD57C45D23BF1E0D41E8C4C330DA919B1`; both builds verified delivery of the third-party notices.
`sw-async-Debug-identity.json`, `sw-async-Release-identity.json` and
`deployment-1541-before-sw-async.json` retain provenance and deployment history.
`solidworks-async-native-summary.json` verifies independent Type100 save and exact
saved-copy module/class/form readback in 2019 SP5 and 2025 SP1.1, with normal exits.

Earlier native trials used `qualified-build/VBAi/Debug/net48/VBAi.dll`, MVID
`3553ced4-f24c-4982-8a33-a593681d867e`, SHA-256
`0266404353075677029F8DA537838768BDFD9F6509E3D7C64AC174B530AB9261`.
The later focused Type100 candidate used MVID
`de014e73-7398-41f8-bc25-7e84115b688a`, SHA-256
`939C74BD1538B06EBEE2269E58BA056F2FB50CF6FF9303ED187F4EAF5114CAB3`
(`type100-candidate.json`); it was not installed. These earlier build identities
do not establish acceptance of the latest source. The initial registry-only
redirection loaded the wrong MVID and is not accepted.

## Executed evidence

Paths below are relative to `artifacts/qualification-v1/` unless stated
otherwise. Counts describe separate runs and must not be added together as
unique test coverage.

| Evidence | Result | Scope and boundary |
| --- | --- | --- |
| Office save regression after async dispatch, `sw-async-office/native.trx` | 3 passed, no failures/skips | Candidate aaf3: two-instance Excel save isolation, Word DOCM and PowerPoint PPTM adapter save/reopen with module/class/form readback. Normal owned-host exits; `runner-observation.json` records no remaining host. This does not rerun Access, Publisher or Outlook on aaf3. |
| Latest candidate whole default suite, `sw-async-global/qualification.trx` | 2,105 passed, 0 failed, 36 conditional skips; 2,141 total | Compiled candidate aaf3; source drift during execution is recorded separately. Native/provider opt-ins are disabled here, not silently marked as passed. |
| Current managed coverage, `sw-async-global/coverage-summary.json` | 32,295/32,459 lines; 32,842/33,205 branches | VBAi managed assembly inside VSTest, including detached UI; external hosts, native renderer and JavaScript are not instrumented. |
| Deferred-save focused regressions, `sw-async-final-focused/focused.trx` | 142 passed, no failures/skips | Owner-thread yielding, delayed Saved transition, cross-session overlap refusal, changed-state/timeout uncertainty, bridge/direct/catalogue validation and privacy. Earlier `sw-async-focused/` had 139 passed and 1 stale synchronous contract expectation; the contract now exercises all schema cases through async dispatch. |
| SOLIDWORKS 2025 metadata, `solidworks-2025-aaf3-metadata/native.trx` | 1 passed, no skips | PID 1236, revision 33.1.1, exact candidate aaf3 and connected VBE add-in. Persistence and UI are separate evidence below. |
| SOLIDWORKS deferred save, `solidworks-async-native-summary.json` | Verified separately in 2019 and 2025 | A single save per fixture; deferred confirmation, then exact saved-copy module/class/form readback. Neither helper save nor macro execution occurs in these persistence trials. Original-file reload after full host restart and signatures remain unqualified. |
| SOLIDWORKS aaf3 UI and execution limits | Form designer captures reviewed; mixed outcome | Synthetic form/label visible under both hosts. 2025 resize before designer passed; 2019 resize with designer open failed with placement restored. A separate 2019 marker macro ran, but post-run SWP hash changed and the unload scenario failed. No complete UI/debugger qualification is inferred. |
| Authorized Codex/SOLIDWORKS read, `solidworks-codex-1541/connected.trx` | 1 passed, no skips | Existing account, model `gpt-6-luna`, low effort, one native read of the manifested synthetic constant module, expected marker and unchanged source. No macro execution, other module access, native write or embedded-chat UI qualification. Original conditional ledger now has 33 successful scenario outcomes across recorded candidate builds; additional qualification failures remain open. |
| SOLIDWORKS/chat correction regressions, `sw-protection-focused/focused.trx` | 209 passed, 0 failed, 1 conditional Ollama UI skip | Candidate 1541; lifecycle, Type100 guards, chat state and tool boundaries. Prior `sw-open-guard-red/red.trx` and `empty-scope-red/red.trx` each reproduced the corresponding missing protection. |
| SOLIDWORKS 2019 metadata, `solidworks-2019-metadata-353d/native.trx` | 1 passed, no skips | PID 56924, revision 27.5.0, loaded 353d; bridge/add-in connection only. The later native Open termination remains separate failed evidence. |
| SOLIDWORKS 2019 native protection and UI, `solidworks-2019-1541-open-refusal.json`, `solidworks-2019-1541-resize.json`, `solidworks-2019-1541-vbe-reopen.json` | Observed refusal and UI checks passed | PID 37308, 1541. No project/file change on refusal; resize/restoration and same-window VBE close/reopen. The navigation UIA TabItem assertion was inconclusive although the reviewed capture shows the expected Monaco document. |
| SOLIDWORKS 2019 Type100 save, `solidworks/type100-2019/native-created-03/` | Immediate adapter result uncertain; persisted-copy readback verified separately | Saved flag false immediately and true later. Native Edit Macro opened an exact byte copy; module/class/form survived. `saved-copy-independent-verification.json` preserves the harness's final Boolean-report error without repeating native input or promoting the uncertain save to product PASS. No macro executed. |
| Historical whole default suite, `scalar-global/qualification.trx` | 2,085 passed, 0 failed, 36 conditional skips; 2,121 total | Debug MVID 353d, unchanged source manifest; opt-in native hosts/providers disabled here. The original activation ledger covers 33 historical conditional scenarios; the added Word Git and UserForm scopes are tracked separately. |
| Historical managed coverage, `scalar-global/coverage-summary.json` | 32,220/32,379 lines; 32,801/33,155 branches | VBAi managed assembly only inside VSTest, including detached UI; excludes measurement of native renderer, JavaScript and external host processes. Not a universal host-compatibility claim. |
| Final combined managed suite, `qualification-final/qualification.trx` | 2,065 passed, 0 failed, 33 skipped | Full instrumented Debug suite on `source-manifest-final.json`; native host, connected-account and live provider opt-ins disabled. Detached UI included. Product DLL hash remained unchanged. |
| Initial instrumented baseline, `baseline/baseline.trx` | 2,003 passed, 1 failed, 28 skipped | Before corrections; post-step UI observation timed out. Its direct cause is not established by the later deterministic queue regression. |
| Completed instrumented checkpoint, `global-final/global.trx` | 2,025 passed, 1 failed, 33 skipped | Only failure: `OpenBindingValidationLockOwnershipAndReleaseMatrix`, Git absolute path limit. Collected before the long-path/standalone follow-ups; not a release pass. |
| Git long-path correction, `git-longpath-green/longpath.trx` | 2 passed | Originally failing binding scenario and real long-cache initialization/commit/readback; no global configuration changes. |
| Standalone/project contracts, `swfix-unit/swfix.trx` | 44 passed | Synthetic/unit guards for the narrow unsaved Type101 path plus regressions. Not native Type100 SWP persistence or observed HRESULT proof. |
| Type100 adapter and dispatch, `type100-focused-green/type100.trx` | 115 passed, no skips | Includes actual selection implementation over fake COM objects, refusal/uncertain outcome, owner-thread Bridge dispatch, direct/catalogue approval revalidation and standalone/project regressions. No native SWP save/reopen. Earlier `type100-focused/` had 81 passed and 1 failed on restoration typing; fixed and rerun. |
| Word path and standalone contracts, `word-save-build/word-standalone-green.trx` | 61 passed, no skips | Word path-unavailable case passed after 2 failing red cases in `word-save-red.trx`; narrow unsaved Type101 guards included. Outlook explicit project metadata compiled; native reruns remain separate. |
| Focused fixes, `fixes-final/fixes.trx` | 73 passed | Algorithm/provider/guard regressions, synthetic and detached scope. |
| Final pre-Enter guards, `immediate-final-guards/guards.trx` | 50 passed | Context/policy changes after command preparation, STA dispatch and no duplicate Enter. |
| Excel operation batch, `excel-final/native.trx` | 15 passed, no skips | Loaded candidate identity, editing/undo, forms, protection, options restoration, procedure operations, persistence, breakpoint and paused values. |
| Additional Office, `office-final/native.trx` | 4 passed | Word, PowerPoint, Access and Publisher; normal exits, no forced termination. Refusals do not qualify missing/blocked save adapters. |
| Final Monaco/Immediate, `native-acceptance/native.trx` | 5 passed, no skips | Real Excel renderer edits/conflicts; reference/hover/completion/format/undo; save cancellation/failure/reopen; privacy; two same-name project Immediate isolation. Normal owned exits. |
| Classic Outlook, `outlook/native.trx` | 1 skipped | Earlier opt-in reached a missing-profile prerequisite. User configuration now makes the prerequisite ready; rerun pending, not passed. |
| Live provider, `ollama-final/ollama.trx` | 2 passed | Production loopback Ollama catalogue, streamed tool response, cancellation and recovery, synthetic content. |
| Activated Excel batch, `activated-excel/excel.trx` | 12 passed, 4 failed, no skips | Final production candidate. Failures: form/array shutdown crashes, persistence shutdown deadline, and format-options revision/restoration. WER diagnosis in `excel-activated-shutdown-diagnosis.json`; fixture lifetime changes compiled separately without claiming native recovery. |
| Focused Excel lifetime rerun, `excel-lifetime-rerun-1/lifetime.trx` | 2 passed, 1 failed, no skips | Arrays and two-process persistence passed normal exit after tracked CommandBars release. Form heap corruption recurred during Workbook.Close(false), before Quit; `form-shutdown.json` and `wer-events.json` retain evidence. No further automatic attempt. |
| Word temporary-path correction, `word-path-final-build/word-path-green.trx` | 65 passed, no skips | Two new failing-before cases reproduced missing/temporary Word backing paths; document identity, PID, native path, format, source and saved-state guards retained. PowerPoint remains strict. |
| Corrected Office acceptance, `office-accepted/native.trx` | 4 passed, no skips | Candidate ce19: Word/PowerPoint adapter-only save and reopen verified. Access/Publisher available scenarios and safe save refusal passed; their save adapters remain absent. Every owned host exited normally. |
| Corrected Outlook acceptance, `outlook-accepted/native.trx` | 1 passed, no skips | Candidate ce19; read-only metadata/startup and normal exit, no mail mutation. |
| Activated Monaco, `activated-monaco/monaco.trx` | 4 passed, no skips | Candidate 5cc: editing/conflicts, live language features, save/reopen and project privacy; normal exits. |
| Activated native palette, `activated-palette/palette.trx` | 1 passed, no skips | Candidate 5cc: conflict/archive/rebase and complete native palette restoration, normal exit. |
| Activated Ollama chat UI, `activated-ollama-ui/ollama-ui.trx` | 1 failed | Candidate ce19: visible-streaming observation timed out while status was already ready. This failure remains recorded despite a later strict pass. |
| Instrumented Ollama chat UI, `ollama-ui-diagnostic/ollama-ui.trx` | 1 failed | Candidate ce19: synthetic transcript contained the visible `No text response.` fallback, with no streamed text or refused tool call. Turn finished around 11.45 seconds; the observation timed out later. Normal test-host cleanup; cause unresolved. |
| Activated Ollama Excel, `activated-ollama-excel/ollama-excel.trx` | 1 passed, no skips | Candidate ce19: real model read an unprompted random marker through project-bound native tools, preserved code and exited normally. |
| Instrumented Excel options rerun, `excel-options-instrumented/options.trx` | 1 passed, no skips | Candidate ce19: full format-options scenario and complete baseline restoration asserted; owned Excel exited normally. The earlier revision drift was not reproduced and its cause remains open. Successful-run baseline attachments were not retained by this VSTest invocation. |
| Native UserForm GitHub attempts, `userform-github-20260929221805/results/userform.trx` and `userform-github-20260929223708/results/userform.trx` | 1 failed in each independent attempt | Candidate ce19: first fixture used an unavailable Designer.Width member; corrected fixture then reached native export and received 0x800AC373. Neither attempt reached push; both owned hosts in the corrected run exited normally. No remote round-trip acceptance. |
| Combined scalar/Word/FRX contracts, `native-scalar-final/focused.trx` | 69 passed, no skips | Candidate 353d. Failing-before setter regressions are retained in `fit-scroll-red` and `scalar-fallback-red`; existing ce19 FRX boundary probes accepted invalid snapshots. Native acceptance is separate. |
| Word same-name Git diagnostics, [native03 TRX](../artifacts/qualification-v1/word-git-native/353-native-03/native.trx) and [phase evidence](../artifacts/qualification-v1/word-git-native/353-native-03/hosts/Word/2cb46193741e415ab5773c1539745096/qualification.json) | 1 failed, no skips | Product 353d unchanged; test-only phased harness. Both production Git captures fail with 0x800AC35C. Independent phases pass: canonical resolution of two Project-named DOCMs, bridge native selection, stale Git binding refusal after owned SaveAs, other-document source preservation, and bridge/external-STA exports to artifact paths with identical SHA-256 `AB893DEF2C283CDC4A0F3B8EE5363508EDF1A1F7B7320528C4FA5703E4B635F5`. Word PID 50144 exits 0 without forced termination. No macro, UserForm or remote operation; no actual chat/Git UI opening claim. Native01 failed on a nonexistent harness command; native02 first exposed Capture failure. Both earlier runs remain preserved with normal exits. |
| Corrected Excel native regressions, `scalar-excel-native/native.trx` | 3 passed, no skips | Candidate 353d: complete UserForm fitting, Variant arrays and persistence now pass with normal owned-process exits. Independent `excel-form-close/353d-fit-scroll-only/case.json` also records normal exit after the isolated scroll operation. |
| Generic native scalar dispatch, `native-scalar/native-353d/qualification.json` | PASS, separate script evidence | Candidate 353d, Excel PID 56904: project Description and component Name use generic dispatch, with direct COM readback and unchanged module source. Label BackColor passed numeric-to-Color conversion, generic dispatch and native OLE readback. Reviewed owned-VBE capture shows the yellow Label and renamed module. Normal exit 0; no macro, save or complete Git qualification. |
| Captured Ollama chat UI, `ollama-ui-wire-353d/ollama-ui.trx` | 1 passed, no skips | Candidate 353d: unchanged strict scenario observed 36 visible characters while busy, Stop/cancellation and the next visible complete reply. Real loopback HTTP with passive bounded wire capture; one synthetic tool call refused. Detached UI and simulated VBE. No LLM production fix; wrapper timing differs, so earlier empty-response failures remain unresolved. |
| Real UserForm GitHub transport, `userform-git-transport-20260929225354/transport.json` | PASS, separate script evidence | Candidate ce19 exported synthetic FRM/FRX, committed and pushed a dedicated branch, then fetched through another repository with strict byte equality and unchanged main. Form-only manifest; no native import or MacroGitOperations acceptance is inferred. |
| Git-fetched UserForm native import, `userform-fetched-import-20260929231432/native-import.json` and `visual-review.json` | PASS, separate script evidence | Candidate 353d, one verified bridge import into an empty disposable workbook. Exact fetched FRX bytes retained; native controls, geometry and independent Designer/Component Caption properties preserved through save/reopen. Imported code has exactly one additional leading CRLF: content accepted under that explicit rule, exact source text equality remains false. Reopened readback exactly matches imported state. Source/import/reopen designer captures reviewed; VBE frame PID ownership verified because Designer.HWnd is zero. Design mode 2 observed before and immediately after import; post-reopen mode was not recorded. Macros/events disabled, no execution request, normal owned Excel exit 0. No MacroGitOperations or raw FRX re-export equality qualification. |
| Activated final-candidate HTTP provider, `activated-ollama/ollama.trx` | 2 passed, no skips | Real loopback model streaming/tool round-trip plus cancellation/recovery with synthetic content. |
| Activated synthetic Git fixture, `activated-git/git.trx` | 2 passed, no skips | Existing-account GET against the explicitly retained private qualification repository ID/main commit; includes fixture safety regression. No PR lifecycle dependency or unrelated private repository. |
| Shown assistant plus Ollama, `ollama-ui/ollama-ui.trx` | 1 passed | Send, visible streaming, Stop and subsequent complete response. Simulated VBE; tools refused. |
| Ollama to real Excel, `ollama-excel/ollama-excel.trx` | 1 passed | Real client and project-bound `read_module` dispatch into Excel COM; marker absent from prompts returned correctly, source unchanged. No macro or embedded UI claim. |
| UI regressions, `ui-regressions/ui.trx` | 7 passed | Shift+Tab, provider focus navigation, debug queue responsiveness and cancellation. |
| Welcome regression, `welcome-red/`, `welcome-green/` | 1 failed before; 3 passed after | French wrapping across panel widths; native before/after captures independently inspected. |
| Reference-cache contracts, `language-cache/language.trx` | 6 passed | Catalogue TTL, draft overlay and reference add/remove/re-add. |
| Local Git, `git-live-local.log` | 17 passed | Real git.exe, local bare remote, conflicts/recovery and synthetic VBE. |
| Connected Git read, `git-live-readonly.log` | 1 passed | Existing identity and read operations; credentials excluded from evidence. |
| Private remote Git, `git-remote-execution.json` | PASS | Creation/retention explicitly authorized; 2 branches pushed and fetched snapshots verified. No existing repository changed. |
| JavaScript, `javascript.log` | 60 passed | Language/editing scenarios, not branch coverage. |
| Native renderer, `native-selftest.log`, `native-loader.log` | PASS | 20 native start/stop cycles and managed loader/hash/ABI checks; no C++ coverage claim. |
| Designer validation, `designers-final/` | 46 surfaces passed | Construction, resizing, editable child components and serialization. |
| Designer metadata | 27 items passed | Evaluated project metadata, not native behavior. |
| Test layout, `mirror-final.json` | 255 mirrors for 313 production files | Inventory only, not measured coverage. |
| Monaco distribution, `monaco-assets.log` | 17 regenerated files matched | Distribution integrity, not renderer execution. |
| Notices, `notices-delivery-final.json` | 14 output files matched hashes in each Debug and Release output | Exact upstream payload and manifest; project license unchanged. |

Earlier native attempts remain recorded. The first combined Monaco run skipped
following scenarios because Excel had not exited. Later reference/save assertions
exposed fixture timing and oracle mistakes. The accepted batch checks process
exit, bounded reference refresh, both editor timers, actual native-save callback
invocation and exact VBE canonical readback, preserving string/comment checks.

## UI and environment

Microsoft 365 x64 is build `16.0.20326.20158`, French UI, on Windows build 26200.
`ui-native-startup/` records automatic Monaco startup, workspace fill, resize
and native Object Browser coexistence. `ui-native-placement-fixed/` records VBE
close/reopen and the repaired welcome card. `display-profiles/` records 4 detached
windows on each of 2 installed displays, both at 96 DPI. Higher/mixed DPI, high
contrast and untested native layouts remain unqualified.

`artifacts/ui-review/screenshots/` contains the detached light/dark review of
synthetic settings, Markdown, streaming, Git and virtualized transcript content.
Capture existence alone is not visual acceptance. Earlier SOLIDWORKS 2019 SP5 trials loaded
MVID `3553ced4-f24c-4982-8a33-a593681d867e`; `solidworks-ui/`,
`solidworks-2019-ui-resize.json` and `solidworks-2019-ui-reopen.json` record bounded
native UI observations. On aaf3, existing Type100 save and saved-copy content passed independently in
2019 and 2025. Reviewed designer captures are in `solidworks-ui/sw2019-aaf3-form/`
and `solidworks-ui/sw2025-aaf3-form/`. They do not qualify the assistant outside
the frames. The 2019 designer-open resize failure remains an explicit UI gap. See [compatibility](compatibility.md) and
the [release qualification tracker](release-qualification.md).

## Measurement boundaries

Build/test commands use isolated outputs as described in [testing](../tests/README.md).
The current `sw-async-global/coverage-summary.json` identifies the successful
default suite, pre-run source manifest, detected later edits and canonical Cobertura output.
Product binaries were not instrumented; collection used the test output copy.
Historical checkpoints in the evidence table are not current-source measurements.

C# collection does not instrument separate Office/SOLIDWORKS processes, the C++
renderer or JavaScript. Conditional skips are not passes. A provider response
does not prove native mutation; a helper save does not prove the product adapter.
No production exclusions were added to improve the metric. Ignored local
artifacts are not public downloadable reports; earlier records remain in Git
history.

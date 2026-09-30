# Version 1.0.0 qualification

**Decision: not qualified for release yet.** This is the current acceptance gate,
not a claim of universal Office compatibility. Test totals and measured coverage
are maintained only in [recorded validation](test-coverage.md).

## Report scope

Qualification checkpoint: 2026-09-30, following the 2026-09-29/30 campaign.
The register below contains all 30 findings, Q-001 through Q-030. Each status
applies only to the stated operation, host and tested candidate; CLOSED does not
qualify an entire application. OPEN and PARTIAL entries remain release gates
where required by the agreed scope.

The last candidate with the native acceptance results below has MVID
`aaf3a555-76d4-4b18-ae09-1e7b3e085934`. Concurrent source changes made after that
build are not covered by its results. The full qualification branch includes the
implementation changes, tests and this register. Its publication checks are
recorded separately in [recorded validation](test-coverage.md); they do not
replace the operation-specific native acceptance or close the remaining gates.

## Required environment

The agreed target is the installed Microsoft 365 x64 suite, followed by
SOLIDWORKS 2019 SP5 and SOLIDWORKS 2025, qualified independently. Include classic
Outlook; record absent applications and applications without a compatible VBE
separately. Windows x86, other Office releases and unavailable applications are
not inferred to pass from this workstation.

Each native result must identify the application build, PID, loaded VBAi MVID,
disposable project and operation. Protocol success is not proof of successful
native mutation, persistence, rendering or shutdown. A refused unavailable
operation can pass a refusal test while the capability remains unqualified.

## Defect and acceptance tracking

| ID | Priority | Finding | Acceptance criterion | Current status |
| --- | --- | --- | --- | --- |
| Q-001 | P1 | Immediate execution could target another active native project after validating only the requested project's mode. | Chat, catalogue and bridge refuse mismatched/missing native selection before execution; verify with two native projects and context changes during command preparation. | Native same-name/two-workbook bridge scenario passed. Final pre-Enter project/mode/policy revalidation on the owning STA passed focused tests; context changes during echo refuse Enter without retry. User-selection races cannot be called atomic. |
| Q-002 | P1 | Git imports accepted identical component names across different types/extensions. | Reject case-insensitive duplicate names before importing or deleting any component. | Fixed with preflight regression. |
| Q-003 | P2 | An atomic editor batch at the size limit could be rejected because of a temporary intermediate length; repeated whole-text copying made multi-cursor batches expensive. | Validate final size atomically, preserve UTF-16 offsets/tied insertion order and benchmark production implementation. | Fixed; failing-before/passing-after regression and comparative benchmark recorded. |
| Q-004 | P2 | SSE concatenation copied growing responses repeatedly; line/body limits were enforced after unbounded buffering. | Bound actual input bytes for SSE, JSON and catalogues; preserve complete tool arguments, cancellation and error semantics. | Fixed with generated-stream regressions and live local-model checks. |
| Q-005 | P2 | A queued breakpoint could consume the only scheduled post-step UI observation. | Retain observation until command ownership is released, without replaying native commands or relying on the periodic timer. | Deterministic failing-before/passing-after regression recorded. |
| Q-006 | P1 | Host tests could report success after uncertain saves, lost forms, stale loaded assemblies or forced termination. | Assert verified persistence, reopen without helper saving, retain form expectation, verify MVID and normal exit. | Assertions now enforce these boundaries. The later Excel activation correctly reports the native crashes in Q-025 and the failed preference restoration in Q-026; earlier successes do not override them. Save-adapter limitations below remain. |
| Q-007 | P2 | External COM temporaries in host fixtures could outlive Quit and obscure shutdown results. | Explicitly release owned collections/windows/commands and require normal exit; distinguish forced termination from an independent crash. | Corrected Office and Monaco fixtures passed normal-exit checks; no forced termination counted as success. |
| Q-008 | P2 | An obsolete UI test called a removed Markdown rendering API. | Exercise the current native Markdown view, streaming, transcript virtualization, settings and Git views with isolated state. | Harness corrected; detached UI pass and captures recorded. |
| Q-009 | P2 | NuGet runtime license/notice payloads were absent from the build output. | Deliver exact upstream texts with provenance and verify output hashes. | Payload added and final Debug/Release delivery hashes verified. |
| Q-010 | P3 | WebView2 creates persistent per-PID profiles without a retention policy. | Define ownership and safe cleanup only after browser processes exit; preserve active/private state. | OPEN; existing storage measured, nothing deleted. |
| Q-011 | P1 gate | Word/PowerPoint adapter acceptance required project-access prerequisites. | Run verified adapter-only save/reopen under a maintainer-approved host configuration. | CLOSED for the tested existing-document save path: `office-accepted/native.trx`, MVID `ce19a20c-9708-4c17-b998-f3415b8e6303`. Both adapters returned Verified=true/Uncertain=false; reopen without helper save verified module/class/form content. Normal exits 0; no policy bypass. |
| Q-012 | P1 gate | Access/Publisher have no VBAi host-document save adapter. | Implement and qualify an adapter, or explicitly narrow the release contract for this operation. | OPEN; safe refusal and helper persistence are distinct results. |
| Q-013 | P1 gate | Classic Outlook initially had no configured profile. | Qualify a read-only scenario in an explicitly configured classic profile without modifying mail or production VBA. | CLOSED for read-only startup/metadata: `outlook-accepted/native.trx`, MVID `ce19a20c-9708-4c17-b998-f3415b8e6303`; exact PID, project inventory, scoped debug state and environment passed, normal exit 0. No account configured by automation, no mail read/sent or VBA mutation. |
| Q-014 | P1 gate | SOLIDWORKS versions require independent native acceptance. | Qualify load, UI, disposable module/form operations, compile/debug, persistence and cleanup in each explicitly selected version. | PARTIAL on aaf3: existing Type100 save and saved-copy module/class/form readback passed separately in 2019 SP5 and 2025 (Q-021). Reviewed designer captures show the synthetic form and label without clipping; assistant behavior is outside those frames. The 2025 resize check passed before opening its form. The 2019 resize check after opening its form failed with `Editor did not adapt`, although original placement was restored; the cause remains unproven. Full debugger, focus/navigation and embedded assistant workflows remain open. The user authorized autonomous host close/relaunch and discarding open work; Both owned hosts exited normally with code 0. |
| Q-015 | P2 gate | The initial instrumented suite timed out on post-step observation; its direct relationship to Q-005 is not proven. | Repeat the complete suite on the corrected source and retain failures/skips honestly. | The later full instrumented run completed with a Git long-path failure (Q-019), not an observation timeout. The clean full instrumented rerun after subsequent corrections passed; conditional native/provider skips remain separate gates. |
| Q-016 | P2 | Shift+Tab accepted a composer suggestion instead of allowing backward keyboard navigation. | Leave backward navigation unhandled while preserving plain-Tab suggestion acceptance. | Fixed; focused regression passed. |
| Q-017 | P2 | The French welcome card is clipped in the narrow native chat panel. | All welcome actions remain visible or reachable by normal scrolling, including after resize. | Fixed by measuring the Designer table at its available width; failing-before/passing-after geometry checks and inspected native recapture confirm all actions are visible. |
| Q-018 | P2 | Monaco tests assumed immediate reference refresh, stopped only one of two synchronization timers and accepted an unrelated pending-edit refusal as a save cancellation. | Respect the bounded reference cache; isolate draft synchronization and prove the native save callback actually runs; verify exact native/readback text. | Stronger fixtures passed the final native batch, including exact VBE canonical readback and renderer reconciliation. |
| Q-019 | P2 | Git for Windows rejects an absolute `--git-dir` beyond its internal path buffer, even when `core.longpaths=true`; the full suite exposed this with a deep cache path. | Keep the verified working directory and pass the repository path relatively; prove real local Git operations at the failing path length without global configuration changes. | Fixed with relative --git-dir and process-local core.longpaths; real Git long-cache commit/readback and the originally failing binding scenario passed focused validation. The final complete managed rerun passed. |
| Q-020 | P1 | In SOLIDWORKS 2019, creating a standalone project succeeds but its unsaved `FileName` getter throws, breaking collection readback and reporting an uncertain result. The same assumption prevents initial standalone save detection. | Recognize the narrowly identified unsaved standalone state; preserve refusal for unrelated getter errors and saved projects; validate first save and reload. | Corrected creation, first SaveAs, synthetic module/class/form, compilation, save and close completed on native 2019/353d. Reload did not complete: the first run was interrupted by the user; the second host exited unexpectedly during Open (Q-030). The original getter HRESULT remains NOT_OBSERVED. No macro ran. |
| Q-021 | P1 gate | The native Save command for an existing SOLIDWORKS Type100 SWP can return before the host's Saved flag becomes true, causing premature uncertain results. | Invoke Save once, yield to the owning UI thread, preserve identity/selection/path/source/metadata guards, and independently read persisted module/class/form content in both versions. | CLOSED for the tested existing-SWP Type100 path on `aaf3a555-76d4-4b18-ae09-1e7b3e085934`. Bounded asynchronous verification returned Verified=true/Uncertain=false in 2019 SP5 (PID 47384, revision 27.5.0) and 2025 (PID 1236, revision 33.1.1). Exact saved byte copies opened once through native Edit Macro retained module/class hashes and the form label: `solidworks/type100-2019/async-save-02/saved-copy-content-verification.json` and `solidworks/type100-2025/async-save-01/saved-copy-content-verification.json`. Pending saves are blocked across sessions on the owning thread; timeout or changed state remains uncertain without retry. Type100 SaveAs, signatures and reopening the original file after a full application restart are not qualified. Earlier 1541 uncertain results remain historical failures. The separate 2019 async-save-01 runtime marker succeeded, but its subsequent SWP hash change failed the unload scenario; that result is not reload proof. |
| Q-022 | P1 | Word VBProject.FileName throws before Save and can identify temporary VBA storage afterward; equating it to Document.FullName rejected a correctly matched document. | Use Word's identity-matched Document.FullName while preserving PID, COM identity, format, source, saved flags and file-byte guards; keep PowerPoint project-path checks strict. | CLOSED for the reproduced defect: native diagnostic proved DirectoryNotFoundException/0x80070003 then ~WRL0002.tmp, with unchanged document path/source and saved flags true. Red/green regressions and `office-accepted/native.trx` on ce19a20c prove adapter save/reopen; Q-024 remains separate. |
| Q-023 | P2 | The Outlook qualification fixture omitted the required Project field for debug_state despite successful add-in loading. | Resolve the unique project explicitly and retain normal-exit and metadata assertions. | CLOSED: corrected fixture passed `outlook-accepted/native.trx` on ce19a20c, with exact project metadata and normal owned-process exit 0. |
| Q-024 | P1 gate | Word VBA backing paths are unsuitable for persisted document identity; complete Git capture also fails during native export into GitTemporary. | Resolve the unique Word document by PID/IUnknown and canonical FullName; preserve raw FileName, refuse stale bindings and qualify capture separately. | PATH CONTRACT VERIFIED on 353d: `artifacts/qualification-v1/word-git-native/353-native-03/hosts/Word/2cb46193741e415ab5773c1539745096/qualification.json` proves two projects named Project resolved by their own DOCM paths, native selection, stale SaveAs binding refusal and unchanged source in the other document. Word PID 50144 exited normally. COMPLETE GIT REMAINS OPEN: production Capture fails with 0x800AC35C, while bridge and external-STA exports to separate artifact files succeed with identical bytes. The test remains failed; actual chat/Git UI opening is not established by the bridge selection evidence. No implicit scope/grant migration or Normal modification. |
| Q-025 | P1 gate | Native Excel form fitting could corrupt memory or prevent shutdown. The installed x64 WinForms Com2PropertyDescriptor.SetValue allocates a 16-byte VARIANT buffer, while its marshaler writes 24 bytes; scroll fitting used this setter. | Avoid that native descriptor setter while retaining validation, conversion, readback and existing control restrictions. Qualify both fitting and the generic scalar dispatch with normal host exit. | CORRECTED on candidate 353d for the tested paths. `scalar-excel-native/native.trx` passes complete fitting, arrays and persistence; the independent scroll-only trial exited normally. `native-scalar/native-353d/qualification.json` verifies project Description, module Name with unchanged source, and Label BackColor through native getters; its reviewed VBE capture shows the yellow Label and renamed module, followed by exit 0. Earlier crashes, dumps and forced cleanup remain failed evidence; array/persistence failures are not automatically attributed to this defect. Other controls and complete Git workflows are not thereby qualified. |
| Q-026 | P1 gate | Activated Excel formatting-options scenario refused a stale revision and its restoration also failed. | Identify the revision drift and rerun the instrumented mutation/restoration scenario without relaxing revision checks. | OPEN; cause not proven. The original font and normal-text foreground were restored through the native bridge; complete options revision matched the retained baseline and Excel exited normally. Test diagnostics now retain failed requests, before/after observations, and every distinct restoration failure. Evidence: `artifacts/qualification-v1/excel-options-recovery-restore-02/report.json`. The instrumented ce19 rerun (`excel-options-instrumented/options.trx`) passed the full scenario, complete baseline restoration and normal exit. The earlier drift was not reproduced, so its cause is not claimed fixed; successful-run snapshot attachments were not retained by that VSTest invocation. |
| Q-027 | P1 gate | UserForm export to the production GitTemporary parent fails natively, and raw FRX serialization varies in timestamp/padding bytes. Opaque FRX contents also need bounded preflight checks. | Qualify native capture/import and recovery, preserve real FRM/FRX through GitHub, and reject missing or manifestly invalid companions before mutation. | OPEN. Candidate `ce19a20c`: positive round-trip attempts stopped at fixture dimensions (corrected), then `Capture` / `Export` (`0x800AC373`) before push/import. Differential probes export the same form successfully under TEMP and E: but fail under GitTemporary, including installed-bridge export. Equal-ACL plain/EFS siblings both work; EFS alone and GUID segment length are not sufficient causes. Consecutive unedited exports differ only in identified CFB timestamp/MS-OFORMS padding fields. Owned probes exit normally. Separately, `userform-git-transport-20260929225354/transport.json` proves production Git commit/push/fresh-fetch with exact FRM/FRX hashes on retained branch `qualification-userform-20260929225354-93ed53dc`, main unchanged; this form-only transport is not native project import qualification. Checkpoint/backup remain intact; no permanent loss is established and no corrupt snapshot was published. Separate candidate `353ddf2a` bridge import from the exact Git-fetched FRM/FRX, native controls/both Caption properties, save/reopen and owned-process shutdown are verified in `userform-fetched-import-20260929231432`; source/import/reopen captures were reviewed. VBIDE adds exactly one leading CRLF to code (content preserved, exact source text differs). This does not qualify the blocked Git coordinator or raw FRX comparison. |
| Q-028 | P1 gate | The real Ollama chat UI scenario intermittently failed to observe streamed text while busy. The ce19 diagnostic ended with the visible fallback `No text response.`, without an observed text fragment or refused tool call. | Preserve the streaming, cancellation and next-send assertions and retain exact synthetic wire evidence. Explain the earlier empty response before claiming a reliability correction. | OPEN, intermittent cause unresolved. `ollama-ui-wire-353d/ollama-ui.trx` passes the unchanged strict scenario on 353d: visible text while busy, cancellation and the subsequent visible complete reply. Wire capture records a refused synthetic tool call and subsequent text; no LLM production fix was applied. The passive wrapper preserves transport configuration but changes timing, so this success does not explain the ce19 failures in `activated-ollama-ui` and `ollama-ui-diagnostic`. Detached UI with simulated VBE only; no embedded-host qualification is inferred. |
| Q-029 | P2 | Empty-project assistant state re-enables Send after a prompt edit because `UpdateBudgetControls` omitted the scope predicate. | Keep Send disabled after scope loss, including draft and busy-state changes; preserve active-turn Stop and normal sending with a selected scope. | Reproduced by a real detached control event in `empty-scope-red/red.trx`. Source correction centralizes the predicate while preserving cancellation. Expanded regression and native/UIA acceptance are tracked separately. `EnsureCurrentScope` still guards dispatch; no privacy bypass is demonstrated. |
| Q-030 | P1 gate | SOLIDWORKS 2019 PID 56924 terminated during `VBProjects.Open` of an owned standalone SWP after successful save and removal. | Prevent invoking the unsafe native API and qualify the host-native macro editing workflow separately. | Mitigation verified on native 2019/1541: explicit refusal, unchanged project collection/file and live host (`solidworks-2019-1541-open-refusal.json`). Standalone Open remains unavailable; full functionality is not restored. User confirms the second termination was involuntary; the first session was manually closed and is not crash evidence. No macro ran and no exact native fault mechanism was established. Native Edit Macro separately loaded an owned saved copy. |

Priority describes impact, not proof of exploitability. Findings based only on
source or detached fixtures retain their native validation requirement.

## UI acceptance

Designer construction, detached WinForms/WPF behavior and native VBE hosting are
separate gates. Exercise startup, close/reopen, docking, resizing, scroll/focus,
keyboard navigation, approval dialogs, cancellation, error display, Monaco/native
synchronization and debugging decorations. Inspect captures rather than treating
their existence as a visual pass. Record the actual language, displays and DPI;
do not infer untested DPI/high-contrast/keyboard profiles.

Detached scenarios use synthetic conversations and isolated settings/history.
Native scenarios use owned Office processes and disposable projects. Do not use
global keyboard injection, weaken host trust settings or run production macros.

The installed Excel layout passed automatic Monaco startup, native code-pane
coexistence, resize/restore and Object Browser layout preservation. The assistant
pane passed native VBE close/reopen; capture review separately identified Q-017.
The detached assistant also passed real loopback Ollama streaming, user Stop and
a subsequent completed response. Its VBE fixture was simulated, so this does not
establish an end-to-end model-driven native mutation.

A separate live Ollama-to-Excel scenario read a random marker through the real
project-bound `read_module` tool and returned it without changing the native
source. The marker was absent from prompts. This establishes a model/tool/COM
read workflow, while mutation approval and embedded assistant dispatch remain
separate acceptance requirements.

Detached display placement was observed on both installed displays at 96 DPI.
This covers window bounds on those displays, not higher or mixed DPI. Light/dark
settings, Git, native Markdown, transcript virtualization and keyboard navigation
have separate detached evidence. No system display or accessibility settings were
changed to simulate an untested environment.

## Git qualification

Authenticated read access, local Git conflict/recovery scenarios and real private
remote creation, commits, pushes and fetched snapshot comparisons were exercised
separately. The maintainer explicitly authorized retaining the synthetic private
repository for future tests. No existing repository was modified and no release,
pull request or issue was published. The retained repository and exact commits are
recorded in `artifacts/qualification-v1/git-remote-execution.json`; this is not a
claim that every Git UI/provider path is qualified.

## Performance and library boundaries

Benchmarks call the production implementations with identical synthetic inputs.
They establish local algorithm costs, not end-to-end application speedups.
Persist raw samples, build identities and assertions alongside the reports.
For a 250,000-character buffer with 2,000 edits, median batch time fell from
777.9953 ms to 0.2972 ms and observed generation-2 collections fell from 399 to
zero per measured operation. For a 16,000-fragment SSE response, median parsing
time fell from approximately 6,118 ms to 61 ms. These in-memory microbenchmarks
exclude network latency, model generation, COM and screen rendering.
Dependency vulnerability results describe the configured feed at audit time;
they are not proof that all bundled or externally installed software is safe.

The current assembly version remains `0.1.0.0`; this is a candidate evaluated
against the v1.0.0 requirements, not a versioned v1.0.0 release. No version bump
was performed.

The source-build distribution still needs a maintainer-selected project license
before public release. The standalone installer remains a later milestone;
updater scaffolding does not establish signed installation/rollback acceptance.
No release, license selection, publishing or credential migration is part of
this qualification run.

## Evidence location

Machine-local TRX, logs, screenshots, manifests and detailed delegated reviews
are under `artifacts/qualification-v1/`; detached UI captures are under
`artifacts/ui-review/screenshots/`. These ignored files are not public downloads.
The maintained [compatibility matrix](compatibility.md) records operation-specific
host outcomes; [testing](../tests/README.md) supplies the standard commands.

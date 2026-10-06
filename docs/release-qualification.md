# Release qualification

[Documentation](README.md)

**Release-wide acceptance is not established.** The selected operation matrices
below have scoped acceptance on their recorded binaries. Combining results from
different candidates does not qualify the latest integrated build.

## Current acceptance summary

| Scope | Decision | Candidate/evidence |
| --- | --- | --- |
| Ollama and embedded Office assistant (Q028) | Accepted for the selected CPU/model/sampling profile and six Office hosts | `85486c00`; [recorded result](test-coverage.md#q028-ollama-and-embedded-office-assistant-qualification-2026-10-06) |
| Word Git and Excel UserForms (Q024/Q027) | Accepted for the complete prepared operation matrix on the real desktop | `c593d6af`; [recorded result](test-coverage.md#q024-word-and-q027-userform-scoped-qualification-2026-10-06) |
| Access/Publisher adapters (Q012) | Accepted existing-document persistence and explicit metadata contract | Frozen `2106fd95` product; [recorded result](test-coverage.md#office-adapter-persistence-q012-2026-10-04) |
| SOLIDWORKS core/editor/local assistant (Q014) | Accepted selected 2019/2025 workflows on the recorded private desktop | Frozen `ddf638b2` product; [recorded result](test-coverage.md#solidworks-native-core-and-assistant-q014-2026-10-04) |
| SOLIDWORKS native creation/publication/reopen (Q020/Q030) | Accepted selected 2019/2025 main-desktop workflows; unsafe generic Open refused | Frozen `08689325` product; [recorded result](test-coverage.md#q-020-and-q-030-native-macro-candidate-2026-10-05) |
| Excel options (Q026) | Accepted selected current-behavior matrix under the maintainer's criterion | `08d7420`; [native boundaries](test-coverage.md#native-options-and-debugger-boundaries) |

Exact MVIDs, hashes, test counts and local proof roots are maintained once in
[recorded validation](test-coverage.md). The original complete-empty Ollama cause,
historical break-mode crashes and other excluded behavior remain unresolved;
scoped acceptance does not claim their causes have been repaired.

## Release criteria

Before claiming a release, identify one integrated candidate and verify applicable
managed, JavaScript, native-renderer, Designer, authenticated-provider and native
host scopes. Record skipped, blocked and unsupported operations explicitly.
The complete managed line/branch target still requires a current measurement.

Native acceptance requires loaded assembly identity, exact disposable project,
operation oracles, original normal process exits, restoration and persistence
readback where relevant. Refusing an unsupported operation can pass a refusal
scenario without providing that capability. Unknown outcomes prohibit retries
of native mutations. Trust policies and existing production documents are preserved.

The agreed host scope includes installed Microsoft 365 x64, classic Outlook and
the independently selected SOLIDWORKS versions. Visio, Project, x86 and other
Office/application builds have no inferred acceptance.

## Finding register

The register preserves Q-001 through Q-030, their acceptance criteria and known
limits. Historical status wording applies to its specified candidate. The summary
above and recorded results are the current operation-specific decisions.

| ID | Finding | Recorded decision | Acceptance boundary |
| --- | --- | --- | --- |
| Q-001 | Execution project identity | Native scenario + guarded dispatch | Revalidate selected project/mode/policy on the owning STA; user-selection races are not atomic. |
| Q-002 | Duplicate Git component names | Regression passed | Reject case-insensitive duplicates across component types before import/deletion. |
| Q-003 | Atomic editor batches | Regression + benchmark passed | Validate final UTF-16 size and preserve tied insertion order. |
| Q-004 | Bounded provider reception | Regression + local provider checks passed | Bound input bytes; preserve cancellation and complete tool arguments. |
| Q-005 | Post-step observation | Deterministic regression passed | Retain observation until command ownership is released; do not replay native commands. |
| Q-006 | Debugger/persistence harness | Selected Excel scope accepted; broader scope open | Declared pages/arrays and product save/reopen pass; trace-cap failure and historical crashes remain unresolved. |
| Q-007 | COM temporary shutdown | Normal-exit checks passed | Release owned references; distinguish crash/forced cleanup from normal exit. |
| Q-008 | Current Markdown UI tests | Detached checks passed | Exercise current rendering/streaming/virtualization APIs; native hosting remains separate. |
| Q-009 | Third-party runtime notices | Payload hashes verified | Deliver upstream license/notice texts with provenance. |
| Q-010 | WebView profile lifecycle | Selected lifecycle accepted | Require editor retirement and matching BrowserProcessExited; retain unknown/locked/legacy profiles. |
| Q-011 | Word/PowerPoint Save adapters | Selected save/reopen accepted | Verify pending sources/form state through adapter Save and independent readback; preserve older failures. |
| Q-012 | Access/Publisher Save adapters | Existing-document contract accepted | One adapter Save, fresh readback and normal exit; NewDocument and unsupported Unicode persistence excluded. |
| Q-013 | Classic Outlook startup | Read-only metadata accepted | Configured profile, explicit project, normal exit; preserve personal VBA and mail. |
| Q-014 | SOLIDWORKS versions | Selected 2019/2025 workflows accepted | Native core, scoped UI/local assistant and normal cleanup on the recorded candidate. |
| Q-015 | Complete managed gate | Recorded candidate accepted | Preserve failures/skips and measured scope; a later binary needs its own complete run. |
| Q-016 | Composer backward navigation | Regression passed | Shift+Tab preserves backward navigation; plain Tab retains suggestion acceptance. |
| Q-017 | Narrow welcome layout | Geometry + native recapture passed | Welcome actions remain visible or reachable after resize. |
| Q-018 | Monaco native assumptions | Selected contracts corrected | Retain reference refresh timing, both synchronization timers and normal-exit assertions. |
| Q-019 | Git long cache paths | Real Git regression passed | Use verified working directory, relative git-dir and process-local long-path configuration. |
| Q-020 | SOLIDWORKS unsaved macro identity | Selected creation/publication accepted | Fresh native Type100 identity, preserved draft and original-file reopen; cold designer requires explicit opening. |
| Q-021 | SOLIDWORKS asynchronous Save | Selected existing Type100 path accepted | One Save; yield/verify guarded state and persisted readback. Timeout remains uncertain without retry. |
| Q-022 | Word temporary VBProject path | Reproduced defect corrected | PID/COM-matched Document.FullName; preserve format, source and file guards. |
| Q-023 | Outlook fixture project selection | Read-only native fixture passed | Resolve a unique Project before debug-state inspection; require normal exit. |
| Q-024 | Word Git/document identity | Selected real-desktop matrix accepted | Canonical same-name isolation, stale SaveAs refusal and installed owner/chat-to-Git; templates/arbitrary SaveAs excluded. |
| Q-025 | Native UserForm scalar setter | Selected Excel scalar/fitting paths corrected | Avoid the x64 descriptor VARIANT overwrite; retain typed validation/readback and normal exit. |
| Q-026 | Native options metadata drift | Maintainer-approved selected matrix accepted | Exact Size mutation and full restoration; historical trigger remains unexplained. |
| Q-027 | UserForm import/recovery/resources | Prepared Excel matrix accepted | Raw FRM/FRX, typed native state, fonts, independent persistence and actual designer review; no universal control claim. |
| Q-028 | Ollama streaming/tools/UI | Selected CPU/model/Office matrix accepted | Strict tools/marker/cancel/recovery/wire/lifecycle; historical complete-empty cause remains unresolved. |
| Q-029 | Send after scope loss | Detached real-control regression passed | Keep Send disabled without scope; preserve active-turn Stop and final dispatch guards. |
| Q-030 | Unsafe SOLIDWORKS Open | Mitigation + selected native workflow accepted | Refuse generic VBProjects.Open before mutation; use verified native Edit Macro reopen. |

## Review and publication

Use [the testing guide](../tests/README.md) to select the applicable scope.
Publish the candidate, test changes and maintained documentation together; do not
replace native evidence with detached screenshots or protocol mocks.
Distribution also requires a maintainer license decision and the later signed
installer/release-payload milestone described in [updates](updates.md).

Superseded checkpoints and investigations remain in Git history. Local artifacts
are retained evidence, not a public package or a license grant.

# Test coverage and qualifications

[Documentation](README.md)

VBAi is tested through managed regressions, editor/UI checks, provider checks and
operation-specific qualifications in real host applications. This page summarizes
all Q001–Q030 findings, including earlier qualifications, rather than only the
latest campaigns.

## Qualification register

**Passed** refers to the recorded test scope. **Accepted** refers to the declared
native workflow and tested candidate. **Partial** means wider behavior remains
unqualified. These results are historical, candidate-specific evidence; combining
passing rows does not qualify the current integrated build or an entire host.

| ID | Area | Recorded result and main limit |
| --- | --- | --- |
| Q001 | Execution project identity | Verified in selected native scenarios; selection races remain a limitation. |
| Q002 | Duplicate Git component names | Passed duplicate-name rejection regressions. |
| Q003 | Atomic editor batches | Passed atomic-edit regressions and the recorded benchmark. |
| Q004 | Bounded provider reception | Passed bounded-input and local-provider checks. |
| Q005 | Post-step observation | Passed deterministic post-command observation regressions. |
| Q006 | Debugger/persistence harness | Partial: selected Excel debugging and save/reopen paths accepted; broader debugger scope remains open. |
| Q007 | COM temporary shutdown | Passed normal-exit checks for owned temporary COM sessions. |
| Q008 | Current Markdown UI tests | Passed detached rendering checks; native hosting is qualified separately. |
| Q009 | Third-party runtime notices | Bundled upstream license and notice payload hashes verified. |
| Q010 | WebView profile lifecycle | Selected WebView shutdown/profile-retirement lifecycle accepted. |
| Q011 | Word/PowerPoint Save adapters | Selected Word and PowerPoint save/reopen workflows accepted. |
| Q012 | Access/Publisher Save adapters | Existing Access/Publisher documents accepted; new-document and unsupported Unicode paths excluded. |
| Q013 | Classic Outlook startup | Read-only classic Outlook startup/metadata scope accepted. |
| Q014 | SOLIDWORKS versions | Selected SOLIDWORKS 2019/2025 editor and assistant workflows accepted. |
| Q015 | Complete managed gate | Complete managed gate accepted for the recorded historical candidate. |
| Q016 | Composer backward navigation | Passed backward keyboard-navigation regression. |
| Q017 | Narrow welcome layout | Passed narrow-layout checks and native recapture. |
| Q018 | Monaco native assumptions | Selected reference-refresh, synchronization and shutdown contracts corrected. |
| Q019 | Git long cache paths | Passed real-Git long-cache-path regression. |
| Q020 | SOLIDWORKS unsaved macro identity | Selected native SOLIDWORKS creation/publication accepted; cold designers must be opened explicitly. |
| Q021 | SOLIDWORKS asynchronous Save | Selected existing SOLIDWORKS Type100 save path accepted. |
| Q022 | Word temporary VBProject path | Reproduced Word temporary-project identity defect corrected. |
| Q023 | Outlook fixture project selection | Passed read-only native Outlook fixture selection. |
| Q024 | Word Git/document identity | Prepared Word Git/document-identity matrix accepted on the real desktop; arbitrary templates/SaveAs excluded. |
| Q025 | Native UserForm scalar setter | Selected Excel UserForm scalar-property and content-fitting paths corrected. |
| Q026 | Native options metadata drift | Selected Excel options/restoration matrix accepted; original drift trigger remains unexplained. |
| Q027 | UserForm import/recovery/resources | Prepared Excel form import/recovery/resource matrix accepted; not a claim for every third-party control. |
| Q028 | Ollama streaming/tools/UI | Selected Ollama model/profile and six Office hosts accepted; historical empty-response cause remains unresolved. |
| Q029 | Send after scope loss | Passed scope-loss Send/Stop regression on actual detached controls. |
| Q030 | Unsafe SOLIDWORKS Open | Unsafe generic SOLIDWORKS Open refused; selected native reopen workflows accepted. |

SOLIDWORKS evidence covers the selected 2019/2025 versions. Office evidence covers
recorded Microsoft 365 x64 hosts, including classic Outlook; it does not extend
to every Office version, Visio, Project or x86. The frozen candidates for the
principal native campaigns are identified in the
[release acceptance summary](release-qualification.md#current-acceptance-summary).
The [native macro evidence index](qualification/q020-q030.json) retains the
machine-readable Q020/Q030 candidate references.

## Managed line and branch coverage

The target remains **100% lines and 100% branches**, with meaningful mirrored tests.
There is no current complete measurement for the integrated `main` build.
Native qualifications and focused passing tests are not coverage measurements.

| Recorded measurement | Source | Lines | Branch outcomes |
| --- | --- | --- | --- |
| Historical complete managed run, 2026-10-01 | `8f2315d04162f55b0956618f96b59d294a3fb681` | 33,562 / 33,755 (99.43%) | 33,975 / 34,487 (98.52%) |

That run recorded 2,324 passed tests, 0 failures and 90 conditional skips. It
measured instrumented managed code, excluding external hosts, JavaScript and
native C++; its percentages are not current release statistics.

## Recent interface and documentation checks

Licensing changes from source `f64c1a12` passed 115 focused runtime, support-module
and coverage-planning tests with no failures or skips. Both application outputs
carried the exact declared license texts; upstream notices remained unchanged.
This was a host-free gate, not a new native qualification or coverage measurement.

The localized-help candidate from source
`43e24549bc7ef4563987fad0afc2b1ce1e8be894` passed 32 focused managed tests, 19
help-builder tests and 21 documentation-checker self-tests. All 13 manually
translated/reviewed CHM editions were compiled: 299 chapters loaded through the
Windows renderer, 533 extracted source/assets matched their inputs, and every
archive contained a full-text index. Native menu loading and standalone viewer
search were not exercised.

The French worked example uses a real disposable Excel macro and separate
implementation/test requests. A fresh-host run passed six tests; a later rerun
returned one unknown outcome and five blocked tests. Repeat-execution acceptance
remains open; a successful earlier run does not erase that result. The committed
[screenshot manifest](help/fr-FR/screenshots/manifest.json) records the real captures.

## Reproduction and evidence

Use the [testing guide](../tests/README.md) to select a test family and
[release qualification](release-qualification.md) for acceptance rules. Local
proof sets retain candidate identities, original test reports and process/recovery
receipts. Full campaign histories are kept in Git history and the corresponding
issues/pull requests, rather than repeated here.

Some original worked-example receipts disappeared when its documentation worktree
was removed concurrently. The retained screenshots do not reconstruct those
receipts. Missing evidence must be reported as missing, and unknown native
mutations must not be replayed as though they had failed safely.

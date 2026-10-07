# Test coverage and qualifications

[Documentation](README.md)

VBAi is tested through managed regressions, editor/UI checks, provider checks and
operation-specific qualifications in real host applications. This page summarizes
all Q001â€“Q030 findings, including earlier qualifications, rather than only the
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

## Installer lifecycle

The maintainer stopped the general campaign and separately authorized installer
qualification on 2026-10-07. After both reviewed documentation branches were
merged, the unsigned 1.0.0 setup was rebuilt from source
`e88528dc81668bb6b7270da262c64bfde25695e4`. That final setup, with SHA-256
`3cb9e1fade176f701fbdae39a0c96f02f6fd11e5e2149457630c1b91ff112c4a`,
passed four real current-user scenarios: fresh installation, same-version repair
with preserved installation identity, uninstallation with preserved user data,
and reinstallation with a new identity. The installed payload hashes, COM paths,
Windows Installed apps entry, marker and unsigned uninstaller were checked.

The reusable harness is `tests/Installer/Invoke-InstallerLifecycle.ps1`; it requires
explicit deployment opt-in and never launches Office or SOLIDWORKS. This result
does not qualify upgrade to a different version, failure recovery, a new UAC
bootstrap on another machine or native-host loading of the installed candidate.
The remaining general/native test campaign was not resumed.

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

## Excel audit correction candidate

The tested working-tree snapshot was based on
`a67873a75ccf678eedd3072aeadd1a859dd1841f`, on `codex/excel-audit-fixes`.
Its frozen build19 product has MVID `99a4a849-b2bb-4555-b92b-5d9aeccd2588`
and SHA-256
`EDFC2B63922DD43C344DF4F8E8B58A00D9A252C6DCD1F9536C319881F929BD96`.
The test MVID is `4397b2fc-e287-452a-8f8d-645cf073d326`; the source snapshot
manifest SHA-256 is
`05D63EDDFEB35CDE051F5FB38D1D0F5BD9C5F06DB1BE672C19B6B04AB8F35571`.
This identifies the actual working-tree candidate without presenting its base
commit as the complete corrected source. Evidence is retained under
`artifacts/excel-audit-fixes-20261007`.

Compilation completed without warnings or errors. The build19 focused gate passed
**177 tests, no failures or skips**; its TRX SHA-256 is
`BEC82CE539C455D235E2EB272AE3B8506411DA2A1BF8257E076AEC0984B6A5D4`.
It targets chat state/activity layout contracts and the selected Git diagnostic
case. The 124 source inputs, 229 frozen files and 180 TEMP files remained
unchanged. Its broader gate passed **5,899 tests, no failures or skips**; its TRX
SHA-256 is
`53232AEEA9EC3CDF71C55C896A085BE77150046FEE95FF71D9DD2DFD7E7AF595`.
The same inventories remained unchanged. The C12 native chat scenario also passed
on this exact build, including the visible collapsed failure status. Build19 differs
from build18 only in the activity-group title and its matching tests: terminal
outcomes now precede a long contextual caption, so end ellipsis cannot hide the
failure badge in a collapsed group. The change was prompted by the actual C11
native capture; C11 is not promoted to full UI acceptance.

The preceding build18 managed gate passed **494 focused tests** and **5,897
broader Unit tests**, with no failures or skips. Its product MVID was
`3821c6c8-741a-4483-91f5-8839d74fd805`; its focused and broader TRX SHA-256 values
were respectively
`8BA6D6BFF5B70F519AA11A72F6DDD0D5A257497BEA4348FBA8DDCDBFE01830AD`
and `5093AC7CF654C06339A39C12E9310D5F2AA19E4842749C8CE63376CCF54072B3`.
The 124 source inputs, 229 frozen files and 180 TEMP files remained unchanged.
Both selected Git-case journals were complete and closed. Live Ollama is
explicitly excluded; managed provider transports are synthetic. No current
line or branch coverage percentage was measured.

The preceding broader gate failed one simulated Git branch-creation case,
reporting changed state immediately after status. Its failed TRX remains retained
with SHA-256
`CA8A427EFE9ECA56A4AC09828CD45F0BDE8230F2817A08AE37D0CF1A10B7EFD2`.
Its cause is not established. Build18 adds an optional, hash-only revision
observer and a test-only command journal to capture the operands of that case.
The observer is null by default and does not change the production admission
verdict. The focused case produced a complete journal: all status/admission
operand pairs matched, and only expected absent recovery/remote references
returned nonzero. Instrumentation adds scheduling overhead, so this passing
observation does not establish a fix for the earlier failure.

The shared correction gates exercise provider/profile capture, settings publication,
public reasoning summaries, contextual tool activities, late callback isolation,
owner-STA save confirmation, Git recovery, native-font transfer, compilation
capability and VBE dispatch guards. Detached controls were also inspected in
light and dark themes. The preceding passive Codex trace correction attributes
the schema-canonical `Project` and `Module` arguments; dispatch and target checks
remain unchanged. C12 accepted the declared build19 chat workflow and its narrow
collapsed failure presentation; it does not qualify every layout or provider.

Native evidence remains candidate-specific:

| Scope | Observed evidence and limit |
| --- | --- |
| N16 Frame/MultiPage, build16 MVID `b267bac7-2c7b-49a3-b918-468925192919` | One native qualification passed with no failures/skips: checkpoint restoration, controlled interruption, explicit rollback and save/reopen. Loaded identity was attested; root font 8.25 points, fractional Frame font 8.27 points, hierarchy, code and pictures were preserved. The original Excel handle exited normally with code zero and registration was restored exactly. The comparator checks complete logical FRX resources; existing recognized serialization padding may differ, so raw FRX hashes are not asserted identical. This covers one synthetic layout. |
| E09 Monaco, build16 | Single integrated Save, exact source, host Saved flag and exact owned-draft retirement passed. The original bootstrap failed its closure deadline; its failed verdict and unobserved Excel exit code remain. After the worker exited and Excel disappeared, registration was restored exactly. A separate reopen under restored installed registration confirmed persisted code, without Save or macro execution, and observed normal original-handle exit. This does not qualify a build16 reopen dispatch. |
| C12 native Codex chat, build19 | Loaded MVID/path/PID attested; visible streaming while busy, one Stop, interruption acknowledgement and exact resumed response passed. Exactly two native reads returned the expected success and missing-module failure, with closed correlated receipts. Root inspected four real captures: contextual action headings were readable and the collapsed failure symbol/status preceded the clipped module context. All three source hashes and the final sole-owned-workbook identity were preserved. The original Excel handle exited normally with code zero; registration was restored exactly and installed files/settings stayed unchanged. No image was taken before Stop and no provider-supplied public reasoning summary was established. Evidence: `nativechat12/root-native-review.json`, summary SHA-256 `BC40E271FF8EF9A498A53E230901BBACBD42C140E1FF28128D13E787210B6949`. |
| C11 native Codex chat, build18 | Loaded identity attested; visible streaming was observed without an image before Stop, followed by one Stop, visible interruption acknowledgement and exact resumed response. Exactly two native module reads returned the expected success and missing-module failure, with correlated closed receipts. All source hashes stayed unchanged; the final sole-workbook check passed, the original Excel handle exited normally with code zero and registration was restored. Root-reviewed captures nevertheless showed the failed activity suffix clipped after a long collapsed caption: core execution passed, UI presentation acceptance failed. An ordinary assistant planning comment was visible; no public reasoning summary is inferred. |
| C08 native Codex chat, build16 | Visible streaming, one Stop, visible interruption acknowledgement and resumed response passed. One actual read returned an unprompted module marker, but the old passive trace misclassified its canonical arguments as `Other`; the case failed before the missing-module action. Separate normal owned cleanup observed exit code zero and exact registration restoration. No failed-activity or public-summary acceptance is inferred from that partial run. |
| Image restoration, build06 MVID `f3371378-4be6-4f62-9e7c-46247d36e736` | Native restoration, explicit rollback, exact resource/font readback and save/reopen passed on the declared synthetic fixture. |
| Test explorer, build08b MVID `e97e90cd-a20a-40f6-9963-cdd03b3d2742` | In-process assembly/path/PID/bitness attestation, support installation, bridge batch, one native UI batch, stale-revision refusal, source preservation and human/LLM reports passed. A real capture showed green success marks and red failures; reopening reused the HWND. One of three production procedures was entered on a coverage copy; no VBA line/branch coverage is claimed. |
| Compilation and signature, build09 | Compilation availability/execution observations were positive, but the original encoding-sensitive oracle failed. Signature dialog read/Cancel preserved unsigned status and exited normally; certificate application and persistence remain unqualified. |

The earlier timeout, strict font refusals, chat harness failures and Monaco Save
failure receipts remain unchanged. Unknown native mutations were not replayed.
An additional workbook of unknown origin was preserved and Excel was closed by
the user. Harness-only failures do not establish product defects; later passing
scopes do not erase original failed verdicts.

These rows do not qualify all Excel behavior or every provider, UserForm,
keyboard/DPI configuration or trust-policy path. Native settings-policy refusal,
Save failure/cancellation and signing persistence remain open. Provider-supplied
public reasoning summaries were not observed in the native chat scenario. The installed payload and personal
settings remain unchanged; candidate registration is temporary and restored
between trials.

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

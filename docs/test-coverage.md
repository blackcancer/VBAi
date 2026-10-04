# Recorded validation

## Office testing subsystem qualification (2026-10-01)

These trials reuse the frozen product built from
`da664f06f8797faaddec0624b7ed3bb10312963b`: MVID
`a799ad84-87a3-4518-9550-3cdda504ca98`, SHA-256
`871104948963C208844F1F62E2C86CAF59F5D5F50CD30933DB7A31C27C140DB4`.
Later changes through `b13ee55` affect qualification fixtures and documentation
only. No production source or product binary changed after the managed coverage
measurement in the Word checkpoint below. Native identity records attest this
candidate in each owned process; fixture assemblies are identified separately.

| Host and exact scope | Actual result | Procedure coverage | Evidence root under `artifacts/qualification-followup` |
| --- | --- | --- | --- |
| Excel, disposable XLSM, `TestSubsystemOnly` | `PASS_WITH_EXPLICIT_SCOPE`; registered batch, actual native explorer actions, reports, stale refusal and normal exit code 0 | Complete 1/3 (33.3333%); original source/probe mapping and event state preserved | `excel-final-testing-only/7eeca1219151492b9ab0a0ac757f0fb5` |
| Word, disposable DOCM | Passed; batch, single, reports, explorer, stale refusal and normal exit code 0 | Complete 1/2 (50%); original preserved, copy closed | Word checkpoint below |
| PowerPoint, disposable PPTM | 1 passed, 0 failed/skipped; batch, single, reports, explorer, stale refusal and normal exit code 0 | Complete 1/2 (50%); original source/revision/test IDs/disk bytes and counters preserved, copy closed | `native-powerpoint-verified-sources/PowerPoint/8f863565eac44b378ec1afff2fc14857` |
| Access, disposable ACCDB | 1 passed, 0 failed/skipped; batch, single, reports, explorer, stale refusal and normal exit code 0; no manual intervention | Unavailable for this host; no measurement claimed | `native-access-autonomous-discard/Access/17b5fd9d859e4f93a801df62e33785b1` |
| Publisher, disposable PUB | 1 passed, 0 failed/skipped; batch, single, reports, explorer, stale refusal and normal exit code 0 | Unavailable for this host; no measurement claimed | `native-publisher-reviewed-discard/Publisher/5cea70175149443f9a8eef078a6ae4bb` |
| Outlook, initially absent personal OTM | 1 passed, 0 failed/skipped; batch, single, reports and stale refusal; original project restored and normal exit verified | Unavailable for this host; no measurement claimed | `outlook-final-testing-only` |

The deliberate batch verdicts are one pass, two failures and one runtime error;
qualification requires these exact outcomes. Both readable and compact JSON
reports agree with the native results. The Excel screenshot attests the global
dark style, green check marks, red crosses and a 1253 x 638 explorer. Its native UI
batch uses the actual owned window and buttons. Other Office rows verify a live
owned explorer window; they do not establish every UI action in each host. Outlook
has no native explorer-action proof in this row. These bridge/UI scenarios do not
exercise a live LLM conversation; the managed testing-tool permission/revision
boundary is covered by the separate managed gate below.

Excel's denominator includes `Workbook_BeforeSave`. Its positive save control is
explicitly `NOT_TESTED_BY_SCOPE`; the event counter/probe observations do not
qualify handler suppression. This testing-only pass does not exercise other VBAi
application workflows. VBA statement and branch measurement remain unavailable.

PowerPoint's security notice was enabled once for the exact owned synthetic copy.
`SaveCopyAs` changed its VBA binary and package metadata; strict offline extraction
verified all three complete module sources against the original and reviewed
fixture before activation. Other package parts matched, and file hashes were
rechecked immediately before the unique UI invocation. The local reader was
[oletools 0.60.2](https://github.com/decalage2/oletools/wiki/olevba), installed only
under ignored qualification artifacts. No global trust setting changed.

Access records one bounded Cancel of the naming dialog for its successfully
created `VBAiOfficeModule`, class `RichEdit20W`, with zero save entries. The dialog
worker verifies original process/start/handle, exact controls/name/thread and
terminal native work; an uncertain cancellation forbids replay or Quit. Publisher
revalidates the complete reviewed source inventory, sole publication, original
process/window ownership and bridge project association before its existing single
Quit. It discards only this row's synthetic changes, invokes no final Save or Saved
setter, and makes no Publisher persistence claim. Other fixtures retain their
existing saved-project guard.

The combined final fixture regression batch reports **33 passed, 0 failed,
0 skipped** in `publisher-test-cleanup-managed/publisher-test-cleanup-managed-final.trx`.
Its tests-only output is `artifacts/build-publisher-test-cleanup`: test MVID
`8d4af179-7ea1-4c7d-88c4-0d83b3d44a99`, SHA-256
`7C67263B71AB1DB7A5E8D1388057DBE48B3A25F2B5F3A918B75E308A82ABF906`.
The Access native row used the earlier tests-only output
`artifacts/build-access-discard-richedit`, MVID
`8337f6cd-ba05-425a-841c-8fe6b03b6b0f`, SHA-256
`AA187800A057A80ADB08895B149777A241EC54632A2E83703D18DC146D55AC6B`.
Both outputs contain hash-verified copies of the same frozen product; no product
rebuild, coverage exclusion or empty regression was introduced.

Earlier PowerPoint security/deadline refusals and the unsupported Publisher fixture
API failure remain failed evidence, even after authorized normal disposal. The
initial Access pass records manual cancellation separately; only the later row
above proves autonomous cleanup. Temporary HKCU/Registry64 registration was restored
and verified after the owned trials, and the installed DLL hash stayed unchanged.
The machine-readable campaign ledger is
`artifacts/qualification-followup/testing-subsystem-final-report.json`.

SOLIDWORKS acceptance on this candidate remains pending explicit selection of a
preloaded instance and disposable SWP. An older candidate's pass is not transferred
to this binary. Visio and Microsoft Project are outside the maintainer-selected
scope; no Ollama qualification was run here. The global test-layout script stops
at the pre-existing `Unit/Bridge/BridgeServer.PathVisibility.Tests.cs` mirror whose
production counterpart is absent. That unrelated layout failure is retained and
was not repaired as part of these testing-subsystem trials.

## Word qualification follow-up (2026-10-01)

The code and test sources at `da664f06f8797faaddec0624b7ed3bb10312963b`
include the typed Word returned-value contract, balanced application/document
leases and copy ownership, queued coverage preparation, and one exact native
focus recovery before arming. Project, source, mode, permission and selection
checks remain mandatory; dispatch is never retried after an uncertain outcome.

| Check | Actual result | Evidence |
| --- | --- | --- |
| Isolated .NET Framework 4.8/x64 solution build | Passed, no errors; four NU1900 vulnerability-endpoint warnings | `artifacts/build-word-exit15` |
| Complete managed VBA testing scope | 537 passed, 0 failed, 0 skipped | `artifacts/coverage/word-exit15-managed/final.trx` |
| Managed executable lines | **3952/3952 (100%)** | Companion Coverlet JSON and Cobertura |
| All managed IL branches | **3792/3792 (100%)**, including 42/42 unmapped branches | `artifacts/coverage/word-exit15-managed/gate.json` |
| Owned-shutdown fixture regressions | 20 passed, 0 failed, 0 skipped | `artifacts/qualification-followup/word-exit15-shutdown/shutdown.trx` |
| Registered Word DOCM scenario | 1 passed, 0 failed, 0 skipped; batch and single outcomes verified | `artifacts/qualification-followup/native-word-exit15-trx/word/word.trx` |
| Word copy-based procedure measurement | Complete **50% (1/2)**; readable/compact reports match; original preserved and copy closed | `artifacts/qualification-followup/native-word-exit15/Word/0eea37bde244427c8e9492ca587be5ba` |
| Owned Word shutdown | One Quit; original handle observed exit code **0** after **6602 ms**, within the explicit 15000-ms fixture bound; no diagnostic collection or forced termination | `shutdown-lifecycle.json` in the same root |

The managed gate includes every `src/VBAi/Testing/*.cs` file and
`src/VBAi/Llm/Chat/LlmVbeTools.Testing.cs`, including Designers and native adapters.
It reconciles both reports from collection
`8e1128bc-82e6-4fad-a818-db138b09c16a` without excluding source, generated classes
or IL branches. Interface-only declarations have no executable sequence points.
The source manifest, post-collector assembly identities, readable `gate.txt` and
machine-readable `gate.json` are retained beside the reports.

The production DLL has MVID `a799ad84-87a3-4518-9550-3cdda504ca98`, SHA-256
`871104948963C208844F1F62E2C86CAF59F5D5F50CD30933DB7A31C27C140DB4`.
The test DLL has MVID `807da707-3d56-475a-a2e1-09fc98043f89`, SHA-256
`A11E66673A1C3F4E6844EDCB1227C40CCC9F5556FD2C613416E17C1D38EB8918`.
Native identity evidence verifies this production DLL and the registered
`VBAi.TestRuntime` callback in the exact owned x64 Word process.

The native trial used Word executable version `16.0.20430.20092` and PID `91900`.
The expected successful test, Boolean failure, swallowed assertion and runtime
error were retained in the batch; single and measured runs passed. Exact live
module code, project/reference revision, test IDs and original disk bytes stayed
unchanged through measurement. Original counters stayed `2,0`; only the copy
entered the selected production procedure. The retained copy/plan are under
`%LOCALAPPDATA%/VBAi/CoverageRuns/fbcd7626b67b4e5cb73c98612f455e21`.
The native explorer measured 1253 x 638 pixels. A later synthetic source edit
refused a stale dispatch and marked the earlier report historical.

The copy's one security activation was limited to the exact owned synthetic
file with matching source/copy hashes. The fixture restored its per-application
security setting. No global trust setting changed. The optional fifteen-second
exit bound affects only the fixture's observation after known Quit; ordinary
fixture observation remains five seconds. `VBAi_TEST_WORD_SETTLED_SCOPE_GC` was
unset and no collection diagnostic ran in this successful trial.

The earlier `7474e4e` trial completed all functional assertions, including measured
coverage and original preservation, but failed the five-second exit observation.
That failure remains preserved under `native-word-focus-recovery`; no Quit or
macro was replayed. A separate `b0fe41a` trial also passed every functional assertion but failed that
same exit gate after testhost-only collection. That experiment did not resolve
the exit delay and does not prove an RCW leak. The fifteen-second observation
trial disables that collection; it does not change product shutdown or repeat
Close/Quit.

The installed DLL remains unchanged. Temporary HKCU/Registry64 candidate
registration is backed up and restored with verification after the owned trial.
No production macro, Ollama scenario, Visio or Microsoft Project qualification
belongs to this follow-up. These observations do not qualify other Word formats,
templates, event-handler combinations or unrelated application operations.

## Ollama pull request after main synchronization (2026-10-01)

Source `ccba639f151eaab238f6982d888922df930ab3d5` merges main `415e16f`
into the Ollama branch. The report conflict preserves both independent
validation records below. The isolated complete solution build succeeds with
no warnings or errors. One focused batch reports **212 passed / 0 failed /
0 skipped**, covering the Ollama wire/model/profile helpers, sampling settings,
HTTP chat client, settings persistence, provider view and localization tests.
Product SHA-256 is
`1F785E7B7C2484FEFC5152E31F69AF333E93074274E9D5944FCB1AE305C9C76C`;
test assembly SHA-256 is
`14F25342792DDF2B2231831A27B95438D66D861DF23A509EBA30A72ED6688A4C`.
Evidence is `artifacts/qualification-v1/followup-20260930/ollama-pr-review/ollama-pr-focused.trx`.
This focused merge verification uses no backend, native host or coverage
collector. The earlier complete and real-provider acceptance remains specific
to candidate `bbb6e6f`; it is not reapplied to the merged binary. The installed
product remains unchanged. A documentation-only publication commit follows.

## Managed VBA testing coverage gate (2026-10-01)

Source revision `51454c61acaef02305ddc28ba6a6041b9532124f` on
`codex/vba-testing-coverage` was compiled and measured in an isolated
.NET Framework 4.8/x64 output. This checkpoint covers the C# VBA testing
subsystem: every `src/VBAi/Testing/*.cs` file and
`src/VBAi/Llm/Chat/LlmVbeTools.Testing.cs`, including Designers, native adapters,
UI handlers, deferred execution, registry verification and the nine LLM tools.
The shared application foundation outside those files is outside this percentage.

| Check | Actual result | Evidence |
| --- | --- | --- |
| Isolated managed build | Passed, no errors; three NU1900 vulnerability-endpoint warnings | `artifacts/build-vba-testing-coverage-final` |
| Focused managed testing/LLM unit classes | 461 passed, 0 failed, 0 skipped | `artifacts/coverage/vba-testing-final/final.trx` |
| Managed executable lines | **3596/3596 (100%)** | Companion Coverlet JSON and Cobertura |
| All managed IL branches | **3608/3608 (100%)** | Raw JSON branch records, including generated classes and async state machines |
| Branches without a mapped source line | 42/42 covered; included in the branch total above | Raw JSON and strict gate reconciliation |
| Coverage gate regressions | 22 passed | `tools/testing-explorer/test_check_managed_coverage.py` |
| Documentation checks | Passed | `python tools/docs/check_docs.py` |

Coverlet collector 6.0.4 instrumented the complete managed assembly with
`--collect:"XPlat Code Coverage"` and
`DataCollectionRunSettings.DataCollectors.DataCollector.Configuration.Format=cobertura,json`.
No exclusion configuration or production coverage exclusion was introduced.
The gate requires both reports from collection
`547b2b14-dd2d-472a-a4e8-cca2f08baf27`, reconciles their exact counters and
inventories, and requires every method line and IL branch to be covered. It does
not merge generated-class rows to hide an uncovered path. All 26 expected source
files enter the check; `VbaTestExplorerService.cs` contains only interface
declarations and has no executable sequence points, rather than being excluded.
The readable summary is `artifacts/coverage/vba-testing-final/gate.txt`, and the
machine-readable result is `artifacts/coverage/vba-testing-final/gate.json`.
The [testing guide](../tests/README.md#managed-vba-testing-coverage-gate) records
the build, test filter and strict gate commands.

The uninstrumented production DLL has MVID
`20a2f16c-c147-4e42-b1d3-a6f2e9415647`, SHA-256
`9A3ABDCCD56B3708C249B6CF4F54C0E62B105EFB362B2EE2765568299CB0337E`.
The test DLL has MVID `5b77c58c-feec-4967-a1e4-b7c53227ecb0`, SHA-256
`4F3FEEE2F0ACFBD4A58BFE21277CD15E6EBA0E4A0A422DFE6F58147430A0BB88`.
The post-collector identities are retained in
`artifacts/coverage/vba-testing-final/assembly-identities.json`.
The source manifest, recorded before this documentation-only checkpoint, is
`artifacts/coverage/vba-testing-final/source-manifest.json`, SHA-256
`7F5D2B3D0DB048C54B36DAECC10FE2EA1E0E5C612259BC40045FEFEBC91C4C36`.
The raw JSON has SHA-256
`27329E56C0493EA223F94BCD545F272AF22CE96A6918A6136CCFB4CC0FCE961B`;
Cobertura has SHA-256
`C64A0B31C0F1958F522876B974D3AF9326D2F81C43630AE09876FA7337F1F9CB`.

New tests exercise rejection and recovery contracts, not just successful calls.
Internal effect boundaries retain real Windows/COM/UI implementations by default;
synthetic VBE objects and owned disposable test-process windows cover failures.
Proven redundant conditions were simplified while preserving the original
uncertain-outcome and cleanup guards. The generated support signature now accepts
the null collections/descriptors already accepted by its generator, and a bounded
serializer overload permits testing the size refusal without a huge allocation;
the application's export limit is unchanged.

This is managed code coverage, not production host qualification or coverage of
user VBA statements/branches. No Office or SOLIDWORKS application was launched,
no production macro ran, no COM registration changed, and no installed DLL was
replaced for this checkpoint. The feature's previously recorded native acceptance
and source-navigation limitations remain outstanding. Ollama qualification is
outside this work and remains assigned to its separate branch.

## Ollama configured candidate acceptance (2026-10-01)

Candidate `bbb6e6f5aca9e47e2b120538e055959ef530e068`, product MVID
`f6f01687-c0a4-4b1c-acab-82cc1dfab2af`, SHA-256
`35E94E67D30819E32790854E317C55D5736DC74EA7EB2461F8301E56F2E6B9D1`,
passes the complete prepared batch: **3482 passed / 0 failed / 115 conditional
skips / 3597 total**, exit **0**. Test assembly SHA-256 is
`D58A5BF1D6B00D5EC8CA917C12C06A9E24CD97D1E5FF642520A19F2AE88925EB`.
The isolated solution build reports no warnings or errors. Source, clean working
tree and both assembly hashes are unchanged at the terminal check.

The verified model is `qwen2.5:7b-instruct`, manifest digest
`845dbda0ea48ed749caafd9e6037047aa19acfcfd82e704d7ca97d631a0b697e`,
on Ollama `0.34.4`, context 8192 and one parallel request. Seven captured chat
bodies verify temperature 0 and top-p 0.8 explicitly sent by the production
client, without logprobs. Cloud processing is disabled; the model cache and
selected literal-loopback endpoint are isolated from personal configuration.

All four activated real-provider scenarios pass with their strict assertions:
exact scalar synthetic tool arguments and result consumption, first-fragment
cancellation/recovery, shown chat streaming/Stop/next-send, and native read-only
Excel module inspection. The language-catalogue completeness check and both
sampling-save failure regressions pass. Failed writes or theme updates restore
the shared in-memory sampling values and preserve the original error; this does
not assert rollback of a partially written settings file.

Excel PID 55000 reads a random marker absent from the prompt through exactly one
real `read_module` call. Full source SHA-256 before, returned by the tool and after
is `6890696832cbbeddce234f43fdea1f19f445e9c9bf95eb99ff2ff1230274030d`.
No macro runs. The unsaved disposable workbook closes and the original process
handle observes normal exit code 0. This is external test-STA dispatch through
production tools, not qualification of the installed bridge or embedded assistant.
The detached chat uses a simulated VBE; SOLIDWORKS and other providers are not
activated in this batch. The installed DLL remains unchanged at `C900BA09...`.

Coverage includes `[VBAi]*`: **98.38% lines / 96.60% branches**
(37,385/37,998 lines and 37,324/38,634 branches). The two collector/deployment
copies contain identical results and are not added together. Conditional native
and authenticated-provider scenarios retain their explicit opt-ins. No coverage
exclusion is added and this is not complete native-host or release qualification.

Evidence is `artifacts/qualification-v1/followup-20260930/ollama-final-acceptance-v2/`
(`full-managed-and-ollama.trx`, `offline-final-acceptance-review.json`, native
Excel readback, synthetic wire captures, coverage and terminal manifests).
The exact owned backend is stopped once after terminal requests; cleanup succeeds
and no backend is retained. Earlier failed trials below remain failed. This
acceptance establishes the selected configuration's stated scope without claiming
the internal cause of historical unobserved generations or universal model reliability.

## Ollama configured integration before catalogue completion (2026-10-01)

Candidate `799ccef288178d266b560158dd3171330b7313af`, product MVID
`919d83d8-40a6-4f1c-8048-2072876c0aa9`, SHA-256
`9B6A998776520FAA2C328A223E50EE604FDD9DA23E9E1245A9D290ED20F3744E`,
uses the isolated verified `qwen2.5:7b-instruct` model, manifest digest
`845dbda0ea48ed749caafd9e6037047aa19acfcfd82e704d7ca97d631a0b697e`,
Ollama `0.34.4`, context 8192, one parallel request, temperature 0 and top-p 0.8.
Seven captured chat bodies independently verify the explicit sampling fields.
No personal settings, credentials or macro source are used as inputs.

All scenarios are prepared before the complete batch. It reports
**3479 passed / 1 failed / 115 conditional skips / 3595 total**, exit **1**.
All four real Ollama scenarios pass: exact scalar synthetic tool arguments and
result consumption, first-fragment cancellation/recovery, shown chat streaming,
Stop and next-send, and native read-only Excel module inspection. The aggregate
fails the complete-language catalogue test because the new settings strings are
not yet present in every language; it is not reported as a green managed gate.

The Excel scenario reads a random marker absent from the prompt through exactly
one real `read_module` call. The full source hashes are equal before and after;
no macro runs and the unsaved disposable workbook is closed. The original handle
observes Excel PID 37324 exiting normally with code 0 after COM release.
This is external test-STA dispatch through production tools, not qualification
of the installed bridge or embedded assistant. The installed DLL is unchanged.
Coverage includes `[VBAi]*`: **98.38% lines / 96.61% branches**, with conditional
native/authenticated-provider scenarios still excluded by their explicit opt-ins.
No coverage exclusion is added.

Evidence is `artifacts/qualification-v1/followup-20260930/ollama-final-acceptance/`
(`full-managed-and-ollama.trx`, `offline-final-acceptance-review.json`, native
Excel readback, bounded synthetic wire, coverage and terminal manifests).
The source and product/test hashes remain unchanged; the exact owned helper
is stopped once after terminal requests, with no backend retained.
The later catalogue completion and sampling-save restoration require a new gate.

The preceding Qwen3 4B aggregate remains failed: the shown chat and cancellation
pass, but the exact echo marker is altered. The independent supported comparison
also alters the marker with temperature 0 or 0.7 and top-p 0.8. The first
logprobs comparison receives HTTP 400 because that model/backend refuses tools
plus streaming plus logprobs; VBAi never sends this diagnostic option. These
controls do not establish the cause of historical unobserved generations.

## Current-main Ollama aggregate before configuration correction (2026-10-01)

The complete suite on test source `f15c48a` and product source main `2f28018`
uses the exact product/test identities in the next section. The batch runs once
against the isolated `qwen2.5:3b` cache with production sampling unchanged.
It reports **3373 passed / 1 failed / 116 conditional skips / 3490 total**;
runner exit is **1**. This is a failed aggregate, not a green managed gate.

| Activated real-provider case | Actual result |
| --- | --- |
| First-fragment cancellation and fresh conversation | Passed |
| Streamed greeting, exact synthetic tool call and result consumption | Passed; wire and pre-assertion ledger both retain scalar `marker=VB_AI_42` |
| Visible chat streaming, stop and next send | Failed before first text; stop and next-send phases are not reached |

The current UI request is 8261 bytes. Its captured SSE body is 433 bytes and
already has empty deltas, `stop` and `[DONE]`, with no text or calls. Diagnostics
record `complete-empty`, and the visible window displays `No text response.`
before becoming ready. The passive tee does not observe transport EOF because
the production parser stops at `[DONE]`; protocol completion is not reported as
socket EOF. Generated tokens are not captured for this live trial, so its internal
cause is not inferred from the separate seeded control.

Coverage is **unavailable for this run**: the artifact driver loads the product
with `ReflectionOnlyLoadFrom` in its own long-lived process to read the MVID.
Coverlet cannot write the DLL while that process retains it and reports failed
instrumentation. The original product and test file hashes remain unchanged at
the terminal check. No percentage is reused from an earlier binary. A disposable
offline copy proves writes succeed before load, fail while the metadata process
lives, and succeed after its exit, with unchanged bytes. A separate corrected
driver reads metadata in a child process that exits before collection; that
correction emits no inference and does not alter this failed record.

Evidence is `artifacts/qualification-v1/followup-20260930/ollama-main-current-capture/`
(`full-managed-and-ollama.trx`, `offline-qualification-review.json`,
`tests-terminal.json`, captured UI/headless bodies and shape ledgers).
The terminal manifest confirms unchanged clean source and payloads. Owned helper
PID `53100` exits after one verified helper-only force-stop with output pumps
terminal; no backend is retained. The installed DLL still has its preceding
`C900BA09...` hash. No native host or tool is exercised. Model/configuration
correction and a subsequent complete acceptance batch remain required.

## Shared Ollama synthetic capture contracts (2026-10-01)

Tests-only source `f15c48a5cabf83ab2159b2b18bc3b092ae045ad9`, based on main
`2f280183fec9572232f196a5b2c4f8f04640df3c`, replaces the fixture-private UI wire
wrapper with shared, separately opted-in UI/headless capture. It adds argument
shape ledgers before the existing scalar assertion, without changing prompts,
sampling, provider parsing or native tool dispatch. All contract scenarios are
implemented before the single focused batch is executed.

The isolated full solution build succeeds with **0 warnings / 0 errors**.
The helper/endpoint batch reports **86 passed / 0 failed / 0 skipped**; it uses
no backend, model inference, native host or coverage collector. Contracts cover
opt-in boundaries, routes/methods, redirect refusal, factory restoration,
credential-header omission, UTF-8 bytes, bounds, synchronous/asynchronous reads,
EOF versus zero-length/early/pending reads, disposal and error transparency,
diagnostic I/O/serialization failure and scalar versus nested argument shapes.

Product MVID is `044522a1-31cc-494c-98e6-46dee00af787`, SHA-256
`1A5B036B899A88D1E4B5315E3082C0689BA810B2BADB87F92A24926E1D251B4F`;
test SHA-256 is
`5C52CF0997F607D9F1AC5BD44B2B36229C80EDBB7A0E91FF079CEC0AC5BB5419`.
Both hashes are unchanged before/after the focused batch. Evidence is
`artifacts/worktrees/qualification-headless-wire/artifacts/headless-wire-validation/`
(`build.log`, `headless-wire-focused.trx`, `focused-terminal.json`). This is a
detached diagnostic-contract gate, not complete managed or live-provider
acceptance. The installed DLL is not replaced.

## Ollama sampling configuration comparison (2026-10-01)

The checkout is synchronized to main `2f28018`. This independent backend
comparison loads no VBAi assembly and does not qualify that product revision.
It uses the previously verified Ollama `0.34.4` executable and isolated
`qwen2.5:3b` manifest digest
`357c53fb659c5076de1d65ccb0b397446227b71a42be9d1603d46168015c9e4b`.
All cache blob lengths and hashes are verified before launch. No model download,
profile edit, authentication or native tool dispatch occurs.

All **12 requests** are prepared before execution: the exact headless echo
prompt/tool schema and the retained detached-UI request, each with seeds 42/73
and temperature omitted, 0 or 0.2. Every request also enables chosen-token
logprobs; the comparison therefore does not recreate the earlier unseeded wire
byte-for-byte. Each request is emitted once, with no retry. The single batch
exits **0** with **12 known HTTP/protocol terminals**, not 12 VSTest passes.

| Fixture | Temperature omitted | Temperature 0 | Temperature 0.2 |
| --- | --- | --- | --- |
| Echo, both seeds | Exact scalar `marker=VB_AI_42` | Exact scalar `marker=VB_AI_42` | Exact scalar `marker=VB_AI_42` |
| UI, seed 42 | Streamed text | `read_module` with an invented module | `read_module` with missing required arguments |
| UI, seed 73 | Complete-empty; malformed generated tool JSON | `read_module` with an invented module | `read_module` with an invented module |

For the new empty response, chosen tokens reconstruct a `discover_tools` call
whose `arguments` is a quoted JSON string containing unescaped inner quotes.
The generated call is malformed JSON; delivered SSE contains neither text nor
tool calls, ends with `stop` and `[DONE]`, and reaches HTTP EOF. This pairs the
model generation with the empty response at the backend boundary. It does not
establish the missing-token historical responses' cause or a VBAi parser defect.
Lower temperature changes this observed outcome but does not establish safe,
correct tool selection. No malformed arguments are repaired or dispatched.

The actual server log confirms temperature **1** when omitted, top-p **1**,
context **4096**, UI prompt **1673 tokens**, echo prompt **181 tokens** and no
truncation. A short context does not explain these captured trials. The exact
Ollama [OpenAI conversion](https://github.com/ollama/ollama/blob/v0.34.4/openai/openai.go#L644)
sets temperature/top-p to 1 when absent from the request, overriding a model's
sampling default for those fields.

Evidence is under
`artifacts/qualification-v1/followup-20260930/ollama-sampling-controls/runs/c455aa1094a54ef3bdbfc05d522fe8ff`:
the terminal manifest, per-request bodies/hashes and chosen tokens,
`offline-sampling-review.json`, and the server's sampling/context log.
The original handle identifies owned helper PID `6024`; after all requests
terminate and no established connections remain, one helper-only force-stop
observes exit `4294967295`, with both output pumps terminal. No Ollama process
remains. This is not native Office shutdown qualification. Q-028 stays open.

## VBA test explorer integration with main (2026-10-01)

The PR combines test-explorer commit
`146296dc0d349413ccbaec67f8cc954360378d3f` with main
`afaa2e950f83e196210640428c57436e5f3ac109` at the maintainer's request.
Documentation and all localization additions from both branches are preserved.
The Office fixture retains main's one-shot shutdown/process-handle and uncertain
command guards together with the test feature's ownership checks, unsettled-test
retention and reviewed Access support prompt. Forced termination remains disabled.

The isolated integration build used production MVID
`ef740dad-5901-4ea3-9fef-4d0dc6d59aad`, SHA-256
`6DEB90199EA6261C74E1B43344FE1327F04F0DFE1749159E36822221C01F3FC8`,
and test MVID `a9b1c606-49db-4bfb-b35a-786619dfae73`, SHA-256
`8C3AF7B0170EAFA836577834E934DBE26F2F27A16C4F9828361FF9FD420AAD10`.
The source manifest is `artifacts/pr-preparation/integration-source-manifest.json`,
SHA-256 `7C22420AAD336E9AF5914C7553DDAFEC3B2A282535A69924522271EF160622BD`.
It records the code tree before this documentation-only checkpoint was added.

| Check | Actual result | Evidence |
| --- | --- | --- |
| Isolated .NET Framework 4.8/x64 build | Passed, no errors; three NU1900 endpoint warnings | `artifacts/build-pr17-integration` |
| Test-subsystem/LLM unit classes plus OfficeVbeFixtureShutdownTests and OfficeOwnedShutdownEvidenceTests | 337 passed, 0 failed, 0 skipped | `artifacts/test-results/vba-tests-pr17-integration.trx` |
| Resource resolution | All keys and values from both sides retained, unique XML keys in each locale | Staged resource validation |
| Documentation and conflict/diff checks | Passed | Documentation checker and Git index checks |

The build command was `dotnet build tests/VBAi.Tests/VBAi.Tests.csproj -c Debug
--no-restore -p:BuildOutputRoot="$PWD/artifacts/build-pr17-integration"`.
The focused test command used that same output with `--no-build`. The two fixture
regressions added during resolution use managed fake documents and the testhost's
own query handle; they launch no Office process. Overlapping earlier filters are
not additional tests of this tree and must not be added to these counts.

This is source integration, not native production acceptance. No native host,
production macro, COM registration change or installed-DLL replacement was used
for these integration checks. The feature's native execution, coverage-host and
Monaco-navigation limitations described below remain outstanding. Unrelated
application qualification was not rerun.

## VBA test explorer draft checkpoint (2026-10-01)

This branch adds the VBA test subsystem; production acceptance remains incomplete.
Its acceptance scope includes discovery, support installation, single/batch runs,
fixtures, assertions, cooperative stop, procedure coverage, explorer actions,
human/JSON reports and the corresponding LLM tools. Unrelated Git/editor workflows
and the rest of VBAi are not acceptance gates for this feature. Visio and Project
are excluded from the requested native test scope.

The PR preparation build used production MVID
`610648ad-a0fd-454c-bda5-5fedb553d743`, SHA-256
`0095FCA1A201DAECBAC36123193850E291CFB988A3B6558C5748D801B2596208`,
and test MVID `f1c3f07c-84c0-4fce-946c-04a0c6e0280b`, SHA-256
`670D1961A80FDBB93F596031ECD5FBFFF91E3594D218A89D261FDB87A3F66667`.
It compiled into `artifacts/build-pr-test-explorer` with no errors and three
NU1900 warnings because the vulnerability endpoint was unavailable.

| Check | Observed result | Evidence |
| --- | --- | --- |
| Before main integration: test-subsystem unit classes and LLM boundary | 324 passed, 0 failed, 0 skipped | `artifacts/test-results/vba-tests-pr.trx` |
| Documentation links | Passed | `python tools/docs/check_docs.py` |
| Last registered Excel trial, earlier candidate `692e236e-278e-41e8-92fd-a9fe88acbf7c` | Four Blocked results before Run; uncertainty false; native-window eligibility refused | `artifacts/test-results/native-explorer-692e236e/subsystem/0b049ed146a747c0b8ba1bab4d4ff84e/` |

The focused unit filter selects VbaTest, VbaCoverage, VbaNativeTest,
VbeTestExplorerService, TestExplorerWindow, TestSupportReviewDialog and
LlmVbaTestingBoundary classes in `VBAi.Tests.Unit`. It excludes native opt-ins and
unrelated application tests. These tests do not measure instrumented .NET coverage
or establish real-host production readiness.

The earlier native candidate has SHA-256
`C9CBEB29D8BD96C4058CBB3AE43810CE6551D682108E6B9A41B150FCA002D329`.
Its source manifest SHA-256 is
`C8AB36014399F916F2AB6A1D05A84EB0A8B5F2FC79F748DC86DB6570F158DF18`.
The executed script SHA-256 is
`57FA19484DA3A79451DCCC7CF59442B7C27F10706FE84B469C49A950EC8C6A3F`.
The explicit `TestSubsystemOnly` scope retained coverage and explorer/report
assertions, while excluding the general Excel BeforeSave positive save control.
That control was recorded as `NOT_TESTED_BY_SCOPE`, not as passing.

No test was dispatched in that native trial. Measured coverage, native explorer
actions and exports were not reached. Close and Quit returned with zero workbooks,
but normal process exit remained unverified at fifteen seconds. No Quit retry or
termination occurred. Temporary HKCU registration was restored and exact
restoration verified. The current diagnostic refinement was not retested natively.

Further acceptance requires real registered execution and the relevant explorer,
report and coverage operations in disposable projects across the requested Excel,
Word, PowerPoint, Access, Publisher, Outlook and user-selected preloaded SOLIDWORKS
scope. Excel/Word/PowerPoint coverage-copy adapters exist; the other hosts have no
implemented coverage-copy path. Native source navigation currently selects the
VBIDE pane; synchronization of the visible Monaco document remains unfinished.
Interrupted navigation fragments are preserved in ignored local artifacts and
are excluded from this PR. Native screenshots are unavailable because the latest
registered trial stopped before explorer actions. No native tests were restarted
while preparing this draft.

## Direct Ollama controls and offline harness correction (2026-10-01)

Four independent synthetic protocol controls use the same Ollama `0.34.4`
binary, isolated `qwen2.5:3b` cache and digest recorded in the following section.
No VBAi assembly, Office host or native tool is loaded or executed. The original
UI request is retained byte-for-byte in A; B only adds `logprobs: true` and
`top_logprobs: 0`, C only disables streaming, and D only supplies an empty tool
array. No seed or temperature is added. These are diagnostic observations,
not a VSTest pass or acceptance of the failed real-provider aggregate.

| Control | Observed terminal response | Captured bytes |
| --- | --- | ---: |
| A: exact original request | `discover_tools` with `Family=all`, no text | 577 |
| B: chosen-token diagnostics | Same recognized tool; 19 chosen tokens / 83 bytes reconstruct its complete `<tool_call>` text | 6,133 |
| C: no streaming | `list_modules` with `{}`; required `Project` is missing; no dispatch | 457 |
| D: no tools | Streamed text, `stop`, no tool calls; content adherence is not qualified | 209,194 |

All four captured bodies have independently verified HTTP EOF and protocol
terminal markers (a complete JSON choice for C, `[DONE]` for the streams).
The original runner exits **1** after A/B: PowerShell 5.1 unrolls the chosen
token's single-byte array `[10]` into a scalar and StrictMode rejects `.Count`.
This is a proven artifact-parser defect, not an Ollama transport failure.
Its original terminal manifest and incomplete token sidecar remain unchanged.
A corrected copy wraps the entire conditional in an array; **8 offline checks
pass**, including original/corrected captures and empty, single, multiple and
absent-byte shapes. An independent Python review also verifies the original
A/B bodies without contacting the backend.

Only the never-sent C/D controls run subsequently, once each. That continuation
exits **0**, using an explicitly reacquired handle for the exact retained backend
PID `14364`, start UTC `2026-10-01T08:14:28.2686818Z`, image and loopback listener
`127.0.0.1:58025`. The original runner's handle is unavailable and is not claimed
as retained by the continuation. After all emitted requests are known terminal
and no established connections remain, one verified force-stop observes helper
exit `4294967295`; no Ollama process remains. This is not normal native-host
shutdown acceptance. No A/B replay, download, authentication, profile edit or
installed-DLL replacement occurs.

Evidence is under `followup-20260930/ollama-controls-prepared/`:
`protocol-controls-summary.json`, `runs/protocol-controls-20261001/`, and
`runs/continuation-C-D-20261001/`. Offline correction evidence is
`ollama-controls-parser-correction/offline-verification/867af08c10e149ca8a3b4f56d0be1e69/offline-report.json`
under the same follow-up root. No coverage collector runs for these controls.
The prior managed coverage and failed activated aggregate remain unchanged.

These new generations do not reproduce the empty delivered response or reveal
the original hidden tokens. Ollama's version-specific source has a generic tool
parser that buffers tagged output and may return no remaining content when an
XML-tagged call cannot be recognized; its HTTP route can emit chosen-token
logprobs while that content is buffered. This establishes a possible mechanism,
not the historical cause. See the official
[tool parser](https://github.com/ollama/ollama/blob/v0.34.4/tools/tools.go#L374)
and [chat route](https://github.com/ollama/ollama/blob/v0.34.4/server/routes.go#L2732).
The valid call in B is recognized and delivered. C demonstrates a separate
missing-argument output, not a VBAi coercion defect. No provider reliability
correction or release-gate closure is claimed; Q-028 remains open.

## Frozen product: embedded Git preparation and real Ollama failures (2026-10-01)

Product source remains `b60996cdf05d816a1f24bea99f2343c166a8bfac`, MVID
`e79c6288-d384-475c-b8bc-276d7caaaf00`, SHA-256
`9ADBCFB96B1F2F4E4FA3066EA26F3CA2E0A2EC7996C85105070A3765BEFC8587`.
These are tests-only changes against copied, frozen product outputs; no product
rebuild, deployment or native host action occurred.

Test source `a28be697085ddd42826e38f6f223b2d3b9cf2770` integrates the actual
`VBAi.GitHub` menu/window scenario. Its complete default batch records
**2,959 passed, 0 failed, 111 conditional skips, 3,070 total**, runner exit **0**,
in **629.223 seconds**. Evidence is
`followup-20260930/managed-v11-embedded-git-tests/full-managed.trx` and its
terminal manifest. Test assembly SHA-256 is
`FA31C7DB243248D740609F77DAE3C0642AAC03CA797A38205B7D9AE6382F3046`.
Both runtime copies of the product and the test assembly are unchanged afterward;
the source is clean and unchanged. The tests-only build has no warnings or errors.
The embedded native scenario is NOT_RUN: owner-thread dispatch, checkpoint
content, independent bridge exports and full FRX comparison are prepared, not
Excel/Git acceptance. Remote publish, import, recovery and disk reopen remain
separate unqualified operations.

Test source `ce3d9da314d4a460a24e0b22084331339e058b05` adds an explicitly
selected literal-loopback port to the detached Ollama cases. The default endpoint
is unchanged; wire capture refuses other servers, credentials and unrelated
routes. The complete prepared guard batch records **65 passed, 0 failed, 0 skipped**,
in **574 ms**, in `followup-20260930/ollama-port-and-embedded-preflight/focused-tests.trx`.
Its scope is the embedded Git protocol/oracle and Ollama endpoint helpers, not a
complete product gate. Test assembly SHA-256 is
`0B84DFC7339D5876AB17926B64889162D2408DDC19FC39E7D5FDF9FB1A51BAD3`.

The subsequent complete suite with the three real Ollama cases activated records
**2,984 passed, 2 failed, 108 conditional skips, 3,094 total**, runner exit **1**.
Both failures are real-provider cases; the aggregate remains failed. The
cancellation/subsequent-request case passes. The tool roundtrip refuses an object
where the supplied schema requires the scalar `marker` value `VB_AI_42`, before
any synthetic tool result is sent. Its response body was not captured, so a
model/backend shape defect or assembly defect cannot be attributed from this
assertion alone. The UI case fails at visible streaming; Stop and its subsequent
send are NOT_RUN. No native tool is executed.

Evidence is `followup-20260930/ollama-v11-reused-verified-model/`:
`full-managed-and-ollama.trx`, `plan.json`, `tests-terminal.json`, `terminal.json`,
`model-ready.json`, `server-stopped.json` and `synthetic-wire/`. The backend is
Ollama `0.34.4`, selected endpoint `127.0.0.1:54579`, model `qwen2.5:3b`, digest
`357c53fb659c5076de1d65ccb0b397446227b71a42be9d1603d46168015c9e4b`.
The model cache is isolated under qualification artifacts; cloud features and
startup pruning are disabled. Product/test hashes and clean source are unchanged
at completion. The original owned helper exits after one explicit force-stop,
exit **-1**; this is not normal Office/SOLIDWORKS exit evidence.

The UI capture `wire-c7229a5235d54589b71f888307cc454c-response.bin` retains the
complete observed **433 bytes**: an empty assistant-content delta, an empty delta
with `finish_reason: stop`, then `[DONE]`, with no text or tool call. The capture
is not truncated; its EOF-read flag is false because the parser stops on `[DONE]`.
The failure snapshot records `complete-empty`, zero text/tool chunks and the
visible `No text response.` fallback with ready status. This reproduces the
observable empty-response defect and locates the empty delivered response at the
Ollama HTTP boundary for this trial. Backend/model internals and the historical
ce19 failure remain causally unproven. No reliability correction is claimed.

Earlier preparation failures remain retained: default-port bind returns Windows
socket error **10048** despite no observed listener or exclusion covering that
port; a new owned port works. The first alternative-port attempt stopped at its
premature readiness timeout. The bounded-readiness attempt prepared the model
with a verified manifest digest but lacked its original CLI exit code; its
failed harness result is preserved. The completed batch reuses that verified
cache without replaying the download. None of these preparation results is a
provider pass.

Raw `[VBAi]*` coverage for both complete batches above is
**34,056/34,250 lines (99.43%)** and
**34,330/34,824 branch outcomes (98.58%, rounded from counts)**, leaving
194 uncovered lines and 494 uncovered outcomes. Their collectors are
`managed-v11-embedded-git-tests/d32694d3-bc35-42ea-8451-9e5373ef5c67/coverage.cobertura.xml`
and `ollama-v11-reused-verified-model/3d97c12c-a384-4fcc-bddd-e4657c3d6717/coverage.cobertura.xml`
below `followup-20260930/`. This managed measurement excludes external hosts,
native C++ and JavaScript. Coverage from the failed live batch does not promote
its failures; Q-028 and the whole-product coverage target remain open.

## Uninstalled compatible Git binding candidate: complete managed and local gates (2026-10-01)

Source `b60996cdf05d816a1f24bea99f2343c166a8bfac`, MVID
`e79c6288-d384-475c-b8bc-276d7caaaf00`, SHA-256
`9ADBCFB96B1F2F4E4FA3066EA26F3CA2E0A2EC7996C85105070A3765BEFC8587`
adds compatible native/legacy document binding lookup to the preceding recovery
candidate. Test assembly SHA-256 is
`BC2EAEDD813451A4A6EA42575EB308383D197312CA6BB5FFE7D38F8A6668CC47`.

`managed-v11-git-scope/full-managed.trx` records **2,918 passed, 0 failed,
110 conditional skips, 3,028 total**, runner exit **0**, in **579.751 seconds**.
Individual `UnitTestResult` outcomes establish the skip count; the TRX summary
reports total/executed/passed but leaves its `notExecuted` attribute zero.
The terminal manifest verifies clean, unchanged source and identical before/after
product hashes. The Debug solution build has no warnings or errors. Raw `[VBAi]*`
coverage is **34,056/34,250 lines (99.43%)** and
**34,333/34,824 branch outcomes (98.59%, rounded from counts)**, leaving
194 uncovered lines and 491 uncovered outcomes. The collector is
`managed-v11-git-scope/80e58a65-fa4a-401b-844b-6ae29909321d/coverage.cobertura.xml`.
The new `MacroGitRepository.Scope` partial and the touched Git operations, LLM
Git and chat-shell partials have complete measured line/branch coverage. No
production coverage exclusion was added. Native/live-provider opt-ins are off;
external hosts, native C++ and JavaScript are outside this managed measurement.
The whole-product coverage target remains open.

The delegated fix `793be32`, integrated as `923910a`, prepares the complete local
matrix for native, legacy, missing and ambiguous bindings, unreadable metadata,
directory/reparse entries, unchanged files and guarded chat/LLM callers. Its
final batch records **109 passed, 0 failed, 0 skipped**, product MVID
`71bc1d22-78bd-44cc-9044-dfaeb9406dcc`, SHA-256
`C0E69F39AEA9315CA4D676477CE487D5FB1F84BBAF832CF934759A17A623BC07`.
Proof is `artifacts/worktrees/qualification-git-scope-case/artifacts/scope-source-proof.json`;
the green TRX is `scope-green-results/scope-compatible-mirrors-green.trx` below
that worktree's artifacts. Six regressions fail with four exact parent call-site
files from `38e3d8e` recompiled alongside the new pure resolver used by the tests.
That binary is not the historical candidate or a complete old-source rebuild.
Some legacy rows fail the exact native-scope assertion even when the old local
action succeeds; the red batch is not six native-export failures. Neither batch
uses Office, COM, real authentication or network transport.

`local-v11-git-scope/terminal.json` records all prepared local checks passing,
runner exit **0**, with unchanged candidate bytes and source: **61 JavaScript
tests**, **20 synthetic native cycles**, exact candidate loader extraction/hash/
ABI/reuse, and **46 detached WinForms Designer surfaces**. The separately built
synthetic C++ binary is distinct from the embedded payload. These checks do not
qualify real-host painting, embedded Git threading, GitHub transport or providers.

Q-015 has managed acceptance for this exact **uninstalled** candidate. The
installed product remains `8f2315d` / `d8f31d57` with its retained hash unchanged.
The nine remaining release gates and all failed historical native results remain
open in their stated scopes.

## Uninstalled recovery marker candidate: complete managed and local gates (2026-10-01)

Source `bdb57e00c8763d84ece627601035b717c1008cf8`, MVID
`f6d47c82-d9d3-4a8e-9f28-79778ae112f1`, SHA-256
`019EA9F7F4A72B3AE92EDC49F3C96A868CD3ED36A340031DF325BD816FD048CA`
includes the preceding Options/CFB corrections and guarded recovery marker
metadata, preparation, rollback and completion. Test assembly SHA-256 is
`6FD6C5625F84C8AB60F79EB35715E1B274459F42474C08CA816A145CB472ABC1`.

`managed-v10-recovery-marker/full-managed.trx` records **2,888 passed, 0 failed,
110 conditional skips, 2,998 total**, runner exit **0**, in **602.314 seconds**.
Individual TRX outcomes independently match the totals. The terminal manifest
confirms identical before/after product hashes and clean, unchanged source.
The Debug solution build has no warnings or errors. Raw `[VBAi]*` managed
coverage contains **34,034/34,228 lines (99.43%)** and
**34,327/34,816 branch outcomes (98.60%, rounded from counts)**, leaving
194 uncovered lines and 489 uncovered outcomes. The collector is
`managed-v10-recovery-marker/79059fd6-eca6-4d1f-95c6-cf6ee52d0483/coverage.cobertura.xml`.
It reports complete coverage for `FormResourcePreflight`, `CompoundFile`,
`MacroGitOperations` and the `MacroGitRepository` partials. There are no new
production coverage exclusions. Native/live-provider opt-ins were disabled;
external host processes, native C++ and JavaScript are not measured by this
collector. The whole-product coverage target remains open.

The delegated recovery commit `f9a221b`, integrated as `035b87c`, prepares real
file/directory markers and disposable local Git repositories, with doubles for
VBE mutations and injected metadata/deletion failures. Five regressions fail on
the two exact parent production files recompiled separately from `d4a0fd5`;
this rebuilt old-source binary is not the historical candidate assembly.
The final recompilation includes all scenarios and records **79 passed,
0 failed, 0 skipped**, with product MVID
`0ac7e107-6c48-4495-b7d4-0d4a8d21f92d`, SHA-256
`7FFC259A243A6CA7D68A7BAC917C5CDF84C180BB58ABEC712D78A3D5643CE800`.
Its proof is `artifacts/worktrees/qualification-recovery-marker/artifacts/recovery-source-proof.json`,
with `recovery-red-results/recovery-parent-directory-red.trx` and
`recovery-final-results/recovery-mirrors-final.trx` below that worktree's artifacts.
The earlier focused run excludes the last prepared additions and is not their
acceptance. These are managed/local-filesystem contracts, without ACL changes,
COM/native import or a claim of atomic filesystem races.

The complete candidate also passes the prepared local batch in
`local-v10-recovery-marker/terminal.json`, runner exit **0**, with unchanged
product bytes and source:

| Check | Result | Scope |
| --- | --- | --- |
| JavaScript | 61 passed, none failed/skipped | Editing, language and synthetic native-export-trace fixtures. |
| Native C++ | 20 cycles passed | Fresh synthetic VBE fixture, scoped drawing, thread guard, destruction, import restoration and GDI balance. This standalone binary is distinct from the embedded payload. |
| Managed native loader | PASS | Exact candidate embedded extraction/hash/ABI and module reuse; no real-host painting claim. |
| WinForms Designer | 46 surfaces passed | Detached construction, resize, editable child properties and serialization roundtrip. |

Q-015 has complete managed acceptance for this exact **uninstalled** candidate.
The local checks do not qualify native Office/SOLIDWORKS, actual GitHub transport,
embedded Git UI/threading or provider behavior. Historical native failures and
the remaining release gates retain their original outcomes.

## Uninstalled exact Options lifetime candidate: complete managed gate (2026-10-01)

Source `d4a0fd5a7662806c141f7e851be126feecb4ca9b`, MVID
`a25eae3f-47b6-45c3-999c-0ff05f630e36`, SHA-256
`8ACE24B2BC85255B8274EE0FC6020228DDB26EE11612CCD8C2CD0A86B26B5798`
combines the corrected native-probe fixtures, exact captured-handle closure for
Options reads/Accept and the full CFB guard contracts. Test assembly SHA-256 is
`432733B0B5D06471A1A7E0EB56B947FE7C8B917D0EBF14244B76104A0AC3D079`.

`managed-v9-options-lifetime/full-managed.trx` records **2,839 passed, 0 failed,
110 conditional skips, 2,949 total**, runner exit **0**, in **506.021 seconds**.
The individual outcomes match the totals. The terminal manifest confirms clean,
unchanged source and identical before/after product hashes. Solution build has
no warnings or errors. Raw `[VBAi]*` managed coverage is
**34,000/34,194 lines (99.43%)** and **34,316/34,808 branches (98.58%)**, in
`managed-v9-options-lifetime/61c7e0b1-62a4-440f-8d4d-54aabcfde059/coverage.cobertura.xml`.
Native/live-provider opt-ins are disabled; native C++, JavaScript and external
host processes are outside this measurement. The complete coverage target
remains open. Q-015 has managed acceptance for this exact uninstalled candidate,
not native release acceptance or a later source revision.

The delegated closure follow-up at source `c34877d`, integrated as `e386f00`,
records **228 passed, 0 failed, 0 skipped** in
`artifacts/worktrees/qualification-teardown-trace/artifacts/focused-options-read-lifetime/options-read-lifetime.trx`.
It verifies both read paths preserve capture/cancellation errors, perform only
one owned Cancel and observe exact destruction. Post-Accept performs observation
only: hidden captured windows remain open; enumeration failures propagate;
no Cancel or replay follows a started validation. Its tested isolated MVID is
`b5d627ea-85da-4875-96b3-97163d33a6f8`, SHA-256
`5E559A25263E43903BDAFA2C67A59A1BB9D3D117E7557B9D3DD7E848B2A2EF3B`.
The global candidate above independently includes these regressions. The failed
preceding full gate and original installed-v6 native Format failure remain
preserved below. The installed DLL is unchanged; Q-026 remains open for native
validation and the earlier unexplained revision/EOF scopes.

The final CFB contract batch at tests-only source `8a3575a`, integrated as
`d4a0fd5`, records **48 passed, 0 failed, 0 skipped** in
`artifacts/worktrees/cfb-guards/artifacts/test-results/cfb-contracts/cfb-contracts.trx`.
Its MVID is `96956356-0bd4-45ae-a2d3-a1b6a807d7ca`, SHA-256
`CE6098CF80C357E6B9C4A9126D09E652E74CC2A34BD0E9596B5EA71D6E4EBF8C`.
No production changes, exclusions or field rewriting are used. The five private
defensive outcomes are exercised as direct helper contracts on a real validated
resource: bounded reads cannot escape into existing envelope/trailing bytes;
invalid sector indices are refused; invalid chain sizes reserve no free sector.
Valid boundary neighbors are also checked. This is explicitly separate from
public malformed-file reachability. The focused collector measures
**38/38 lines and 24/24 branch outcomes** for `FormResourcePreflight`, and
**181/181 lines and 196/196 branch outcomes** for `CompoundFile`.
The complete global collector also reports 100% lines/branches for these two
classes. It does not extend that result to the entire product or native hosts.

## Uninstalled Options cancellation candidate: complete gate failed (2026-10-01)

Source `dc5d3d064882ab7220d9210e6616793c83550f31`, MVID
`a066b3d6-d211-42a1-b665-904f31d2c6d2`, SHA-256
`76978BE05E394D525875291D3D2819608DF87DD0893F975126F98C52D73779F2`
adds one guarded Options cancellation with bounded exact-handle closure
verification. A throwing Write or started Accept is retained without cancellation;
known pre-write failures preserve primary and cancellation errors separately.

The delegated prepared focused batch at source `37bb0d7` records **57 passed,
0 failed, 0 skipped** in
`artifacts/worktrees/qualification-teardown-trace/artifacts/pure-options-cancel/options-cancel.trx`.
That focused pass is not full acceptance: the combined candidate's
`managed-v8-options-cancellation/full-managed.trx` records **2,764 passed,
3 failed, 110 conditional skips, 2,877 total**, runner exit **1**, in
**568.522 seconds**. Its terminal manifest confirms unchanged product bytes
and a clean, unchanged source revision throughout the run. The Debug solution
build has no warnings or errors. Native/live-provider opt-ins were disabled.

The failures are `VbeDebugWindowsSystemTests` cases
`NativeDialogCaptureReadsVisibleControlsAndNativeWatchMessages`,
`UiaOptionsReadNativePatternsWithoutReadingPasswordsOrDisabledControls` and
`UiaGeneralTabRequiresOneSelectableTabAndReadsEachRadioSelection`. Their
Options/Cancel fixtures no longer satisfy the new native ownership guard.
The guard/fixture contract requires correction and a new complete gate;
neither the preceding focused pass nor the earlier candidate's full pass
qualifies this assembly. This candidate is **uninstalled** and Q-015 is open
for it. The installed v6 and preceding combined candidate retain only their
separate acceptance scopes below.

The failures are explained by the injected scene contract: it omitted the
`OptionsWindowEnabled` seam for simulated Cancel handles; its thread/PID model
also needed faithful unknown-handle behavior. The third test passed a VBA error
dialog to Options.Close, which the new identity guard correctly rejects.
Tests-only commit `45a9511`, integrated as `d1c960c`, models the genuine UIA
dialog root and injected child ownership, restores all modified delegates and
asserts refusal without a posted message for non-Options dialogs. The complete
prepared focused batch records **120 passed, 0 failed, 0 skipped**, including
all three failed methods, in
`artifacts/worktrees/qualification-teardown-trace/artifacts/focused-options-fixture/options-fixture-full-focused.trx`.
It uses unchanged frozen product `ced62724` from source `37bb0d7`. It is not a
new complete gate. A read-only review additionally identifies visible-only
closure checks after Accept and in both reads, plus read-error masking by
finally Close; these still require correction before deployment.

### Compound-file guard coverage follow-up

Tests-only source `2c6d9b2`, integrated as `083db3c`, prepares all reachable
missing `CompoundFile` outcomes through real public resource preflight and
comparison. Its single instrumented focused batch records **35 passed,
0 failed, 0 skipped** in
`artifacts/worktrees/cfb-guards/artifacts/test-results/cfb/cfb.trx`. Product MVID
is `7b103e8f-06c6-4060-ac60-e811f2c7f96d`, SHA-256
`8EADB7EFA462D8E5DCB6C41CF940B91BEAA7FAD370F26915B24913284AB25093`;
no production code or exclusion changes. The collector under
`5debd30c-b055-4e32-940f-3a5b81008381/coverage.cobertura.xml` measures this
focused scope only: `CompoundFile` has **0 uncovered lines and 191/196 branch
outcomes**. It covers 22 previously missing outcomes; the five remaining are
defensive size/sector/field-bound guards unreachable from the public path after
its preceding checks, with explicit proofs in `artifacts/cfb-review/evidence.json`.
Those outcomes remain reported rather than excluded. This is not whole-product
coverage or native acceptance.

## Uninstalled combined candidate: complete managed run (2026-10-01)

Source `fa7955ff9957138924ba911df446a9270e377dc5` combines bounded,
atomic UserForm storage-graph comparison and scalar-setter failure-phase
diagnostics. Product MVID is `a8a35043-3353-4917-812a-6e020da7a049`, SHA-256
`431F27F1507E1E0D50DEEFDE657885EF6B8601938CF048BB8B7BB09BC53BE6D4`.
This candidate is **uninstalled**; the installed v6 identity below is unchanged.

`managed-v7-final-combined/full-managed.trx` records **2,703 passed, 0 failed,
110 conditional skips, 2,813 total**, runner exit **0**. The terminal manifest
records **565.722 seconds** and identical before/after product hashes. The
candidate manifest binds the compiled test assembly. Test-only guard edits made
while this frozen run executed are qualified separately below; this run does
not claim their compiled execution.

Raw `[VBAi]*` managed coverage is **33,927/34,122 lines (99.42%)** and
**34,238/34,756 branches (98.50%)**, in
`managed-v7-final-combined/3c4c5213-fff3-4c60-8604-7eca5a71df61/coverage.cobertura.xml`.
Native/live-provider opt-ins were disabled. Native C++, JavaScript and external
host processes are outside this measurement. Q-015 has managed acceptance for
this exact candidate; native release gates and the complete coverage target
remain open.

`frx-drift-analysis/final-combined-graph-comparison.json` invokes the actual
production snapshot comparison in this assembly on all twelve retained native
before/unchanged pairs. All compare equal, including Frame/MultiPage. Raw FRM/FRX
files are unchanged. The graph parser rejects unknown, malformed, missing or
extra streams atomically; picture payloads and meaningful properties remain
significant. Unsupported layouts retain strict comparison. This offline result
does not change the failed installed-v6 matrix or execute native import,
recovery, persistence or authenticated transport.

Tests source `2f186993a159b085f228d9c7f8b19a5ebffbbab9` is compiled separately
against the unchanged frozen product above. Test SHA-256 is
`BC256576C21842339CABDC81EE1A1C051B2F7F3E27E78FD6040D86A5FFEF42DD`.
`combined-test-guards/guards.trx` records **215 passed, 0 failed, 0 skipped** in
one prepared batch. This includes grammar, scalar failure phase, Format
lifecycle/bootstrap and sequential import guards. A pending recovery marker or
unreadable marker retains the native host after an import failure; final evidence
failure cannot mask the scenario error or skip context restoration. Format uses
explicit owned bootstrap and refuses inherited token manifests before launch.
No host is launched by this focused batch. Neither test-only lifecycle correction
changes the shared fixture or explains the original native failures.

The read-only fixture review found that the actual recovery callback used
`File.Exists`, which can hide access errors as absence despite the helper's
simulated exception guard. Tests source `3a1076a` now observes real marker metadata:
any existing file/directory requires retention; only FileNotFound or
DirectoryNotFound establishes absence. Other metadata errors remain primary
diagnostic evidence. `recovery-observer-guards/recovery-observer.trx` records
**22 passed, 0 failed, 0 skipped**, runner exit **0**, including real synthetic
file/directory/missing paths and simulated access/I/O failure propagation.
Its candidate manifest binds the test hash and unchanged frozen product.
The authenticated native scenario remains NOT_RUN.

### Installed v6: full Format matrix stops at a retained native dialog

`format-native-v6-corrected/format.trx` records **0 passed, 1 failed, 0 skipped**
on installed v6, tests source `2f186993`. Explicit `/x /automation` bootstrap
verifies owned Excel **48192**, start **2026-10-01T03:50:04.2993576Z**, and loaded
MVID equal to the installed product. No path/token manifest is supplied. The
tests-only payload references the original installed assembly; it cannot qualify
the uninstalled combined candidate. An earlier driver path-decoding error
occurred before preparation or host launch and is not a native attempt.

The font write has a confirmed terminal reply and independent readback. The
subsequent empty-size catalogue refuses `12` with the exact expected error, but
independent complete native window enumeration finds the owned `Options`
`#32770` dialog **visible**. The guard correctly stops dispatch and retains the
host and COM references. Twenty durable phase records preserve the baseline,
committed font restoration entry, refusal and native observation. Palette,
category, margin, stale-revision and complete restoration phases are **NOT_RUN**.
No restoration or Close/Quit follows the refusal. At this checkpoint the VSTest
console has written its failed result, but the runner has not returned a terminal
exit; do not treat the TRX as normal process completion. Q-026 remains open.
An additive independent read-only window observation at
**2026-10-01T04:11:51.7847671Z** confirms the same owned PID/start and no remaining
`#32770` dialog. The initial visible dialog was observed immediately after the
refusal; it is not evidence of permanent failure to close. The original failed
test and retained-host decision remain unchanged. No bridge, COM, input,
restoration or cleanup occurs during this later observation, retained in
`retained-dialog-later-observation.json`.

### Q-026 inactive-desktop preparation (2026-10-03)

Source `180aee899941027c16bece8bef61dfcef47f7f22` prepares the focused Q-026
harness against the frozen installed candidate, not against a rebuilt product.
Its actual product MVID is `d2c3601b-893d-4e84-b9e3-c172da7e2437`, SHA-256
`C4D095D9427379AC8F2A82D047C8780637A0E815173241976F2C86664E777644`.
The focused harness/helper and the complete solution build completed without
warnings or errors. The source working tree was clean when the campaign was
prepared; subsequent documentation edits do not alter its frozen files.

`artifacts/q026-20261003-native/managed/managed.trx` records **68 passed,
0 failed, 0 skipped**, covering the private-desktop and Format lifecycle guards.
This is focused managed validation, not coverage or native acceptance.
The one-shot campaign has not dispatched a native Format command: it waits
for another qualification's retained Access PID 6880, start
`2026-10-02T22:27:55.5717465Z`. The owning Q-012 receipt records a returned Quit
without observed process exit; Q-026 neither replays its cleanup nor terminates it.
Once that owner finishes recovery, the frozen worker additionally requires
30 continuous seconds without another VBE host before its own Excel launch.
Q-026 remains **OPEN / native NOT_RUN** at this preparation checkpoint.

### Q-026 native Format failure and recovery (2026-10-03)

The frozen `180aee8` campaign subsequently executed the complete Format method
once on the inactive desktop `VBAiTests_24227bdb782d40e087294b262600a190`.
Excel PID 32152, start `2026-10-02T22:42:37.0468311Z`, version
`16.0.20430.20092`, loaded the preparation candidate identified above.
`artifacts/q026-20261003-native/native/format.trx` records **0 passed, 1 failed,
0 skipped**. Durable receipts verify baseline read stability, font and three
normal-text palettes, another-category foreground, and the empty-size refusal
with independently observed closure and unchanged complete state. The margin
write then returned `Ok=false` / an invalid-object-state error; its Options
dialog remained open. No margin retry, stale-revision trial, automatic
compensation or campaign Close/Quit followed that unknown write outcome.
Full restoration and qualification are therefore failed in that TRX.

Independent recovery captured the exact owned Options dialog, cancelled it once
on a private-desktop thread and verified its absence. An initial default-desktop
cancellation guard refused before any native message; that refusal was not a
lost cancellation. The first compensation attempt likewise refused before any
write because Windows PowerShell interpreted a UTF-8 receipt as ANSI. With
explicit UTF-8 decoding, `recovery-utf8/terminal.json` proves the five positively
committed entries returned to baseline, including every tab/category/catalogue
and revision `544477535d391a457fc66947a3ddb3da7ffb430e6212289c6586a6fcb65f3e05`.
The failed margin remained at its original `On` value and was not replayed.
`recovery-shutdown.json` records one disposable-seed Close, one returned Quit
and normal exit 0 without forced termination. This uses a fresh recovery handle;
it does not qualify the original campaign's retained-handle shutdown.

The corrective source replaces Options-checkbox UIA Toggle with a guarded
two-state Win32 Button read/click/read sequence. A matching value sends no click;
an unreadable initial state refuses before dispatch; a failed or unmatched
readback never triggers another click. The initial isolated-build
`q026-checkbox-fix-tests/checkbox-guards.trx` records **11 passed, 0 failed,
0 skipped** for the pure transition/failure contracts. The broader private-desktop regression
record `q026-checkbox-dispatch-unit-v2/results/checkbox-native-guards.trx` has
**201 passed, 1 failed, 0 skipped**: its new synthetic button fixture exposed an
ambiguous duplicate accessible name. After correcting only that fixture,
`q026-checkbox-dispatch-unit-v3/results/checkbox-focus.trx` records **80 passed,
0 failed, 0 skipped**, including the real Win32 transitions and lifecycle guards.
The solution builds without warnings/errors in `q026-checkbox-candidate`.
These are local control/managed tests, not acceptance of the repaired Excel path.
The new checkbox-only Excel method must pass on a fresh owned host before a
complete Format qualification is attempted. Q-026 remains **OPEN** here.

### Q-026 corrected-candidate prelaunch guards (2026-10-03)

Source `5ab57ea84beebd47934ba5fc2da8e639bc2d050a` prepared a margin-only
diagnostic with candidate MVID `24d4809c-d427-4412-b746-793d59fc056a`, SHA-256
`37D06CC3D25CC455A3AE6F128756921003577188E1B740882CF17A25DCD6AC48`.
The separate `q026-margin-native/managed/managed.trx` and
`q026-margin-native-v2/managed/managed.trx` each record **81 passed, 0 failed,
0 skipped** for their focused managed gate. Both workers then exited with code
1 before any native launch. The diagnostic worker receipt in the second run
identifies a registered-CodeBase mismatch; no Excel or preference write occurred.
An independent limited-worker observation saw the original product path,
although the development shell read the candidate path. This is a registration
context observation, not a proven explanation of historical Options revision drift.

The development-context registration was restored from its exact backup.
A subsequent limited-context activation refused before mutation because another
campaign had just opened Word PID `42988`. Its ownership remains with that
campaign; Q-026 neither closes it nor overwrites its registration. The repaired
Excel path remains **NOT_RUN** pending an exclusive native-test interval.

### Q-026 candidate loading diagnostic (2026-10-03)

The subsequent `q026-margin-native-v3` campaign at source `639b488` used the
same corrected candidate identified above. Its managed gate recorded **81 passed,
0 failed, 0 skipped**. The native TRX recorded **0 passed, 1 failed, 0 skipped**:
owned Excel PID `26764`, start `2026-10-03T08:16:27.2217572Z`, remained in add-in
loading before any preference dispatch. An independently read VBA dialog said
the add-in could not load and offered its removal. Explicit recovery answered
No once, preserving registration; bootstrap then failed terminally because the
bridge was unavailable. No Format scenario is qualified by that run.

The independent class-activation diagnostic returned HRESULT `0x80070002` and
loaded no VBAi assembly with the helper's percent-encoded CodeBase. The helper
now matches `Install-VBAi.ps1` and the other repository registration tools by
preserving the unescaped Unicode `file:///` path. After guarded restoration and
reapplication, `q026-class-activation-v2.json` recorded successful activation of
the exact candidate MVID and path; its private worker exited normally. This
diagnoses the qualification helper's loading error, not historical revision drift.

`q026-margin-native-v3/bootstrap-shutdown.json` records normal recovery exit 0,
one returned seed Close and one returned Quit, without termination. This
bootstrap-only recovery required terminal worker/boot failure, no preference
receipts, dialog absence, the exact saved seed and its unchanged SHA-256. An
earlier exclusive file-hash read refused before Close; the shared read then
verified the same seed hash. The fresh recovery handle does not qualify the
original failed campaign's shutdown lifecycle. All Excel processes were absent
after recovery. The next native diagnostic remains separately **NOT_RUN** here.

At source `af7f352`, `q026-margin-native-v4` recorded **81 managed passed,
0 failed, 0 skipped** and **0 native passed, 1 failed, 0 skipped**. Owned Excel
PID `15268`, start `2026-10-03T08:30:50.6679365Z`, failed terminally before Format
dispatch because the bridge was unavailable. Independent limited-context
inspection then found the existing VBAi LoadBehavior DWORD was 0. This run did
not test the repaired preference path. Its `bootstrap-shutdown.json` verifies
one returned seed Close, one returned Quit and fresh-handle normal recovery
exit 0; all earlier failed results remain failed.

The qualification activation helper now supports an explicit, reversible
autoload opt-in, preserving the observed DWORD 0 in
`q026-limited-registration-v3.clixml` before enabling only the installed VBAi
add-in. Execution checks that value before native launch. No Office trust
setting or registered identity is changed. The next diagnostic must verify
actual loaded-candidate identity before any preference scenario is accepted.

### Q-026 current native acceptance (2026-10-03)

Source `9400d654a11affaa931271cc4efcdc4dde05a80c` freezes both final campaigns
on clean branch `codex/q026-qualification`. Both use product MVID
`24d4809c-d427-4412-b746-793d59fc056a`, SHA-256
`37D06CC3D25CC455A3AE6F128756921003577188E1B740882CF17A25DCD6AC48`, and
Excel `16.0.20430.20092`. The loaded native MVID and all frozen file hashes were
verified independently of registration readback.

| Campaign / recorded artifact | Actual result | Exact scope |
| --- | --- | --- |
| `artifacts/q026-margin-native-v5/managed/managed.trx` | 81 passed, 0 failed, 0 skipped | Focused managed gate, not the full repository suite |
| `artifacts/q026-margin-native-v5/native/format.trx` | 1 passed, 0 failed, 0 skipped | Fresh owned Excel PID 44076, start `2026-10-03T08:36:34.7590580Z`; real On-to-Off margin transition, independent readbacks, exact complete restoration and normal exit |
| `artifacts/q026-full-native-v3/managed/managed.trx` | 81 passed, 0 failed, 0 skipped | Same focused managed gate for the independently launched complete campaign |
| `artifacts/q026-full-native-v3/native/format.trx` | 1 passed, 0 failed, 0 skipped | Fresh owned Excel PID 22396, start `2026-10-03T08:44:40.2579485Z`; full planned Format matrix, refusals, restoration and normal exit |

The complete campaign verifies initial full-state stability; font, foreground,
background and indicator mutations; another category's foreground; a real margin
transition; and an intentionally stale revision refusal. Its size catalogue is
empty: only the expected refusal is accepted, **not a size mutation**. Both
refusals have independently observed dialog absence and structurally unchanged
complete readbacks. All six positively committed entries were restored in
reverse order, with explicit categories and fresh revisions. No uncertain native
action was replayed.

The independent `complete-restoration-independent.json` comparisons find every
tab, control, catalogue and palette equal to baseline, with revision
`544477535d391a457fc66947a3ddb3da7ffb430e6212289c6586a6fcb65f3e05` and matching
PID/start/MVID metadata within each campaign. The margin comparison finds only
`/Tabs/1/Controls/11/Value` changed On to Off. The separate full-campaign font
comparison finds only `/Tabs/1/Controls/8/Value` changed. The restored margin
snapshot also matches the next fresh Excel's complete initial snapshot; this is
an actual observation across normal exit and a new host, not a restart claim
derived from `CommitRequested` alone.

Each fixture retains its original native launch handle and asserts its normal
exit independently of the separately queried process. Both shutdown receipts
record exit 0, returned Close/Quit and no forced termination. Both private desktop
terminals report original-worker exit 0 with input desktop `Default` and no
desktop switches. Terminal qualification receipts have no primary, restoration
or shutdown error. No SOLIDWORKS or other-host acceptance is implied.

`artifacts/q026-release.json` records release at `2026-10-03T09:15:13.5259898Z`,
after owned Excel exit and absence of competing hosts. The limited-context
restore verified the exact prior Q-024 CodeBase and LoadBehavior DWORD 0.
The current instrumented Format matrix is **ACCEPTED**. Q-026's historical
stale-revision cause remains **UNPROVEN**: the old failed run is not converted
to a pass by these new results, and its stricter causal-closure criterion remains
open. Further blind repetitions would not supply its missing historical receipts.

### Q-026 historical binary diagnostic (2026-10-03)

The original failed Format test retained a complete initial baseline in
`qualification-v1/activated-excel`, revision
`544477535d391a457fc66947a3ddb3da7ffb430e6212289c6586a6fcb65f3e05`, file SHA-256
`EA0C00167ECB4431DA0E06FA1D0EAA386722D63B7E5EBF1B7B7F062A7F196DDF`.
Its test fixture checked the loaded product MVID before that capture. The frozen
original product is MVID `5cc513d1-5569-4835-bf6c-cf70a18274fb`, SHA-256
`A12378FB826CAAD6C0BBE79BF09C52332760E6C32FDB4F65138876CB90F5F167`.
Portable-PDB document checksums match the relevant options implementation and
Format test at source `5c860a3`; `q026-original-source-provenance.json` retains
those hashes. The exception stack points to the test's catch/rethrow line,
not to a particular preference. Later recovery snapshots show font and normal
text foreground changes, but are not the two snapshots compared at the failing
revision guard. Those guard snapshots remain unavailable.

Source `eb89a3a` introduces a test-only, opt-in CLR collector for the unchanged
historical binary. `q026-historical-trace-preflight/trial-08/preflight.json`
verifies a complete synthetic options graph, exact request fields, debugger
detach and normal helper exit. This preflight activates no Office host, performs
no target function evaluation and changes no preference. Framed captures are
decoded offline and can be hashed using the original .NET Framework
JavaScriptSerializer/UTF-8/SHA-256 algorithm. Empty or incomplete traces cannot
establish revision drift.

| Source / retained artifact | Actual result | Scope |
| --- | --- | --- |
| `eb89a3a`, `q026-historical-prefix-campaign-v1` | Prelaunch refused; no native scenario | Canonical registered path differed from a noncanonical plan path; guarded registration release, no preference writes |
| `e1e158c`, `q026-historical-prefix-campaign-v2/managed/managed.trx` | 71 passed, 0 failed, 0 skipped | Historical-compatible managed subset; excludes the later native-checkbox implementation |
| `e1e158c`, `q026-historical-prefix-campaign-v2/native/format.trx` | 0 passed, 1 failed, 0 skipped | Owned Excel PID 37156, start `2026-10-03T10:23:57.9740740Z`; scenario stopped before palettes when immediate native enumeration still found Options |
| `7203f492`, `q026-recovery-validation/managed/managed.trx` | 71 passed, 0 failed, 0 skipped | Same historical-compatible subset; both diagnostic projects build without warnings or errors |
| `7203f492`, `tools/probes/tests/test_read_q026_clr_guard.py` | 6 passed, 0 failed | Offline framing, Unicode, truncation and collector-error checks; no native acceptance |
| `7203f492`, `q026-historical-prefix-campaign-v3/managed/managed.trx` | 71 passed, 0 failed, 0 skipped | Historical-compatible managed subset on the frozen original product |
| `7203f492`, `q026-historical-prefix-campaign-v3/native/format.trx` | 1 passed, 0 failed, 0 skipped | Owned Excel PID 39752, start `2026-10-03T10:46:21.2979545Z`; font, honest size refusal and foreground/background/indicator without Query, four reverse compensations, complete restoration and original-host-handle exit 0 |

The failed native prefix loaded the exact historical MVID on inactive desktop
`VBAiTests_6bcc4563e670469d823b4ece448ef906`. Its sole positively confirmed
preference change was the font. The expected size refusal occurred before any
size write; independent enumeration found a visible owned Options dialog at
that instant. A later read-only enumeration verified its absence without Cancel
or further input. The detached trace contains no complete stale-guard capture.
The first recovery precheck stopped before any write because it counted both
the size label and combo box. The corrected precheck selected the unique editable
control and verified its unchanged baseline value before the font compensation.

`recovery-font-once-v2/terminal.json` verifies the complete baseline revision and
tab structure after one positively confirmed font compensation, with no replay
of the failed size operation. `recovery-shutdown.json` records one returned
seed Close, one returned Quit and normal exit 0 through a fresh recovery handle;
it is not original-launch-handle lifecycle acceptance. The debugger detached
normally with exit 0. `registration-release.json` records release at
`2026-10-03T10:43:09.3011741Z`, restoring the prior Q-024 CodeBase and autoload
setting. The old desktop helper recorded retained private handles after the
failed campaign; the recovery issued no forced termination. Source `7203f492` allows future
helpers to release retained handles after independently observing original-child
exit and an empty private desktop, without replaying native cleanup.

The subsequent prefix starts from the same complete baseline on inactive
desktop `VBAiTests_5eedf5ed3b2b4e4b8d3579cbf532a4e2`. Its size refusal observes
Options still present, then absent after 54 ms across two native enumerations,
without bridge requests or input in between. All three palette changes preserve
the observed normal-text category. The independent
`complete-restoration-independent.json` comparison verifies complete Tabs and
revision equality, with identical recorded PID/start/MVID metadata. Its terminal
has no primary, restoration or shutdown error; the fixture verifies exit 0 on
the original Excel launch handle. The CLR collector detaches normally with exit
0 and yields no stale-guard capture. The product SHA-256 remains unchanged.
`q026-historical-wrapper-v3.json` verifies guarded registration restoration and
absence of Excel at `2026-10-03T11:06:59.0147355Z`.

This diagnostic prefix includes additional independent reads, a debugger and
inactive-desktop isolation; it does not reproduce the original run's timing or
qualify the original full matrix on the old product. Both external scheduled
launchers lack a terminal receipt and later report `0xC000013A`; their stopping
cause and original-worker normal exit remain unverified. This is separate from
the native test runner's verified result and original Excel-handle exit. No
cleanup is replayed and no global keyboard input or desktop switch is used.

Q-026's historical causal criterion remains **OPEN**. Delayed Options closure
does not establish why the original complete revision differed, and recovery
does not turn the failed native prefix into a pass.

### Q-026 diagnostic crash and targeted capture (2026-10-03)

Historical campaign `q026-historical-prefix-campaign-v4`, source `474d4c2`,
loads the unchanged original product on inactive desktop
`VBAiTests_285a6e65f7664bf38cceec8f1357fc18`: Excel PID 17064, start
`2026-10-03T16:06:40.7009034Z`. The font and three palette writes pass their
readbacks; three palette compensations complete. The subsequent
`read_vbe_options` request expires after emission at
`2026-10-03T16:24:19.9324626Z`. No font compensation, cleanup or retry is sent.
The CLR trace has no complete stale-guard capture and records an access
violation. Windows Application Error event 1000 independently identifies
the same PID/start, exception `0xC0000005`, and EXCEL.EXE offset `0x10CFBDE`.
The owning worker exits with code 1. Its direct GUI launcher and private-desktop
terminal observe that original-worker exit normally; they do not establish a
normal exit for the crashed Excel. Windows also records a Publisher crash
during the native interval, so exclusive host use throughout that trial is
not established. Neither crash cause is inferred from these events.

Office subsequently starts Excel PID 42408 with `/restore` on desktop `Default`.
After explicit user reservation for Q-026 recovery, exact native attachment
finds only the unchanged saved disposable seed. One Close and one Quit return;
the fresh recovery handle observes exit 0. This successor is not the original
campaign host and cannot supply its missing lifecycle evidence.

`q026-historical-prefix-campaign-v4/separate-owned-recovery` owns a further
isolated Excel PID 12836, start `2026-10-03T16:49:11.7267442Z`, on desktop
`VBAiTests_11784aa1e68847bfaa1de1b91f99f792`, with the exact historical MVID.
The new host's complete snapshot differs from the original baseline only in
the positively committed font. Three palette entries already match and are
not written again. One font compensation with a fresh revision, independent
readback and complete Tabs equality restores baseline revision `544477...`.
Its original native launch handle observes exit 0; the original private worker
and direct GUI launcher also exit 0 without desktop switches. At
`2026-10-03T16:50:55.9361372Z`, preference recovery and guarded restoration of
the previous Q-024 registration are complete. These recovery receipts explicitly
keep the original campaign failed.

| Source / artifact | Actual result | Scope |
| --- | --- | --- |
| `474d4c2`, `q026-historical-prefix-campaign-v4/managed/managed.trx` | 71 passed, 0 failed, 0 skipped | Historical-compatible managed subset |
| `474d4c2`, `q026-historical-prefix-campaign-v4/native/format.trx` | 0 passed, 1 failed, 0 skipped | Historical prefix; restoration read expires and owned Excel crashes |
| Intermediate working tree after `474d4c2`, `q026-v4-followup-build/managed/managed.trx` | 73 passed, 0 failed, 0 skipped | Host-exclusivity guards before the later warmup additions; test assembly SHA-256 `F249020528DC07A83CA801B736768160BE718B7ECCB25C7C6C3AED31DD4E4EBE` |
| Intermediate working tree before `f9f5ccc`, `q026-guard-il-harness-build/managed/managed.trx` | 74 passed, 1 failed, 0 skipped | Retained failure before the synthetic warmup-model correction; test assembly SHA-256 `75FC2E03C30E8081E1C6DEBC5CF3526270700D2DADE62B2DBAE7208C957213CB` |
| `f9f5ccc`, `q026-guard-il-harness-green/managed/managed.trx` | 75 passed, 0 failed, 0 skipped | Corrected focused lifecycle model, exclusivity and nonmutating guard warmup |

Source `f9f5ccc` checks exact ownership and competing VBE hosts before every
Format dispatch. A changed interval retains the host and forbids further reads,
writes, compensation and cleanup. Refusals require complete structural equality
as well as unchanged revision. No production revision check is relaxed.

The earlier exception collector stops on all CLR exceptions before filtering
the message and substantially alters timing. A new mode binds only the exact
stale-guard IL branch in the frozen binary. Pure metadata inspection identifies
`ldstr` at IL offset 233 (`0xE9`); the collector verifies its literal, product
MVID/hash, uniquely bound native address and complete matching synthetic
preflight before attachment. Normal UIA exceptions are not intercepted. The
guard warmup uses one intentional refusal targeting an absent property and
verifies dialog destruction and complete unchanged state before attachment.
Debugger breakpoints are diagnostic code changes, not preference writes or
target function evaluations.

`q026-guard-il-preflight-v1` retains a failed PowerShell JSON comparison after
a complete capture; an independent structural comparison finds no differing
fields. Windows PowerShell had decorated the root array with ETS properties.
The corrected comparison uses the original .NET Framework serializer on both
JSON graphs. `q026-guard-il-preflight-v3/preflight.json` verifies the complete
synthetic snapshot and request at the exact IL branch, with debugger detach and
original-helper exit 0. It activates no Office host. The new mode is prepared;
it does not establish the historical revision drift or qualify native Office.

The targeted capture follows the primary [SOS documentation on GitHub](https://github.com/dotnet/diagnostics/blob/main/src/SOS/Strike/sosdocs.txt)
and Microsoft's [breakpoint command documentation](https://learn.microsoft.com/en-us/windows-hardware/drivers/debuggercmds/bs--update-breakpoint-command-).
Microsoft's [UI Automation threading guidance](https://learn.microsoft.com/en-us/windows/win32/winauto/uiauto-threading)
requires a separate MTA client thread for interactions with the client's own UI;
the bridge already uses a separate worker. The [VBE-Themes project's registry mapping](https://github.com/vicsar/VBE-Themes)
also documents common VBE settings, supporting the need for an exclusive host
interval. These references guide the investigation; they do not prove either
the original stale revision or this later crash. Q-026 remains **OPEN**.

### Q-026 targeted native capture interrupted by another host (2026-10-03)

Source `35d3c429070e4b184d8811b8035a1bb5b67c6d99` freezes
`q026-historical-prefix-campaign-v5` on the unchanged historical product. Its
managed subset records **75 passed, 0 failed, 0 skipped**, and its native
scenario records **0 passed, 1 failed, 0 skipped**. Excel PID 10228, start
`2026-10-03T17:26:47.2289879Z`, is owned on inactive desktop
`VBAiTests_2caea500bf17437faa8129082c137de8`.

The nonmutating guard warmup verifies the known refusal, actual dialog absence
and a complete unchanged readback. The collector binds the exact native
stale-guard branch and reaches ARMED with `ExactGuardILBreakpoint` mode.
The matrix captures its complete baseline, then the next exclusivity observation
finds competing Access PID 41756, start `2026-10-03T17:28:41.747923Z`.
The campaign stops before the font pre-write read and before every matrix
preference write. The positive commit ledger is empty; no unknown native
operation is replayed. The collector detaches normally with exit 0. This real
guard stop validates the protection, not the historical cause or the interrupted
matrix. The external GUI helper retains the private desktop until recovery.

After Access disappears, `isolation-stop-preferences/terminal.json` verifies a
fresh complete baseline read on that same owned Excel, with no compensation
write. `isolation-stop-shutdown.json` records one returned Close, one returned
Quit and exit 0 through a fresh recovery handle; it is not original-campaign
handle proof. `isolation-stop-release.json` confirms host and previous
registration release at `2026-10-03T17:33:11.8109153Z`. The original-worker
terminal exits with code 1 and keeps the native scenario failed. Recovery uses
the narrowly checked pre-dispatch isolation-stop path subsequently added to
`Recover-Q026Preferences.ps1`. An exclusive interval has been requested before
another native attempt. Q-026 remains **OPEN**.

### Q-026 reserved-slot interruption before the scenario baseline (2026-10-03)

Source `f9a6d9a80d82a107d12146207f9a82bbcb82de15` freezes
`q026-historical-prefix-campaign-v6` on the same unchanged historical product.
Its managed subset records **75 passed, 0 failed, 0 skipped**; the native
scenario records **0 passed, 1 failed, 0 skipped**. Owned Excel PID 33088,
start `2026-10-03T17:43:45.2020947Z`, runs on inactive desktop
`VBAiTests_9e52898a7a24477ba494469fbcd7991d` with the exact frozen loaded MVID.

The nonmutating warmup captures complete revision `544477535d391a457fc66947a3ddb3da7ffb430e6212289c6586a6fcb65f3e05`,
verifies the expected stale refusal, dialog absence and unchanged full readback.
The collector arms the exact guard branch. Publisher PID 40760, start
`2026-10-03T17:45:17.9813208Z`, then appears during the explicitly reserved
interval. The immediate pre-dispatch observation stops the scenario before its
baseline read, with a null failed request and an empty commit ledger. No
preference write is issued. Collector detachment is verified with exit 0; no
natural stale-guard capture occurs. The owned host and registration are initially
retained pending competing-host release and an independently verified shutdown.

The recovery validator now narrowly supports this earlier stop using the exact
complete verified warmup readback. The matching offline PowerShell evidence
gates in source `c5ff8fb` record **17 passed, 0 failed**, with no Office activation or native
dispatch (`q026-warmup-recovery-gates.json`). They refuse foreign PID/start/MVID,
missing or changed structures despite equal hashes, uncertain replies, known
writes, intervening native dispatch, positive ledgers, unknown requests/errors,
unverified closure and ambiguous receipts. The native failure stays failed;
Q-026 remains **OPEN**.

After Publisher is released, `isolation-stop-preferences/terminal.json` records
a fresh complete readback equal to the verified warmup baseline. No compensation
write is sent. `isolation-stop-shutdown.json` observes one Close, one Quit and
normal exit 0 through the fresh recovery handle, not the original campaign
handle. Host and previous registration release complete at
`2026-10-03T20:42:37.6784270Z` (`isolation-stop-release.json`). The original
GUI worker exits with code 1 and preserves the failed native scenario; the
separate recovery GUI worker exits with code 0. No forced termination occurs.

### Q-026 inter-case gap and normal release (2026-10-03)

Source `047d8c7b011c688e6e7f8064d407278fdd1b85fc` freezes
`q026-historical-prefix-campaign-v7` on unchanged historical MVID
`5cc513d1-5569-4835-bf6c-cf70a18274fb`. The focused managed subset records
**75 passed, 0 failed, 0 skipped**; the native scenario records
**0 passed, 1 failed, 0 skipped**. The campaign first waits for Q-012 Access
PID 191444, then observes a quiet inventory interval and starts before the
user's response establishes that Q-012 is still running. A quiet interval
between native cases is not release of the other qualification batch.

Owned Excel PID 136028, start `2026-10-03T20:46:39.6119508Z`, runs on inactive
desktop `VBAiTests_7d0d627923dc4fe7aa8f8446504ca342`. Its guard warmup and
complete baseline are verified. Publisher PID 104540, start
`2026-10-03T20:48:21.0478323Z`, appears during that baseline read. The next
pre-dispatch observation stops before the font pre-write read and all matrix
preference writes, with an empty commit ledger. The collector detaches normally
with exit 0; no natural stale-guard capture occurs.

After that Publisher process disappears, a fresh complete recovery read equals
the retained baseline, with no compensation. One Close and one Quit return;
the fresh recovery handle observes normal exit 0. Host and prior registration
release complete at `2026-10-03T20:50:50.7553273Z`. The original GUI worker
exits 1 and the separate recovery worker exits 0; no process is forcibly ended.
The failed native scenario is preserved. No further native campaign is scheduled
pending the user's explicit confirmation that Q-012 has ended. Q-026 remains
**OPEN** for the historical causal criterion.

### Q-026 content-free combo observations (2026-10-04)

Source `16ffa13` adds optional native/UIA Options combo observations through
the existing bounded inspection trace. The isolated Debug/net48/x64 production
candidate is MVID `9dc786a9-7128-41db-b21e-5bce14c0935f`, SHA-256
`9F005722C6ED4749F428AD44E111FA4917EC468EC6A1990B96B0C311CA11C7CB`.
At this checkpoint it is **not installed or native VBE-qualified**. The first test-project build
fails for missing test namespace imports; the corrected build passes without
warnings/errors (`q026-combo-observation-build-final.log`).

`q026-combo-observation-managed-v1/combo.trx` records **34 passed, 0 failed,
0 skipped**. The scope includes the Options revision/write guards, bounded
inspection trace, real disposable standard Win32 combos and simulated UIA
providers. A genuinely empty list records zero native entries before/after its
single expansion; a separately simulated native owner populates its list on
the documented dropdown notification, and the existing reader sees those
entries without another opening or any selection/edit notification. Empty,
populated and fallback paths preserve their exact values and close only the
temporary list they opened. Unavailable logging cannot suppress the native
ownership refusal; mixed trace events retain the existing cap and exclude
exception messages, control labels and choice values.

The managed worker runs on inactive desktop
`VBAiTests_2bb0865561ae456483fff2311c51073c`, with no Office activation,
input-desktop switch or owned foreground observation. Its original child and
GUI launcher both exit normally with code 0; its scheduled launch task is
exported and removed. The user's existing SOLIDWORKS is untouched. Native
font-size/catalogue diagnosis and the historical revision-drift acceptance
remain **NOT_RUN / OPEN** pending an explicitly released host interval.
These observations add evidence for the next native trial, not a causal
correction or a new coverage percentage. Notification behavior is documented
by [Microsoft](https://learn.microsoft.com/en-us/windows/win32/controls/cbn-dropdown)
and its [published Win32 documentation source](https://github.com/MicrosoftDocs/win32/blob/docs/desktop-src/Controls/about-combo-boxes.md).

### Q-026 focused font/size diagnostic preparation (2026-10-04)

Harness source `edb8e81` adds `-Scenario FontSizeCatalogue`. It uses the same
exact-ownership, revision, positive-commit ledger, complete-restoration and
normal-exit gates as the full Format matrix, while excluding palette, margin
and deliberate stale-revision mutations. Both real catalogue choices and an
empty-catalogue refusal are exercised in the managed lifecycle model.

`q026-font-size-catalogue-native-v1/managed-preflight/managed.trx` records
**88 passed, 0 failed, 0 skipped** against the frozen `16ffa13` product
(MVID `9dc786a9-7128-41db-b21e-5bce14c0935f`). This focused harness build and
the launcher build have no warnings/errors. These are managed/model and owned
standard-control tests, not native VBE acceptance.

The preflight runs on inactive desktop
`VBAiTests_62e75c8c3c354a189aad35aa0bf307b1`. The original worker and GUI
launcher exit with code 0; input-desktop switches and owned foreground
observations are zero. The scheduled task is exported and removed. No Office
host is activated and registration is unchanged. The frozen native plan and
one-shot wrapper are prepared but **NOT_RUN**, pending explicit release of the
native interval. Historical causality and real font-size mutation remain
**OPEN**; the existing full-matrix acceptance belongs to its earlier candidate.

### Q-026 native size catalogue and historical ownership stop (2026-10-04)

After explicit user reservation, harness `38ee39f` runs the frozen `16ffa13`
product (MVID `9dc786a9-7128-41db-b21e-5bce14c0935f`) on inactive desktop
`VBAiTests_58d9f0494ad14604aece8d5b3b6e6a39`. The preceding v1 wrapper
refuses registration before any Office activation because the actual limited
worker has LoadBehavior 2. The explicit loading opt-in now accepts that existing
DWORD and restores it exactly; no registration is created or trust policy changed.
The elevated reader's different registration values are not substituted for the
worker's actual values.

`q026-font-size-catalogue-native-v2` records **88 managed passed, 0 failed,
0 skipped; 1 native passed, 0 failed, 0 skipped**. Owned Excel PID 46072,
start `2026-10-04T09:10:49.0020351Z`, Office `16.0.20430.20092`, verifies
font mutation, closed independent readback, empty-size refusal without replay,
complete restoration and exit 0 from the original retained launch handle 2112.
All Tabs independently match the original revision `544477...`; 51 phase
receipts and 14 command records are retained. The original worker and GUI
launcher exit 0 with no desktop switch or foreground observation. Registration
is restored and the terminal launch tasks are exported and removed.

The new trace contains 20 actual native Size observations: control ID 4911,
style `0x50010302`, native string combo, count 0 before and after its single
expansion, selected index -1, expanded state true and final expanded state false.
All reads complete without a native error and belong to PID 46072. This occurs
before and after the font change. The empty catalogue is therefore observed
through Win32, not merely missing UIA descendants. No size entry is invented;
**size mutation remains unqualified**. Independent evidence is retained in
`q026-font-size-native-v2-audit.json`. This is scoped font/refusal acceptance,
not full Format acceptance on this diagnostic candidate or historical causality.

The separate historical v8 campaign uses unchanged product MVID
`5cc513d1-5569-4835-bf6c-cf70a18274fb` and passes **77 managed tests** but
fails its single native test during the known pre-write guard warmup. Owned
Excel PID 38836 starts `2026-10-04T09:29:28.6293854Z`. The exact expected
guard refusal is received, but window enumeration cannot read a window owner;
it does not establish retained Options visibility. No preference commit or
collector attachment occurs. An independent later observation confirms Options
absence. This failed case is retained and is not a natural stale-drift reproduction.

Source `388c916` permits recovery only for this identified historical binary,
the exact sentinel request and terminal pre-write refusal, unique complete
baseline, unchanged ownership and an empty commit ledger. It still requires
independent closure and a fresh complete matching snapshot before cleanup.
Its pure-data validation records **25 passed**; the existing recovery-baseline
gates record **17 passed**, with no native dispatch. Actual recovery reads the
full unchanged `544477...` baseline without a compensation write, closes the
owned seed once and calls Quit once; a separately retained recovery handle
observes exit 0. The wrapper then fails its postcondition because it checks an
absent `Exited` field instead of the receipt's explicit normal-exit state.
Cleanup is not replayed. A registration-only actor verifies the existing receipt
and restores the original registration at `2026-10-04T09:42:38.4397840Z`.
Neither recovery nor this outer-wrapper correction qualifies the original test.

A passive observation during v8 records an enabled Size combo on its owning
dialog while the Code Colors list has focus. It sends no input or bridge call;
this does not prove focus caused the empty catalogue. An earlier observer runs
after PID 46072 has already exited and returns no rows; it is not live evidence.
Historical revision drift and real size mutation remain **OPEN**.

### Q-026 exact historical guard campaign (2026-10-04)

Campaign `q026-historical-prefix-campaign-v9` uses harness `fd183b0` and the
unchanged historical product MVID `5cc513d1-5569-4835-bf6c-cf70a18274fb`
(SHA-256 `A12378FB826CAAD6C0BBE79BF09C52332760E6C32FDB4F65138876CB90F5F167`).
It records **83 managed passed, 0 failed, 0 skipped; 1 native passed, 0 failed,
0 skipped**. Owned Excel PID 69328 starts `2026-10-04T09:49:33.4942230Z`
on inactive desktop `VBAiTests_900b63e8d1a14661b713c435714629b1`.

All four historical font/palette commits and their readbacks succeed. All Tabs
and complete revision `544477...` are restored, independently audited in
`q026-historical-v9-audit.json`. Evidence includes 126 phase receipts and 36 command
records. The exact guard breakpoint at IL offset 233 is armed; debugger PID 77856
detaches with exit 0 before cleanup. The original Excel launch handle 2224 observes
exit 0, and both original worker and GUI launcher terminate normally without an
input-desktop switch. Registration is restored; the terminal task is exported
and removed. The three terminal v8 tasks are separately exported and removed.

The retained log contains **zero natural guard captures**. The offline decoder
previously treated its fixed `Q026_GUARD_IL_ARMED` status receipt as JSON; it now
accepts only that exact receipt and still rejects malformed or interrupted frames.
Its focused offline suite records **9 passed**. This decoder correction does not
change the historical binary or create evidence of a guard hit. The intentional
pre-attachment warmup refusal is not a natural revision-drift reproduction.
Historical causality and real size mutation remain **OPEN**.

The next dedicated `SizeFocus` diagnostic is implemented and builds against the
frozen content-free observation candidate. It makes no preference write and does
not yet have native acceptance at this checkpoint. Its single dialog-local focus
message follows [Microsoft's WM_NEXTDLGCTL contract](https://learn.microsoft.com/en-us/windows/win32/dlgbox/wm-nextdlgctl);
its catalogue expansion uses [CB_SHOWDROPDOWN](https://learn.microsoft.com/en-us/windows/win32/controls/cb-showdropdown).
Neither message establishes that this VBE build populates Size choices; the native
observation must decide that hypothesis.

### Q-026 native Size focus and current-candidate preparation (2026-10-04)

`SizeFocus` harness `1d2e107` against frozen product MVID
`9dc786a9-7128-41db-b21e-5bce14c0935f` records **94 managed passed, 0 failed,
0 skipped; 1 native passed, 0 failed, 0 skipped** in `q026-size-focus-native-v2`.
Owned Excel PID 176968 starts `2026-10-04T10:31:39.6471352Z` on its inactive
desktop. The actual native catalogue changes from zero entries to nine after
one dialog-local focus operation, then remains at nine after expansion. Its
observed choices are `8, 9, 10, 11, 12, 14, 16, 18, 24`. The edit value remains
`10`. This proves focus-dependent population on this host/build; the values
are observations, not a portable allowed-size range.

The Options menu executes once on the fixture's owning STA. The separate private
worker selects the observed Format tab, posts one `WM_NEXTDLGCTL`, observes the
owning GUI thread's focus and cancels once. The exact dialog is destroyed.
All Tabs and revision `544477...` remain unchanged. There are 22 phase receipts;
original launch handle 2176 observes normal exit 0. Worker/launcher exit 0,
registration is restored, and `q026-size-focus-native-v2-audit.json` independently
checks the complete baseline and action receipts. No preference write occurs.

The preceding v1 at `facd68a` retains **1 native failed** before any Options
invocation because the harness sends the unsupported name `vbe_options` instead
of `read_vbe_options`. Its exact inert refusal is retained, not counted as native
acceptance. A recovery actor's first attempt stops before dispatch because it
expects an outer terminal that is pending while Excel remains open. Its next
attempt performs one complete read, then rejects its comparison because Windows
PowerShell decodes an unmarked UTF-8 baseline as ANSI. An independent structured
comparison finds zero differing Tabs and equal complete revisions. The final
actor reads those existing receipts explicitly as UTF-8, observes dialog absence
and closes the owned workbook/Quit once, with exit 0 from a recovery handle.
Neither read nor cleanup is replayed; registration is restored. The original
native failure remains failed. These terminal tasks are exported and removed.

Production `c1cb2c3` prepares only an empty recognized native Size control. It
validates ownership, parent dialog, enabled state and unchanged edit value,
dispatches to the owning GUI thread once and requires bounded focus observation.
The unlabelled write verifier also recognizes observed control ID 4911. No size
value is invented and the exact-choice/revision guards remain enforced.
Content-free opt-in trace metadata adds focus-attempt and post-focus count fields.

The final focused suite at `08d7420` records **9 managed passed, 0 failed,
0 skipped**, including the real owned Win32 focus/population regression and
unlabelled write verification with unknown-choice refusal. Build completes with
zero warnings/errors. Frozen current product MVID
`7156af5b-941c-4452-9e78-1381cb69af0d` has SHA-256
`950D86EDEF3F90DC3A8EE274627AC75C08012BF158F55140817ECEC9FABF85EC`.
The single complete Format campaign `q026-full-size-native-v1` is started after
that gate, including real Size mutation and restoration; its native acceptance
is **PENDING** at this checkpoint. Historical causality remains **OPEN**.

### Q-026 offline snapshot review (2026-10-03)

Branch `codex/q026-qualification` starts from `origin/main` at `4b382b9`,
independently of the complete Q-006 publication in PR #20. No native host was
launched and no preference was changed in this initial Q-026 follow-up.
The prepared case matrix is `tools/tests/q026-scenarios.json`; all its native
cases remain **PLANNED_NOT_RUN**, including full restoration and normal exit.
Execution must first use the verified inactive-desktop isolation infrastructure.

The offline comparator rejects truncated/ambiguous receipts, missing hashed
control fields, malformed revisions and unverified dialog closure. It compares
all recorded tab/category/palette fields, catalogue order and value types;
recorded host identity remains separate from preference equality. It does not
recompute the .NET serialization hash or prove native restoration/cause.
`be2e25f` records **7 passed, 0 failed** in
`artifacts/q026-initial-review/offline-tests.log`. The identity/count correction
at `8a8c7f8` records **8 passed, 0 failed** in `offline-tests-identity.log`.
These are Python diagnostic tests, not a .NET/native run or coverage measurement.

The actual historical inputs are the complete durable `options-0004`, `0006`,
`0010` and `0015` receipts below
`E:/Développement/AddIn/CodexVBA/artifacts/qualification-v1/followup-20260930/format-native-v6-corrected/phases/options-evidence-8c1703a33ef44a6593935b7292ee563c/`.
They identify Excel PID 48192, start `2026-10-01T03:50:04.2993576Z`, product MVID
`d8f31d57-8612-465e-871c-93a62f2b3eae` and test MVID
`7c4a6bf2-f4a0-4234-b196-8ec0fcd0af96`. This is historical frozen evidence, not
execution of the currently installed product.

- `baseline-to-before-write.json`: equal complete revision and tab structure
  between baseline capture and the font's preceding read.
- `font-change.json`: the recorded revision changes; the sole structural
  difference is `/Tabs/1/Controls/8/Value`, from `Consolas (Occidental)` to
  `Courier New (Occidental)`.
- `font-to-size-read.json`: equal complete revision and structure between
  independent font readback and size catalogue read.
- `baseline-to-size-with-identity.json`: preserves matching recorded PID/start/
  product metadata while reporting the expected font difference.

Both baseline and later size receipts already show `Taille :` value `10`, no
choices and selected index `-1`. This does not establish why the catalogue is
empty, but prevents attributing its first appearance to this font write.
No spontaneous revision drift is observed in these compared reads. They do not
explain a different older stale-revision failure, replay any uncertain operation,
or establish complete baseline restoration. Q-026 remains **OPEN**.

## Current installed v6: complete managed acceptance (2026-10-01)

Product source `8f2315d04162f55b0956618f96b59d294a3fb681` includes the Monaco
active-document status correction and explicitly authorized, disabled-by-default
path/effective-token diagnostic. Installed MVID is
`d8f31d57-8612-465e-871c-93a62f2b3eae`, SHA-256
`C900BA09D92DA7CF50CC09033C63F5226DD04C18426CC386B30AF86EB0BA0941`.
`artifacts/qualification-v1/followup-20260930/deployment-v6.json` records the
exact installed hash and previous-payload/registration backup. The complete
instrumented run returned **2,324 passed, 0 failed, 90 conditional skips,
2,414 total**, runner exit **0**, in **496.570 seconds**. Individual TRX outcomes
match the counters; the product hash is unchanged after the run.
`managed-v6-path-diagnostic/candidate.json`, `full-managed.trx` and `terminal.json`
retain exact source/test identities and terminal evidence.

Raw managed `VBAi` coverage is **33,562/33,755 lines (99.43%)** and
**33,975/34,487 branches (98.52%)**, from
`managed-v6-path-diagnostic/d25f9f82-6146-4466-8c24-de444ccf25de/coverage.cobertura.xml`.
Native/live-provider opt-ins were disabled; native C++, JavaScript and external
host processes are not measured. Q-015 closes for this exact source/MVID/hash,
not the conditional host/provider scopes. Coverage remains below the complete
line/branch target. Earlier complete passes retain their own candidate scope.

### Current v6: terminal owner-STA path diagnostic

`path-visibility-v6-native/path-visibility-native.trx` records **1 passed,
0 failed, 0 skipped**, with runner exit **0** in `terminal.json`. Owned Excel
**39452**, native owner thread **50268**, STA and loaded v6 MVID are verified.
The testhost's independently recorded PID/thread are **47476 / 15440**.
The retained `hosts/25af21d96c284f149922af9015ea4fdf/path-visibility.json`
compares both contexts before/after the single read-only bridge observation.

Both direct GUID directories below LocalAppData and TEMP, and both fixed
synthetic files, are visible through managed/native attribute reads. Attributes
match the testhost before and after; synthetic file hashes are unchanged.
`Directory.Exists=false` for the file rows is expected directory-only API
behavior, not invisibility. All four native attribute calls succeed; their
LastError values are not failure evidence. Owner-thread OpenThreadToken returns
ERROR_NO_TOKEN (1008), so the recorded effective token is the primary token.
Testhost/owner user SID, medium integrity and AuthenticationId match; before/after
metadata is retained without credentials or token mutation. Owned Close/Quit and
normal exit **0** are recorded under `hosts/a1cfa70532ca45439cce49fa97f8b6e7/`.
No macro, export, source write, policy or permission change is performed.

This qualifies the diagnostic only. It does not reproduce or explain the
preceding `d5e25e25` export path/name-not-found observation in another owned Excel
process, establish an EFS/ACL/token cause, or qualify native UserForm export.
Q-027 remains open.

### Current v6: ordered native scalar acceptance

`scalar-v6/terminal.json` records the ordered native campaign on installed
product source `8f2315d`, the v6 MVID/hash above, and tests source `304ee02`.
Each stage's `scalar.trx` records **1 passed, 0 failed, 0 skipped**, runner exit
**0**, in a fresh owned Excel 16.0.20326.20158 instance:

| Stage | Owned PID | Page and native observer evidence | Inspection correlation |
| --- | --- | --- | --- |
| `01-unsupported` | 8284 | Offset 3 / Limit 3: array, Variant and object declarations are Skipped; no command 229 or Quick Watch observer runs. | `4813e824d9be481798d7b500f3438005` |
| `02-one-long` | 44124 | Offset 0 / Limit 1: Long 42; exactly one command 229 / native observer cycle. | `acff5fe3512243d6a48ea246c285df4e` |
| `03-full-page` | 57412 | Offset 0 / Limit 6: Long 42, expected String and true Boolean; three unsupported declarations remain Skipped, with exactly three observer cycles. | `815eaa8f3b4c4e07b3d4704f5884b450` |

Each `native-phases.jsonl` has one correlated enqueue, owner-STA callback,
validated context, CoreTerminal and Terminal, without phase errors. The
per-host startup JSON confirms the loaded installed assembly and exact
PID/executable/start identity; each shutdown JSON records Close/Quit, no force
termination, and normal exit **0**. The passing tests assert pagination, project,
module/procedure and source SHA, preserved selection/focus and break mode before
the verified cleanup Reset. Each native inspection is issued once, without replay.

This qualifies the declared scalar page on the current v6 product, not complete
runtime Locals enumeration, other local types or another host. Historical
`7b5f11d8` break-mode stalls and the post-Reset crash remain unexplained; Q-006
remains open for those causes and the separate failed lifecycle scopes.

### Current v6: failed native Monaco fixture attempts

`monaco-native-v6/monaco-native-v6.trx` and
`monaco-native-v6-v2/monaco-native-v6.trx` each record **0 passed, 1 failed,
0 skipped**, runner exit **1**, on the same installed v6 product. The first
tests source is `7a0d0e3`; source `304ee02` is the second attempt's actual source,
recorded in its additive `test-source-correction.json`. The original second
driver's invalid `TestSource={}` remains preserved.

The first attempt is refused because its select_code request supplies only one
of the paired StartColumn/EndColumn fields. The second reaches the real installed
embedded editor: its retained UI observation identifies owned Excel **57348**,
the visible editor and selected caption `VBAProject · ModuleClosedScope`. The
test incorrectly compares this decorated caption with a bare module name and
fails before closing the other project. The closed-project/live-status scenario
is therefore **NOT_RUN**, not an accepted UI test. Owned Excel **57180** and
**57348** both close normally with exit **0**, without replay or force termination.
These preparation failures remain failed, and neither qualifies the subsequent
closed-project status or complete embedded UI behavior.

### Current v6: terminal Excel core and editor batches

`excel-core-v6/excel-core-v6.trx` records **10 passed, 5 failed, 35 conditional
skips, 50 total**, runner exit **1**, on product `8f2315d` / installed v6 and
tests `304ee02`. `terminal.json` retains the exact filter and identities; only
`VBAi_RUN_EXCEL_TESTS` is enabled. The separate scalar and diagnostic scopes above
are excluded. The editor/language/save/privacy wrappers and native export/Git
probes lacking their additional opt-ins are skipped, not qualified or rejected
because an existing host was detected in this batch.

Passed scopes include owned workbook/bridge Save, public procedure and private
parameter/class-member rename with native behavior/undo, matching-project native
execution refusal, UserForm inside-dimension fitting, native format-choice options
roundtrip/restoration, project-tree/protection-dialog observation and read-only
toolbox pages. The complete batch remains failed:

| Failed case | Exact observed limit |
| --- | --- |
| Project protection save/reopen | Excel 46192 does not exit within the bounded post-Quit wait. Later absence does not provide a normal exit code; no force termination is recorded. |
| Native Variant array values | Excel 45228 returns both native operations and final source readback, then exits abnormally with `0xC0000409` during cleanup. |
| Native ParamArray values | Excel 58352 returns the native value operations and final source readback, then exits abnormally with `0xC0000409` during cleanup. |
| Extended options roundtrip | The pipe closes after emission; delivery is uncertain and the request is not retried. Later normal host exit does not qualify that mutation or full-state restoration. |
| Monaco native breakpoint | The test times out waiting for the real WebView2/Monaco control; owned cleanup exits normally, but the breakpoint scenario is not accepted. |

`exact-owned-excel-crash-events.json` matches both array-test PID creation times
to their startup evidence. Application Error events identify exception
`c0000409`, faulting module/path `unknown`, and the same fault address. Their
matching WER records corroborate BEX64/StackHash, not a cause.
`exact-owned-wer-availability.json` records only the two exact report archives:
each retains Report.wer but no event-linked temporary minidump remains. Quit
return and terminal value responses are not normal-exit evidence. No transport,
fixture-release, theme or host defect is established by these records alone.

The separately opted-in `excel-editor-v6/excel-editor-v6.trx`, tests `3fb1665`,
records **0 passed, 1 failed, 3 conditional skips**, runner exit **1**.
The first module-roundtrip wrapper owns Excel **36420** and reports final exit
`0xE0000002`; its exact Application Error event instead records a prior
`c0000005` in `combase.dll` 10.0.26100.9549, offset `0x19de28`. Both observations
are preserved in `exact-owned-excel-editor-crash-events.json`; the exception
and final exit codes are distinct. The remaining language/save/privacy cases
are refused by existing-Excel guards while the process remains briefly visible.
They are NOT_RUN, and this older scenario lifetime has no durable per-host
shutdown JSON. Later process absence does not qualify normal exit. These are
editor-wrapper results, not acceptance of the actual installed embedded UI.

### Current v6: Access metadata setter and persistence failures

These native campaigns use the installed v6 product identified above and tests
source `f39bb27`. `metadata-setters-v6/metadata-setters.trx` records **0 passed,
1 failed, 5 conditional skips**, runner exit **1**. The bridge HelpContextID
setter returns successfully for **321** in owned Access **49496**. Descriptor,
CLR binder and raw IDispatch getters all return 321 after the planned setter and
after the single verified product Save; the raw VARIANT canary remains intact.
Quit returns, but the process does not exit after COM release. Fresh-disk reopen
is **NOT_RUN**, and the original save/reopen test remains failed.

`metadata-setters-v6/terminal.json` retains copied `METADATA_GETTER_BATCH_TERMINAL`
and `ProbeAdditionalMutations=0` / `ProbeAdditionalSaves=0` labels from the earlier
read-only driver. Those fields do not describe this setter campaign's planned
work. The per-fixture `adapter-only-progress.json` ledger is authoritative: one
planned setter and one verified `save_host_document`, with no replay.
`terminal-label-correction.json` adds this correction while preserving the
original driver JSON and failed TRX.

The separate `metadata-setters-v6-external/metadata-setters.trx` records
**0 passed, 3 failed, 1 conditional skip**, runner exit **1**. These setters run
on the external fixture STA through the mapped COM object, rather than the
in-process bridge. Each executed case has one setter and one verified product
Save, with identity evidence before/after and no replay:

| Case | Live setter/getter evidence | Fresh disk and lifecycle outcome |
| --- | --- | --- |
| HelpFile, production CLR setter | Access 42136 returns normally; all three getters agree on 173 characters / 346 BSTR bytes after mutation and Save. | Fresh Access 43812 returns an altered 86-character / 172-byte BSTR through all three getters. Both hosts exit normally with code 0; exact metadata persistence still fails. |
| HelpFile, raw IDispatch PUT | Access 44760 records one PUT with HRESULT 0, VT_BSTR and intact argument/result VARIANT canaries; all three getters agree on the same live value. | Fresh Access 49364 returns the same altered BSTR shape as the CLR case, with exact getter equality and intact canary. Both hosts exit normally with code 0; persistence fails. |
| HelpContextID, production CLR setter | Access 57628 returns normally; all three getters return 321 after mutation and Save. | Quit returns without process exit; fresh-disk reopen is NOT_RUN and the test fails. |
| HelpContextID, raw IDispatch PUT | Conditional skip; setter NOT_RUN. | Persistence and normal exit are not qualified for this case. |

The HelpFile failures are therefore not confined to the production CLR setter.
The consistent raw/managed getter observations do not establish a conversion
fix or justify decoding the altered BSTR. HelpContextID's successful live value
does not qualify disk persistence or explain the earlier setter HRESULTs.

Each nonexiting HelpContextID host has a separately authorized, terminal owned
cleanup record. `metadata-setters-v6/authorized-owned-cleanup.json` records a
stable retained database with SHA-256
`AC43D9A45AB6EE36073656934A18EC9B3E0B27BE4062959D3286F5DC3103054F`;
`metadata-setters-v6-external/authorized-owned-cleanup.json` records
`3594DD1D3B40C34DB70A3B87966EF3629026E5B64998D7C2FB46692DB303723F`.
Both retained copies are `retained-before-force.accdb`, observed stable before
termination rather than claimed atomic snapshots. Forced exit is observed for
49496 and 57628, but each exit code is null (**NOT_OBSERVED**); normal exit is
not qualified. No additional Quit, mutation or Save is recorded. Cleanup does
not promote either original failed test or add disk-reopen evidence. Q-006 and
Q-012 remain open for these scopes.

### Current v6: Publisher preparation failure

`publisher-adapter-v6/publisher-adapter-v6.trx` records **0 passed, 1 failed,
5 conditional skips**, runner exit **1**, on the same product/tests pair.
The startup ledger verifies the sole disposable publication and owned PID
**36568**, and `project_persistence_status` reports HostAvailable and
IdentityVerified with the exact publication path. The sole VBE project is
pathless `Project`; `debug_state` returns null SelectedProject and ActiveModule.
The active-project startup guard fails before baseline binding, source/property
edits or product Save. The fixture's initial publication bootstrap save is
separate from adapter acceptance. `native-selection-inspection-correction.json`
preserves the external read-only scripts but marks their project inventory
UNVERIFIED: Publisher.Application.VBE is unavailable, so the earlier null-as-empty
reports are not proof of an empty native inventory. The actual host bridge's
project/persistence rows remain authoritative. Whether selection is missing or
the fresh-project guard is too strict remains unproven; Q-012 stays open.

`authorized-owned-window-close.json` then records one guarded WM_CLOSE after
PID/start/window, loaded MVID and exact saved-publication checks, with source
inventory preserved. Owned PID 36568 exits normally with code **0**, without
force termination, additional Save or macro execution. This later cleanup does
not convert the startup failure into adapter save/reopen acceptance.

### Current v6: Publisher adapter follow-up with harness failures

`publisher-adapter-v6-v2/publisher-adapter-v6-v2.trx` records **1 passed,
5 failed, 0 skipped, 6 total**, runner exit **1**, on the same installed product
and tests `3fb1665`. The module/class adapter-only Save/reopen case passes with
exact source readback and normal initial/fresh host exits. Each metadata/reference
case records one verified product Save and a fresh-disk reopen, but fails the
common harness assertion comparing original selector `Project` with the
reopened canonical publication path. Description, HelpFile and reference
addition by GUID/file/removal therefore stop before their final disk-state
verification callbacks; their complete tests remain failed.

Per-fixture ledgers under `hosts/Publisher/` record both owned normal exits with
code **0** in every case, without forced termination or helper Save. This
supersedes the earlier preparation-only limit for the module/class operation,
not the original failed batch or the unverified metadata/reference readback.
The selector mismatch is a tests-only qualification defect; it is not evidence
that those native values were lost or that they persisted correctly. Q-012
remains partial, with HelpContextID and the remaining adapter scopes open.

### Current v6: embedded Monaco status and reviewed native capture

`monaco-native-v6-v3/monaco-native-v6.trx` records **1 passed, 0 failed,
0 skipped**, runner exit **0**, on the installed product above and tests
`687a93d`. Owned Excel **38944** retains the closed project's decorated tab
while the live project's selected tab has a healthy synchronized status for
at least three seconds across synchronization timers. Both exact source
snapshots, the live project's path/identity, selection and design mode remain
unchanged. Close/Quit exits normally with code **0**, without forced cleanup.

`hosts/1e70312422db4bde9e560b60f81f5fbc/native-monaco-closed-project.json`
records the actual installed embedded editor rather than a detached control.
Its `native-monaco-live.png` is a real capture, SHA-256
`73E2021E44C0C1772BD0FDE2C68C6077DD83536BFEDDB4AC0C2E9B7C0DE57BB2`.
The additive `native-monaco-live.png.visual-review.json` records visual review
of the selected LiveScope tab, retained ClosedScope tab, expected function
source and healthy localized synchronization status. The original automated
sidecar remains unchanged: accessibility did not expose rendered source, so
pixel review and native SHA readback are separate evidence. Earlier preparation
failures above remain failed; other editor-wrapper and SOLIDWORKS scopes remain
open.

### Current v6: inherited export failures and paired owner-STA observations

`inherited-exports-v6/inherited-exports-v6.trx` records **1 passed, 3 failed, 0 skipped**,
runner exit **1**, tests `687a93d`, on this same installed product. Fresh
COM-activated Excel instances **26120**, **29932** and **50988** each receive
one terminal failed export below LocalAppData, LocalAppData/VBAi and
GitTemporary respectively, with `Objet spécifié introuvable`. Excel **46332**
exports the synthetic form below TEMP successfully. All four hosts exit
normally with code **0**. The failed reports' stage label
`ONE_NATIVE_EXPORT_PENDING` is stale: their received `Ok=false` responses are
terminal native errors, not unobserved delivery. Those original labels remain
preserved; no failed export is retried.

`paired-export-v6/paired-export-v6.trx` records **2 passed, 0 failed,
0 skipped**, runner exit **0**, tests `3605605`. Each fresh owned Excel uses
explicit `/x /automation` bootstrap with the fixed synthetic manifest. Actual
owner-STA PID/native TID/MVID and effective-token metadata immediately precede
one export, with no intervening host command or COM call. Excel **34892**
exports below the LocalAppData GUID child; **44872** exports below TEMP.
Both owner observations see all four synthetic targets, with medium primary
tokens after ERROR_NO_TOKEN and unchanged synthetic hashes. FRM files are
**479 bytes**, FRX companions **2,584 bytes**; exact file hashes, source and
native Label properties are retained. Both hosts close normally with code
**0**, without permissions, attribute, trust-policy or token changes.
The report's interim `SUCCESS; awaiting normal owned exit` text remains
preserved; its independent Shutdown record proves the subsequent normal exit.

These are different launch contexts and fresh instances, not a controlled
causal explanation of the inherited-export failures. Immediate sequencing is
not atomic proof of an unchanged token context during export. This scoped
explicit-bootstrap export acceptance does not qualify the complete Git
capture/normalization/import/recovery or current-provider transfer. Q-027
remains open.

### Current v6: semantic Publisher reopen and remaining metadata failures

Tests `3f23205` retain original native state comparisons while proving reopen
identity from owned process, actual loaded candidate, publication path,
project name and persistence identity. A legitimate name-to-path selector
change is accepted; the original selector-assertion failures above remain
failed.

`publisher-semantic-v6/publisher-semantic-v6.trx` records **1 passed,
1 failed, 5 conditional skips**, runner exit **1**. Description persists
exactly after normal close and fresh reopen. HelpContextID in owned Publisher
**60080** returns `Ok=false`, HRESULT **0x9CFD3148**, yet terminal read-only
readback records **321** and `Saved=false`. The setter is not replayed and no
product Save follows the error. The teardown guard refuses unsaved cleanup,
retaining the original process handle and recording zero Quit entries.
Live sources, native identity and properties are then preserved before one
authorized window close and one exact owned `Ne pas enregistrer` invocation.
`authorized-discard-terminal.json` observes normal exit **0**; this cleanup
does not qualify the failed setter or disk persistence.

The independent fresh-case batch
`publisher-semantic-v6-rest/publisher-semantic-v6.trx` records **5 passed,
1 failed, 0 skipped**, runner exit **1**, excluding the failed HelpContextID
scope. Module/class/form, Description and reference addition by GUID/file and
removal pass full source/metadata/reference readback after verified product
Save and fresh native reopen. All initial and fresh hosts exit normally with
code **0**, with one Quit entry and no forced termination. HelpFile fails exact
metadata readback: its reopened string is altered, including expected `E`
read as `㩅`. This resembles the Access observation but establishes no shared
cause or safe decoding fix. Q-012 remains partial for metadata and lifecycle;
no whole-host acceptance follows.

`publisher-helpfile-storage-v6/helpfile-storage.json` and `cfb-streams.json`
add read-only storage evidence for the retained failing publication, whose
before/after SHA-256 is
`313CE321139B3768B8E602F5E49369F9FBDB30E79BF5AEBA1E4940D62F16D272`.
The bounded CFB reader extracts the compressed VBA `dir` and `PROJECT` streams.
PROJECTCODEPAGE is 1252; both PROJECTHELPFILEPATH fields contain the exact
expected path as identical 174-byte MBCS sequences, and the PROJECT HelpFile
line also matches. This follows the
[Microsoft record specification](https://learn.microsoft.com/en-us/openspecs/office_file_formats/ms-ovba/b1e1f51f-6bef-49fe-b6a9-76e174d51b0d).
The fresh getter's 87 UTF-16 characters reproduce those 174 stored bytes when
encoded as UTF-16LE. Correct path storage is observed; whether native loading,
BSTR construction or marshaling causes the altered exposed string remains
unproven. The original test stays failed, without heuristic decoding or a
production correction. The reader assembly is a detached diagnostic helper,
not evidence of an additional installed product candidate.

`access-raw-context-v6/access-raw-context-v6.trx` records **0 passed,
1 failed, 0 skipped**, runner exit **1**, tests `3605605`, completing the
previously unexecuted raw HelpContextID setter case on fresh Access **38712**.
One IDispatch PUT returns HRESULT **0**, VT_I4 **321**, intact VARIANT canaries
and unchanged identity. Getter readback and one verified product Save succeed,
but the host fails the bounded exit wait after its single Quit; disk reopen is
NOT_RUN. The retained lifecycle ledger preserves the original handle and
records no cleanup replay. Separately authorized owned cleanup preserves
`retained-before-force.accdb`, SHA-256
`5A9C1C09AE11EA80FCB1F4A59489F147232272054412F2C0376D2DF29CF03547`,
then observes forced exit **-1**. No additional Quit, setter or Save is issued;
normal exit and metadata persistence remain unqualified.

### Current v6: Word/PowerPoint adapter and diagnostic guard batches

`word-powerpoint-v6/word-powerpoint-v6.trx` records **1 passed, 1 failed,
0 skipped**, runner exit **1**, product `8f2315d` and tests `3605605`.
PowerPoint module/class/form adapter-only Save and fresh disk reopen pass;
owned cleanup records one Quit and normal exit **0**. Word also reaches its
fresh source/form readback, but final owned Word **28700** fails the bounded
exit wait after its single Quit and retains its original process handle.
The original complete case stays failed. Subsequent process absence is
recorded separately with exit code **NOT_OBSERVED**, no force termination or
additional Quit/mutation; it does not qualify normal shutdown.

`diagnostic-guards-v6/diagnostic-guards.trx` records **30 passed, 0 failed,
0 skipped**, runner exit **0**, tests `28c0565` against the frozen installed
product. These are detached paired-export and teardown-gate/controller guards,
not native Office or global coverage evidence. The prepared CDB teardown
preflight remains failed in both `teardown-preflight/` and
`teardown-preflight-v2/`: the initial fatal handler is first-chance-only;
the next helper exposes no exception context and exits zero after debugger
detach. Normal helper detach/exit passes, but neither failed preflight permits
Office attachment. Their original reports/logs remain preserved. No native
array teardown diagnostic has run at that earlier checkpoint.

The next `diagnostic-guards-v6-v3/diagnostic-guards-v3.trx` records
**26 passed, 0 failed, 0 skipped**, runner exit **0**, tests `3a3c266`;
these remain detached guards, not a new global coverage result.
`teardown-preflight-v3/` still fails: CDB reports the fatal exit while `.ecxr`
cannot provide live-event registers. The corrected live-register collector's
`teardown-preflight-v4/4eeec955f15941d183729e0ed5f00edb/preflight.json`
passes actual helper checks. Normal helper **47732** detaches and exits **0**;
synthetic helper **56496** captures the exception record, live registers and
stack, then both target and debugger exit **0xC0000409**. This matching fatal
exit is separate from normal detach and does not qualify Office behavior.

The subsequent single owned native array trial, tests `cc813dc`, records raw
`excel-array-teardown-v6/excel-array-teardown-v6.trx` **1 passed, 0 failed,
0 skipped**, runner exit **0**, with Excel **38036** exiting **0**.
Its additive `diagnostic-validity-correction.json` marks qualification
**INVALID_DIAGNOSTIC_INTERVENTION**, without rewriting the original TRX:
default CDB settings stop at a first-chance access violation in `clr+0x72eb0`
after the original Close/Quit/release sequence. The controller deadline then
issues `qd`; detachment can mark the outstanding exception handled and affect
its progression. No fatal stack is captured, no causal fix is established and
the original uninstrumented array/ParamArray failures stay failed. There is no
procedure or native-cleanup replay, host termination, full dump or global
policy change. A corrected exception-forwarding preflight is required before
further native diagnostic acceptance.

After the exception-forwarding correction, tests `e10cd46` pass the detached
`diagnostic-guards-v6-v5/diagnostic-guards-v5.trx` batch: **49 passed,
0 failed, 0 skipped**, runner exit **0**. This includes paired terminal-state
guards and the complete collector/helper controller guards, not additional
global product coverage.

`teardown-preflight-v5/ccf20d0db4ea4459b248851dddb5f718/preflight.json`
passes all four actual helper trials together: normal detach/STOP and exit
0; fatal 0xC0000409 and software fatal 0xC0000005 with exact target/debugger
exit matching and executed exception/live-register/stack evidence; locally
handled software first-chance AV forwarded unhandled to the application,
followed by normal exit 0. No global CLR policy is changed. CDB safe-stop
proof requires the exact owned injected breakpoint rather than a pending
application fault. The helper and common-script hashes are bound in the
report; older preflights cannot authorize this controller.

`excel-array-teardown-v6-v2/excel-array-teardown-v6-v2.trx` records
**1 passed, 0 failed, 0 skipped**, runner exit **0**, owned Excel **56808**,
with normal exit **0**. The exact controller records **NOT_REPRODUCED**:
39 first-chance AV events are forwarded, no fatal exception is captured,
and the process exits without deadline expiry or controller stop/detach.
This is an instrumented scoped result, not an explanation or correction of
the original uninstrumented crashes. No native invocation/cleanup is replayed.

The separate uninstrumented prepared two-case batch,
`procedure-values-standalone-v6/procedure-values-standalone-v6.trx`, records
**2 passed, 0 failed, 0 skipped**, runner exit **0**, with the teardown
diagnostic variable absent. Fresh owned Excel **47732** and **59352** preserve
the asserted native ParamArray/Variant-array values, arity/bounds, invocation
counts and final source readback, and both exit normally with code **0**.
Their distinct startup identities and per-host shutdown records remain under
`hosts/`; PID reuse in another earlier helper trial does not establish shared
process identity. This accepts the current-product standalone scopes only.
The earlier core batch and its exact crash events remain failed and unexplained;
an isolated pass does not make the full campaign green or close Q-006.

### Current v6: complete prepared explicit-launch form layout matrix

`userform-explicit-localgit-v6/userform-explicit-localgit-v6.trx` records
**0 passed, 12 failed, 0 skipped**, runner exit **1**, on tests source `8480aea`
and the unchanged installed/referenced v6 product MVID/hash above. All prepared
layouts were executed in separately owned `/x /automation` Excel instances:
Label/Button, TextBox, ComboBox, ListBox, CheckBox, OptionButton, ToggleButton,
ScrollBar, SpinButton, TabStrip, Image and Frame/MultiPage. This test-local launch
selection does not change the shared Excel fixture or activate token reading.

Each case verifies its synthetic native layout and captures the genuine designer,
then fails the first comparison of two unchanged production `VbaGitProject.Capture`
snapshots. In the retained Label/Button pair, sources, manifest and FRM hashes
are identical; the same-length FRX hashes differ. Captures therefore reach the
production GitTemporary directory in this launch context, but logical revision
stability fails. No meaningful-property-change, Git commit, checkpoint import,
rollback or helper Save/reopen phase is reached. Those later phases are
**NOT_RUN**, not failed native mutations or accepted recovery. No remote Git
operation or macro execution occurs. The retained source-designer PNGs have now
been visually reviewed: each synthetic form and target control is visible without
an error dialog. `source-designer-visual-review.json` binds observations to each
original image hash. This reviews the source layout only; it does not establish
visual acceptance after import, recovery or reopen, which remain NOT_RUN.

All per-host shutdown records verify normal exit **0**, without Close/Quit
errors or forced termination. Owned PIDs are 49560, 45604, 57368, 51208, 41316,
15220, 48004, 20548, 23068, 56920, 23632 and 54756; startup records bind each to
its distinct image/start/MVID. The terminal report and raw before/unchanged files
remain under `artifacts/qualification-v1/followup-20260930/userform-explicit-localgit-v6/`.
This reproduces a comparison failure beyond the previously scoped single-export
passes; Q-027 remains open. No normalization rule is relaxed by this scenario.

### Current v6: options failure located before its intended mutation

Read-only analysis of the original `excel-core-v6` command ledger locates the
ExtendedOptions failure in owned Excel **17592**, start
`2026-10-01T00:29:48.3078893Z`, loaded v6 MVID. Sequence 3 is `read_vbe_options`,
emitted at `00:29:53.3080108Z`; the connection closes at `00:30:38.5511464Z`,
after 45.243 seconds. The intended checkbox toggle is never emitted, and the
Docking iteration is **NOT_RUN**. Its existing finally block reads successfully,
requests the already-selected margin-indicator value (On to On), then reads
successfully; the host exits normally with code **0**.

The successful response records are truncated before their `OptionsVersion`
field, so identical retained prefixes do not independently establish complete
baseline equality. Source control flow implies the in-memory restoration
assertion completed without replacing the original read failure; this is an
inference, not complete persisted revision proof. EOF alone does not identify
a native UI call, serialization failure, deadlock or pipe-write timeout.
The original scenario remains failed and Q-026 remains open.

### Uninstalled StdFont prototype and options lifecycle guards

Source correction `839550c` consumes the documented packed StdFont representation
without normalizing any font property. The isolated pre-commit prototype is MVID
`277e30ac-8f58-44e8-ac63-cd0dbf49b8b7`, SHA-256
`CC3F30F52C6C9FFBC1CE7756F383385934D47AA0B11BF6D2127472E20F3A032A`.
It remains **uninstalled**; the installed v6 binary is unchanged. Retained native
exports are reanalysed with each actual managed assembly, without COM activation.
On v6, every layout's comparison differs only in FRX. On the prototype, the
Label/Button pair compares equal; all remaining layouts still retain the complete
original logical streams when their grammar is unsupported.

Every differing Label/Button raw byte is assigned to an already-excluded CFB
timestamp or a documented site/name-padding region. Source, manifest, FRM,
control-object data and font values remain unchanged. The bounded physical and
logical maps, exact assembly comparisons and primary format references are in
`frx-drift-analysis/diagnostic-report.md`. The original prototype JSON reused an
installed-candidate scope label; its additive
`prototype-comparison.identity-correction.json` corrects that label while
preserving the original report and hash. No native prototype acceptance is claimed.

Tests source `ff3925b` adds the test-local options helper without changing the
shared Excel fixture. A pre-write failure emits no restoration write, and unknown
delivery retains the exact host without another bridge/Close/Quit request.
Only a verified closed/committed write permits revision-guarded restoration.
Durable bounded semantic summaries preserve the full options revision separately
from omitted catalogues and retain primary/restoration/cleanup failures.

The prepared combined managed batch
`std-font-options-managed/std-font-options-managed.trx` records **35 passed,
0 failed, 0 skipped**, runner exit **0**, using this isolated prototype and tests
source `ff3925b`. It checks documented StdFont limits, every font byte remaining
significant, malformed/truncated/unknown layouts, input immutability and the
options lifecycle/error guards. This is a focused managed batch, not a complete
suite, coverage measurement, native options rerun or deployed-candidate result.
The preceding structural-only batch remains separately retained. Q-026 and Q-027
remain open pending their native causes and complete acceptance scopes.

### Current v6: guarded options mutation and durable restoration evidence

The isolated `options-guard-native-v6` trial and the subsequent
`options-durable-native-v6` trial each record **1 passed, 0 failed, 0 skipped**,
runner exit **0**, against the unchanged installed v6 product. Owned Excel PIDs
53912 and 48980 exit normally with code **0**. Their successful TestContext
attachments were not retained durably. The second driver's copied build path
also selected tests source `93bb870`, rather than its recorded `313e0b6`.
`test-build-scope-correction.json` preserves that provenance correction without
changing the original terminal/TRX; neither trial proves the durable-summary fix.

The separately prepared `options-durable-native-v6-v2` trial uses frozen tests
source `313e0b6`, test assembly SHA-256
`A1D92AD7D9B366C0637EE2C00E9658E52B8D44CC39C7653DA4816D73CE6D094E`,
and the same installed v6 MVID/hash. It records **1 passed, 0 failed, 0 skipped**,
runner exit **0**. Excel PID 47212 exits normally with code **0**, with no
Close/Quit errors or forced termination. Its owned GUID root retains 38 bounded
phase summaries, including the complete baseline options revision
`544477535d391a457fc66947a3ddb3da7ffb430e6212289c6586a6fcb65f3e05`.
Editor Format margin indicators and Docking Immediate Window each change from
checked to unchecked and back; both final full revisions match the baseline.
All writes have verified closed/committed replies before guarded restoration.
No token diagnostic or shared fixture lifecycle change is active.

These fresh disposable trials qualify this installed-candidate checkbox scope.
They do not explain the original pre-write EOF or qualify every option/category;
Q-026 remains open. No uncertain request is replayed.

### Uninstalled combined flat-control grammar prototype

Source `317fa80` combines StdFont, six MorphData controls and ScrollBar,
SpinButton, TabStrip and Image grammar. Its isolated product MVID is
`c5e41581-6a45-4512-bfd9-b1f41106bf4a`, SHA-256
`CA1A461F65C426605C9EDD530BFF59444EED518F6D207B95BE583358CF55ECBB`.
The Debug/net48/x64 build reports no warnings or errors.
`combined-controls-managed/combined-controls-managed.trx` records **115 passed,
0 failed, 0 skipped**, runner exit **0**, in one prepared focused batch covering
form padding grammar, snapshot guards and options qualification lifecycle guards.
This is not a full-suite run or a coverage measurement.

`frx-drift-analysis/combined-controls-comparison.json` invokes this actual assembly
against all retained native before/unchanged snapshot pairs without COM activation.
Eleven flat layouts compare equal. Frame/MultiPage still differs and retains
unsupported container streams; complete graph validation remains outstanding.
The prototype is **uninstalled**. This offline proof does not change the original
failed native matrix, execute its later Git/import/recovery/persistence phases,
or close Q-027. Raw exports remain intact and transport bytes are not normalized.

The follow-up prepared batch at source `10484fe`,
`prepared-sequential-guards/prepared-sequential-guards.trx`, records **125 passed,
0 failed, 0 skipped**, runner exit **0**. It adds absent/default MorphData guards
and nested timeout/I/O/cancellation classification for the new sequential native
GitHub fixture. The actual uninstalled assembly is MVID
`a9cf8552-de7a-4648-a26e-dff2049677da`, SHA-256
`6509AC97B98B908FEFCA3D39FB9D9496F4F02DACA93EE78E1D551CE85ADC861D`;
`prepared-sequential-guards/candidate.json` also binds the test assembly hash.
An initial test namespace shadowed existing infrastructure names and failed
compilation; the corrected source builds without warnings/errors before this
single test batch. No coverage percentage or full-suite result is claimed.

The new authenticated scenario is **NOT_RUN**. It prepares source capture and
normal closure before fixed synthetic-repository push/fetch, then launches a
distinct target for guarded import/backup, native readback and helper save/reopen.
It verifies exact remote bytes and unchanged remote main. Uncertain native
delivery retains the exact host without another dispatch or cleanup; uncertain
push is never replayed. The case needs separate native/authenticated opt-ins and
a matching installed candidate. Preparation/compilation is not Q-024/Q-027
acceptance, and the existing simultaneous-host scenario remains separate.

### Preceding Monaco-status candidate: complete managed acceptance

Source `d7a1c75d90d840949cc78606f2ae55b14e93ec55`, MVID
`6a74af33-b2fd-4ff9-9133-475b53ae9e42`, SHA-256
`FD39665BE5CCB4E5DE8166526892F760D8D4DF4D240A63C5DC7EBB9D98A660FB`, passes
its complete instrumented suite: **2,298 passed, 0 failed, 88 conditional skips,
2,386 total**, runner exit **0**, in **515.164 seconds**. The product hash is
unchanged after execution. `managed-monaco-document-status/candidate.json`,
`full-managed.trx` and `terminal.json` retain exact provenance; this isolated
candidate was not the installed product during that run.

Raw managed `VBAi` coverage is **33,385/33,571 lines (99.45%)** and
**33,866/34,351 branches (98.59%)**, from
`managed-monaco-document-status/452f31de-8922-460c-8b51-a643e84dfd46/coverage.cobertura.xml`.
Native/live-provider opt-ins were disabled; C++ renderer, JavaScript and external
host processes are excluded. Q-015 closes only for this candidate; it does not
qualify the later v6 diagnostic binary or achieve the complete line/branch target.

Separate checks on this same source/product record **46 Designer views** passing
load, resize, child-property edits and serialization roundtrip in
`designer-monaco-document-status/designers.json`, and **60 passed JavaScript
language/editing checks, 0 failed or skipped** in `monaco-status-js.log`.
Those checks do not measure native VBE embedding or application lifecycle.

## Preceding v5: complete managed acceptance (2026-10-01)

Initial product and test source `2e751618aa8d5be08ca0fa84d6ceff4d0b1312dc`
implements the stable Access application/database-path/mapped-project identity
guard while retaining owner-PID, VBE, selection, mode/protection, source and
metadata checks. The isolated build completed without warnings or errors.
Candidate `f9a36c85-1d9c-4a53-8645-06c99617b12c`, SHA-256
`76524AC884247BEBC03B0A73E36AA1C96414CB650D98E9E74F11BCA5A4D068D1`, was
installed; `candidate-v5.json` and `deployment-v5.json` retain exact provenance
and deployment/backup records.

After tests-only harness correction `f0874e6d85aa67c5d240b45bf332975585ef2026`,
the complete unfiltered instrumented suite returned **2,294 passed, 0 failed,
88 conditional skips, 2,382 total**, runner exit **0**, in **603.371 seconds**.
Product source remains `2e751618aa8d5be08ca0fa84d6ceff4d0b1312dc`; the frozen
product was copied into an isolated output, not rebuilt. Product and test hashes
were unchanged after execution, and individual TRX outcomes match the counters.
`managed-harness-complete/candidate.json`, `full-managed.trx`, `summary.json` and
`terminal.json` retain exact product/test identities and terminal evidence, under
`artifacts/qualification-v1/followup-20260930/`.

Managed `VBAi` coverage is **33,378/33,564 lines (99.45%)** and
**33,863/34,347 branches (98.59%)**, from raw integer counters in the summary and
`managed-harness-complete/db754ce0-cea0-4b11-9a45-9545cd5a4afe/coverage.cobertura.xml`.
Native/live-provider opt-ins were disabled. C++ renderer, JavaScript and external
host processes are outside this measurement; conditional skips do not qualify
those scopes. Q-015 is closed for this exact product/test pair. Coverage is below
the requested complete line/branch target, and any subsequent product correction
requires a new complete gate. The subsequent Monaco-status candidate above
has its own complete result; current-v6 managed acceptance is recorded above.

### Initial v5 complete run: explained harness failures

The initial complete instrumented run is terminal **FAILED: 2,279 passed,
8 failed, 88 conditional skips, 2,375 total**, in **484.103 seconds**.
`managed-access-guard/full-managed.trx`, `summary.json` and `terminal.json`
retain the individual outcomes, nonzero runner exit and unchanged product hash.
The failures are a missing debugger-script regression fixture and seven local
Git scenarios whose nested bare-remote paths exceed the usable Git path length.
These diagnosed harness failures remain failures of the complete run; a
tests-only correction packages the exact debugger-script fixtures and gives
owned local Git fixtures shorter paths on the same volume. The subsequent
complete run above passes without changing production or global Git settings;
it does not change the original failed result.

The failed run measured managed `VBAi` only:
**33,359/33,564 lines (99.39%)** and **33,826/34,347 branches (98.48%)**,
from raw integer counters in the summary and
`managed-access-guard/e295ba72-cd41-4381-8d60-0441b45bd84b/coverage.cobertura.xml`.
Native/live-provider opt-ins were disabled; C++ renderer, JavaScript and external
host processes are not measured. Coverage does not convert this failed run
into acceptance, and historical percentages below do not replace the current
successful-run metrics.

### Preceding v5: partial SOLIDWORKS 2025 native qualification

The explicitly selected SOLIDWORKS 2025 instance is **PID 35136**, actual COM
revision **33.1.1**, with the exact installed v5 MVID/hash above. Launch and ROT
records are `solidworks-2025-v5-launch.json` and `solidworks-2025-v5-rot.json`.
The independent add-in loading check
`solidworks-2025-v5-vstest/solidworks-2025-load-v5.trx` records **1 passed,
0 failed, 0 skipped**; this is separate native evidence, not another complete
suite or a lifecycle qualification.

The host-created disposable fixture is
`artifacts/qualification-v1/solidworks/type100-2025/v5-01/Qualification2025.swp`.
Its `host-macro-creation.json` and `type100-qualification.json` retain ownership,
module/class/form creation, stale-source refusal, compilation, one verified
product Save and exact live source/label readback. Exactly one synthetic native
run reports `Ran=true`, `Error=0`, a verified marker and observed project unload.
The strict execution trial nevertheless **FAILED** whole-file preservation:
the saved SWP SHA-256 changes from
`5706B3F514ADB1835CE9B5ABAE053B3478AB73AD4556D302236B1EDE33AFC476` to
`ADA1AB0B57F6AE39FF887CD5C0789C53D841045F874F32AF2A0A957F369FD0D2`.
The cause of this binary change is not established.

A separate, single native Edit Macro reload targets the observed post-runtime
hash. `native-open-after-runtime.json` and
`post-runtime-content-verification.json` in the fixture directory verify the
original module/class source hashes and persisted label, with no additional
mutation or macro run. This is content persistence evidence; the original strict
trial remains failed and whole-file preservation is not qualified.

Reviewed `solidworks-2025-v5-code.png` shows the selected synthetic class and
`Value = 42` rendered in Monaco. `solidworks-2025-v5-code-resize.json` records
workspace fit after resize and verified restoration of the original placement.
The capture also shows a stale closed-project warning despite the selected live
class. The later Monaco-status correction has detached acceptance above;
complete embedded assistant/UI acceptance remains pending on current v6.
The reviewed `solidworks-2025-v5-real-designer.png` shows the actual designer and
persisted label after a successful `open_form` response. Earlier artifacts named
`solidworks-2025-v5-designer` were mislabeled: their attempted command was unknown
and code remained visible. The correction record preserves that failed setup.
The resize helper measures Monaco bounds even when the designer is selected;
its later successful resize is not a measured designer-resize pass.

`solidworks-2025-v5-before-close.json` preserves both owned projects' live sources
and form metadata before authorized discard. The exact owned instance then
closes through one `ExitApp`, with no force termination and exit **0**, recorded
in `solidworks-2025-v5-normal-close.json`. SOLIDWORKS 2019's separate abnormal
termination, strict SWP preservation and complete debugger/assistant acceptance
remain unresolved; Q-014 remains open.

### Preceding v5: terminal Access adapter campaign

`access-adapter-v5/access-adapter-v5.trx` and `terminal.json` record
**5 passed, 2 failed, 0 skipped, 7 total**, runner exit **1**, on source
`2e75161` and the exact installed v5 MVID/hash above. The Microsoft 365 Access
executable version is **16.0.20326.20158**. Each scope preserves the original
adapter response, source/property/reference snapshots and owning-process identity
in `hosts/Access/<fixture>/adapter-only-progress.json` and `qualification.json`.
Exactly one adapter Save follows the prepared mutation; no subsequent compile
or helper save contaminates fresh disk readback.

| Native scope | Outcome | Initial / fresh-reopen PID and cleanup |
| --- | --- | --- |
| Active-module-only edit | PASS | **15500 / 58004**; exact source hashes, both normal exits **0**. |
| Module and class edits | PASS | **53808 / 51736**; distinct pending module/class hashes retained on disk, both normal exits **0**. |
| Description | PASS | **5436 / 59200**; source and exact metadata readback, both normal exits **0**. |
| Scripting reference addition by GUID | PASS | **41548 / 47852**; exact source and installed reference identity readback, both normal exits **0**. |
| Scripting reference addition by file | PASS | **36380 / 38676**; exact source and installed reference identity readback, both normal exits **0**. |
| HelpFile path | FAIL | **49492 / 55512**; adapter Save and normal closes/reopen complete, but fresh metadata contains `㩅` where `E` was expected at index 9. Both normal exits **0** do not qualify changed metadata. |
| Scripting reference removal | FAIL | **52940 / NOT_RUN**; Quit returned but the initial host did not exit before the fixture deadline. No fresh reopen or disk-persistence acceptance. |

The successful scopes establish the stable Access save guard on this candidate;
they do not qualify HelpFile, HelpContextID or reference removal. For the retained
removal fixture, `authorized-retained-host-cleanup.json` records one explicitly
authorized force termination, exit **-1**, no Quit replay and no qualification
pass. The exact disposable database is retained as
`retained-reference-removal-Disposable.accdb`, SHA-256
`FA91E9A79C1C6F3501054E4FB023DEB4D610BEA21BFD9E4BF9B49D06841E8EA3`.
Q-012 remains partial; v5 Publisher acceptance is not established, and no
current-v6 adapter acceptance follows from this campaign.

### Preceding v5: terminal Access metadata getter diagnostic

`metadata-getters-v5/metadata-getters.trx` and `terminal.json` record **1 passed,
2 failed, 0 skipped, 3 total**, runner exit **1**, in **38 seconds**, against
frozen product `f9a36c85` / SHA-256 `76524AC884247BEBC03B0A73E36AA1C96414CB650D98E9E74F11BCA5A4D068D1`
and tests-only source `7645a5d`. The fresh read-only baseline passes; HelpFile
save/reopen and HelpContextID mutation remain failed scenarios.

Before and after the existing single adapter Save in owned Access **54560**,
the descriptor, CLR binder and raw IDispatch HelpFile getters agree on the exact
synthetic path: **164 characters / 328 BSTR bytes**. Fresh disk reopen in **37504**
returns **82 characters / 164 BSTR bytes** through all three getters, with the
same garbled Unicode. Raw BSTR bytes match the original path's host ANSI encoding;
the raw VARIANT canary remains intact. These observations localize the altered
value beyond a single managed getter conversion; they do not establish the
native persistence mechanism or justify heuristic string repair.

Runtime type information in the HelpContextID scope **26652** confirms both
getter and setter use **VT_I4**, consistent with descriptor `System.Int32`.
The existing setter nevertheless fails with **0xD09072B8**. Fresh baseline
**26204**, both HelpFile processes and the HelpContextID process exit normally
with code **0**. Getter snapshots are retained under `metadata-getters-v5/hosts/Access/`,
including BeforeExistingMutation, AfterExistingMutation, AfterExistingSave and
FreshDiskReopen phases. The getter probe adds no mutation or save to the existing
scenarios; no conversion heuristic or product fix is applied. Q-012 remains
partial, and this preceding-product diagnostic does not qualify current v6.

### Preceding v5: Publisher recovered-document startup failure

`publisher-adapter-v5/publisher-adapter-v5.trx` records **0 passed, 1 failed,
5 skipped, 6 total**, runner exit **1**, on the same installed v5 product and
initial test source `2e75161`. The first fixture fails before the adapter Save:
its baseline module already exists in the selected project. Cleanup refuses a
changed active-document identity; the other scenarios are not executed.

Readonly native evidence identifies recovered Publisher **55000** and its child
**57036**. The parent's document catalogue contains the prior owned
`office-adapter-v4` publication and the new disposable `publisher-adapter-v5`
publication. The visible titles distinguish recovered and newly saved documents.
The parent's bridge exposes `pub4F7A.tmp`, not a proven association to the new
file. Its startup fallback accepted the sole non-template project despite an
unavailable host-document association. This is a qualification-fixture defect,
not evidence that the product save adapter failed. No child bridge command is
delivered: connection fails before request transmission.

`retained-native-application-readonly.json`, the genuine owned-window snapshots
and `before-owned-close.json` preserve the association and recovered sources.
Two guarded close messages are posted once. The recovered publication's exact
owned save prompt is discarded through its accessibility Invoke pattern;
`authorized-discard-terminal.json` records parent exit **0**, no Save and no
force termination. The child is absent after its close, but its exit code is
**NOT_OBSERVED**. This cleanup does not qualify any skipped Publisher operation.
The fixture must prove document/window/project association before mutations;
automatic process rebinding and deleting recovery data are not corrections.

### Preceding v4: native export path trace

The integrated disposable non-Office debugger preflight
`cdb-preflight-integrated/592180985be8419bab03d52907cd56b5/preflight.json`
is **PASS**: paired synthetic native file-call tracing, exact debugger/script
hashes, observed attachment and verified detach, helper survival and normal
helper/debugger exits. This is tool qualification, not an Office export result.
The subsequent tests/tool correction `54cd6e8` reads pending/preflight JSON as
UTF-8 explicitly, preserving accented paths in Windows PowerShell. That source
correction is not another native export trial or a passing product result.

| Evidence relative to `artifacts/qualification-v1/followup-20260930/` | Outcome | Exact scope |
| --- | --- | --- |
| `native-export-trace-v4/native-export-trace-v4.trx` | **0 passed, 1 failed, 0 skipped** | Owned Excel **49660**, preceding product `d5e25e25`; UTF-8 JSON path decoding failed during trace setup before any debugger attachment or native export. `setup-failure-reconciliation.json` preserves zero attachment/export attempts. The separate `authorized-normal-cleanup.json` records one NativeOM Close/Quit, normal exit **0**, no force; it does not turn the test into a pass. |
| `native-export-trace-v4-v2/native-export-trace-v4.trx` | **0 passed, 1 failed, 0 skipped** | Fresh Excel **52952**, exact preceding `d5e25e25` MVID/hash, tests-only source `2e75161`; exactly one native export returns `Objet spécifié introuvable.` No successful FRM/FRX capture is established. CDB **4496** detaches verifiably with exit **0**, no forced stop; owned Excel closes normally with exit **0**. |

The fresh trial retains `native-export.json`, `trace.cdb.log` and
`trace.lifecycle.json` under
`native-export-trace-v4-v2/exports/HostBridge-LocalAppData-8f3de0909bbe41c1845ba70e0f3bf185/`.
Six paired native calls report **0xc000003a (path not found)** or
**0xc0000034 (name not found)**, including VBE's `CreateFileA` path and managed
parent opens. The exact ASCII GUID child still exists and root synthetic
write/read succeeds; parent/child encrypted attributes are retained, while the
synthetic EFS metadata query itself reports access denied. This localizes the
observed failure to path visibility in that host, without proving its cause or
attributing it to EFS, ACLs or tokens. The controlled export is not replayed.
Q-027 remains open, and this preceding-product trace is not v5 native acceptance.

## Preceding v4 candidate: complete managed pass (2026-09-30)

Product and test source `b77a782` adds the recovery-message catalogue correction
and two conditional scalar diagnostic pages to the preceding candidate.
The isolated `build-v4` assembly has MVID
`d5e25e25-e3b4-4be0-ad45-a2dceb7be6b1` and SHA-256
`06F9767B6970973B333E5D255210896335DA062C78F4F4116CBF150DD7E121EE`.
Compilation completed with no warning or error. This assembly was installed:
`deployment-v4.json` records the matching installed SHA-256, previous payload
backup and registry exports. Native loaded-MVID acceptance is recorded below;
deployment alone does not qualify Office or SOLIDWORKS operations. It has now
been replaced by the v5 candidate above.

The complete instrumented default suite finished with **2,234 passed,
0 failed, 83 conditionally skipped, 2,317 total**, in **9 minutes 24 seconds**.
The individual TRX outcomes agree with the counters and the product SHA-256
remains unchanged after execution. The previously failing catalogue scenario
passes in this complete run. This closes the managed Q-015 scope for this
compiled candidate; it does not qualify the conditional native scenarios.

Evidence is under `artifacts/qualification-v1/followup-20260930/`:
`candidate-v4.json`, `managed-v4/full-localization-corrected-managed.trx` and
`managed-v4/45350edf-916f-446c-94cb-98992c6dea7c/coverage.cobertura.xml`.
Coverage measures only managed `VBAi`: **33,343/33,529 lines (99.45%)** and
**33,838/34,323 branches (98.59%)**, from the raw integer counters.
Native Office/SOLIDWORKS and live-provider opt-ins were disabled; the native
C++ renderer and JavaScript are outside this measurement.

Separate checks on this source passed **60 JavaScript tests**, the embedded
native-renderer extraction/hash/ABI/module-reuse contract, and test-layout
validation (**265 dedicated mirrors for 323 production files**). Layout
presence is not a coverage measurement. The later crossed-DACL fixture merge
also compiles in an isolated tests-only output against this unchanged product,
but is not included in the complete run's test source revision.

## Preceding corrected candidate: complete managed failure (2026-09-30)

Product and test source `fd65c57` includes guarded asynchronous Access save
verification, preservation of import/recovery and option/restoration errors,
bounded scalar-inspection phase logging and observer reads, and native
active-module following in Monaco. The isolated `build-v3` assembly has MVID
`f04a35b8-1f1e-4b2a-9851-fe9e6afd79ab` and SHA-256
`1ADFE986A1D2BA167AAF28749A0307DE1844B751EA16FEAE6882FE74723B56C1`.
It is not installed and has no native-host acceptance yet.

The complete instrumented default suite finished with **2,232 passed,
1 failed, 81 conditionally skipped, 2,314 total**, in **9 minutes 14 seconds**.
`AllLanguagesHaveCompleteEmbeddedCataloguesAndRecognizableMenus` failed on an
untranslated new Git recovery message in the Spanish catalogue. This is a
failed complete run; focused successes and the older complete pass do not
qualify this candidate's Q-015 gate.

Evidence is under `artifacts/qualification-v1/followup-20260930/`:
`candidate-v3.json`, `managed-v3/full-corrected-managed.trx` and
`managed-v3/c5f24eac-2120-4ab5-afbe-aada09afc2f6/coverage.cobertura.xml`.
Coverage measures only managed `VBAi`: **33,343/33,529 lines (99.45%)** and
**33,838/34,323 branches (98.59%)**, calculated from the raw integer counters.
Native Office/SOLIDWORKS and live-provider opt-ins were disabled. Skipped
scenarios, native C++ and JavaScript are not covered by this acceptance result.

## Preceding installed v4: native follow-up (2026-09-30/2026-10-01)

The accepted scalar pages used then-installed product `d5e25e25` and tests-only
source `c68f85b`, followed by the complete-page scenario from `b81b317`.
The earlier startup failure is retained separately.
The source changes after `b77a782` are not included in the complete managed
run above and do not constitute a new coverage measurement.

| Evidence relative to `artifacts/qualification-v1/followup-20260930/` | Outcome | Exact scope |
| --- | --- | --- |
| `scalar-v4-skipped/scalar-skipped.trx` | **0 passed, 1 failed, 0 skipped** | Fixture startup threw NullReferenceException before opening VBE or checking the loaded add-in. The explicit bootstrap records PID **50200**; its later absence is not an observed exit code. This remains failed preparation evidence. |
| `scalar-v4-skipped-v2/scalar-skipped-v2.trx` | **1 passed, 0 failed, 0 skipped** | Owned Excel PID **8828**, Microsoft 365 x64 **16.0.20326.20158**, exact loaded MVID `d5e25e25`; array, Variant and object rows are skipped without QuickWatch. Native enqueue/STA/context/terminal phases, unchanged source, identity, selection and mode are verified. Normal owned Close/Quit, no forced termination, exit **0x00000000**. |
| `scalar-v4-long/scalar-long.trx` | **1 passed, 0 failed, 0 skipped** | Owned Excel PID **59460**, same Office build and exact loaded MVID; one Long scalar is read through native QuickWatch. Command 229, observer read, continuation and terminal phases complete with unchanged source/identity/selection/mode; normal exit **0x00000000**, no forced termination. |
| `scalar-v4-full/scalar-full.trx` | **1 passed, 0 failed, 0 skipped** | Complete page from tests-only source `b81b317`, compiled against frozen product `d5e25e25`, owned Excel PID **26384** on the same Office build. Exactly one Offset=0/Limit=6 request reads Long=42, the exact quoted String probe and Boolean=True/Vrai, then skips array/Variant/object with no values. One correlation records three Command229/observer/continuation cycles and a terminal outcome; identity, source, mode and selection remain unchanged. Normal exit **0x00000000**, no forced termination. |
| `office-adapter-v4/office-adapter-v4.trx` | **ABORTED campaign; 17 planned, TRX total 9/executed 5: 1 passed, 4 failed; 4 individual NotExecuted rows** | Tests-only source `c68f85b`, installed `d5e25e25`. Publisher Description passes; Access Description/HelpFile save identity checks and Access/Publisher HelpContextID setters fail. Raw counters report notExecuted=0 despite the four individual NotExecuted rows, which remain unexecuted. In-flight Word has no result in the TRX; the remaining planned cases are NOT_RUN. Authorized termination of testhost **23892**, Word **60408**, Access **54632** and Publisher **58748** is cleanup, not a normal-exit pass. |
| `access-identity-readonly/access-readonly-identity.trx` | **1 passed, 0 failed, 0 skipped** | Tests-only diagnostic `eb84f2b`, owned Access 16 PID **3156**, loaded product `d5e25e25`. Five simultaneously retained CurrentProject wrappers have distinct IUnknown identities while database path and mapped/selected VBProject identity remain stable; Application/VBE identity is also recorded. No source/property edit or adapter Save is invoked; balanced references and normal exit **0** are recorded. This is identity diagnosis, not persistence acceptance. |
| `solidworks-2019-v4-vstest/solidworks-load-v4.trx` | **1 passed, 0 failed, 0 skipped** | Owned 2019 SP5 PID **51376**, revision **27.5.0**, exact loaded `d5e25e25`; VBE inventory and VBAi.AddIn connected state verified. No macro execution or shutdown acceptance is included. |

Each accepted scalar page retains its own `inspection.jsonl` plus exact owned
`hosts/<fixture>/startup.json` and `shutdown.json`. The initial fixture failure
was corrected by querying the process image through its retained native handle.
The complete-page trial closes the concrete missing declared-page test scope,
but does not explain the historical `7b5f11d8` bridge stall or combase.dll crash.
It does not enumerate every runtime local type or qualify another host. Q-006
remains open for unresolved historical failure and host lifecycle evidence.

Preceding-v4 SOLIDWORKS native bootstrap is retained in
`solidworks-2019/stage-f4667f08b6b641b7ba00a0cc6d59bb71/`. The module/class/form
readback and copied disk hash pass before UI navigation. Designer and code
resize/restoration pass in `solidworks-2019-v4-designer-resize.json` and
`solidworks-2019-v4-code-resize.json`; reviewed code/designer/returned-class
PNG captures are recorded alongside them. The later class source has an `on`
prefix, and the maintainer reports possible diverted keyboard input. The
cause remains unproven. `solidworks-2019-v4-ui-evidence.json` preserves the
changed source and records PARTIAL acceptance, with no harness source writes,
macro executions or source restoration after drift.

The preceding-v4 2019 instance did not exit normally. After the authorized single
ExitApp request, PID **51376** remained stopped in its native debugger at heap
corruption **0xc0000374**. Address/module observations include `ntdll.dll`,
`ucrtbase.dll`, `mfc140u.dll` and `sldappu.dll`, without resolved symbols; these
frames do not establish the originating defect. The single authorized forced
termination exceeded its ten-second wait while the debugger retained its target.
An initial PID lookup reported absence, but process-name/debugger observations
still showed termination pending; that lookup is not final shutdown proof.
After verifying Visual Studio PID **49796**, its owned utility solution and sole
debug target **51376**, `Debugger.Stop(false)` was invoked once and returned.
The final `dte-readonly-49796-9a9faefd314049d995a6b814abb3ebb2.json` records design
mode and an empty process collection, with no remaining SOLIDWORKS process by
name confirmed in `solidworks-2019-v4-cleanup-terminal.json`. Exit code remains
**NOT_OBSERVED**, and no additional kill was issued.
Retained records are `solidworks-2019-v4-normal-close.json`,
`solidworks-2019-v4-native-crash-frames.json`,
`solidworks-2019-v4-authorized-forced-cleanup.json` and
`solidworks-2019-v4-forced-cleanup-reconciliation.json` and
`solidworks-2019-v4-debugger-cleanup.json` and
`solidworks-2019-v4-cleanup-terminal.json`. None is normal shutdown
acceptance; Q-014 and the relevant lifecycle gate remain open.

Word read-only CDB evidence is
`office-adapter-v4/word-60408-readonly-stacks-v2.log` and its companion JSON.
The nonsuspending/noninvasive inspection ends with debugger exit zero and the
same host still alive; the STA snapshot is in FM20 overlay/visibility handling.
It does not explain the native stall. Tests-only containment source `8d1ee9e`
has **11 passed, 0 failed, 0 skipped** in the agent's
`artifacts/test-results/office-adapter-containment/office-containment-final-pure.trx`.
Those fake-dispatch regressions do not retroactively qualify the blocked batch
or exercise a native host. The installed product hash remains unchanged.

The Office run is now terminal after the maintainer-authorized forced cleanup,
recorded in `office-adapter-v4/authorized-forced-cleanup.json` and
`authorized-retained-host-cleanup.json`. The TRX ResultSummary is Failed;
ABORTED describes the interrupted campaign, not a rewritten TRX outcome.
An absent Word result and unexecuted cases do not become passes or fixture
skips. The later read-only Access probe and normal exit are independent of the
failed batch: `access-identity-readonly/hosts/Access/` retains
`access-identity-probe.json` and its companion qualification/shutdown stages.
It proves that CurrentProject wrapper identity can change without a database
or selected VBProject change. A guarded product correction and fresh adapter
save/reopen acceptance are still pending; no identity guard is removed by this
diagnostic result.

## Historical native and fixture follow-up (2026-09-30)

The product installed for the following historical trials was clean `0ddb0dd`, with
MVID `7b5f11d8-f184-4302-834a-572e92a6ab81` and SHA-256
`332C5B6FADFBB2247A38FE671FC352419E7415A995FA8EBC99CF33A90DD8F3F7`.
Only the test assembly was rebuilt for the subsequent fixture commits; the
product hash was independently rechecked. It has since been replaced by the
preceding installed `d5e25e25` candidate above. Evidence paths in this
section are relative to `artifacts/qualification-v1/followup-20260930/`.

The Excel core batch on test source `558d73c` finished with **13 passed,
1 failed, 0 skipped** in `native-excel/excel-native-core.trx`. The successful
fixtures retain startup identity and normal shutdown evidence. The scalar
inspection case on owned PID **27704** reached break mode and opened Locals,
then its bridge operation failed to respond. Its cleanup failure masked the
original inspection error in the old harness. External COM recovery verified
Reset to design mode before Close/Quit, but Windows event 1000 then records
that same PID crashing in `combase.dll` with `0xc0000005`. It is **not** a normal
shutdown pass. Request history, recovery stages and crash events are retained
in the native TRX attachments, `local-scalars-recovery.json` and
`local-scalars-crash-events.json`. No uncertain operation was replayed.
The subsequently restored empty Excel PID **56748** was identified by its
`/restore` command line and absence of a workbook window, then closed normally
through the owned WindowPattern. Its exit code was zero, recorded in
`restored-empty-excel-normal-close.json`; this does not change the original
crashed process's result.

The adapter-only Office batch on test source `558d73c` finished with **2 passed,
1 failed, 1 skipped** in `office-adapter-only/adapter-only.trx`. Word PID
**14600**, and Publisher PIDs **47216** then **56484**, loaded that exact product
on Microsoft 365 x64 **16.0.20326.20158**. A single adapter save, source-hash
comparison, form/Label preservation, disk readback without a post-save helper,
and exit code zero were verified. Access PID **49324** returned an uncertain
save result because its immediate Saved check was false; the next read-only
observation and all delayed observations showed the project and components
saved. Disk reopen was prevented by a fixture shutdown timeout. The second
Access case was skipped because that instance still existed. These facts do
not qualify Access persistence or retroactively turn the uncertain response
into a verified save. Per-stage records are under `office-adapter-only/hosts/`.

The second adapter-only batch on test source `d8c4747`, with balanced Access
cleanup references, finished with **3 passed, 2 failed, 0 skipped** in
`office-adapter-v2/adapter-v2.trx`. Word, Publisher and PowerPoint pass. The
PowerPoint trial uses PID **35604** on the same Office build and product MVID,
and exits with code zero. Both Access cases now complete fresh disk readback:
active-module-only uses PIDs **48724/23984**; module-plus-class uses
**34008/14508**. Each pending source hash survives without helper Save or
compilation, and each process exits with code zero. Their only retained
scenario failure is the original adapter's unverified response; later
observations and persistence do not retroactively alter that response.
The separate Word/PowerPoint save/reopen scope therefore closes Q-011 for
this candidate, while the Access response correction and additional metadata
scenarios remain open in Q-012.

The first native local-Git layout matrix on test source `b286de1` finished with
**0 passed, 12 failed, 0 skipped** in `userform-layouts/native-layouts.trx`.
All owned Excel instances exited with `0x00000000`. Ten cases stopped before
snapshot/import at the screenshot helper's zero designer HWND. Image stopped
at its external Picture assignment; Frame/MultiPage stopped at a fixture
control-count assumption. These results diagnose fixture preparation and do
not exercise or qualify the remaining Git/form operations. No remote action
occurred. Corrected screenshot ownership and scalar-error evidence have
separate pure regressions: **14 passed** and **5 passed**, respectively;
the Git fixture-cleanup regressions separately passed **3 tests**. These are
not native-host acceptance. Both fixture-contract groups were also executed
together in `fixture-evidence-followup/fixture-evidence.trx`, confirming the
same pure scenarios against the parent test assembly.

The second layout matrix on test source `cadfec3` finished with **0 passed,
12 failed, 0 skipped** in `userform-layouts-v2/native-layouts-v2.trx`. Eleven
cases reached production snapshot capture and failed at native Export with
`0x800AC373`; Image failed at an external StdPicture getter after successful
in-host image installation. The corrected Frame/MultiPage hierarchy checks
passed their preparation stage. Every owned process exited normally with
code zero. Import, recovery, comparison and remote Git acceptance were not
reached.

The dispatch/destination probes on test source `164d1da` finished with
**4 passed, 2 failed, 0 skipped** in `export-probes/export-probes.trx`.
External STA and in-host bridge Export both passed in fixture temporary and
evidence directories, and both failed in the production GitTemporary location.
Their detailed attachments are under the qualification worktree's own
`artifacts/qualification-v1/followup-20260930/export-probes/` directory.
The paired encrypted/plain sibling probes on test source `ab9d5eb` then passed
**2 tests** in `efs-probes/efs-probes.trx`. These outcomes refute a general EFS
incompatibility; they do not explain the production directory failure.
The subsequent C/E volume and parent/leaf inheritance matrix on test source
`eef343e` passed **4 tests**, with no failure or skip, in
`efs-volume-probes/efs-volume.trx`. Both C and E destinations retained their
encrypted attribute and exported successfully. The failure is specific to the
existing production location; neither volume nor generic EFS inheritance is
established as its cause. This tests-only assembly was built in the export
probe worktree against the unchanged `7b5f11d8` product.

The four inherited-storage probes on test source `3f6e805` finished with
**1 passed, 3 failed, 0 skipped**. Export passed in the user TEMP directory
and failed beneath LocalAppData, LocalAppData/VBAi and GitTemporary. This
widens the observed failing boundary beyond GitTemporary; it does not prove
an encryption or permission cause. Each scenario retains query-only token
observations for its testhost and exact owned Excel process. Its TRX and JSON
are under the export-probe worktree's
`artifacts/qualification-v1/followup-20260930/ancestor-probes/`, not the central
evidence directory. No existing parent encryption, ACL or token was changed.

The crossed-DACL matrix on historical `7b5f11d8` finished with **1 passed,
1 failed, 0 skipped** in central `crossed-dacl-v2/crossed-dacl-v2.trx`.
The user TEMP destination with LocalAppData's raw DACL exported successfully
(owned Excel PID **43636**); the LocalAppData destination with TEMP's raw DACL
still failed (PID **59524**). Each disposable child retained its native parent,
EFS metadata and verified copied DACL without permission escalation or changes
to existing parents. Both processes exited normally with **0x00000000**.
The precise descriptor, synthetic access check, original export response and
shutdown are retained in each `HostBridge-*/native-export.json`. This rules out
the observed child-DACL differences alone as the cause of the destination
failure; it does not establish the underlying cause or qualify native Git.

The explicit SOLIDWORKS 2019 SP5 instance, PID **47344**, revision **27.5.0**,
loaded the same `7b5f11d8` candidate. The native connection test passed
**1 test**, with no failure or skip, in `solidworks-2019/vstest/solidworks-load.trx`.
The disposable copied fixture's source/form/hash readback passed in
`solidworks-2019/stage-1c168f3037ac4cda8a6b567f45af024d/`.
Designer resize and exact restoration passed in that stage's
`designer-resize.json`. The reviewed designer capture shows the synthetic
label without clipping; the return-to-code capture reveals an empty Monaco
shell despite correct native module selection. Captures are under
`../solidworks-ui/followup-2019-designer-7b5/` and
`../solidworks-ui/followup-2019-return-code-7b5/`.

A distinct native-created macro passed module/class/form editing, stale-hash
refusal, compilation and one verified adapter save. The synthetic marker ran
once and unloaded; the strict scenario nevertheless failed because SWP bytes
changed. Separate guarded native Edit Macro and source/class/label readback
passed against the observed post-execution hash. Original failure evidence is
retained, and no whole-binary preservation or standalone-project acceptance
is claimed. These artifacts are under
`../solidworks/type100-2019/followup-7b5/`, including
`type100-qualification.json`, `native-open-after-runtime.json` and
`post-runtime-shared-file-readback.json`.
The later authorized close completed normally with exit code **0** and no force
termination, recorded in `solidworks-2019/normal-close.json`. Before closing,
one fixture had a whitespace-only live-source difference and Saved=false.
`solidworks-2019/preserve-before-close/preservation.json` retains the exact live
sources and form before/after one verified product Save, with no source write
or macro run. That save does not assert fresh-disk reopen or upgrade the earlier
binary-preservation failure. This session still loaded historical `7b5f11d8`;
it is not native SOLIDWORKS acceptance of the subsequently installed product.

The unchanged strict loopback Ollama scenarios passed **3 tests** with passive
wire capture and **3 tests** without the capture wrapper, recorded in
`ollama/ollama-native-transport-ui.trx` and `ollama-unwrapped/ollama-unwrapped.trx`.
They verify visible streaming, Stop and the next completed reply in detached
controls with simulated VBE. They do not explain the historical ce19 empty
response, whose failed wire payload was not retained; Q-028 remains open.

## Complete managed qualification follow-up (2026-09-30)

The clean source commit `0ddb0dda880a843be3fdc5c4dbb99c7d120034a7` was built in
an isolated output. The tested assembly has MVID
`7b5f11d8-f184-4302-834a-572e92a6ab81` and SHA-256
`332C5B6FADFBB2247A38FE671FC352419E7415A995FA8EBC99CF33A90DD8F3F7`.
The complete suite with XPlat coverage collection finished with **2,160 passed,
0 failed, 36 conditionally skipped, 2,196 total**, in **8 minutes 30 seconds**.
Every previously retained failing scenario passed in this complete run. This
closes Q-015 for this candidate; the historical failure causes remain unproven.

Evidence is local under `artifacts/qualification-v1/followup-20260930/`:
`candidate.json`, `managed/full-managed.trx` and
`managed/c390190e-2869-4f4b-8a8e-b63703a41244/coverage.cobertura.xml`.
The TRX individual outcomes agree with the counters. The assembly hash remained
unchanged and the source worktree was clean after execution.

Coverage measures only the managed `VBAi` assembly: **33,090/33,267 lines
(99.47%)** and **33,553/33,977 branches (98.75%)**, calculated from the raw
integer counters. Native Office/SOLIDWORKS and authenticated/live-provider
opt-ins were disabled; their skipped scenarios do not qualify those paths.
The C++ renderer and JavaScript are outside this measurement.

Separate checks passed on the same source: **60 JavaScript tests**, the synthetic
native renderer's **20 start/stop cycles**, the managed/native loader contract,
and documentation validation (**36 maintained Markdown files, 171 local links,
0 errors**). These are separate boundaries from native host acceptance.

## Open qualification: managed failures and owner-thread corrections (2026-09-30)

This checkpoint supersedes the older managed-run status below, without replacing
its historical evidence. Paths in this section are relative to
`artifacts/test-results/open-qualification/`. No native host acceptance or
whole-suite pass is inferred from the focused follow-ups.

### Latest diagnostic candidate and interrupted complete run

At this historical checkpoint the diagnostic product candidate was **installed**: MVID
`ae8a0978-db91-40f3-b8f7-8957d1a6b7b7`, SHA-256
`B80B6886DBAAEE61204B5FD6D91DE9752241B303F2833744C195A3C0CF740F26`,
compiled under `artifacts/build/office-options-diagnostics/VBAi/Debug/net48/`.
After the user closed SOLIDWORKS and absence of Office/SOLIDWORKS hosts was
verified, installation completed with exit code **0**. Independent readback of
`bin/Debug/net48/VBAi.dll` matches that SHA-256. The full preceding `95576771`
payload, seven registry exports and installation log are retained under
`artifacts/installation-backups/office-options-diagnostics-20260930/`.
The log's final result is **Registration=OK**, **ComActivation=OK**,
**OnConnection=NOT_TESTED**, **ChatMonaco=NOT_TESTED**. Installation and standalone
COM activation do not establish native add-in or panel acceptance.

`options-category-diagnostics.trx` records **20 passed, 0 failed, 0 skipped**
in `WritableOptionsTests`: one new regression for the precise requested
category, UI Automation readback and native-index diagnostic, plus 19 existing
cases. Separately, `office-options-diagnostics.trx` records **38 passed,
0 failed, 0 skipped**, all in `VbeOtherHostPersistenceTests`. A class-name
filter mismatch excluded the options tests from that second run; its filename
does not demonstrate options coverage. Neither group establishes a native
format-options fix or successful Office persistence.

The native-disabled complete managed run with the XPlat coverage collector was
**INTERRUPTED at the maintainer's request** when development was paused. The
runner received Ctrl+C and exited with code 1; no final
`options-diagnostics-full-managed.trx`, complete counters or completed coverage
result was produced. This is an incomplete run, not a passing qualification or
a newly diagnosed test failure. The interruption record is
`options-diagnostics-full-interruption.json`.
`options-diagnostics-full-source-before.json` records **1,159 files**, the
candidate SHA-256 above and baseline `ffb4984e24c18f793006d2a8a98ce816ccb54382`.
The interruption comparison found **0 changed source inputs**. The six failures
from the preceding complete run below remain recorded and Q-015 remains open.
The planned Access and native options trials on this installed candidate were
not started before the pause.

### Final managed candidate and retained failures

`final-full-managed.trx` completed with **2,140 passed, 6 failed,
36 conditionally skipped, 2,182 total**. Individual TRX outcomes agree with these
counters. The runner-reported duration is **28 minutes 27 seconds**; the TRX
start/finish interval is approximately 28 minutes 49 seconds. This is a failed
whole-suite result, with no release or universal compatibility acceptance.

The compiled candidate is MVID `096b2e2b-73fb-4d97-8bd8-a4abda6d78eb`, SHA-256
`C0516746881D3A7D94C8BC29C25209A6F4513D6DFD078794F90270C9B32B8BFC`.
`final-source-manifest.json` identifies this binary and **813 source entries**
against baseline `ffb4984e24c18f793006d2a8a98ce816ccb54382`. No source modification
was recorded during the run. The candidate includes the subsequent
`StartUiAction` UI-wrapper preparation correction; it is no longer a pending
change outside the tested binary.

`0b7811bc-257a-48de-8f74-4ef787376842/coverage.cobertura.xml` records
**33,016/33,201 managed lines (99.44%)** and
**33,489/33,927 managed branches (98.71%)**, calculated from the raw counters.
Coverage belongs to this failed run and candidate. It excludes measurement of
external Office/SOLIDWORKS processes, the C++ renderer and JavaScript.

All six failures remain recorded; their causes have not been established by
this run. In particular, a timeout is not classified as a harmless fixture
failure without further evidence.

| Failing scenario | Observed result |
| --- | --- |
| `ConflictPreviewAndOursTheirsTextResolutionMatrix` | Disposable GitWindow operation timed out. |
| `NativeCredentialChildReceivesNoninteractiveInputAndNeverLaunchesGcm` | Expected `OperationCanceledException`, received `IOException`. |
| `ResolveReloadAndRestoreWithoutAmbientContextKeepRendererOwnershipAndArchivedDraft` | Recovery action did not complete; the displayed status reported synchronization with VBA. |
| `CommitSynchronizationStaleStateAndOperationStatusMatrix` | Disposable WinForms test thread timed out. |
| `BranchCheckpointRemoteAndMergeActionsUseNativeGitAndReturnToLiveState` | The fixture cache path was in use by another process (`IOException`). |
| `HistoryCheckpointAndModuleRestoreNavigateActualCommitSnapshots` | Disposable WinForms test thread timed out. |

### Installation and first native Access attempt

The later `office-identity-git-cancellation-verified.trx` focused run passed
**31 tests, 0 failed, 0 skipped** on MVID
`95576771-3adb-4991-99e6-d47e593a5fc6`, SHA-256
`5CACC24A593FF763FDB2A6ED58FAF5F82112B18023B590B8B2440B5F03917E29`.
It covers Office persistence identity/refusal contracts and Git credential-input
cancellation. The native credential-child scenario and deterministic injected
input failure both passed: an uncancelled IOException remains the same error,
while cancellation returns OperationCanceledException. The preceding
`office-identity-git-cancellation.trx` remains **30 passed, 1 failed**: killing a
child did not deterministically cause an IOException on a buffered pipe; the
revised regression injects that failure explicitly while retaining the real
child integration scenario. This does not clear the complete-run gate.

Candidate `95576771` was previously installed. The previous payload, seven registry
exports and installation log are retained under
`artifacts/installation-backups/office-followup-verified-20260930/`.
The installation log records Registration and COM activation OK; its installed
DLL SHA-256 was independently read back at that checkpoint as
`5CACC24A593FF763FDB2A6ED58FAF5F82112B18023B590B8B2440B5F03917E29`.
The historical native Access bridge below confirms its exact loaded MVID and PID. These
checks do not imply successful native persistence or a passing complete suite.

The later guard-diagnostic candidate is **compiled, not installed**: MVID
`93e7596d-4950-41c7-b809-7d33dd573d75`, SHA-256
`80DF77E6385A61B227528354E55A5BBC34A7083775A4AA07C7FAFD31FF484D2A`,
under `artifacts/build/office-save-diagnostics/VBAi/Debug/net48/`.
`office-save-diagnostics.trx` records **38 passed, 0 failed, 0 skipped**:
29 existing cases and nine regressions for exact failed-guard reasons
(`ProjectIdentity`, `ProjectPath`, `HostPath`, `HostSaved`, `ProjectSaved`,
`SourceSha256`, `FileFormat`, `FileExists`, `FileLength`). The diagnostic change
preserves ordered short-circuit verification and save guards, without replaying
the mutation. It adds no native persistence acceptance. This candidate was not
deployed; the later `ae8a0978` installation is recorded above.

Installation of `93e7596d` was attempted, but the script refused the newly
running SOLIDWORKS instance, PID **769136**, started at **17:07:35** on
2026-09-30, before registration or installed-payload mutation. The directory
`artifacts/installation-backups/office-save-diagnostics-20260930/` contains
the copied prior payload, but installation is **NOT_RUN** and no installation
log was produced. At this refused deployment, independent readback of the
installed DLL still gave
`5CACC24A593FF763FDB2A6ED58FAF5F82112B18023B590B8B2440B5F03917E29`.
This refusal is separate from the subsequent successful `ae8a0978` installation.

`native-access-bounded-dialogs.trx` records **0 passed, 1 failed, 0 skipped**.
The detailed record is
`native-access-bounded-dialogs/Access/487f93d3b7c34ed3a1bd8cf977254355/qualification.json`,
owned PID **783612**, candidate `95576771`. Unlike the preceding blocked attempt,
the bounded dialog worker stopped and the guarded adapter save was invoked
**once**. Its response has `Ok=true`, `SaveInvoked=true`, **`Verified=false`** and
**`Uncertain=true`**. Independent before/after observations retain
`ProjectSaved=false`, with verified project/document identity and the same file
length. A successful protocol response is therefore not proof of persistence.
The owned process then failed to exit within the fixture's five-second cleanup
deadline, and reopening was refused. No save retry or successful round-trip is
claimed.

The later `owned-disposable-access-recovery.json` in that same fixture directory
records fresh confirmation of PID **783612** and its exact disposable database
path before a single `Quit(acQuitSaveNone)`. The process exited normally with
code **0**, without forced termination. This separate recovery does not change
the failed native test or demonstrate that the adapter persisted its edits.

The subsequent
`persisted-inspection-3805a42b13704486a644a25ecadc46bd/inspection.json` in that same
fixture directory records inspection of an exact copy of the retained database.
The original and initial copy SHA-256 are both
`9A761CD59740C55ADC65F62DF215A4E95E5591928A73AA613DAA93AA928ACA9D`, and
`OriginalUnchanged=true`. Fresh Access **16.0**, owned PID **791016**, reopened
the copy with `ProjectSaved=true`; inspection performed no save or compilation,
then `Quit(acQuitSaveNone)` completed with normal exit code **0**.

Both edited sources match their latest pre-save `read_module` records in
`qualification.json`: `VBAiOfficeModule` has SHA-256
`dba6dfe936c1263da64a2f23fd8f9a3fc3430d7039c3939bad15ad424fb390d1`, and
`VBAiOfficeClass` has SHA-256
`1eeed8d60e9d15eb42979777b2225c2af8aaff837b5e0a26f4f62e1d23efe7b5`.
The inspection's `AdapterPendingEditPresent=false` for the class checks the
module-specific marker; it is not evidence that the class edit is missing.
The original campaign compiled the project after the uncertain adapter save
and later quit Access. This snapshot proves the final retained source contents,
but cannot isolate adapter-only persistence or explain the earlier
`ProjectSaved=false`. The failed native trial and its qualification gate remain
open. A test-fixture extension for delayed, read-only observations after an
uncertain Access save now compiles; it has not been executed against the native
host, and no delayed-state result is claimed.

Two focused reruns of the four previously failing GitWindow scenarios completed:
`git-window-retained-failures-diagnostic.trx` records **4 passed, 0 failed,
0 skipped**, and `git-window-retained-failures-coverage.trx` records **4 passed,
0 failed, 0 skipped** with the XPlat coverage collector enabled. They cover
conflict resolution, branch/merge actions, checkpoint synchronization and
history/module restoration. These are two runs of the same four scenarios, not
eight distinct tests. Neither their isolated success nor the focused collector
replaces `final-full-managed.trx`: its **six failures remain recorded**, their
full-run causes unresolved, and the complete-suite acceptance gate remains open.
No new global coverage percentage or current Access/Publisher adapter acceptance
is inferred.

The earlier `096b2e2b` candidate was deployed using the installation and verification
scripts, with its TLB regenerated. The previous DLL payload and seven registry
key exports are retained under
`artifacts/installation-backups/20260930T111550Z-0a8e84657f204d70b0a74d83248d41e8/`.
Its `deployment.json` and `installation.log` record **Registration=OK** and
**ComActivation=OK**; initial **NativeOnConnection=NOT_TESTED** remains a separate
gate. Registration and standalone COM activation do not prove native chat or
Monaco initialization.

`native-access-final.trx` records **0 passed, 1 failed, 0 skipped** for
`AccessDatabaseRoundTrip`. Startup did not reach an identified VBA project; the
original startup failure was obscured by cleanup errors (`CurrentProject`
unavailable, then the owned process still alive). The unversioned
`Access.Application` registration selected `Access.Application.8` through CurVer,
and the launched binary had native file version **8.0.4122**. Microsoft 365 Access
**16.0.20326.20158** and the versioned `Access.Application.16` registration are also
installed. This attempt therefore does not qualify the installed Access 16
adapter, nor establish its failure.

The later recovery record
`native-final/Access/3f6c90b4ed874f52bd1b97332ce13fec/owned-access-diagnostic.json`
belongs to the freshly revalidated owned PID **742936**: no database was open,
`Quit(acQuitSaveNone)` was invoked, and the process exited normally with code 0.
It demonstrates recovery of the empty test instance without forced termination;
it does not convert the failed native test into a pass. The completed native
attempts below replace the previously pending statuses, while preserving this
initial failure as separate evidence.

### Completed native follow-up on the installed candidate

The Office `status` replies in the JSON records below identify the installed
`096b2e2b-73fb-4d97-8bd8-a4abda6d78eb` assembly, connected bridge and 64-bit host.
They provide per-host connection evidence beyond the installation-time
`NativeOnConnection=NOT_TESTED` snapshot. They do not establish complete UI or
debugger compatibility. Counts describe separate runs and are not combined with
the managed suite or historical native passes.

| Evidence | Recorded result | Scope and boundary |
| --- | --- | --- |
| `native-word-powerpoint-accessible.trx` | **1 passed, 1 failed, 0 skipped** | PowerPoint document round-trip passed; Word failed after a native form request timed out. |
| `native-excel-final.trx` | **1 passed, 2 failed, 2 skipped** | Native form fitting passed; Monaco diagnostic removal and native format-options selection failed. Save and language scenarios were skipped to preserve existing Excel processes. |
| `native-access-publisher-ms365.trx` | **0 passed, 2 failed, 0 skipped** | Access did not reach the adapter-only save trial; Publisher refused document/project path binding and failed shutdown/reopen. |

**PowerPoint:**
`native-final/PowerPoint/0431a384903e48798619537570f12968/qualification.json`
records PID **786336**, module/class/form editing, guarded adapter save,
navigation/compilation and exact saved-presentation readback. The independent
`BeforeAdapterSave`, `ImmediatelyAfterAdapterSave` and `AfterAdapterOnlyReopen`
observations retain the same source SHA-256
`55b66cebccc619827cf03e46c64191934b2cf390f1ded78ce2a3b250d0fd4684`;
document/project Saved change from false to true and remain true after reopening.
The owned process exited normally with code 0. This passes this PPTM scenario,
not every PowerPoint feature or the entire Office campaign. An earlier
`native-final/PowerPoint/d82ba2a55110405aa96966d37831a5fc/qualification.json`
attempt failed on the programmatic-project-access prerequisite and remains
separate evidence.

**Word:** `native-final/Word/f1a0cfbba49543338a5b7f2d6ca76917/qualification.json`
records PID **775964**. Inventory, references, module/class editing and stale-write
guards passed before the `create_form` response deadline expired. Delivery was
uncertain and the request was not retried. Later reads did not answer; cleanup
received `RPC_E_CALL_REJECTED`, and the owned process was retained without forced
termination at the end of this attempt. No save/reopen or native-form acceptance
is inferred. The earlier project-access refusal in
`native-final/Word/731d23aebe2c4fec8f99c67d12a0ac54/qualification.json` is distinct
from this later failure.

The user subsequently closed Word and reported a restart/debug crash window.
`native-final/Word/f1a0cfbba49543338a5b7f2d6ca76917/shutdown-event.json` retains
Application Error **1000**, recorded at **2026-09-30 15:50:32 +02:00**, for
WINWORD.EXE 16.0.20326.20158. The event identifies PID **0xbd71c (775964)**,
KERNELBASE.dll and exception **0xe0434352**, report
`0112a1dd-cb98-4878-81f1-696d8d7dcb12`. The cause is **NOT_ESTABLISHED**.
No agent forced termination occurred; the user's closure and recorded crash do
not convert the failed native trial into normal-exit or operation acceptance.

**Excel:** the Monaco scenario retained **one compiler diagnostic marker after
the native source was corrected**, where zero was expected. Its TRX output also
records abnormal shutdown of owned PID **778204**, exit code **0xE0000002**;
the primary assertion and shutdown failure are both retained without assigning
a crash cause. The format-options scenario reported that the native Code Colors
category did not retain its selected value. Its baseline, command requests,
outcomes and available restoration evidence are retained under
`native-final/ExcelFormatOptions/options-evidence-cafec7c7fcf7418ba7640130cf9cd870/`
for PID **778476**. That failure does not prove complete restoration. The native
form-fitting pass does not qualify the failed Monaco/options scenarios, and the
two skipped scenarios do not establish save or language acceptance.

**Access 16:**
`native-final/Access/e33c934c1fe54fb1bfb5e976464e5b92/qualification.json`
records connected candidate `096b2e2b` in PIDs **764612**, then **776040**. The
save-dialog handler did not stop before the required adapter-only mutation, so
that guarded trial was not performed. The later helper-assisted reopen could
not find `VBAiOfficeModule` and failed independently. Both owned processes exited
normally with code 0. These observations do not demonstrate adapter persistence;
they also do not turn the separate Access 8 startup failure into a pass.

**Publisher:**
`native-final/Publisher/03f6d549e2134a598edbd1e716337cf0/qualification.json`
records PID **781804**. The persistence adapter returned `HostAvailable=false`
because the document path did not identify the selected VBIDE project. A safe
refusal does not pass the required save/reopen scenario. Although later project
Saved readback was true, shutdown did not complete within the owned fixture's
deadline; the process was retained without forced termination and reopening was
refused. Native save/reopen and normal-exit acceptance therefore remain open.

The partial current successes above do not replace failures from the same runs.
Earlier Office passes later in this document remain tied to their own source
manifests and binaries; they cannot be promoted to acceptance of this candidate.

### Publisher pathless-project correction and later native refusal

The subsequent Publisher candidate is MVID
`de3c5a79-3962-49dd-9fe3-a65c18900e13`, SHA-256
`C7429807E5A81647338FBA0A648DD9F462EB4806E5F28F6F016813FEE38F8327`.
It was installed with the previous payload and registry exports retained under
`artifacts/installation-backups/publisher-pathless-20260930/`. Its installation
log records Registration and COM activation OK; native connection is separately
confirmed by the later Publisher bridge replies. The installed DLL hash was
read back for this documentation update. The preceding `096b2e2b` whole-suite
and coverage results do not measure this newer binary.

`publisher-pathless-contracts.trx` records **28 passed, 0 failed, 0 skipped** for
the focused contracts. `native-publisher-pathless.trx` records **0 passed,
1 failed, 0 skipped**, with test duration **28.07 seconds**. The native refusal
occurred before the corrected pathless-project guard could be exercised:
`HostAvailable=false`, with reason `No running Publisher application was
verified as belonging to this VBE PID.` The run therefore neither qualifies the
native pathless fix nor proves that this fix failed after application resolution.

`native-publisher-pathless/Publisher/89dbc5059ae64d37b7ff8dea10e05163/qualification.json`
records the exact candidate loaded in owned PIDs **775408**, then **786636**.
Both processes exited normally with code **0**, without forced termination.
The later native save/reopen subscenario reports PASS but also records
`HelperSaveInvoked=true` and `AdapterOnlyClose=false`: it validates only that
helper-assisted round-trip, not persistence through the production adapter.
The overall native test remains failed, and the earlier Publisher binding and
shutdown failures are retained as separate attempts. No new global pass or
coverage measurement is claimed for this follow-up.

### Earlier resource-guard candidate and focused owner-thread repairs

The preceding full instrumented run `resource-guards-full-managed.trx` returned
**2,131 passed, 12 failed, 36 conditionally skipped, 2,179 total**. These counts
were checked against individual TRX outcomes as well as the summary counters.
It tested candidate MVID `99ca80c4-2bbc-4ddd-b4d0-1d3957610ba4`, SHA-256
`C74ABC336867F78160D19B11A8CA67E4C0953EA60C8ED82247037A25CAA65EFD`.
`resource-guards-source-manifest.json` records baseline commit
`ffb4984e24c18f793006d2a8a98ce816ccb54382`, the candidate identity and 813 source
entries. The run remains failed; subsequent source changes and isolated successes
do not turn it into a passing full-suite result.

The associated
`d33a8f97-ef32-46e9-8613-ce48fbfb18aa/coverage.cobertura.xml` reports
**32,979/33,158 managed lines** and **33,468/33,919 managed branches**
(approximately **99.46%** and **98.67%**). These are measurements of that failed
run and compiled candidate, not of the later working tree. They do not measure
the native renderer, JavaScript or code executing inside Office/SOLIDWORKS.

The twelve failures include obsolete fake FRX data, a long-cache fixture path,
a case-sensitive Word temporary-path expectation, browser mocks that did not
follow the per-editor profile contract, and real Monaco UI ownership failures.
The fixture corrections retain the production resource/path guards. Separate
failing-before tests establish loss of the owning editor thread across awaits
when no ambient synchronization context is available; this is a production
defect, not a reason to dismiss all failures as fixture timing. The corrected
paths cover Monaco tools, diff/close, conflict resolution, native reload and
archived-draft restoration. The complete original TRX is retained.

| Evidence | Recorded result | Scope and boundary |
| --- | --- | --- |
| `monaco-owner-thread-red.trx` | 0 passed, 1 failed | Reproduces renderer callbacks leaving the owning editor thread during a Monaco read without ambient context. |
| `monaco-ui-callback-thread-red.trx` | 0 passed, 1 failed | Reproduces wrong-thread renderer access in the diff/close callback path. |
| `monaco-recovery-thread-red.trx` | 0 passed, 1 failed | Reproduces the same ownership defect in recovery, beginning with conflict resolution. |
| `suite-fixtures-owner-thread-green.trx` | 48 passed, 0 failed, 0 skipped | Git fake-FRX/long-path, Word identity/path, browser contracts and the initial owner-thread regression; a focused repair group. |
| `webview-profile-lifecycle.trx` | 3 passed, 0 failed, 0 skipped | Profile lifetime and cleanup guards; not native Office acceptance. |
| `monaco-tools-real-owner-thread.trx` | 3 passed, 0 failed, 0 skipped | Detached real WebView2/Monaco tool and synchronization scenarios after the owner-thread correction. |
| `monaco-ui-owner-thread-final.trx` | 55 passed, 1 failed, 0 skipped | Intermediate group still timed out in `RealRendererEditsSynchronizesRejectsStaleReplacementsAndDisplaysDiff`; the filename does not imply success. |
| `monaco-recovery-owner-thread-green.trx` | 57 passed, 0 failed, 0 skipped | Recovery, diff/close, tool contracts and detached real renderer scenarios, including the previously timing-out renderer case. Separate runs are not summed as unique coverage. |

The completed recovery-focused group above tested
`artifacts/build/open-qualification-recovery-final/VBAi/Debug/net48/VBAi.dll`,
MVID `c7238910-aa98-4c9f-aa9d-afa5e45068f3`, SHA-256
`E77D84759312F8FDACDF83683BE5B5AAB15506D9A8297A2274591077E02A393B`.
The DLL hash was read back while documenting this checkpoint. This focused result
is not a whole-suite or coverage measurement. The later `StartUiAction`
UI-wrapper preparation correction is included in candidate `096b2e2b` and its
failed full run recorded above. The earlier native results below remain tied to
their own binaries.

## Full-branch publication check (2026-09-30)

Source checkpoint: `b6778560e9370b0bc15da74d9d7426a71b93d89c`, including the
qualification changes and the merge of main's `b6eff51` correction. The merge
introduced no additional production or test content relative to the compiled
working tree. Debug solution compilation completed without warnings or errors:

```powershell
dotnet build VBAi.sln -c Debug -p:BuildOutputRoot="$PWD/artifacts/qualification-v1/pr-full-build" --no-restore -v:minimal
dotnet test tests/VBAi.Tests/VBAi.Tests.csproj -c Debug --no-build -p:BuildOutputRoot="$PWD/artifacts/qualification-v1/pr-full-build" --logger 'trx;LogFileName=qualification.trx' --results-directory artifacts/qualification-v1/pr-full-tests -v:minimal
```

The resulting `VBAi.dll` SHA-256 is
`494567E025CAB4F3869FE82222CA4063582B56550A75C2FB9C227F99050B84BC`.
The complete default run returned **2,107 passed, 1 failed, 36 conditionally
skipped, 2,144 total**, in 8 minutes 37 seconds. These counts were verified against
individual TRX outcomes, not only its summary counters. No coverage collector
was enabled for this run.

`CommitSynchronizationStaleStateAndOperationStatusMatrix` failed because the
restore action reported that VBA changed during the operation instead of
confirming restoration. One isolated rerun on the same binary passed (1 passed,
0 failed, 0 skipped), recorded in `pr-full-git-recheck/git-restore.trx` using
`--filter FullyQualifiedName~CommitSynchronizationStaleStateAndOperationStatusMatrix`.
The cause is unresolved; the isolated success does not replace the failed full
run or establish that the issue is only a test timing problem.

Additional publication checks passed: 60 JavaScript tests, test-layout validation
(258 dedicated mirrors for 316 production files), and syntax parsing of the
17 added PowerShell probes. Probes were not executed against native hosts.
The 13 upstream license/notice text payloads match their manifest hashes in Git;
`.gitattributes` preserves their exact bytes. Whitespace validation excludes
these verbatim upstream texts, which retain their original whitespace.

Native Office/SOLIDWORKS and authenticated-provider acceptance were not repeated
for this publication check. The conditional skips remain unqualified on this
binary. The historical native results and coverage below remain tied to their
own source manifests and compiled candidates. The branch PR does not change the
release decision or close the outstanding qualification gates.

## Native qualification campaign checkpoint

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
| Word same-name Git diagnostics, native03 TRX (local artifact: `artifacts/qualification-v1/word-git-native/353-native-03/native.trx`) and phase evidence (local artifact: `artifacts/qualification-v1/word-git-native/353-native-03/hosts/Word/2cb46193741e415ab5773c1539745096/qualification.json`) | 1 failed, no skips | Product 353d unchanged; test-only phased harness. Both production Git captures fail with 0x800AC35C. Independent phases pass: canonical resolution of two Project-named DOCMs, bridge native selection, stale Git binding refusal after owned SaveAs, other-document source preservation, and bridge/external-STA exports to artifact paths with identical SHA-256 `AB893DEF2C283CDC4A0F3B8EE5363508EDF1A1F7B7320528C4FA5703E4B635F5`. Word PID 50144 exits 0 without forced termination. No macro, UserForm or remote operation; no actual chat/Git UI opening claim. Native01 failed on a nonexistent harness command; native02 first exposed Capture failure. Both earlier runs remain preserved with normal exits. |
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
The latest complete managed candidate is identified at the top of this page.
The historical `sw-async-global/coverage-summary.json` identifies its own successful
default suite, pre-run source manifest, detected later edits and canonical Cobertura output.
Product binaries were not instrumented; collection used the test output copy.
Historical checkpoints in the evidence table are not current-source measurements.

C# collection does not instrument separate Office/SOLIDWORKS processes, the C++
renderer or JavaScript. Conditional skips are not passes. A provider response
does not prove native mutation; a helper save does not prove the product adapter.
No production exclusions were added to improve the metric. Ignored local
artifacts are not public downloadable reports; earlier records remain in Git
history.

## Open qualification follow-up (2026-09-30)

Baseline `ffb4984`; work on `fix/open-qualification-gates`. The test project and
its production dependencies built in Debug/net48/x64 with no warnings/errors at
`artifacts/build/open-qualification`. Candidate MVID:
`9924660b-8b89-46de-9910-6dcddad1d158`; SHA-256:
`3A62B3C0D7F65367B05400AE08948C677901B0E0CA66FFC8D71EAA4ADFF69A62`.
This isolated candidate was not installed into Office or SOLIDWORKS.

`webview-profile-lifecycle.trx` records **3 passed, 0 failed, 0 skipped** under
`artifacts/test-results/open-qualification`. Cases cover both retirement/exit
orders, a wrong browser PID, a replacement browser invalidating old exit evidence,
retention after uncertain initialization, a locked cache file, and preservation
of another profile and unknown prior data. The real WebView2 case opens two Monaco
windows, closes each normally, observes profile removal after runtime exit, and
checks the other window remains ready with a responding renderer. Its VBE is
simulated; this is real browser lifecycle evidence, not Office-host acceptance.
Test layout and whitespace checks also passed. No full managed suite or coverage
measurement was run for this change.

An independent disposable Excel export diagnostic ran on Office 16.0 build
20326, PID 734824. Native module and UserForm exports succeeded once to each of
GitTemporary, system TEMP and a fresh artifact directory, with nonexisting output
files. Excel exited through normal Close/Quit with exit code 0. Local evidence:
`C:/Users/jvc/Documents/Codex/2026-09-25/bo/artifacts/open-qualification/export-paths/bdd7091d986b4ee79adee3bb5e20620d/export-paths.json`.
No macro ran, no host trust setting changed and no existing user project was
used. This diagnostic did not call the Git coordinator or qualify the Word path;
Q-024/Q-027 remain open. Successful raw exports on this workstation do not explain
the historical export failures on another qualification environment.

### FRX preflight and scope follow-up

The next isolated build is `artifacts/build/open-qualification-frx`, still based
on `ffb4984` plus the uncommitted follow-up source. Debug/net48/x64 build passed
without warnings/errors. Production candidate MVID:
`7b2423c3-eb9b-4b25-a622-aaf48c03c9b4`; SHA-256:
`EC2A0B469EB0AE49F69777F713A7EA75FAB85DC2A2C5F4F796D9541E8DCB1D81`.
It is not installed. `frx-scope-options-focused.trx` records **19 passed,
0 failed, 0 skipped** with this build: bounded OLE/CFB preflight and snapshot
validation, detached real-control Send/Resume/Stop transitions and the existing
options revision/category guards. The CFB cases include mini streams, normal
version-4 streams, DIFAT extension, allocation aliasing/cycles, lengths and
truncation. No full-suite or coverage result is claimed.

The modified native options fixture compiled but was not rerun: it now retains
requests, outcomes and successful before/after observations in a unique optional
durable output directory and targets the exact observed category during palette
restoration. This does not identify the historical revision drift or qualify the
native restoration. The scope tests use a simulated VBE with real detached UI
controls; embedded-host/UIA acceptance remains open.

Local production-capture evidence is below
`C:/Users/jvc/Documents/Codex/2026-09-25/bo/artifacts/open-qualification/`:

- `production-git-excel/328828a277d14d629dd812d60bbb9715/qualification.json`:
  candidate `9924660b`, Excel PID 748272, Office 16.0 build 20326; unedited
  snapshots differ only in the FRX file, and guarded Apply refuses before import.
- `production-git-excel-preflight/6c28bee352ad49dfbf7f20f3a8217f86/qualification.json`:
  candidate `7b2423c3`, Excel PID 748240; native FRX exports pass structural
  preflight, while the raw-comparison import remains refused.
- `production-git-excel-preflight-corrupt/7335b7df267e41ecbef2dea0978223a3/qualification.json`:
  the same candidate, Excel PID 752408; a nonempty native FRX with its CFB
  signature deliberately damaged is rejected at snapshot construction, with the
  synthetic Label caption unchanged. The intact exports remain byte-preserved.

All three owned Excel processes exited normally through Close/Quit with exit
code 0. They used disposable module/form controls and no macro execution. The
isolated production assembly ran on an external STA against native COM; these
results do not qualify the installed bridge, embedded Git UI, GitHub transport,
successful import/recovery or persistence. An external Word harness could not
obtain its project and did not reach production Capture; its owned processes
closed normally, but it adds no Word acceptance. Q-024/Q-027 remain open.

`git-synchronization-matrix-accepted-env.trx` separately records **1 passed** on
candidate `9924660b` for the previously intermittent managed Git/UI matrix. Earlier
attempts here failed before exercising that scenario: the launcher inherited
both `PATH` and `Path`, then the sandbox refused GitTemporary creation. The
successful isolated rerun uses a child-only normalized environment and the
authorized scratch directory. It does not explain or override the historical
full-suite restore failure.


### Logical form comparison, native import and host/stream follow-up

The isolated `open-qualification-logical-forms` build passed Debug/net48/x64
without warnings/errors. Candidate MVID `3eda62bd-9a90-4c91-bb8a-d299e1c9cc15`,
SHA-256 `EFF0AD2FE67011F15E2E571DA7D2DF74D7770D8952841912A53B109FDA84A3FC`.
`logical-forms-host-stream-focused.trx` records **78 passed, 0 failed,
0 skipped**: CFB/form comparison and snapshot/adapter guards, Git revision,
stream/client protocols and Office save contracts. Its native Excel import
restored controls but failed final FRM equality because VBIDE inserted a leading
code line. This failure remains recorded, not replaced by the managed result.

The next `open-qualification-form-import` candidate is MVID
`a37d53f0-bff8-4bee-889c-e35e27f29777`. Its build passed without warnings/errors;
`form-import-host-stream-focused.trx` records **82 passed, 0 failed, 0 skipped**,
including the initial prefix fix, final Publisher pre-invocation guards and the
detached HTTP/chat callback test. Native form import still failed: unlike the
export, native CodeModule.Lines does not include the final export line terminator.
A separate readback diagnostic established that exact difference. Pure fake
modules had not reproduced this native representation, so their passing prefix
tests alone were insufficient acceptance.

The corrected `open-qualification-form-readback` build passed without
warnings/errors. Candidate MVID `82942b5d-0369-4f4e-9247-25841e205b90`, SHA-256
`9D6F8E1F3003BD6BBE6EAD4D55CF6F6E539FA37DAB06C68F130247FBB5B2AEDE`.
`native-form-prefix-readback-focused.trx` records **10 passed, 0 failed,
0 skipped**, limited to the changed project import adapter and its exact code,
intentional blank-line, concurrent-change and identity guards. No full-suite or
coverage result is claimed for these follow-ups; none is installed.

Local native evidence is below
`C:/Users/jvc/Documents/Codex/2026-09-25/bo/artifacts/open-qualification/`:

- `production-git-excel-logical/6bebea026d8242a09bb4a162f59f6617/qualification.json`:
  3eda62, Excel PID 754072; three unedited captures compare equal, native controls
  restored, but final exact FRM equality fails on the extra code line.
- `production-git-excel-form-import/f501e90f368543f7bac3745032192637/qualification.json`
  and `production-git-excel-form-code-diagnostic/85eb36e9248c4e4a96d276ffaa8ed05c/qualification.json`:
  a37d53, PIDs 753740/755696; retained prefix failures and precise native
  CodeModule.Lines observations.
- `production-git-excel-form-readback/bb3705fd489549aba695e580d2d45f2e/qualification.json`:
  82942b, PID 753672; repeated captures, nonempty corrupted-FRX refusal, guarded
  production Apply, exact FRM/logical FRX readback, native controls and helper
  save/reopen all pass.
- `production-git-excel-coordinator/741bf1f9c4ef4a2c8a2a95bafc6c5061/qualification.json`:
  the same candidate, PID 756468; production local-Git checkpoint restore,
  before-state backup, explicit rollback, final explicit restore, cleared
  recovery marker and owning managed STA continuity pass. Original FRX bytes
  survive checkpoint commit/read exactly. Helper save/reopen retains matching
  snapshot and both native captions. No remote access or publication occurred.

All these owned Excel processes used Office 16.0 build 20326 and exited normally
through Close/Quit with exit code 0. Only disposable forms/modules were changed;
no macro ran and no trust policy changed. The candidate assembly ran on an
external owning STA against native COM. This does not qualify the installed
bridge, embedded UI, current-candidate GitHub transfer, unsupported control
layouts or historical export failures on another machine. Q-027 remains open
for those separate requirements.

Access/Publisher save adapters and stream diagnostics are compiled and covered
by the stated contract/protocol tests, not native adapter/provider qualification.
Access has no invented document Saved flag; first SaveAs remains unavailable.
The diagnostics retain counters and filtered terminal metadata, excluding
provider text/prompts/tool arguments, and add no automatic retry. Q-012/Q-028
remain open until their native/historical acceptance requirements are met.

### Workspace geometry and conservative resource guards

The `open-qualification-office-ready` candidate, MVID
`c7bda9b7-236f-4785-9b87-cab0f8bcbef0`, passed its isolated Debug build without
warnings/errors. SHA-256
`EAAAD53F7EC1F54062E604153E506773185218470441C701A5CDC911EFD8F8FD`.
`form-class-identity-preflight-focused.trx` records **24 passed, 0 failed,
0 skipped** for known-form CLSID gating, strict UTF-16 CFB names and snapshot
comparison. Access/Publisher persistence fixtures compiled but were not run.

`workspace-hidden-red.trx` reproduces stale hidden Monaco bounds with the native
designer selected in a disposable detached Windows MDI fixture.
`workspace-hidden-green.trx` records **9 passed, 0 failed, 0 skipped** after the
geometry correction, including child ordering, focus and Object Browser behavior.
This is not proof of the historical SOLIDWORKS 2019 failure mechanism or native
acceptance in either required SOLIDWORKS version.

An independent review found that opaque resources could start before an OLE blob,
or use multiline declarations missed by the old comparison extraction.
`opaque-resource-red.trx` demonstrates the false equality before correction.
The comparison now falls back to raw FRX bytes for any opaque resource reference.
`resource-workspace-guards-green.trx` records **34 passed, 0 failed, 0 skipped**
across snapshot/form/revision guards and workspace geometry. The final isolated
`open-qualification-resource-guards` build passed without warnings/errors:
MVID `99ca80c4-2bbc-4ddd-b4d0-1d3957610ba4`, SHA-256
`C74ABC336867F78160D19B11A8CA67E4C0953EA60C8ED82247037A25CAA65EFD`.
It also compiles stronger adapter fixtures that change source immediately before
Save and retain a created form expectation even if subsequent configuration fails.
The subsequent Publisher fixture correction calls `Application.Quit` directly,
exactly once, only after fresh verification of the saved VBA project and saved
publication, sole document, expected path, retained PID and IUnknown identity.
It does not call `Document.Close`, which could create a replacement blank
publication. A failed guard or Quit retains the owned instance and COM references
without retry or forced termination. At this resource-guards checkpoint, native
Access/Publisher persistence and Publisher normal-exit acceptance were
**NOT_RUN** for the corrected fixture. The later native attempts and failures
are recorded separately above. The resource-guards candidate was not installed;
the later installed candidate is identified in the current checkpoint.

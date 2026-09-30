# Recorded validation

## Current installed v5: complete managed acceptance (2026-10-01)

Initial product and test source `2e751618aa8d5be08ca0fa84d6ceff4d0b1312dc`
implements the stable Access application/database-path/mapped-project identity
guard while retaining owner-PID, VBE, selection, mode/protection, source and
metadata checks. The isolated build completed without warnings or errors.
Candidate `f9a36c85-1d9c-4a53-8645-06c99617b12c`, SHA-256
`76524AC884247BEBC03B0A73E36AA1C96414CB650D98E9E74F11BCA5A4D068D1`, is now
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
the requested complete line/branch target, and any subsequent product correction,
including the pending Monaco status correction, requires a new complete gate.

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

### Current v5: partial SOLIDWORKS 2025 native qualification

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
class. Its correction and complete embedded assistant/UI acceptance are pending.
The host remains alive at this checkpoint: final cleanup and normal exit are
**NOT_RUN**, and SOLIDWORKS 2019's separate abnormal termination is unresolved.
Q-014 therefore remains open.

### Current v5: terminal Access adapter campaign

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
Q-012 remains partial; current-v5 Publisher acceptance is not established.

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

# Test coverage and qualifications

[Documentation](README.md)

VBAi is tested through managed regressions, editor/UI checks, provider checks and
operation-specific qualifications in real host applications. This page summarizes
all Q001-Q030 findings, including earlier qualifications, rather than only the
latest campaigns.

## Latest managed check

### Shared workspace resource loading in COM hosts

The native follow-up on product MVID `9da28f76-35c9-4be3-91d8-1bbe92e4cfee`
exposed a chat startup failure: compiled resources requested
`System.Resources.Extensions` version 4, while the deployment contained version 8.
Office does not use the add-in's application binding redirects. The error dialog
delayed VBE startup, and the first Excel scenario subsequently failed its normal
exit check. The four-case Excel group recorded three failures; its remaining case
was interrupted. The following thirty-four Office/SOLIDWORKS cases were not run.
Owned processes were released, including forced cleanup; this is not native
qualification. Registration restoration was refused after the registration changed,
and the current checkout's registration was preserved.

The workspace now resolves only the known legacy reader identity while constructing
its controls or applying localized resources. The resolution handler is removed in
`finally`; unrelated assembly requests retain the host's normal resolution path.
A standalone process without binding redirects reproduced the failure in all
thirteen UI languages on the preceding assembly and passed all thirteen on the
corrected assembly. No controls were shown during that diagnostic.

The corrected snapshot has product MVID `c9e422cb-37df-414a-aabb-26202d45ce38`.
Its build completed without errors and adds five reader-identity regression cases.
The complete **7,965-case** managed bank passed with **7,964 successes, no failures
and one conditional provider skip**. Its 2,316 frozen inputs remained unchanged;
all preceding methods and rows were retained. The original worker and testhost
exited normally with code 0, the complete gate was accepted and its temporary
task was removed. A preceding sandbox attempt could not connect VSTest to its
testhost and executed no regression tests.

The consolidated follow-up applies the same explicit reader scope to the chat,
Git, LLM settings and About constructors. Opening those forms already succeeded
on the preceding candidate in the standalone diagnostic; their compiled resources
still name the legacy reader, so their loading now has an explicit scope as well.
Product MVID `a45b89d4-c048-4024-b09c-67b10075e3f0` compiled without errors and
passed **65 standalone resource activations** across the five surfaces and thirteen
languages, without application redirects or displayed windows. Its complete
**7,971-case** managed bank passed with **7,970 successes, no failures and one
conditional provider skip**. All 2,318 frozen inputs remained unchanged, with
all preceding rows and the six added cases retained. Original worker and testhost
exit codes were 0; the strict gate was accepted and its task removed. No new
line/branch coverage percentage is claimed.

The **38-case native follow-up** on that exact candidate is terminal:
**four passed, five failed, one was inconclusive and twenty-eight were not run**.
Excel source persistence with an independent reopen and the catalogue revision
check passed. The registered test explorer passed in Word and Excel. The Options
batch failed during target validation; the three chat cases failed discovery,
and PowerPoint's explorer case failed presentation closure after completing its
functional checks. SOLIDWORKS was not reached. Owned processes and the temporary
task were released; forced PowerPoint cleanup does not qualify normal shutdown.
Changed registration and provider preferences were preserved rather than
overwritten. This campaign does not establish native acceptance.

The follow-up aligns the chat observer with the actual shared Designer control
identities and its independent Stop button, including a hidden Stop whose native
handle is created on the first turn. Hidden observations retain identity checks;
clicks still require visible, enabled, owned controls. Options inspection now
checks native availability as well as UIA state, and batch refusals identify the
phase and target. On 2026-10-13, product MVID
`3fa985de-a971-47e7-910b-735f7a337a0f` compiled without warnings or errors.
Its isolated combined **406-case targeted bank passed without
failures or skips**, including the Options regressions and approval catalogues
constructed in all thirteen UI languages. The four observer files were integrated
without modifying the product chat. These targeted results do not establish
Office/SOLIDWORKS qualification.

The consolidated candidate, product MVID
`7c8f4fb0-4466-4e40-86bc-ba9af81ea446`, also changes the disposable PowerPoint
fixture's shutdown sequence: a final inventory must identify exactly the owned
presentation and its design-mode VBA project before the saved flag is staged and
one application Quit is requested. It does not retry an uncertain Close or Quit.
Its complete **8,051-case** managed gate passed with **8,050 successes, no failures
and one conditional Ollama UI skip** on 2026-10-13. All 2,256 frozen inputs and
all preceding rows were retained, including eighty added cases. Worker and
testhost exited normally with code 0; the strict gate was accepted and the owned
task removed. The fixture changes are integrated, but its PowerPoint shutdown
and the complete Office/SOLIDWORKS bank still require native qualification.

The subsequent native bank on that same product is terminal: **six passed,
three failed and twenty-nine were not run**. All four Excel editor/persistence/
Options/signature-inspection cases passed, with their scope released and temporary
registration restored. The Word and Excel test-explorer cases also passed. The
three assistant cases reached the real embedded chat but failed the streaming
observation. The rendered numbered reply used separate short text controls, while
the observer required a longer reply in one control. PowerPoint cleanup also
refused its inventory because the fixture used an incorrect design-mode value.
Both are test-harness defects being corrected; the original failed verdicts remain
unchanged. The remaining Office and SOLIDWORKS cases were not reached.

The owned PowerPoint instance and private provider processes were released after
the provider proxy reported no active requests. Forced cleanup does not establish
normal PowerPoint shutdown. Registration and provider preferences changed during
the bank, so restoration preserved those concurrent values. The terminated
coordinator's temporary task was removed. No broader native acceptance or new
line/branch coverage percentage is claimed.

The two harness corrections then passed a single **43-case targeted bank**, with
no failures or skips. This bank used a separately compiled test assembly and the
unchanged product DLL from the accepted complete managed candidate. It verifies
split-list observation, exclusion of prompt echoes, design/run/break mode guards
and the guarded shutdown sequence. It does not yet establish native chat recovery
or normal PowerPoint exit.

### Native Office recheck after harness corrections

The fresh ten-case bank on product MVID
`7c8f4fb0-4466-4e40-86bc-ba9af81ea446` ended with **one pass, two failures and
seven unexecuted cases**. Word verified visible streaming, one Stop, cancellation
acknowledgment, a completed next reply, a marker read from the native VBA module,
unchanged source/references and fixture shutdown. Excel and PowerPoint verified
streaming, Stop and recovery, but the model did not call a tool for the marker
request and instead contested the supplied module or project. These failures
remain failures; the marker oracle was not relaxed.

PowerPoint also exceeded the original normal-exit deadline. Its later absence
does not reconstruct a normal-exit receipt. A real Word capture shows overlapping
transcript cards; functional workflow acceptance does not qualify this visual
behavior. All 2,479 frozen inputs were unchanged. Original controller/provider
exits were observed, the owned task was removed, and the retained private Ollama
processes were released after requests settled and all hosts were absent.
Registration restoration refused a changed registration, and preference recovery
preserved changed values; the writer is unconfirmed. This bank is not complete
Office/SOLIDWORKS acceptance. Its retained evidence is under
`artifacts/native-followup-20261013v/office-i10-startupfix2-20261013v/`.

### Detached transcript performance recheck

On 2026-10-13, the frozen product MVID
`7c8f4fb0-4466-4e40-86bc-ba9af81ea446` passed all **27 targeted performance
workloads**, with no failures or skips. The activity, text-sizing and row-layout
sources used by those workloads match the current checkout. This bank used actual
RichEdit controls and detached WPF/WinForms hosting without opening Office or
SOLIDWORKS or contacting a provider.

For 100 completed siblings and one streamed activity over 32 frames, expanded
groups measured 274.392 ms total and 30.993 ms p95 at 340 pixels, and 101.851 ms
total and 14.618 ms p95 at 780 pixels. Controls and selections remained intact.
These machine-specific observations supersede the older candidate's workload
cost for this tested scope; they do not prove smooth real-host rendering or
physical-DPI acceptance. The retained TRX and measured rows are in
`artifacts/chat-performance-recheck-20261013/`.

### Shared workspace integration follow-up

An isolated source snapshot with product MVID
`6f7b5ebb-af33-4d4c-90fa-d49e8d252a92` completed **7,939 Unit cases**:
**7,927 passed, eleven failed and one conditional provider skip**. All 2,305
recorded inputs remained unchanged. The original worker exited with code 1 and
the original testhost with code 0; its temporary task was removed. Concurrent
checkout edits were preserved. The complete gate is not accepted.

The follow-up corrects the shared library's SDK resource-build configuration,
the Arabic resource's missing framework assembly alias and direction reset when
switching languages. It adds the eight missing workspace usage hints and one
manually translated hint in all thirteen UI languages. Markdown regressions now
inspect the controls actually displayed, including captured reference targets,
formatted text, table cells, code-copy actions and reported callback failures.

The next snapshot, product MVID `c30d862b-4eb0-40ee-ad47-c136126d6a75`,
completed **300 affected cases**: **298 passed, one failed and one conditional
provider skip**. Its inputs remained unchanged and its original testhost exited
normally with code 0. The remaining failure exposed an enabled delete action
without a selected conversation; the adapter now disables it initially and lets
the existing session-state checks enable it later.

Product MVID `9da28f76-35c9-4be3-91d8-1bbe92e4cfee` contains that final fix
and has a prepared complete **7,960-case** Unit bank. Compilation succeeded with
zero errors; NuGet reported an unavailable vulnerability-feed lookup. The full
bank completed with **7,959 passed, no failures and one conditional Ollama UI
skip**. All 2,306 frozen inputs remained unchanged, all preceding methods and
their rows were retained, and the twenty-one added cases were present. The
original worker and testhost both exited normally with code 0. The strict
complete managed gate is accepted for this snapshot. Its temporary task was
removed; no source was restored and no process was stopped during cleanup.
This check does not establish Office/SOLIDWORKS or release qualification,
and no new line/branch coverage percentage is claimed.

### Shared conversation project context

The current follow-up build has product MVID
`535e5b57-be57-4d76-bd26-6e0f31f091d8`. The chat and synthetic provider
fixture now share the same project-context formatter. It preserves exact saved
document paths, Unicode selectors and unsaved-project names; an unbound request
does not inherit a project label. The synthetic fixture also binds its catalog
to that project. Native dispatch remains prohibited and its marker is still
withheld until the exact module-read result.

All **169 targeted regressions passed**, including seven new context contract
cases and the Codex instruction-refresh checks. Compilation completed with zero
warnings and errors. The original worker exited with code 0 and VSTest recorded
testhost PID 38440 exiting with code 0. Both recorded actor/host PIDs were absent
at scope release; 1,698 frozen inputs were checked and the owned task removed.
This is a targeted check, not a new full managed or native qualification.

The complete **7,930-case** bank on this compiled candidate subsequently recorded
**7,927 passed, two failed and one conditional provider skip**. Both failures
occurred while tearing down negative Object Browser UIA scenarios. The worker
exited with code 1 and the original testhost with code 0. Seventeen source/test
inputs changed during the run; the maintainer confirmed another session was
editing the shared checkout. Cleanup preserved those changes, verified the
original binaries and removed the owned task. This result does **not** qualify
the changed checkout; the complete managed gate remains rejected.

A separate **26-case UIA regression bank passed** against that same immutable
product DLL after correcting fixture teardown. Negative tests deliberately
report an unavailable native HWND; teardown now exposes the actual owned root
identity only while disconnecting its providers, then restores the simulated
identity and clears the destroyed HWND. Cached-element invalidation, original
owner-thread exit and the five-second deadline remain asserted. The isolated
project rebuilds only these fixtures/tests, and Object Browser fixture windows
are shown without activation. This targeted result does not replace the failed
complete bank or establish Office/SOLIDWORKS behavior. No new line/branch
coverage percentage is claimed.

The passive synthetic HTTP capture also passed **65 offline contract cases**,
including its explicit host-context scope, bounded evidence and factory
restoration. The six synthetic host-context scenarios now use that capture to
record the actual declared tools and request context. These contract checks do
not establish successful live model tool use. One sandbox runner attempt was
aborted before executing any test because VSTest could not connect to its
testhost; the successful contract run used the regular Windows runner.

One subsequent six-context wire diagnostic recorded **four passed and two
failed**. Every captured chat request contained the eight core tool schemas,
including `read_module`, and its exact project selector; the captures were not
truncated. Access requested `monaco_open` despite the scenario requiring the
module read, and Outlook returned the preceding interruption text without a tool
call. Those failures remain recorded. The owned Windows Ollama backend and its
runners exited, and diagnostic inputs remained unchanged. This evidence narrows
the failure to the observed provider replies for those requests; it does not
prove that all provider behavior or native host operation is correct.

The same nine preloaded Ollama scenarios then recorded **five passed and four
failed**, with no skips. Several marker requests were answered without the
required read, including a claim that the available read tool was unavailable.
Those failures remain recorded. A preceding instructions-only candidate had
recorded four passed and five failed before the missing project-context prefix
in the fixture was identified. Neither result establishes native-host behavior
or proves the cause of the model's refusal. Model, sampling, reasoning and
production response deadlines were unchanged. No Office/SOLIDWORKS host was
opened; the private backend and recorded runners exited and its task was removed.

### UI Automation fixture lifecycle order

The follow-up build has product MVID
`34d1b949-ef95-47f3-aa50-5d6ac585eae1`. The test fixture destroys its
owned window before releasing the window's provider map and disconnecting its
providers, while preventing provider republication. The five-second cleanup
deadline, cached-element invalidation assertions and owner-thread exit checks
are retained.

All **159 targeted cases passed**: the preceding 147 regressions and twelve UIA
fixture checks. The formerly failing oversized scenario completed in 3.008
seconds, including its assertions and cleanup. The original worker exited with
code 0; VSTest recorded testhost PID 92604 exiting with code 0, without a forced
termination. The recorded actor and testhost PIDs were absent at scope release;
the owned task was removed after verifying 1,688 frozen inputs. This check did
not retain an original testhost handle, rerun the complete managed bank, measure
line/branch coverage or establish native-host/release qualification.
Compilation completed with zero warnings and errors.

### Procedure submission regression follow-up

The follow-up build has product MVID
`6b8977f7-5207-448a-a379-2a65bf4bbb85`. Worker cancellation now preserves
the original exception and cancellation token while retaining the canceled task
state. The UIA regression observes native HWND focus messages for root providers
and provider focus callbacks for child fragments; neither path is omitted.

All **147 named regression cases** passed, including the six preceding failures
and two additional cancellation-token cases. The original worker and testhost
exited normally with code 0, the observer retained the original host handle, and
the owned task was removed after verifying the recorded generations were absent.
Compilation completed with zero warnings and errors.

The complete **7,923-case** managed bank finished with **7,921 passed, one failed
and one conditional provider skip**. The failure is the oversized UIA fixture's
five-second teardown deadline: 3,687 provider disconnections had completed at the
deadline. The original worker exited with code 1 and the testhost with code 0;
1,679 frozen inputs were verified and eleven recorded generations had exited
before the task was removed. The complete managed gate is **not accepted**.
No new line or branch coverage percentage is claimed.

A separate nine-case synthetic provider diagnostic used Windows Ollama 0.34.4,
the same Qwen3.5 9B model manifest, and a 32,768-token context. It recorded **five
passed and four failed**, with no native Office or SOLIDWORKS host opened. Three
requests expired during initial model loading; PowerPoint's synthetic marker
scenario returned text without the required tool call. Cancellation/recovery,
the synthetic tool round trip, the displayed-chat scenario and the Publisher and
Outlook synthetic markers passed. GPU allocation was observed, but this does not
establish native-host qualification or a comparative performance benchmark.
The private server and its recorded runners exited, the task was removed and
1,685 frozen inputs were verified. Shared services and personal settings were
unchanged. This diagnostic remains **not accepted** as a complete provider bank.

The same nine scenarios were then run on the UIA lifecycle candidate with one
empty-prompt model preload, keeping the model, sampling settings, reasoning
setting and production response deadline unchanged. Preloading completed in
74.883 seconds; the scenario bank completed in 1 minute 38 seconds with **seven
passed and two failed**, and no request timeout. Access refused the module read
based on the document extension; Publisher claimed a read without a tool call.
The assertions rejecting both answers were retained. The private backend and
its recorded runners exited; 1,695 inputs were verified and its task removed.
This is a synthetic provider diagnostic, not native Office qualification.

### Preceding procedure submission candidate

The current isolated build has product MVID
`d695a454-2291-46e4-8614-559cd1a2ec34`. It captures failures inside the
Immediate-window worker, checks project identity, source revision and design mode
again on the owning VBE thread before submission, and discovers a unique text
provider by its capabilities when the pane does not expose a Document control.
Ambiguous providers are refused. These changes have not yet been qualified in a
native host.

The complete managed bank recorded **7,914 passed, six failed and one conditional
provider skip**, across **7,921 cases**, including 27 additional scenarios.
The failures exposed two cancellation-state regressions and four inaccurate
focus observations for root UIA providers. Every preceding method was retained.
The original worker exited normally with code 1 and the testhost with code 0;
1,676 frozen inputs were unchanged, all eleven recorded generations exited and
the owned task was removed. The complete managed gate is **not accepted**.
Compilation completed with zero errors and five NU1900 vulnerability-feed warnings.

### Current Word path-read candidate

The 2026-10-12 working-tree candidate has product MVID
`c5d2f1d4-8d05-4249-876f-b98ab63ed543`. Application lookup reads the host
classification and current process ID once per lookup rather than once per native
window. Word revision-path reads query only the document directory and full name;
save-state reads still use the separate document-state contract. Fresh final
application ownership and document/project identity checks remain in place.

The complete managed bank recorded **7,888 passed, one failed and one conditional
provider skip**, across **7,890 cases**. All **12 additional cases** passed and
every preceding method was retained. All 112 named ownership, timing and Monaco
regressions passed. The failing navigation scenario exceeded its original
five-second fixture-owner shutdown deadline; its retained error identifies UIA
provider teardown, not a Word native operation. The precise teardown stage is
not yet established.

The original testhost exited normally with code 0 and the worker with code 1.
All seven recorded generations exited; 1,668 frozen inputs were unchanged and
the owned task was removed. The complete managed gate is **not accepted**.
The new 38-case native bank is prepared but has not been dispatched, so no
performance improvement or native qualification is claimed for this candidate.
No new line or branch coverage measurement is claimed.

The follow-up build `51fb4924-3bea-4644-b0a2-4b6e694af920` separates the
five navigation fault cases and records the teardown stage and provider counts
observed at the original deadline. It retains that deadline and waits for the
original owner thread to finish before another fixture starts. Its complete
managed bank recorded **7,893 passed, zero failed and one conditional provider
skip**, across **7,894 cases**. Every preceding method was retained, and all
117 named regressions passed, including all five navigation fault cases.
The original worker and testhost exited normally with code 0. All seven recorded
generations exited; 1,671 frozen inputs were unchanged and the owned task was
removed. Compilation had zero errors and ten NU1900 vulnerability-feed warnings.
The complete managed gate is accepted. This successful execution does not
establish the cause or correction of the preceding intermittent timeout.
Its native bank verified the four Excel saved-source, unsigned-signature and
VBE-options cases, with five normal host exits and verified COM restoration.
The registered Word test-explorer case also passed: the complete four-result
batch, individual run and three procedure-coverage copies completed without an
uncertain or stale report. Each copy preserved the original revision, counters
and open-document catalogue. Word exited normally with code 0 after one Quit.
This proves those operation-specific cases. The subsequent Word chat scenario
failed before visible streaming: the upstream model returned no response bytes
before the observation deadline. Excel chat streamed, acknowledged interruption
and completed the next send, but failed the native-marker read. Its provider
response contained no tool-call deltas and its visible answer fabricated module
content instead of reading the undisclosed marker. The Office group completed
all twelve scenarios: the registered test-explorer cases passed on Word, Excel,
PowerPoint, Access, Publisher and Outlook, while all six chat cases failed.
The five chat failures after Word concern the undisclosed native marker.
The original failure verdicts are retained. The Office group exited normally
with code 1, released its physical host scope and restored its COM registration.
The complete native bank finished with **21 passed, 13 failed and four not run**
across 38 scenarios. The saved-fixture preparation and close cases passed on
SOLIDWORKS 2019 and 2025. Other SOLIDWORKS cases failed file-preservation or
execution checks; unchanged decompressed module source does not satisfy the
separate original-file byte-preservation check. A procedure worker stopped in
the debugger before the native request completed. Its failed result was retained.

All six group scopes were released, temporary COM registrations and provider
preferences were restored and verified, and no native host remained. Some failed
host generations required forced recovery; this is not normal-exit qualification.
Review verified 1,902 frozen inputs and 28 child plans before source release.
Complete native qualification and release acceptance are **not accepted** for
this build.

### Preceding Word snapshot ownership candidate

The 2026-10-12 working-tree candidate has product MVID
`510801de-5ca9-4d15-9db9-28bd054cca9b`. Word host-path reads now balance
application, document and project getter acquisitions, retaining the fresh PID
and project-identity checks. Refused ROT/NativeOM lookups also release their
acquisitions. Snapshot diagnostics separately observe host-path, module-source
and reference reads; they remain disabled by default and do not cache revisions.

Its complete managed bank recorded **7,877 passed, zero failed and one conditional
provider skip**, across **7,878 cases**. All **39 additional scenarios** passed,
and every preceding case was retained. Named-row review verified all 100 diagnostic
and ownership regressions, including the earlier Monaco policy failure. Both
original processes exited normally with code 0, without forced termination,
unhandled owner posts or an AppDomain-unload failure. All seven recorded
generations exited; 1,664 frozen inputs were unchanged and the owned task was
removed. Compilation had zero errors and ten NU1900 vulnerability-feed warnings.

The complete managed gate is accepted. Its 38-case native bank passed the four
Excel scenarios, including five normally exited host generations and verified
COM restoration. Word's first registered batch again exceeded its unchanged
observation deadline; its original pending verdict is retained.

The detailed Word observations localize the snapshot cost to host-path acquisition:
144 host-path reads averaged 377.55 ms, compared with 1.44 ms for module reads and
0.69 ms for reference reads. Complete project snapshots averaged 381.36 ms;
14 native dispatches averaged 1.43 ms. Nested timings overlap and are not additive.
These observations identify the expensive path; they do not yet establish which
application/document getter inside that path is responsible.

Recovery normally closed the exact disposable Word generation with code 0,
without force or macro replay. COM and all temporary provider preferences were
restored in the original interactive context. Review verified 1,893 frozen inputs
and 28 child plans before source release. Later Office/SOLIDWORKS cases were not
run; complete native and release acceptance remain pending. No new line or branch
coverage measurement is claimed.

### Preceding Monaco reply-publication candidate

The 2026-10-12 working-tree candidate, based on source revision
`213231485c3bba0251357357a09392b750b8c24f`, has product MVID
`9f9b014f-2086-41b2-b482-4cc0e3448e35`. Its STA tool entry point keeps
COM/UI work and cleanup inside the owner envelope, while publishing the finished
reply independently of the caller's borrowed synchronization context.

The complete managed bank recorded **7,838 passed, zero failed and one conditional
provider skip**, across **7,839 cases**. Every preceding method was retained.
All **24 additional Monaco scenarios** passed, covering ready renderer replies
and renderer failures under holding, plain and absent synchronization contexts.
The preceding failing Monaco policy scenario and all 37 instrumentation scenarios
also passed. This verifies those managed regressions; it does not prove the sole
cause of the earlier intermittent timeout or native renderer behavior.

The original worker and testhost both exited normally with code 0, without forced
termination, an unhandled owner post or AppDomain-unload failure. Review verified
1,658 frozen inputs; all seven recorded generations exited and the owned task
was removed. The complete managed gate is accepted. Compilation had zero errors
and ten NU1900 vulnerability-feed warnings.

Its 38-scenario native bank accepted the four Excel saved-source, unsigned-signature
and VBE-options scenarios. All five original Excel processes exited normally,
and COM restoration was verified. The first Word scenario failed during startup,
before its explicit qualification operations: Word reopened the prior bank's owned
`Disposable.docm` alongside the new inert seed, violating the one-document guard.
Read-only NativeOM inspection identified both exact disposable paths and their
saved state. The mechanism that reopened the older document remains unproven.

Recovery issued one normal Quit to that exact Word generation, observed exit code
0 and verified both files were unchanged. COM registration and all six temporary
provider preferences, including the reasoning preference, were restored in the
original interactive Windows context. Review verified 1,884 frozen inputs and
28 child plans; the owned task was removed and no host remained. The failed Word
verdict is retained. Later Office and SOLIDWORKS scenarios were not run;
complete Office/SOLIDWORKS and release acceptance remain pending.

A fresh independent native bank on the same candidate then passed the Excel group
and Word's one-document startup guard. Word's first registered batch still exceeded
its unchanged sixty-second observation deadline. Its retained report contained
three correlated results; automatic retry/reset and normal fixture disposal were
refused because the original run remained pending.

Opt-in observations recorded 142 project snapshots averaging 426.64 ms (maximum
531.87 ms), and 14 native dispatches averaging 1.03 ms (maximum 3.94 ms).
Nested spans overlap: these figures must not be added, and they do not explain all
of each test's roughly fifteen-second reported duration. They identify repeated
snapshot acquisition as a measured cost requiring further localization.
Recovery issued one normal Quit to the exact disposable Word instance and observed
exit code 0, without forced termination or macro replay. The original native worker
exited with code 1. COM and temporary provider preferences were restored; 1,887
inputs and 28 child plans were unchanged before source release. Later Office and
SOLIDWORKS cases remain unexecuted, and the complete native gate remains unaccepted.
No new line or branch coverage measurement is claimed.

### Preceding native-timing diagnostic candidate

The 2026-10-12 candidate with product MVID
`2115ea40-3088-484a-b8fc-f56cb725156a` adds opt-in, bounded stage observations
for native test execution and live project snapshots. Diagnostic failures do not
replace the original native result or cause a retry. Logs exclude source text,
project paths, method arguments and exception messages; their opaque recorder
identity distinguishes services sharing a process.

Its complete bank recorded **7,813 passed, one failed and one conditional provider
skip**, across **7,815 cases**. All **37 additional instrumentation scenarios
passed** and every preceding method was retained. The failure was
`MonacoPolicyAndDefaultSessionFactoriesRejectEveryMissingOrUnapprovedContext`:
the second ready-state `monaco_open` exceeded the existing 15-second owner-task
deadline. Its cause remains unproven; the complete managed gate is not accepted.
The original testhost exited normally with code 0; the runner exited with code 1
for the failed assertion. All seven recorded generations exited, the owned task
was removed, and 1,654 frozen inputs were unchanged before source release.

Compilation completed with no errors and ten NU1900 warnings because NuGet's
vulnerability feed was inaccessible. The retained Word standard-module sources
were also measured without COM or a host: 30 warm discovery samples had a p95
of 0.9147 ms. This does not measure live COM acquisition or explain the native
Word delay. No native bank was dispatched for this failed managed candidate;
real Office and SOLIDWORKS acceptance remains pending.

### Preceding managed candidates

On 2026-10-12, the resumed full bank for product MVID
`381d3be3-4fe0-4900-b3bd-e4d8ea1888ce` recorded **7,713 passed, five failed and
one conditional provider skip** across 7,719 cases. All preceding cases were
retained, including the additional shutdown, asynchronous-event and connection
failure regressions. Frozen-input review verified 1,596 inputs and both exact
assemblies. The complete managed gate remains **not accepted**.

The failures were a modal Git branch-switch deadline, an empty-toolbox UIA
observation, two native-diagnostic UIA observations and an Access queued-save
confirmation assertion. The two diagnostic readers encountered a WinForms
arithmetic-overflow error message rather than the fixture's expected message.
Read-only Git metadata showed the branch and baseline update before its timeout;
that evidence does not establish subsequent refresh or terminal completion.

The original worker and test host both exited with code 1. The test-host exit
was observed through a retained generation-specific process handle and matched
the VSTest exit callback. VSTest did not force this shutdown, but its retained
stderr reports an unhandled late owner-dispatcher post. Normal successful host
closure is therefore unqualified. All recorded campaign generations are absent,
the exact owned task was removed and the source freeze was released. Failed Git
scratch evidence was preserved. Neither native nor release acceptance is claimed.

The subsequent candidate (product MVID
`99b9a6a7-e197-41ac-a7b2-0f3fe2af144e`) recorded **7,766 passed, two failed and
one conditional provider skip**, across 7,769 cases. All preceding cases and all
50 additional scenarios were retained; the additional scenarios passed. They
cover binary Git batch reads, low-32-bit synthetic UIA handling, recreated-HWND
fixture readback, deterministic Access deadlines and draining admitted Git work
after a failed fixture deadline. Both original processes exited with code 1,
without forced termination or an unhandled owner post. Review verified 1,616
frozen inputs; the owned task and recorded generations were released. The two
remaining failures were a five-second editor-event drain and a test expectation
corrupted by a UTF-8/Windows-1252 transfer. This complete gate was not accepted.

The editor-event candidate (product MVID
`3df32133-7436-497b-843d-fe50cfc3a7d0`) restores the affected test file's UTF-8
encoding and pumps queued editor events on their owning context before deciding
that it is quiescent. The original five-second budget remains unchanged. Its
complete bank recorded **7,772 passed, zero failed and one conditional provider
skip**, across 7,773 cases, including four additional event-lifetime scenarios.
All preceding cases were retained and 1,620 frozen inputs were unchanged. The
original runner exited with code 0 and the original test host with code 1,
without forced termination or unhandled owner-post stderr. The host trace records
`CannotUnloadAppDomainException` during MSTest's final AppDomain unload. The
successful assertions therefore do not establish the complete lifecycle gate.
All seven recorded generations exited; the owned task was removed and the source
freeze released. Native and release acceptance remain unavailable.

The provider-cleanup candidate (product MVID
`667167c9-d44a-4feb-8f0a-ab52810e580c`) adds explicit disconnection of the owned
synthetic UI Automation providers and their HWND event map, prevents republishing
the same provider during release, and observes the original provider thread's
exit. The fixture's non-activation option also applies to its WinForms window.
Its complete bank recorded **7,775 passed, zero failed and one conditional
provider skip**, across 7,776 cases. The three provider-lifetime scenarios passed,
including cached-provider invalidation and original owner-thread exit. Review
verified all 1,623 frozen inputs. The runner exited with code 0; the retained
testhost again exited with code 1 after MSTest's AppDomain-unload failure. Provider
cleanup is therefore verified locally, while its sole responsibility for the
global unload failure is contradicted. All seven recorded generations exited,
the owned task was removed and the source freeze was released. The complete
managed lifecycle gate remains not accepted.

Process isolation exposed additional test-dispatch problems. Ordinary async
tests must use MTA independently of explicit STA or pumped WinForms fixtures.
Conflicting ordinary/UI method attributes were removed without removing their
scenarios. The subsequent candidate (product MVID
`7a75a56d-3dc1-49bb-bee3-274fc3708c31`) recorded **2,173 passed and zero failed**
before an ordinary synthetic HTTP test stopped making progress. Only its exact
owned testhost was force-stopped; these interrupted results are not a complete
gate. All seven recorded generations exited, 1,639 frozen inputs were verified,
and the owned task and source freeze were released.

The asynchronous-isolation candidate (product MVID
`bf5a7f21-8901-46a1-b0f7-0ce8cfbf8747`, working tree at base commit
`213231485c3bba0251357357a09392b750b8c24f`) isolates ordinary asynchronous unit
tests in 23 classes on separate MTA threads, retaining each explicitly selected
UI dispatcher. Two regressions check that an unpumped fixture context cannot
contaminate the next invocation. The full **7,778-case** bank retains all preceding
scenarios. Its isolated solution build completed with no warnings or errors.
The bank recorded **7,583 passed, four failed and one conditional provider skip**
before a scenario-level asynchronous session read stopped making progress.
The four failures identify a direct STA-marker invocation, a WPF test missing
STA metadata, and an unowned continuation test borrowing residual UI context.
The exact original testhost was force-stopped after eight minutes without
progress. All seven recorded generations exited and 1,642 inputs were verified
before source release. These interrupted results are not a complete gate.

The STA-dispatch candidate (product MVID
`3f7030d9-2fa3-47b1-9e04-817b73f61344`) dispatches the legacy STA fallback on
an actual STA thread, marks the WPF scenario explicitly, and isolates the
continuation and session-reader fixtures. The async declaration audit now also
includes `Scenarios/` and fully qualified task types. No test body, case or
existing deadline was removed. Its isolated solution build completed with no
warnings or errors. Its complete bank recorded **7,777 passed, zero failed and
one conditional provider skip**, across 7,778 cases. All cases were retained,
including the previously blocked session read. The original runner exited with
code 0; the retained testhost exited with code 1 without forced termination,
AppDomain-unload failure or unhandled owner-post stderr. A cancellation test
invoked the progress window's real process-exit action instead of its owned
result observer. All seven recorded generations exited and 1,645 frozen inputs
were verified before source release. The complete lifecycle gate was not accepted.

The current candidate (product MVID
`22aaf833-eb26-4d59-b6bd-aa3c3b9c7b83`) isolates that updater cancellation result
and asserts that both cancellation and disposal preserve the test process exit
code. The external updater's production exit behavior remains intact. Its isolated
solution build completed with no warnings or errors. Its complete bank recorded
**7,777 passed, zero failed and one conditional provider skip**, across all
**7,778 cases**. The original runner and retained testhost both exited normally
with code 0. No forced termination, AppDomain-unload failure or unhandled owner
post was observed. All preceding methods were retained and 1,648 frozen inputs
were unchanged. All seven recorded generations exited, the exact owned task was
removed and the source freeze released. The complete managed lifecycle gate is
**accepted**. The dedicated x64 testhost runs without a child AppDomain and
resolves assets from the physical test assembly. Native and release acceptance
remain pending for this candidate.

The preceding complete bank also recorded 20 managed transcript performance
measurements, including eight activity workloads at widths of 340 and 780 pixels.
Their individual p95 values ranged from 7.478 to 53.056 ms; the highest was the
expanded, narrow view with 100 activity items. These host-free control workloads
do not qualify real Office rendering or establish a frame-rate guarantee.

### Preceding shutdown and asynchronous-event bank

On 2026-10-11, the shutdown and asynchronous-event candidate (product MVID
`381d3be3-4fe0-4900-b3bd-e4d8ea1888ce`, built from the working tree at base commit
`213231485c3bba0251357357a09392b750b8c24f`) recorded **7,718 passed, zero failed
and one conditional provider skip**, across 7,719 cases in one complete run.
All preceding cases were retained. The 38 additional regressions cover admitted
asynchronous event lifetimes and independent initialization-failure recovery.
Terminal review verified 1,596 frozen inputs and the exact product and test
assemblies. No unhandled dispatcher-post exception was observed in this run.

The original worker exited successfully, but VSTest exceeded its default
test-host shutdown budget and reported forced cleanup. The host log recorded
that exit had begun; no normal host-exit callback was captured. The complete
managed lifecycle gate is therefore **not accepted**. Passing assertions do not
establish normal test-host termination or native application shutdown.

A subsequent bank used the same compiled candidate with a process-only,
five-second VSTest shutdown budget. It was **interrupted at the maintainer's
request** after 558 passing results were recorded. Its partial TRX is preserved;
it is not a complete run and does not resolve the shutdown gate. The exact owned
test host was stopped, the original worker's failed exit was observed, all
recorded campaign processes are absent and its launch task was removed.
The maintainer resumed testing on 2026-10-12; that completed bank is recorded above.
Neither completed or interrupted bank opened Office or SOLIDWORKS, and this candidate
has no new native qualification or line/branch coverage measurement.

### Preceding owner-completion bank: lifecycle scope correction

On 2026-10-11, the owner-completion candidate (product MVID
`5e482222-f690-457e-96f9-660a875f487a`) recorded **7,665 passed, zero failed and
one conditional provider skip**, across 7,666 cases in one complete run. All 51
new regressions passed: 33 owner-STA completion cases and 18 actual Git modal
imports with a caller context that cannot dispatch continuations. The preceding
identity-drift timeout case also passed. These assertion results and the original
worker's normal exit remain valid. Review of the retained VSTest log subsequently
found a test-host shutdown timeout, so the earlier acceptance statement did not
prove normal test-host termination. The complete managed lifecycle gate is
**not accepted** under the explicit normal-host-exit requirement.

Terminal publication now belongs to the returned owner-STA operation task and
runs synchronously after dispatcher cleanup. Distinct operation, cleanup and
publication failures remain recorded. This proves the tested lifecycle contract;
it does not establish that the earlier intermittent timeout had only this cause.

Independent review retained all preceding cases, verified 1,858 frozen inputs and
observed the original worker's normal successful exit. All six observed process
generations and the launcher actor are absent, and the exact owned launch task
was removed. The source freeze is released. No Office or SOLIDWORKS host was
opened by this bank, and native qualification remains pending for this candidate.

### Preceding window-inventory bank

On 2026-10-11, the window-inventory follow-up candidate (product MVID
`446a484f-20f8-4a15-a3e7-e53cfe1ad76a`) recorded **7,613 passed, one failed and
one conditional provider skip**, across 7,615 cases in a single complete run.
All 30 new window-lifetime regressions passed. They cover destroyed windows,
reused handles, stable process/thread identity, retained partial observations,
one bounded fresh read, enumeration bounds and desktop loss. A pre-existing Git
modal-handoff identity-drift case exceeded its 40-second operation deadline.
That failure remains preserved; the complete managed gate is **not accepted**.

Independent terminal review retained every preceding case, verified 1,856 frozen
inputs and the original worker's normal failed exit. All six observed process
generations and the launcher actor are absent, and the owned launch task was
removed. This run opened no Office or SOLIDWORKS host. The prepared native bank
has not been dispatched for this candidate.

### Preceding accepted managed bank

On 2026-10-11 local time, the complete toolbar-ownership follow-up candidate (product MVID
`30d9a687-4e66-4776-9d04-b30bc9dfe55e`) passed **7,584 cases with no failures and
one conditional provider skip**, across 7,585 cases. Independent terminal review
verified all 1,854 frozen inputs, the original worker's normal exit and the
current assembly used by every case. Both preceding full-bank catalogs were
retained, including every case from the immediately preceding accepted bank. The
complete managed gate is **accepted**; native and release acceptance remain
separate.

The preceding context candidate's five failures and original assemblies remain
preserved. Corrected fixtures capture actual source revisions, refuse a missing
live revision and check the size limit on fresh context. Module discovery and
selection retain the source-read revision; selecting a module before discovery
finishes pins that revision explicitly. Sending captures references and draft
attachments before reentrant COM reads, and persists the same captured context
that was dispatched. Queued attachments use detached snapshots. The latest bank
retains all 7,561 cases from the preceding accepted bank and adds 24 toolbar
shutdown ownership cases. These cover indexed and managed collections, reverse
deletion, repeated acquisitions of the same object, exact session tags, partial
getter failures and release/deletion failures without replay. The cleanup uses
the existing reference scope and leaves the borrowed VBE intact. The original
shutdown crash was not reproduced in this candidate's four Excel checks; this
does not establish a complete host or release qualification.

The preceding admission cases for classic Excel shutdown remain included.
Disabled diagnostics perform no
image read; enabled diagnostics require an owned process, durable evidence,
matching loaded-host identity and the retained Excel executable. Earlier
first-chance debugger checks remain in the catalog. Five synthetic helper
scenarios also passed, including a four-record capture bound across six
forwarded exceptions and executed module inventories for all five original processes.
Those helper checks opened no Office host and do not qualify Excel shutdown.

Reference inventory, editor-window ownership, language-reference ownership and
cross-assembly dynamic/serialization scenarios remain in the complete catalog.
The single skip is the explicitly opted-in live Ollama chat case, not an
additional failed or silently omitted managed scenario.

The worker and launcher exited normally. All six observed worker/descendant
process generations and the launcher actor were absent, and the exact owned
launch task was removed. Complete historical descendant capture is not claimed.
This bank opened no Office or SOLIDWORKS host and did not replace the installed
assembly. Compilation reused the previously built native renderer. No new
line/branch percentage, native qualification or release acceptance is claimed.
The repository layout checker still reports an older unmatched
`BridgeServer.PathVisibility.Tests.cs` counterpart.

## Latest native check

### Current candidate: Excel accepted, Word execution pending

On 2026-10-12, product MVID `22aaf833-eb26-4d59-b6bd-aa3c3b9c7b83`
passed the four prepared Excel scenarios for saved module/class/form source,
read-only reopening, unsigned VBA signatures and complete VBE option restoration.
All five owned Excel processes, the group worker and its observer exited normally
with code 0. The temporary COM registration was restored and independently checked.
This accepts these four scenarios only; first certificate selection and native
chat rendering performance remain unqualified.

The complete native preparation retained 38 scenarios in 29 plans. It corrected
six stale `BuildProps` aliases, verified the actual MSBuild target of every plan,
and added guards for normal cleanup after startup refusal and complete native
stderr capture. Excel ran first. The following Word batch produced three
correlated VBA results but remained pending at its original sixty-second
observation deadline. Its failure is retained; no command was replayed and the
deadline was not extended. Other Office and SOLIDWORKS scenarios were not run.

The owned Word process was stopped after evidence capture. COM registration, all
six temporary provider preferences and the model-specific reasoning preference
were restored and verified from the same interactive Windows context as the
native worker. A terminal-context observation showed different registry and
AppData values; it did not establish a concurrent deployment. All native hosts
are closed, the owned tasks were removed, and 1,861 frozen inputs plus 28 child
plans were unchanged before source release. The complete native bank is not
accepted. Evidence is retained under
`artifacts/native-harness-fixes-20261012/` and the associated native-context-L
campaign directories.

### Preceding candidates

On 2026-10-11, the repaired native recipes ran the complete declared 32-case
matrix against the owner-completion candidate (product MVID
`5e482222-f690-457e-96f9-660a875f487a`): **25 passed and seven failed**, with
no case omitted. The complete native gate remains **not accepted**.

The saved-fixture scenarios passed in SOLIDWORKS 2019 and 2025. Strict storage
checks still found changed macro-file bytes after explorer use in both versions,
despite successful native test execution. One SOLIDWORKS 2019 close/discard
outcome remained uncertain; its missing original terminal, host/IDE exit and
registration receipts were retained separately from scoped recovery evidence.
These failures were not converted to passes or replayed. Read-only inspection of
the three failed macro files with complete post-exit artifacts found differences
in the project's `CMG`, `DPB` and `GC` metadata and the removed fixture module's
name in the compiled `_VBA_PROJECT` stream. Original module-source containers
were retained. These observations narrow the investigation; they do not prove
whether native edit/run/remove serialization or VBAi caused the byte changes.
The strict file-preservation failures remain open.

All four Excel checks passed. The Office matrix covered Word, Excel, PowerPoint,
Access, Publisher and Outlook. The chat cases in Access and Publisher timed out
waiting for the visible unprompted native marker; the other cases passed. These
two marker replies completed with nonempty provider answers and no tool calls;
their visible claims about missing tools or module arguments contradicted the
supplied requests. Access displayed the exact upstream answer. No native source
read was attempted in either marker turn, so these failures do not demonstrate
a native read failure or a lost chat refresh. Protected signing/tampering remains outside this
matrix and is not qualified by it.

Terminal review verified 2,037 frozen inputs and 22 child plans, original
coordinator/provider/proxy exits, and every group scope release. All 93 observed
process generations and the launcher actor are absent; the exact owned task was
removed and no Office or SOLIDWORKS host remains. Historical descendant capture
is not claimed complete. The source freeze is released. No release acceptance
or new line/branch coverage percentage is claimed.

### Preceding native recipe admission failure

The subsequent owner-completion candidate (product MVID
`5e482222-f690-457e-96f9-660a875f487a`) stopped at native recipe admission.
Its declared matrix records **one failed preparation case and 31 not run**;
no native test result was produced. The first saved-SOLIDWORKS group still
required terminal Office receipts, although the coordinator scheduled Office
later. The failure location is inferred from the frozen execution order and
absent prerequisite files; the original leaf error stream was not captured.

The original coordinator, helper, provider, group, leaf and proxy process handles
all recorded normal exits. Terminal review retained the missing native receipts
and reverified 2,034 frozen inputs plus 22 child plans. All six recorded process
generations and the actor are absent, the owned task is removed, no Office or
SOLIDWORKS host remains, and the candidate COM overlay is absent. Cleanup made no
registry write and did not claim a successful registration restoration. The
failed scope is released; the complete native gate remains **not accepted**.

### Preceding toolbar-ownership native bank

On 2026-10-11, the toolbar-ownership candidate (product MVID
`30d9a687-4e66-4776-9d04-b30bc9dfe55e`) recorded **seven passed, two failed and
23 not run** across the complete declared 32-case bank on the real desktop.
Independent terminal review verified 2,027 frozen inputs and retained the
original failures. The bank is **not accepted**.

Excel passed catalogue, persistence, unsigned-signature and batched format-option
checks; the two instrumented shutdowns exited normally without reproducing the
earlier crash. Word's explorer and assistant workflows and Excel's explorer also
passed. Excel's assistant failed the visible-streaming deadline: the upstream
response supplied reasoning but no answer text. PowerPoint's assistant failed
when a window disappeared during the test harness's HWND inventory. Its retained
owned process required forced cleanup; normal shutdown is not qualified.

The remaining Office cases and all SOLIDWORKS groups were not run. All 42 observed
process generations are absent and the owned launch task was removed. Concurrent
COM registration, provider preferences and reasoning changes were preserved;
readback confirmed the temporary candidate was no longer registered. Exact prior
registration restoration is not claimed. Complete historical descendant capture,
protected signed/tampered workflows and release acceptance remain unqualified.

### Preceding toolbar and context bank

On 2026-10-11, the complete declared follow-up bank used product MVID
`7008de30-10af-41e2-9164-533209b6cfbc` on the real desktop and recorded
**24 passed, seven failed and one not run** across 32 cases. Independent terminal
review checked 2,024 frozen inputs and the original process exits. The bank is
**not accepted**.

| Scope | Passed | Failed | Not run |
| --- | ---: | ---: | ---: |
| Excel catalogue, persistence, unsigned signature and format options | 3 | 1 | 0 |
| Six Office Test Explorers and six embedded assistant workflows | 10 | 1 | 1 |
| SOLIDWORKS 2025 lifecycle controls | 2 | 3 | 0 |
| SOLIDWORKS 2019 lifecycle controls | 3 | 2 | 0 |
| Both SOLIDWORKS saved-fixture preparation and reopening groups | 6 | 0 | 0 |

Excel's persistence case ended with a CFG fast-fail during shutdown. The
debugger captured four preceding invalid-address reads in the CLR. Its original
module inventory maps the referenced addresses into VBEUI. Eight observed
shutdown/disposal invocations returned, with 128 receipt hashes verified; this
does not prove that native references were released or establish the faulty COM
object. The unsigned-signature, catalogue and batched format-options cases passed.

Five SOLIDWORKS lifecycle cases failed their raw-file preservation checks even
though their native test oracles and normal process exits completed. Their
failures remain preserved. Separately, both versions passed saved-fixture
preparation and reopening in fresh owned instances, with direct and VBE-close
shutdown paths. These results do not replace the failed raw-file oracles.

Office's Outlook assistant passed streaming, cancellation and next-send checks,
but failed to display the required native marker. The final provider response
contained one character after a successful `read_module` result carrying the
marker declaration. Cleanup also reported `RPC_E_DISCONNECTED`. Outlook's Test
Explorer case was consequently not run. The other five Office hosts passed
their explorer and assistant workflows.

Client, plan and observer used the same isolated qualification model. Runtime
sampling observed a loaded 32k context, but did not establish per-request
allocation or a context-truncation cause. All hosts and 102 observed campaign
process generations are absent, the owned task is removed, and the user's
Visual Studio was preserved. Complete historical descendant capture is not
claimed. Protected signed/tampered workflows and release acceptance remain
unqualified.

### Preceding native bank

On 2026-10-11 local time, Native269 used isolated Build270 (product MVID
`7faa6f49-b73c-46e1-949f-c741d4f4f8ee`) on the real desktop. Its complete,
previously declared bank recorded **19 passed and seven failed** across 26 cases.
Independent terminal review verified all 1,864 frozen inputs and the original
worker, provider and group exits. The complete native bank is **not accepted**.
Earlier failed banks and their original receipts remain preserved.

Excel passed the catalogue/focus, product-save/reopen, unsigned-signature and
complete-format-options scenarios. All six Office Test Explorer scenarios also
passed. Excel's three coverage-copy cycles preserved the original and detached
every copy, leaving only the original workbook and VBProject. These results are
specific to the recorded candidate and installation.

Four embedded Ollama assistant scenarios passed: Excel, PowerPoint, Access and
Publisher. Word passed streaming while busy, cancellation and next-send recovery,
but its final response did not expose the disposable project's native marker.
Its original process exited normally within the single five-second observation
after Quit; the previous shutdown failure did not recur in this case. Outlook
failed the visible-streaming deadline. Both complete assistant cases remain
failed. Cooperative message-pumping scenarios do not qualify chat responsiveness
during non-pumping CAD work.

SOLIDWORKS 2019 and 2025 passed their blank-project lifecycle checks; 2019 also
created its own saved seed. Both modified Test Explorer fixtures preserved source
and references before exit, but their saved macro's raw bytes changed during
normal shutdown. Closing the VBE window before shutdown also failed byte-level
preservation in both versions: it hid the editor without confirming that the
modified project had been discarded. These fixtures installed test/runtime
modules, edited source and removed their owned modules without saving. They do
not isolate execution and inspection of an already saved test fixture.

A fresh SOLIDWORKS 2025 process reopened the exact resulting file without
executing a macro and found unchanged source, references and project properties.
Its preservation verdict remains failed because the original binary had already
changed. Logical readback does not replace byte-level preservation.

Protected first-signature selection and signed-file tamper verification were
explicitly **not run** in this unattended bank. No certificate was created and
no cryptographic qualification is claimed. The implementation still requires
operator selection in the protected Windows Security picker.

All campaign-owned Office and SOLIDWORKS instances, their dedicated IDEs and the
provider observer are released. The final ownership check verified the absence
of all 170 captured process identities. The owned launch task was removed and
the user's Visual Studio instance was preserved. Temporary COM registration and
preferences were reviewed and restored within the actual test actor. Original
failed outcomes were retained without replaying a native mutation. These findings
are operation-specific and do not constitute complete VBE, provider, signature
or release acceptance.

## Earlier native check

On 2026-10-10, Native239 exercised the earlier Build246 candidate (product MVID
`dc007102-3a28-4df5-a0da-6e9bfcfbe842`) on the real desktop:
**17 passed, three failed and three were not run** across its 23 declared cases.
Independent terminal review checked all 1,822 frozen inputs. The complete native
bank is not accepted, and protected first-signature/tampering workflows remain
outside this bank's scope.

The initial Excel group passed its catalogue/focus, save/reopen, unsigned-signature
and complete-format-options scenarios. All six Office Test Explorer scenarios
passed. In particular, Excel's three distinct coverage-copy cycles preserved the
original source and file bytes, and each copy detached completely: only the
original workbook and VBProject remained. This positive result is specific to
this candidate and the recorded Office installation; Microsoft updated the
Office executables before the bank was sealed, so the ownership correction alone
has not been established as the exclusive cause.

Five Office assistant scenarios passed. Word passed streaming while busy,
cancellation and the next response, but its final reply did not expose the value
read from the disposable marker module. The available provider metadata does not
establish whether the failure originated in the tool request, its result or the
model's answer. No native action was replayed.

In SOLIDWORKS 2025, both blank-project shutdown scenarios passed. The Test Explorer
scenario restored source and references before exit but changed the saved macro's
raw bytes; its preservation verdict remains failed. SOLIDWORKS 2019 did not become
ready for its exact-PID COM connection while a startup dialog was visible. Its seed
setup failed and the three dependent cases were not run. The dialog's meaning
could not be established by the bounded read-only control inspection.

All campaign-owned hosts and observers were released, temporary COM registration
was restored and the launch task was removed. SOLIDWORKS 2019 and its dedicated
Visual Studio required forced recovery after the startup failure; this establishes
cleanup, not normal-shutdown qualification. The original failed scope review is
retained alongside the independent recovery review. The user's Visual Studio
instance was preserved. Original failures from earlier candidates remain recorded.

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

## Transcript rendering and shutdown follow-up

The earlier accepted complete Unit candidate tested on 10 October 2026 is based on source revision
`213231485c3bba0251357357a09392b750b8c24f`. Its product MVID is
`393858b2-7006-4b61-8609-957a6c869428`, with SHA-256
`1ECFFE1893B327DB9DA0467F6C9AE304DB64A370E773E7586993065ECB7AEEBD`.
Its complete Unit gate recorded **6,803 passed, no failures and one conditional
live-provider skip**. Independent review checked the original runner exit,
frozen input hashes and 56 required regression groups, including all 16 shutdown
cases. Evidence is retained in
`artifacts/excel-audit-followup-20261008/completion-managed144/`.

The candidate avoids RichEdit access for unchanged activity rows and preserves
explicit disclosure choices when captions change. A historical serial comparison
on the same actual-control workload matrix passed 27 scenarios on each product,
but timing was mixed: one narrow expanded activity workload improved from
213 ms to 202 ms total while its p95 increased from 21.4 ms to 25.6 ms. This is
neither a general speedup nor a real-host frame-rate qualification. The original
comparison remains in `completion-transcript-comparison126/` under the same
artifact root.

Shutdown handles an already-detached WebView2 controller without retrying its
native close. Embedded tool windows stop site tracking and dispose their child
forms in place; ordinary floating-window detachment remains separate. Unexpected
failures still propagate.

SOLIDWORKS 2025 SP1.1 loaded this exact candidate on the owning STA and reached
the unchanged-source/reference/disk baseline. Its first blank-macro control
returned from Monaco disposal and site-tracking shutdown, then exited with code
one while the trace was inside `chat.Dispose`. The original host exit was
observed; the utility IDE exited normally, registration was restored and the
launch task was removed. **One scenario failed and two were not run.** The
precise native cause and post-exit raw storage remain unqualified. Evidence is
retained in `solidworks-storage-main145/root-failure-terminal-review.json`.

The previous product's Office matrix recorded eight passing cases, one failing
Outlook chat shutdown and three cases not run. The Outlook chat completed its
streaming, Stop and next-send oracles, but `Quit` failed with `0x800706BE`;
later process absence does not prove a normal exit. Registration and preferences
were restored and the failure is retained in
`completion-native141/root-terminal-review.json`. The next twelve-case Office
matrix on the product identified above recorded **nine passed, one failed and
two not run**. The Outlook chat passed with a normal exit; the Outlook explorer
completed its execution and project-restoration checks but failed at `Quit`
with the same RPC error. The two PowerPoint cases were not dispatched. All
owned hosts were subsequently absent, and scoped registration and preferences
were restored. Evidence is retained in `completion-native146/root-terminal-review.json`.

The subsequent focused product has MVID
`cdc5bf29-e598-4b5b-b1d9-849a9b8c0855` and SHA-256
`699905E3CFCEC2EEBE20E7513747B126DCE6880300A820C7701F8CC2A7F22D01`.
Its complete affected-family gate recorded **430 passed, no failures and one
conditional live-provider skip**. Independent review verified the original
runner exit, frozen inputs, exact assembly and 24 required regression groups
(`completion-managed168/root-terminal-review.json`). This is a focused gate;
the preceding full Unit run remains historical.

The subsequent complete Unit run on this same candidate recorded **6,879 passed,
one failed and one conditional live-provider skip**. Independent review verified
the original runner exit code of 1, exact assembly, 1,357 frozen inputs and 75
required regression groups (`completion-managed172/root-failure-terminal-review.json`).
The failing rollback-authority case (`preview`, `after`) expected the Git/VBA-state
refusal but received the source-change refusal. Its cause is not established and
the full managed gate is not accepted. This run does not measure line or branch
coverage and does not qualify native-host behavior.

An admitted SOLIDWORKS trace located an invalid COM reference in the native
container's ambient-background paint callback during child-handle destruction.
The shared chat/test-explorer container now owns a local background color and
copies its form's background on attachment. On that product, the blank
SOLIDWORKS storage control passed with normal process exit, unchanged sources
and references, and identical macro-file bytes (`solidworks-storage-main169/`).
The independent Explorer control subsequently timed out on a stale-revision
refusal request and retained its uncertain state; the remaining control was
not dispatched. After the maintainer authorized stopping campaign-owned instances,
the exact retained SOLIDWORKS and utility Visual Studio processes were stopped.
The original controller observed their exits; scoped registration was restored
and the launch task was removed. Independent terminal review preserves **one
passed, one failed and one not run** (`solidworks-storage-main169/root-failure-terminal-review.json`).
Forced recovery does not qualify normal shutdown or settle the uncertain request.

Independent review of the completed blank control verified its original host
and utility-process exits, scoped registration restoration, 1,587 frozen inputs,
and all 132 raw shutdown receipts. Runtime cleanup, component disposal and native
handle destruction returned. The existing diagnostic reader rejected the newly
added nested chat stages; this reader failure is retained separately from the
normal-process exit and storage evidence. The corrected stage contract is now
integrated; the recorded native result still belongs to the earlier product.

An isolated reader update passed **178 managed cases** against the frozen product
binary, including validation of the captured blank-control trace
(`receipt-reader-study174/results/root-terminal-review.json`). A separate local
Git authority study passed **15 cases**, preserving the original refusal and
zero-import assertions (`git-authority-study175/results/root-terminal-review.json`).
The earlier full-suite rollback failure remains unexplained; the isolated pass
does not replace it.

An additional producer regression showed that an ancestor invocation could record
its active child's exception. The proposed guard restricts attribution to the
innermost active invocation. The isolated baseline reproduced all four failures;
the corrected study passed **183 cases**, with an explicit check that the fixture
used the modified study type (`producer-scope-study182/results/root-terminal-review.json`).
An earlier study was rejected because its fixture still bound to the unmodified
product type. These reader and producer changes are now integrated with 60 new
reader cases and four producer cases. The producer change affects diagnostic
attribution; it is not evidence of a fix for the native Explorer timeout.

The subsequent complete Unit context study retained the same product binary and
the original case catalog, adding the reader cases and Git refusal diagnostics.
It recorded **6,939 passed, one failed and one conditional live-provider skip**
(`managed-context-study178/results/root-terminal-review.json`). All added reader
cases passed. The earlier Git case passed with its expected refusal and unchanged
source, but this does not establish the cause of its original failure. The new
failure was a modal-diagnostic STA continuation deadline. The full gate remains
unaccepted. A traced study of the complete affected dispatcher, modal-diagnostic
and inspection-trace families passed **58 cases** and verified the original STA
for both continuations (`dispatcher-study183/results/root-terminal-review.json`).
The full-context timeout remains unexplained; these passing scoped results do
not replace the failed complete run.

The integrated working-tree candidate was compiled with **no warnings or errors**
into `completion-build185/`. Its product MVID is
`8115f36b-6b85-4c25-a5da-a4e82d3cc4cf`, with SHA-256
`07905A2E48E730C948337ACD53043C90452B262F1311C01E9972F93509CCF6EF`.
The first integrated complete run, `completion-managed186/`, ended without an
original exit receipt or final TRX. Its runner and testhost were independently
verified absent; its 3,785 observed result notifications remain partial evidence,
not an accepted gate. The termination cause is not established.

The durable complete run, `completion-managed189/`, recorded **6,943 passed,
one failed and one conditional live-provider skip**. Independent review verified
the complete original catalog, all 64 added cases passing, frozen source/binary
inputs and 75 required regression groups. Its sole failure was the pure Node
debugger-script test: the task worker could not resolve `node.exe`. The original
worker exited with code 1 and its owned task was removed. This failed gate remains
recorded; it is not current native acceptance.

The same **6,945-case Unit catalog** passed in `completion-managed190/` with
explicit SHA-checked paths for Node, Git, PowerShell, Python and .NET. The paths
apply only to the test process; no persistent environment setting is changed.
The product/test assemblies are unchanged. The result is **6,944 passed, no
failures and one conditional live-provider skip**. Independent review checked
the complete catalog, all 64 added cases, 75 required regression groups, frozen
inputs, runtime resolution and the original worker's normal exit with code zero.
Its launch task was removed. This is an accepted managed gate, not a line or
branch coverage measurement.

The twelve-case Office matrix on this candidate finished in
`completion-native188/`, on the real desktop with disposable projects: **nine
passed, two failed and one was not run**. The test explorer and assistant passed
in Word, Excel, Access and Publisher; the PowerPoint explorer also passed.
PowerPoint streamed, acknowledged Stop and completed the next assistant turn,
but its displayed answer did not contain the required recovery marker. Its
fixture exited normally; the failed oracle remains failed. Outlook failed during
Quit with `0x800706BE`, and its explorer was not dispatched. Outlook's process
was absent at later inspection; this does not reconstruct a normal-exit receipt.
Independent review verified the nine passing cases, all 1,567 frozen inputs and
scoped COM/preferences restoration. The owned Office task was removed. The
auxiliary provider controller exited with code one after the shutdown request;
its proxy PID and listening port were absent, but its normal-shutdown receipt was
missing. This diagnostic limit is retained separately from native results.

The three independent SOLIDWORKS controls were dispatched once in
`solidworks-storage-main187/` after Office cleanup: **one passed, one failed and
one was not run**. The blank-project direct-exit control preserved source,
references and raw macro bytes, with normal original-handle exits. The ordinary
VBE-close control failed its settlement oracle: the host exited with code zero,
but the required VBE-close completion with the host still alive was not observed.
The typed debugger snapshot showed Design mode and no targets at failure; it
does not support a debugger-break explanation for this case. Explorer was not
dispatched. Independent review checked all 1,596 frozen inputs, preserved the
missing original case/bank terminal receipts, and verified COM restoration and
owned-task removal. No native action was replayed.

Neither matrix inherits native acceptance from the preceding product. The 17
actual chat captures are visual observations, not frame-rate or layout acceptance.
An isolated, windowless inspection of build185's message controls at widths 300,
500 and 800 pixels measured 18 extra pixels from three visible empty metadata
rows. This narrow result does not explain all larger gaps seen in hosted captures
or qualify WPF layout performance.

The build193 correction candidate (MVID
`85f05472-19a7-4ef1-9d0a-8d4e87f7da5d`, product SHA-256
`BDE58B3915466CA6014CDAB92E13704629948DB4E8624FB99EC544704E05CD7C`),
compiled with no warnings or errors. Empty message metadata rows now hide and
reappear when their contents change. The qualification fixture observes its
retained Outlook process handle after a failed Close/Quit, while preserving the
original failure and refusing further COM actions or OTM recovery. The VBE-close
helper now posts the system-menu `SC_CLOSE` command, with continued host lifetime
still required by its native oracle.

All 24 added scenarios were implemented before dispatching the complete
6,969-case Unit catalog in `completion-managed194/`: **6,968 passed, no failures
and one conditional live-provider skip**. Independent review preserved the full
previous catalog, checked all 24 added cases and 75 regression groups, verified
1,526 admission inputs and the original worker's exit with code zero. Its owned
task was removed. This is an accepted managed gate; it does not measure line or
branch coverage. The earlier Office/SOLIDWORKS results remain specific to build185
and do not qualify build193.

The complete native matrix was prepared before dispatch: twelve Office cases,
four SOLIDWORKS 2019 cases and three SOLIDWORKS 2025 cases, on the main desktop.
For build193, Office (`completion-native197/`) finished with **10 passed, one
failed and one not run**. Explorer and streamed chat scenarios passed in Word,
Excel, PowerPoint, Access and Publisher, including normal host shutdown. Outlook
passed its streaming, cancellation and recovery assertions, but `Application.Quit`
failed with `0x800706BE`. Its retained process handle subsequently observed exit
code zero; this does not qualify the failed Quit. The Outlook Explorer scenario
was not run. Temporary preferences and COM registration were restored.

The bounded provider observer ran for that batch and exited normally. It recorded
SSE counts, content digests and recovery-marker predicates without storing reply
text, reasoning, tool arguments or credentials. Its eight standalone parser tests
cover fragmented UTF-8, split markers, incomplete streams, malformed payloads and
privacy; they are distinct from the native chat results.

SOLIDWORKS 2019 SP5 (`solidworks-storage-main199/`, revision `27.5.0`) finished
with **one failure and three not run**: the fresh year-local seed producer exceeded
the normal-exit deadline, so no accepted seed was supplied to the remaining
controls. SOLIDWORKS 2025 SP1.1 (`solidworks-storage-main198/`, revision `33.1.1`)
finished with **one passed, one failed and one not run**. The saved blank control
preserved its bytes and exited normally. The Explorer stale-revision request timed
out with an uncertain outcome. Read-only inspection found an owned Visual Studio
inconsistent-line-endings dialog naming `VbeTestExplorerService.cs`; its impact on
other failures is unproven. Authorized cleanup stopped the blocked owned host and
debug utility, preserved the original failures and backups, and restored COM
registration. No uncertain native mutation was replayed. Neither system-close
control received native acceptance. All campaign-owned hosts and tasks were
released; the unrelated development IDE was preserved.

The accepted Unit194 run also contains **27 transcript performance controls**,
all passed. At width 340, appending 131,072 characters over 64 frames took
152.372 ms with a 3.973 ms p95; updating 100 expanded activities took 198.124 ms
with a 21.310 ms p95. These local workload observations use simulated controls,
not a live Office/SOLIDWORKS chat. The native captures use `PrintWindow`, so they
do not establish on-screen composition quality, frame latency or physical DPI
acceptance.

Earlier six-host chat evidence (`completion-native132/`) and shutdown failures
(`solidworks-storage-main135/` and `solidworks-storage-main138/`) identify their
own products. They are not current-candidate acceptance. Target-only screenshots
are separate visual observations and do not qualify dark themes, a physical DPI
sweep, native frame latency or chat responsiveness during a busy CAD operation.

No current line/branch measurement or complete release acceptance is claimed.

The follow-up build201 compiled without warnings or errors (MVID
`3422af36-bc63-4003-a46f-d83cff84f015`, product SHA-256
`C1D4D733E4CC1B704A0192DB8504A97B597632F9F2C1B41DCB93F5DD839F551A`).
`VbeTestExplorerService.cs` now has consistent CRLF line endings. Shutdown
diagnostics use Unicode Win32 attribute reads to avoid expected missing-path
exceptions during atomic publication. Only missing-file and missing-directory
errors are tolerated; other read errors still fail admission, and existing linked
ancestors remain refused. This reduces observer work; it does not establish the
cause of the SOLIDWORKS shutdown deadline failure.

All 17 new regression scenarios were written before the complete Unit202 batch:
**6,985 passed, no failures and one conditional live-provider skip** out of 6,986
cases. Independent review verified the full previous catalog, every added case,
1,527 admission inputs, the loaded test assembly and the original worker's exit
code zero. Its owned task was removed. The provider skip is
`LocalOllamaShownChatStreamsStopsAndCompletesNextSend`. This is an accepted managed
gate with no line/branch measurement.

The complete native follow-up matrix was prepared and sealed before dispatch on
build201. Its 19 scenarios ended with **13 passed, three failed and three not
run**: Office203 contributed 10/1/1, SOLIDWORKS 2019 bank204 2/1/1, and
SOLIDWORKS 2025 bank205 1/1/1. Independent reviews verified their immutable inputs
and original coordinator exits. These results do not complete native acceptance.

The unresolved failures are Outlook's `Application.Quit` (`0x800706BE`), a
SOLIDWORKS 2019 fixture rejected because its newly generated macro still contains
`Sub main`, and an emitted SOLIDWORKS 2025 test command whose response deadline
expired while its debugger was paused. The blank-macro guard was preserved; no
uncertain command was replayed. Recovery closed only owned processes, restored
the scoped COM registration and removed the temporary tasks. Forced recovery of
the 2025 host does not qualify normal shutdown or recover its unsaved state.

Supplemental captures observed the actual chat panel on the main desktop in all
six Office hosts. The Word image shows overlapping history text and gray areas
in the composer. Those captures followed `PrintWindow`, so its possible effect
on the visible surface has not been isolated. Rendering correctness, native
frame latency, DPI coverage and responsiveness during CAD work remain unqualified.
Evidence is retained in `completion-native203/`, `solidworks-storage-main204/`
and `solidworks-storage-main205/` under the audit artifact root.


The next candidate, build211 (MVID `2dd9848e-ce13-4e12-8ee3-6de216115889`),
returns expected test-source revision refusals directly at the bridge boundary.
The strict internal command API and its exception contract remain unchanged.
Owner-thread, project, revision and execution-authorization checks are retained.
A fixture-only planner admits removal of the observed SOLIDWORKS-generated
default procedure using an exact source checksum; arbitrary code is refused.
Neither this planner nor its managed tests executes a macro.

The complete Unit212 batch ended with **7,028 passed, no failures and one
conditional live-provider skip** out of 7,029 cases. All 43 added scenarios passed,
the complete Unit202 catalog was preserved, and the exact assembly, immutable
inputs and original worker exit were independently verified. Its temporary task
was removed. The Office213 follow-up finished with **7 passed, 4 failed and
1 not run**. Failures were a Word visible-stream timeout, an Excel coverage
copy remaining in `VBProjects` after its workbook closed, a PowerPoint
next-response display timeout, and Outlook `Application.Quit` returning
`0x800706BE`. The Outlook explorer case was not run. The hosts, relay and
temporary tasks were released, and temporary settings were restored.

Fifteen accepted actual-screen observations reproduce gray composer rectangles
and transcript overflow; the observer uses `CopyFromScreen` without calling
`PrintWindow` or activating a window. These observations establish visible
defects, not their cause or rendering performance.

The SOLIDWORKS217 (2019) follow-up ended with **1 failed and 3 not run**:
startup encountered a toolbar-information dialogue and exceeded readiness
admission before creating a macro. The owned host required authorized forced
cleanup; that cleanup does not qualify normal shutdown. An earlier controller
failure, before any host opened, remains separately recorded.

SOLIDWORKS218 (2025) ended with **1 passed, 1 failed and 1 not run**. The blank
macro lifecycle passed. The installed test-explorer functional test passed,
including stale-source refusal, without the earlier debugger pause. Its
separate raw-storage check failed after normal shutdown because the saved
macro bytes changed despite logical source/reference restoration. The final
case was not run after the frozen-input check detected that change. Scoped
registration restoration and removal of the owned tasks were verified;
all campaign hosts were released. This does not qualify native shutdown across
all hosts, unchanged raw storage, line/branch coverage or a release.

Build222 (MVID `5fb6fa1f-fa53-4fa4-bf00-0f1ac2e685dc`) adds native viewport
clipping for transcript cards and opaque themed composer backgrounds. Full
Unit223 completed with **7,032 passed, no failures and one conditional
live-provider skip** out of 7,033 cases. All four added rendering scenarios
passed, including native child-window region bounds during scrolling, resizing,
scale changes and unload/reload. The full Unit212 catalog, frozen inputs,
assembly hashes and original worker exit were verified; the temporary task was
removed. These tests establish the implemented clipping and background contracts;
the loaded add-in's rendering, responsiveness and complete host lifecycle still
require native qualification. Subsequent source changes are outside this result.

Build224 (MVID `7461deb2-7b3d-4d73-91d5-b0c490f62d99`) includes the optional
Outlook cleanup observer and draft-preserving project navigation. Full Unit225
ended with **7,076 passed, two failures and one conditional live-provider skip**
out of 7,079 cases; 45 of the 46 added cases passed. The failures were delayed
popup-close state and an existing Monaco recovery continuation remaining busy.
The catalog comparison explicitly accounts for the diagnostic's newly admitted
PowerPoint process-name row. Frozen inputs, original exit and task removal were
verified. The gate failed: no native deployment or release acceptance follows
from this batch.

Build226 (MVID `014820b0-1b49-471c-b4b9-1595267343f6`) fixes the delayed
popup-close race and adds bounded traces around the existing Monaco recovery
continuations. Full Unit227 completed with **7,079 passed, no failures and one
conditional live-provider skip** out of 7,080 cases. Review verified the complete
Unit225 catalog, both previous failing cases, frozen input and assembly hashes,
the original worker's normal exit and removal of its temporary task. The new
regression exercises repeated popup close/reopen and disposal of its observer.
The Monaco recovery case passed, but its earlier intermittent failure's cause
remains unproven. These results qualify the managed gate only; native host
behavior and line/branch coverage require separate evidence. Terminal review is
retained in `completion-managed227/root-terminal-review.json` under the same
artifact root.

The next native bank on this candidate ended with **4 passed, 3 failed and
21 not run** across its 28 declared scenarios. The original Excel batch selected
only seven cases: two format-option selectors used an obsolete namespace. The
three failures were two abnormal Excel shutdowns (`0xC0000409`, faulting module
unknown) and a first-signature certificate confirmation timeout. The retained
signature fixture required authorized forced cleanup. Its temporary certificate
and the completed launch task were removed. Snapshot restoration refused a
concurrent COM registration change and preserved it; the original COM baseline
was not restored. No Office or SOLIDWORKS follow-up group ran, and no uncertain
mutation was replayed. Review is retained in
`completion-bank228/root-terminal-review.json`. Subsequent filter admission uses
actual compiled test metadata; this does not qualify the missing native cases.

The archived Windows Error Reporting records for both crashes identify fast-fail
parameter `0xA` (`FAST_FAIL_GUARD_ICALL_CHECK_FAILURE`). Neither archive contains
a dump or call stack, so the failing indirect call and its cause remain unknown.
The bounded report is retained in `completion-bank228/wer-crash-review.json`.
The first-signature timeout also requires a separate interpretation: the current
implementation waits for a person to confirm the certificate in Windows Security.
The unattended batch did not supply that interaction. Its failed result remains
unchanged and does not, by itself, demonstrate a signature implementation defect.

The disabled-by-default native teardown collector now also admits the exact
unsigned-workbook and module/class/form persistence scenarios. It retains the
owned process, candidate, nonce, debugger state and single-cleanup checks; it
does not sign a project, disable Windows protections or replay a native action.
This is diagnostic preparation, not a correction of the crash or native
qualification. Its new managed scenarios must pass before a diagnostic host run.

Build229 retains the exact product assembly from Build226 and updates the test
assembly for this collector extension. Full Unit230 completed with **7,097
passed, no failures and one conditional live-provider skip** out of 7,098 cases.
Independent review preserved the entire Unit227 catalog and verified all 18
additional admission, missing-readback and single-cleanup scenarios, input and
assembly hashes, the original worker's normal exit, and removal of its task.
Evidence is retained in `completion-managed230/root-terminal-review.json`.
These are managed results, not native acceptance or a line/branch measurement.

The next native campaign was prepared as a single 23-scenario batch: the 21
previously unexecuted scenarios plus two new shutdown observations using fresh
disposable workbooks. Its selectors are resolved against the compiled test
assembly before launch. The first-signature case is kept outside this unattended
batch because its protected certificate confirmation is an explicit interaction.
All four synthetic debugger preflight scenarios passed before Excel dispatch;
this validates the collector, not Office or SOLIDWORKS behavior. Independent
terminal review of `completion-bank231/root-terminal-review.json` verified all
1,795 frozen inputs and recorded **13 passed, four failed and six not run**.
The Office submatrix passed 11 of its 12 scenarios; Excel's completed coverage
copy remained in the VBE project collection. The other failures were Excel's
abnormal process exit during persistence, an unavailable Office signature
verification component, and SOLIDWORKS 2019's exact-PID ROT readiness timeout.
The three dependent 2019 cases and the three 2025 cases were not dispatched.

The Excel crash diagnostic captured a second-chance `0xC0000409` with parameter
10 during CLR shutdown, after 318 first-chance access violations. The failing
callback and underlying cause remain unidentified. The signature failure
reported `VerifierUnavailable` for `msosipx.dll`; it does not establish an invalid
signature or corrupted workbook. SOLIDWORKS 2019 required separately authorized
forced closure; the campaign's Visual Studio utility then closed normally.
Original failures and missing lifecycle receipts remain retained. Scoped COM and
provider settings were restored, the observer exited normally, the finished
task was removed, and an independent process inventory found no remaining host.
Neither cleanup nor the passing rows establish complete native qualification.

Build232, product MVID `5f5aa3ad-a73f-47b2-baf1-194fdf734008`, balances temporary
NativeOM window and rejected application acquisitions in Excel application
resolution. All eleven new ownership and release-failure regressions were written
before the complete Unit233 batch was dispatched. That accepted gate retained the
complete Unit230 catalogue. Its original worker exited normally and its task was
removed. The independent review retained its initial admission error and then
enforced the full baseline regression counts, rather than the smaller counts of
additions. The relationship between the corrected reference leak and the retained
Excel failures remains unproven.

Build234, product MVID `bf59945e-f98d-472f-9308-11cbcdfaa715`, adds discovery,
trust validation and scoped loading of the current registered Office VBA SIP in
`mso.dll`. The complete Unit235 gate passed 7,210 tests with zero failures and one
conditional provider skip. All 102 new registration, path, policy, fallback and
module-lifetime scenarios passed; the complete accepted Unit233 catalogue and its
regression group sizes remain present. Independent review verified the frozen
candidate inputs, original worker exit and removal of its owned task. This is
managed evidence, without a new line/branch coverage measurement. The 23-scenario
Native236 bank completed on this same frozen candidate: 15 scenarios passed,
three failed and five were not run. Its first four Excel
scenarios passed: complete Format restoration, catalogue-focus metadata,
unsigned-file verification through the registered `mso.dll` VBA SIP, and product
save/reopen preserving modules, classes and forms. The group exited normally,
released its hosts and restored COM registration. Eleven of twelve subsequent
Office scenarios passed. Excel's repeated-copy scenario failed after the second
copy: the workbook collection contained only the original, but the VBE retained
a pathless project. The first cycle detached correctly; the third was not reached.
SOLIDWORKS 2025 reached the exact-process COM connection, then failed before the
product checks because the harness had not staged its required seed manifest.
SOLIDWORKS 2019 did not establish its exact-process ROT connection within the
startup deadline; an owned modal was observed, but its meaning remains unknown.
Dependent SOLIDWORKS scenarios were not run.

All campaign hosts and its Visual Studio utility were released; the 2019 host
required authorized forced termination, which does not qualify normal shutdown.
The provider observer exited normally and the owned launcher task was removed.
The final scoped COM restore refused concurrent changes without writing: current
AddIn/Chat registration points to `bin/Debug/net48/VBAi.dll`, with TestRuntime
registration absent. Those changes were preserved, not reported as a verified
restore. Independent terminal review verified all 1,800 frozen inputs and retained
the original failures and missing receipts. This does not qualify signed or
tampered files, establish the cause of earlier shutdown failures, or accept the
complete native workflow.

## Excel audit correction candidate

### Audit continuation on 9 October 2026

The continuation is based on `5eb46e9c0a5542a64b946825cc9133a4ae19e7bb`,
with the first correction commit `6c359286` and additional working-tree fixes.
The first completion candidate, product MVID
`939d7fe6-64f5-483a-8808-191ac3fd3f7b`, passed 5,616 Unit tests with no failures
and one live-provider opt-in skip. Its original native bank passed two cases
and failed five; two Options cases were not selected because their declared
namespace differed from the filter. The original failures remain retained.

The second completion candidate has product MVID
`cf4c973b-041e-455d-a436-deaf1dcd2ea2` and SHA-256
`018FABB93479CFB133A41AE1EA0A30DA5B0EC37AAA651BA44CC1AF2019582F42`.
Its affected managed gate passed 596 tests without failures or skips. Its real
desktop native bank executed nine cases: seven passed and two failed. Passed
cases cover cold RunSub targeting, native BeforeSave cancellation and acceptance,
module/class/form save and fresh-process reopen, embedded Monaco live scope after
another project closes, read-only Options inspection and complete Format
round-trip/restoration. Cancellation preserved the exact disk baseline;
acceptance counted one event and persisted the requested changes. The root-reviewed
Monaco capture contains the actual retained tabs, selected live module and native
synchronization status. Owned Excel and Word fixtures shut down normally; scoped
COM registration and unchanged test-owned preferences were restored and verified.

Both Ollama assistant cases failed before visible streaming while busy; Stop,
next Send and the subsequent native read were therefore not accepted. A later
local endpoint probe found no listener on port 11434. This observation does not
establish the sole cause of the earlier failures or compare model quality.
Evidence is retained under
`artifacts/excel-audit-followup-20261008/completion-native2`, including
`native-completion.trx`, original-handle lifecycle receipts, registration readback
and `root-monaco-review.json`.

Expanded detached activity groups were a measured performance problem on that
historical second candidate. With 100 completed siblings and 32 streamed frames, the
340-pixel viewport recorded 10.599 seconds total and an 884-millisecond 95th
percentile. These are detached WPF/WinForms timings, not real-host frame rates.
The subsequent constrained-size cache and grouped Format batch were built as
product MVID `bd0395fb-8e8c-4c1f-862a-2b5f6ffede91`, SHA-256
`672AD1968AB528341478BF1F842F990291E0EEB1928903367C644ACD89F96F64`.
Their affected gate passed 864 tests with no failures and one unactivated live
provider skip. It includes complete batch/restoration fault matrices, legacy
message compatibility, nested tool validation, edit-policy guards and cache
invalidation against independently constructed controls. On the same detached
100-sibling, 32-frame workload, the 340-pixel viewport recorded 197.906 ms total
and a 21.950 ms 95th percentile; the 780-pixel viewport recorded 99.522 ms total
and a 14.874 ms 95th percentile. These observations retain content, selection
and control-identity checks but are not native rendering acceptance.
The intermediate failed gates and abnormal compiler exit remain retained;
the successful gate is `completion-managed5/managed5.trx` under the continuation
artifact root. The prepared real-desktop bank then passed all three selected
native workflows with no failures or skips: Excel assistant, Word assistant and
complete grouped Format. The Format case staged 33 edits before one OK, verified
them independently, rejected the stale revision, restored the confirmed edits
with one OK and recovered the exact complete baseline. Both assistant cases
observed visible streaming while busy, one Stop, cancellation acknowledgement,
successful next Send and the hidden marker from the disposable module. Sources
and references remained unchanged. All owned hosts and the original worker
exited normally; registration and test-owned settings were restored. Frozen
inputs remained unchanged. Evidence is retained in `completion-native3`, including
`root-terminal-review.json` and the original native TRX. This is the declared
workflow scope, not a complete Office qualification or native frame-time claim.
Physical DPI/docking, current SOLIDWORKS behavior and signature application remain
outside this acceptance. Native signature verification is blocked by the missing
Microsoft VBA SIP provider. No current complete line or branch coverage was
measured.

### Office continuation and installed Excel execution

On the `bd0395fb` candidate above, the next prepared Office bank passed the
PowerPoint and Access assistant workflows. Publisher passed visible streaming,
Stop acknowledgement, next Send and native marker readback, but its original
process exit was not observed within the five-second shutdown deadline. Its
workflow therefore failed; the six later rows were not run. A later process
disappearance does not replace original-handle exit evidence. Registration and
the test-owned preferences were restored. Evidence is in `completion-native4`.

Subsequent installed Excel coverage trials exposed an inaccessible second VBE
project after the disposable workbook closed. Independent readback confirmed one
open workbook and intact original source and disk bytes; this did not satisfy
the stricter VBE project-absence oracle. Those failed trials are retained in
`completion-native5` and `completion-native6`.

Balancing the acquired explorer, compilation and Excel returned-value COM
references produced candidate MVID `f29f6248-9887-407d-bac6-05e4d6972145`, SHA-256
`9486853AF7759C558779192F71AEEE7D4713F0FC0F85976FDE3714D5902F8438`.
Its affected managed gate passed 836 tests without failures or skips
(`completion-managed8/managed8.trx`). The installed Excel case then passed the
complete registered callback workflow: four independently checked batch outcomes,
one selected test, real copy coverage with one of two production procedures
entered, original source/revision/disk preservation, absence of the closed copy
from both workbook and VBE collections, native docked explorer and stale-revision
refusal without another invocation. Owned Excel exited normally.

That prepared twelve-case bank next failed the Word unsaved-source oracle:
the owned application's ForceDisable policy correctly refused earlier than the
expected save-required guard. Word also exited normally. The original worker
exited with failure; registration and all six test-owned preferences were restored
and verified. The result is one passed case, one failed case and ten not-run
cases, retained in `completion-native7`. The Word scenario prerequisite was
corrected to the same temporary owned-application ByUI policy used by its saved
copy scenario, with restoration after a settled outcome; no Trust Center policy
is changed. Native event effects, Publisher shutdown and the remaining host rows
are not accepted by this partial bank. The source snapshot is based on commit
`21323148` with the scoped working-tree corrections recorded by the frozen plan.

The subsequent candidate has MVID `a74f74b0-0adc-4bee-bf99-4663a0b4f0b7`,
SHA-256 `C0528359B06FA1FC00B7D8A1A607688629E4F8530D84908DF360C2BA6D859279`.
Its affected managed gate passed 1,033 tests with no failures or skips
(`completion-managed10/managed10.trx`); this is not a whole-solution coverage
measurement. `completion-native8` again passed the complete installed Excel case.
Word passed the unsaved-source refusal and registered batch/single assertions,
but copy execution encountered the native macro security warning and did not
complete within the qualification deadline. The original bank remains failed:
one passed case, one failed case and ten not-run cases. Word copy coverage and
application event effects are not qualified by this run.

A separate recovery declined the warning for the exact owned copy, observed the
terminal uncertain result without replay, and closed owned Word normally. Its
recovery process handle observed exit code zero; this does not replace the
missing original qualification exit receipt. An initial COM restoration check
outside the native actor saw development registration and refused restoration.
A separate check in the original limited actor found the exact applied snapshot;
restoration there verified the complete prior HKCU state. The different contexts'
registry visibility was observed, but its underlying mechanism is not established.
The six test-owned preference baselines had been held only in worker memory and
were unavailable after its exit; their original restoration must not be claimed.

The next worker persists those six nonsecret values before writing and separates
preference recovery from COM restoration failure. `completion-native9` failed
registration preflight before opening any host; all ten remaining native cases
were not run. Its durable preference baseline and all six restored fields were
verified. Evidence for the successful same-actor registration recovery is in
`native8-registration-recovery`. These setup/recovery results do not qualify any
additional native scenario.

The remaining ten-case continuation (`completion-native10`, same product MVID)
failed its first PowerPoint case while opening the copy: the native macro security
warning appeared before the coverage plan could be written. The preceding
registered batch and single assertions passed, but the complete case did not.
The result is zero passed cases, one failed case and nine not-run cases. Recovery
requested one cooperative Stop, declined the warning for the owned copy and
observed terminal cancellation without replay. Normal Quit returned but did not
produce an observed process exit within the recovery deadline. The exact owned
process was then forcibly retired as separate cleanup; this is not normal host
lifecycle acceptance. COM registration and the six test-owned fields were verified
against the durable baseline in the original limited actor; no preference save
was needed because those six baseline values already matched. Word and PowerPoint
copy coverage still require a prepared native consent scenario. Remaining Office
hosts and the current SOLIDWORKS candidate are not qualified by this continuation.

The next managed candidate has MVID `eb89355b-8e06-4fe3-8a95-d421ce38b0e6`,
SHA-256 `835922677F8026EDA8017992050CA648B762480A9A26CE31FE6A0C230DCF4490`.
Its affected gate passed 1,424 tests with no failures or skips
(`completion-managed12/managed12.trx`), including grouped Options application and
restoration, native-theme controls and PowerPoint acquisition ownership.
PowerPoint returned-value targets and coverage copies now balance explicitly
acquired application, collection, enumerator, presentation and project references
on their owning thread. Borrowed project arguments are not released again.
Known preflight refusals release their acquisitions; uncertain invocation, copy
preparation or close retains references without repeating the native mutation.
Tests also cover release failures and service dispatch cleanup.

The first complete managed gate failed one new ownership assertion because the
test counted its own collection getter as a production acquisition. The corrected
aggregate passed; the original failure is retained in
`completion-managed11-complete/managed11-complete.trx`.

The same product candidate then passed 1,473 affected managed tests with no
failures or skips (`completion-managed13/managed13.trx`). This gate adds the
file-specific synthetic-copy consent and observer acquisition/recovery matrices;
it launches no native hosts. The subsequent twelve-case real-desktop bank
(`completion-native11`) failed its first Word case, leaving eleven cases not run.
Scoped consent returned once without a trust-policy change. The copy reported
complete procedure coverage of 50% (one of two procedures), and source, counters
and copy-absence assertions passed before the event check. The event oracle failed:
no opening callback was recorded, and closing arrived on an RPC thread instead
of the observer's STA. These partial results do not qualify the complete case.
The preservation receipt was not yet durable at that assertion boundary; a later
test-fixture change moves it before the independent event oracle.

Word and the original worker exited normally through retained original handles;
COM registration and all six test-owned preferences were restored and verified
in the original limited actor. No native mutation was replayed. PowerPoint normal
exit, remaining Office cases and the current SOLIDWORKS candidate remain open.
These are affected test gates, not whole-solution line or branch measurements.

The next test-fixture revision passed all 1,485 affected managed cases with no
failures or skips (`completion-managed14-complete/managed14-complete.trx`), with
the product MVID/SHA unchanged. The observer now synchronously processes borrowed
event paths on its original STA and records the RPC delivery thread separately.
Synthetic callback tests check return timing, cancellation preservation and
failure propagation without Office activation. The first aggregate failed one
new assertion expecting an exception wrapper; WinForms returns the original
exception. That failed aggregate remains in `completion-managed14/managed14.trx`.
Its fresh native bank (`completion-native12`) verified the synchronous Word
closing callback on the original STA, including the copy path, unchanged
cancellation and separate RPC delivery thread. Complete copy coverage and durable
original-source/counter/document-absence preservation also passed. The opening
callback remained absent, so the complete twelve-case bank still reports zero
passed cases, one failed case and eleven not-run cases. Word and the original
worker exited normally; same-actor COM/preferences restoration and all frozen
input hashes were verified. No uncertain mutation was replayed. Missing Word
opening events and the remaining native acceptance items are still open.

The macro-free Word lifecycle control fixture passed all 1,554 affected managed
cases without failures or skips (`completion-managed16/managed16.trx`), with the
product identity unchanged. Its fresh bank (`completion-native13`) refused startup
because the generated template added another non-Normal project. No lifecycle
control or test macro ran: the bank records zero passed, one failed and eleven
not-run cases. Separate recovery closed the exact owned Word normally and verified
COM and all six preference restorations; that recovery handle does not replace
the missing original qualification exit receipt. Frozen inputs remained unchanged.
The fixture now selects Word's canonical document `HostPath` rather than guessing
from a sole non-Normal project or Word's temporary VBA backing-file path.
The complete affected gate for that selection correction passed 1,566 tests with
no failures or skips (`completion-managed17/managed17.trx`). This remains a managed
regression result, not native lifecycle acceptance or a coverage measurement.
The next native bank (`completion-native14`) verified canonical Word selection
and reached the copy prerequisite after batch/single execution. Its diagnostic
source fingerprint failed because the new helper denied Word's existing write
handle. No lifecycle control opened. The complete bank remains zero passed,
one failed and eleven not-run. Word and the original worker exited normally;
COM/preferences restoration and frozen hashes were verified. The fingerprint
reader now permits the existing write handle and still checks the source again
after each control; this fixture correction requires its own managed gate.
That complete affected gate passed 1,570 tests without failures or skips
(`completion-managed18/managed18.trx`), including a real held write handle,
later-content-change detection and denied-sharing refusal without Office.
Its native continuation (`completion-native15`) completed the hidden and visible
macro-free controls under ForceDisable with suppression enabled. Both recorded
only the closing event, preserved the source bytes and restored their application
policy. The first unsuppressed control refused before mutation because Word's
loaded-template inventory contained personal `Normal.dotm`; the fourth control
was not run. Visibility alone does not explain the missing opening event in these
controls, and suppression causality is not established. The full bank remains
zero passed, one failed and eleven not-run. Original Word and worker exits,
COM/preferences restoration, task removal and unchanged frozen inputs were all
verified. Safe template isolation is required before another unsuppressed trial.
The next fixture admits Word's own startup Normal object only after independent
metadata checks prove its exact COM identity, design mode, no protection and one
empty document class. A name/path match alone is insufficient; no code text or
personal template file is read directly. The scope is rechecked immediately before
opening. Its complete affected managed gate passed 1,592 tests with no failures or
skips (`completion-managed19/managed19.trx`). Native causality remains unproved.
The fresh bank (`completion-native16`) completed both suppressed macro-free
controls, but refused the unsuppressed control before its flag change: the
generated DOTX structure reader denied Word's live write handle. Its result
remains zero passed, one failed and eleven not-run; original exits, restoration
and frozen inputs were verified. The fixture now reads both generated package
structures through one bounded, hash-verified immutable snapshot reader and
records partial template metadata before later refusals. It does not alter the
missing-opening-event oracle.
The complete affected gate for this package-reader correction passed 1,607 tests
without failures or skips (`completion-managed20/managed20.trx`), including live
write sharing, immutable snapshots, changed hashes, denied access, invalid ZIPs,
size bounds and inspector failures. This is still managed evidence only.
The real four-control matrix in `completion-native17` then verified the missing
event's cause for these macro-free controls: suppressed hidden/visible openings
recorded only BeforeClose, while both unsuppressed controls recorded Open then
BeforeClose. Native receipts independently proved the exact Normal object empty,
with no code-text or direct template-file read. All control policies were restored.
The actual macro copy still ran with suppression enabled and failed its unchanged
opening-event oracle, despite complete coverage and original preservation. The
full bank remains zero passed, one failed and eleven not-run; original exits,
restoration and frozen inputs were verified. A separate one-shot event scope now
uses that reviewed template boundary for the real copy and restores suppression
before normal cleanup. Its managed/native qualification remains pending.
The latest affected managed gate passed 3,047 tests with no failures and one
conditional local-provider skip (`completion-managed89/managed89.trx`), including every
setup/restoration boundary, failed reference release, scope loss, the final
post-activation template/security review, uncertain copy retention, observer
parent lifetime, observation deadlines, fresh UI Automation snapshots, bounded
native selector readback, read-only transcript recycling and stable native
file-dialog admission, explicit Word shutdown-observer admission, cleanup
receipt validation, the one-shot ordinary VBE-close rejection matrix and actual
compilation of its helper at the Windows PowerShell C#5 language boundary.
This gate uses `completion-build88`, product MVID
`5aa85cc8-9d53-45d7-adcf-7799ed778871`, SHA-256
`F17961F863E51B4CF9232580B9997D60997BCFFD6BD32676620B0172832A0E7C`.
The HTTP deadline diagnostic, activity rendering and Ollama reasoning-selection fixes change the candidate identity;
earlier native passes are historical evidence. The test assembly SHA-256 is
`32D671E5FD5D641C210468A3A7F200A29B374AAD83E666390CB29092A246B68E`.
The deadline matrix covers expiry at headers and streaming/buffered body reads,
client lifetime cancellation, coinciding Stop/deadline signals, the next request,
unrelated transport cancellation and protocol errors. The 120-second production
limit is unchanged, and no automatic request retry was added. The diagnostic is
manually translated in every interface language. The activity tests verify that
six hundred fragments across two unrealized rows cause one group publication
at the frame boundary, preserve final metadata, and cannot reinsert invalidated
rows. Existing realized-row selection and disclosure tests also pass. These
checks establish bounded collection updates, not native frame rate or absence
of flicker in every host. The earlier managed75/managed79 failures remain retained;
the preceding corrected complete aggregate was managed83, not a filtered passing rerun.
Managed89 additionally verifies Ollama capability declarations, positive admission
of default/off controls, malformed/unsupported selections rejected before transport
construction, immutable streaming/JSON request selection, per-model and conversation
restoration, unchanged tool-round selection and qualification-profile isolation.
When a catalogue replaces an unavailable model, its previous model's effort is
discarded before restoring the newly selected model's preference. The first
managed86 aggregate failed this restoration check and an existing resume fixture
using an unsupported synthetic Ollama effort. Both failures remain retained; the
complete managed89 aggregate passed after the restoration correction and use of
the provider's actual default option in that fixture, with its resume/disposal
oracles unchanged. No current native acceptance follows from the managed result.
The full affected gate includes the chat-state regressions. Its eight new polling
scenarios verify reentry rejection, suspended timer ownership, recovery after
exceptions, scope/runtime deferral, disposal during a read, and busy/hidden
transitions before another host read. They also verify that closed project scopes
are reconciled before project-specific collection acquisition, that unavailable
catalogues defer those reads, and that events during acquisition are reconciled
before the next collection is requested. This ordering correction has passed the
managed gate; its native replay remains pending. The conditional skip is
`LocalOllamaShownChatStreamsStopsAndCompletesNextSend`; real provider checks belong
to the separately opted-in native bank. Managed64 and managed65 failed before
executing tests. The second bootstrap's log records a denied local socket;
managed66, managed68, managed70, managed83, managed86 and managed89 ran outside the restricted sandbox, without changing
firewall rules.
The added observer regressions cover native control-class filtering, hidden
popup settlement without repeated expansion, original popup identity and
immediate rejection of foreign ownership. The preceding managed54 run failed
at the VSTest connection boundary before executing tests; it remains a failed
bootstrap record and is not part of this passing aggregate.
The final selector regressions also verify one accessibility focus before one
expansion, owner revalidation between both actions, refusal before effects,
intent publication failures and uncertain-action retention without a retry. These
are helper tests; the separately frozen Word follow-up failed as recorded below.
The latest gate also covers the complete composer write/readback refusal matrix,
errors after application, owner loss at every boundary, text-free diagnostics,
intent publication and panel-specific native filtering. The managed59 runner
failed to connect to testhost within its unchanged startup deadline and executed
zero tests. It remains a failed bootstrap record; managed60 ran the entire
affected scope again with diagnostic logging and local runner communication
available. The latest regressions additionally verify an explicit Word-only
testhost collection after returned Quit and external-reference release: scope is
captured before shutdown, uncertain Quit/release/execution prevents collection,
failure is retained without replay, and collection time consumes the original
exit-observation deadline. The separately frozen native40 comparison failed as
recorded below. The complete native41 Office bank uses the changed production
candidate with both explicit Word GC diagnostics disabled. No additional native
acceptance is inferred from the managed gate.
Its native-resource scenarios additionally exercise delayed positive completion,
preflight refusal, dispatched uncertainty, timeout, disconnection, break mode,
failed release, stale preparation and wrong-thread disposal/retention. The prior
failed aggregate remains in `completion-managed51`; its combined refusal and
release diagnostic was corrected before the complete scope was run again.
The DirectUI save-question regression matrix additionally covers immutable
full-path bindings, physical and accessible command agreement, foreign or
incomplete dialogs, ownership loss, uncertain delivery and two captured questions
that can advance only after the addressed original dialog disappears. It also
includes the observed Windows title-bar inventory and normal-owner-exit proof
for the last queued question; neither can issue another command. Admission
additionally requires three complete reads and a monotonic quiet interval;
changed inventories reset that interval, including delayed questions after an
apparently empty close boundary. Its
read-only observer protocol now requires matching parent/host identities,
candidate hashes, a complete primitive response and exit through the retained
creation handle before native input. Temporarily unavailable observation cannot
advance a pending response or resend it; it breaks the admission quiet interval.
These contracts are covered by the affected gate. The
native close/reopen pair is reviewed separately below; managed tests alone are
not evidence that those native questions have been handled successfully.
The same complete gate first failed in `completion-managed42` because the
restricted process denied temporary-file publication; that failed aggregate is
retained. The successful run used the authorized Windows context and the same
assemblies, without changing assertions or activating native hosts.
Its source is base `213231485c3bba0251357357a09392b750b8c24f`
plus the working-tree product/fixture changes recorded in the frozen gate. It does
not by itself qualify a native copy.
The corrected Word scenario passed in `completion-native18`: the same copy
produced Open then BeforeClose on the owning observer thread, procedure-entry
coverage was complete, original source/counters were preserved, suppression was
restored once and the original Word process exited normally. The next PowerPoint
scenario failed before copying: its counter reader supplied a borrowed application
to a resolver that owns and releases each acquisition. That fixture now preserves
the native resolver and disposes its returned target, retaining uncertain native
dispatches. This PowerPoint correction still requires native verification. The
bank remains one passed, one failed and ten not-run; both original host exits,
registration/preferences restoration and frozen inputs were verified.
In `completion-native19`, the corrected PowerPoint case and all five installed
Test Explorer cases passed, as did the Excel assistant case. Word's assistant
also completed streaming, cancellation, recovery and native marker/source checks,
but its original process did not exit within the measured shutdown wait after
`Quit` returned. Its complete case therefore remains failed: six passed, one
failed and five not-run. No second `Quit` or forced termination was performed.
After no hosts remained, a separate same-user Limited recovery restored the exact
COM overlay and six preferences. Both original worker exits and task removals
were verified, while Word's original host exit remains unobserved. UIA client
lifetime is under investigation; the functional evidence does not replace the
failed lifecycle oracle.

The subsequent Word observer isolation bank (`completion-native20`) expired
before the observer finished: one failed case and five not-run cases. The
observer later completed all four functional oracles, but the original parent
could not verify its creation-handle exit or complete native teardown. These
late results do not turn the case into a pass. Separate recovery verified the
unchanged marker, empty document class and sole disposable document, then closed
Word once and observed normal exit through a recovery query handle. The COM
overlay and six preferences were restored; both scheduled tasks were removed.
The final managed guards reject observations that return after their deadline
and new UI actions after the original parent exits. The test observer now fetches
control properties together in a fresh element-scoped UI Automation snapshot
for each inspection and rechecks ownership after the read. The Send readiness
check uses one snapshot per poll instead of several reads of the same button.

The fresh-snapshot bank (`completion-native21`) failed before any prompt was
written: the model-catalogue observation exceeded its deadline. The original
observer's creation-handle exit was verified, but Word rejected Document.Close
with RPC_E_CALL_REJECTED before the fixture entered Quit. The bank remains one
failed and five not-run cases. A separate source-verified recovery used one
application Quit without repeating Document.Close and observed normal exit
through a recovery query handle. The scoped COM overlay and six preferences
were restored and both owned tasks were removed.

A read-only probe of that exact disposable assistant found native selection
reads of 0.2032 ms for Ollama and 0.8383 ms for the selected model. These are
control-read timings, not chat rendering or provider-response benchmarks.
The updated observer binds fixed controls through their actual native HWNDs,
reads drop-down selection with bounded native messages and restricts composer
searches to the original prompt host. All refusal scenarios were implemented
before running the affected managed matrix. The next bank (`completion-native22`) passed Word's catalogue deadline and
observed visible streaming, one Stop, terminal idle UI, original observer exit
and normal Word exit through its retained fixture handle. It failed while
reading the cancellation acknowledgement: a transient RichEdit child disappeared
between the existence and ownership reads. A separate typed-variable collision
in the bank's failed-case receipt reader stopped the later cases, despite the
verified normal exit. Its result stays one failed and five not-run cases; COM
and all six preferences were restored and the owned task was removed.

The latest observer discards missing or detached read-only rows from an
observation, while still rejecting foreign ownership and changed fixed controls.
The remaining observation deadline bounds further reads; no native action is
replayed. The bank template now separates its XML test receipt from its JSON
host-lifecycle receipt. These corrections passed the prepared affected managed matrix.

In `completion-native23`, Word passed visible streaming, cancellation
acknowledgement, the next Send and the unprompted native-module marker. Original
source/references and the observer's creation-handle exit were verified. Word's
single Quit returned, but its original fixture handle did not observe exit in
the configured 5,000 ms window. The complete case remains failed: one failed
and five not-run cases. Later absence is not an original exit receipt. After no
hosts remained, a separate same-user Limited worker restored the exact COM
overlay and six preferences; both original worker exits and owned task removals
were verified. No second Quit was issued.

The separate `completion-native24` bank passed all five remaining Office cases:
PowerPoint, Access, Publisher and Outlook assistant workflows, and the installed
Outlook test explorer. Each assistant case verified streaming, cancellation,
next Send, native marker readback and unchanged source/references. Original host
shutdown receipts, worker exit, six preference restorations, COM restoration and
owned task removal were independently checked. All 1,457 frozen inputs remained
unchanged. These results do not replace Word's failed shutdown verdict.

SOLIDWORKS preparation initially stopped before any file mutation while Windows
created, replaced or temporarily exposed multiple save dialogs. Those explorer
cases remain not run; normal owned-process exits and scoped restoration were
verified. The reusable file-dialog helper now requires two ready matching reads
before admission and retains the exact binding for subsequent deliveries.

`solidworks-explorer-main4` passed the complete installed explorer test on
SOLIDWORKS 2025: the deliberate four-outcome batch (one pass, two assertion
failures, one runtime error), the passing single test, human/machine reports,
native docked window, stale-revision rejection and exact source/reference
restoration. The file baseline was unchanged at the test's cleanup boundary.
Original SOLIDWORKS, Visual Studio and worker exits were normal; COM restoration
and task removal were verified. All 1,471 frozen inputs remained unchanged.

End-to-end disk preservation is still unqualified: after host exit, the SWP's
bytes differed from its verified pre-exit baseline. Offline comparison found
the same module source but changes in internal project/compiler streams; this
does not prove the complete cause or byte-exact preservation. Two earlier
PowerShell preparation hashes were unavailable because of the open file's
sharing mode; the native test used a shared-read stream and independently
verified its baseline and backup. These limits are retained in the terminal
review and are not replaced by retrospective hashes.

The later Word diagnostic (`completion-native25`) again passed all four assistant
oracles but failed its original shutdown deadline. Its trace request was rejected
because the product diagnostic only admitted Excel; the load log independently
records Word's OnDisconnection. Missing diagnostic events do not establish that
cleanup was not invoked. The failed original verdict is retained; scoped COM and
six preferences were restored and both owned tasks were removed.

The diagnostic now also admits exact WINWORD process names while preserving the
request's process generation, loaded product, hash and owner-thread checks.
This changes the product identity to MVID `c034068b-9a54-40db-81f1-5a0c7ee6ccf8`,
SHA-256 `FBA7C1F703D36FBF0CB8DE5767E5C2B0B1E124F0967150D4B0C0344CF7BF9B05`.
Earlier native results remain tied to their previous candidate.

In `completion-native26`, the corrected diagnostic recorded returned
OnBeginShutdown/OnDisconnection callbacks and their returned Dispose calls, with
no recorded cleanup fault. Its first event arrived 6.4 seconds after Quit;
the observed cleanup sequence then took under one second. The independently
validated late receipt does not change the original five-second exit failure or
prove the cause of the delay. All assistant oracles passed; COM and six preferences
were restored, owned tasks removed, and hosts released without replay or force.

The isolated transcript workload (`completion-performance32/performance32.trx`)
passed all 27 cases on the preceding product candidate: streamed text, retained
activities, histories up to 2,000 entries, SQLite save/reopen and idle layout.
For 131,072 characters at 340-pixel width, incremental updates took 137 ms total
and 3.5 ms p95 versus 618 ms and 19.5 ms for complete replacement. Expanded
activity groups with 100 siblings at that width recorded 29.2 ms p95 and 35.1 ms
maximum. Background compilation overlapped this measurement; these local timings
are not controlled native-host frame-rate or real-DPI acceptance.

The host catalogue now balances acquired Word window and document collection
references through bounded, one-based traversal instead of leaving implicit COM
enumerators alive. The product candidate is MVID
`cec1d7e2-5c95-4643-9d29-6204e4ac154e`, SHA-256
`CA3ED8AAEE85D3B77365A735A702EBB238F94EB5F6E0EAD50922FBCEFFA73213`.
In `completion-native27`, Word's assistant passed all four functional oracles,
preserved source and references, and exited normally through the original retained
handle in 3,326 ms within the unchanged 5,000 ms bound.

The subsequent Word explorer scenario failed during cleanup. Windows recorded an
access violation in `VBE7.DLL` for the original Word process; its crash reporter
launched an empty recovery instance. This does not establish the cause of the
crash. The original verdict remains one passed, one failed and ten not run.
The empty recovery instance was closed normally once, its separate query handle
observed exit code zero, and scoped COM registration plus six preferences were
restored. All five campaign tasks were removed; the independent terminal review
verified 1,520 frozen inputs. Recovery does not qualify the original failed case.

In `completion-native28`, the fixture replaced final RCW release with balanced
getter-reference release. The five explorer cases passed independently in Word,
PowerPoint, Excel, Access and Publisher. Word and PowerPoint preserved the original
document and independently measured procedure-entry counters; their copies
produced Open then BeforeClose on the owning thread. Original normal exits were
verified, including Word's unchanged five-second bound. This result qualifies
those cases; it does not establish the cause of the preceding Word crash.

Excel's subsequent assistant case failed while observing the Default window
inventory after Send. The uncertain turn was retained without another Send/Stop,
and the six later cases were not run. A separate read-only recovery proved idle
UI, unchanged source/references and the sole disposable workbook before one
normal close and exit-code-zero observation through a recovery query handle.
That recovery does not qualify the failed case. The same-user actor restored COM
and six preferences, all three tasks were removed, and the independent terminal
review verified 1,522 frozen inputs. The bank remains five passed, one failed and
six not run.

The subsequent main-desktop inventory correction permits one fresh read-only
enumeration after a structurally valid partial observation with no native error.
It does not retry malformed inventories, lost ownership or native mutations.
In `completion-native29`, the Excel and Word assistants and the Word and
PowerPoint explorers passed. Excel's explorer failed a separate cleanup oracle:
the copy had left Workbooks, but a pathless project remained in VBProjects.
The original Excel process nevertheless exited normally with code zero. The
bank restored COM registration and all six preferences and its task was removed.

`completion-native31` executed the seven never-run cases from that same frozen
matrix, without replaying preceding scenarios. Both Access and Publisher explorers,
all three remaining document-host assistants, and both classic Outlook scenarios
passed. Original normal host exits, unchanged source/references, restoration and
task removal were independently checked. The combined twelve-case result remains
**eleven passed and one failed**, with the Excel ghost-project failure preserved.
The intermediate preparation attempt `completion-native30` stopped before any
host or test launch because it reused a diagnostic nonce; its configuration was
restored before the corrected fresh preparation.

The next working-tree candidate, MVID `d72c14c6-78bb-486c-baa7-5ff9956f00c7`
and SHA-256
`E5AF9EF4C5BBE7B9A83CC06C48C1AA791EDD27C3EA6BA607712CB66CCDFDB126`,
suspends passive Monaco navigation throughout
owned coverage-copy preparation, compilation, execution and cleanup. This prevents
automatic following from acquiring temporary runtime modules while native calls
pump messages. Its managed tests exercise nested scopes, refusal, cancellation,
uncertain execution, close errors and a service disposed before its queued run
starts. They passed in the complete affected gate; the earlier failed aggregate
is retained. In `completion-native32`, all three affected copy scenarios passed
on that exact candidate: Excel, Word and PowerPoint. Excel's copy was absent
from both Workbooks and VBProjects, with only the unchanged original present.
Word and PowerPoint preserved their originals and independently verified the
copy events and measured procedure counters. Original normal exits, COM and six
preference restorations, task removal and 1,522 frozen inputs were independently
verified. This qualifies those copy workflows; it does not establish the sole
cause of Excel's preceding intermittent ghost project or qualify an entire host.

The first SOLIDWORKS storage bank stopped before host launch or COM registration:
the newly added helper compiled under the managed project's C# version but
Windows PowerShell rejected its throw expressions. Both original worker exits
and owned task removal were verified; all native cases remained not run. The
helper now uses compatible syntax, with an actual compiler regression check.

The corrected five-case SOLIDWORKS storage matrix was prepared before execution
on the same product candidate. Both installed explorer cases passed their test
assertions and source/reference restoration. The first three storage cases
failed raw-byte preservation after normal exit, including both blank controls
that executed no macro. Offline comparison found identical stream payloads in
those blank files but changed container bytes; the explorer case changed five
internal streams while preserving both original modules' compressed source.
These observations do not qualify byte preservation.

The fourth storage case stopped on a DirectUI save question unsupported by the
legacy prompt reader. Read-only UI Automation identified two questions naming
the exact disposable macro. Separate cleanup discarded each once, verified the
original SOLIDWORKS/Visual Studio exits and restored COM; it did not turn the
failed close into a pass. The fifth, native reopen case was not run. Independent
review therefore records zero storage passes, four failures and one not-run
case in `solidworks-storage-main2/root-terminal-review.json`. All frozen inputs
were unchanged, the owned task was removed and every host was released. Raw-byte
failures remain intact; no original file was overwritten from its backup.

The first affected close/reopen pair then confirmed ordinary VBE closure, but
the same save questions appeared later during SOLIDWORKS `ExitApp`. That shutdown
boundary still used the legacy reader and stopped. Read-only inspection also
identified Windows nonclient title-bar controls and a UI Automation Value pattern
on the instruction icon, explaining a refusal before any new-controller input.
Separate cleanup discarded each captured question once without repeating the
already-entered `ExitApp`; original host/IDE exits, scoped COM restoration, all
four task removals and unchanged frozen inputs were independently checked.
The cleanup worker itself exited with an error, so recovery is not qualified.
The pair remains one failure and one not-run, despite passing installed explorer
assertions. The raw file still changed. The prepared replacement uses the same
guarded reader at both shutdown boundaries and retains the original raw-byte
and reopen oracles.

The following pair retained a preflight failure when a second fully owned save
question appeared between binding and delivery. No original discard was sent,
the installed explorer assertions passed, and the reopen remained not-run.
Separate cleanup bound both current questions; original normal host/IDE exits,
COM restoration and task removal were verified, with the original failure intact.
The cleanup worker exceeded its own observation bound and failed; its result is
not promoted. The next prepared pair adds a quiet interval both before binding
questions and before accepting an empty ordinary VBE-close boundary.

That quiet-interval pair captured both original questions before input, then
failed a bounded native-text read while observing the first delivered response.
The original first response was never repeated. Separate cleanup independently
matched only the remaining second binding, answered it once, and requested one
normal `ExitApp`, which the original bank had not entered. The cleanup observer
exceeded its deadline and failed; original bank handles subsequently proved
normal host/IDE exits and COM restoration. Both owned tasks were removed and
all frozen inputs were checked before later documentation edits. The installed
explorer assertions passed; the close/reopen pair remains one failure and one
not-run. Read observation during control destruction and observation-client
lifetime require correction; no native preservation acceptance is claimed.

The next pair (`solidworks-storage-main6`) performed accessibility acquisition
in separate owned processes, which exited before their snapshots were accepted.
Both cases completed ordinary VBE closure and normal original-handle
SOLIDWORKS/Visual Studio exits, with scoped COM restoration. The installed
explorer assertions passed. All frozen inputs were unchanged and the owned task
was removed. This supports the declared shutdown workflow without establishing
the cause of earlier observation-client failures.

Both preservation cases remain failed. The first exit changed the raw macro
bytes; reopening that exact file preserved its post-exit bytes, sources and
references but could not satisfy the original-byte oracle. The full project
descriptor comparison also failed: its only differences were process-generated
enum type names for Mode, Protection and Type; their values were unchanged.
This diagnostic does not replace the frozen property-comparison verdict.
Two reader workers failed during original host exit; their errors and original
creation-handle exits remain recorded separately. Neither failed read was
accepted as a native snapshot. No action was replayed and no backup was used to
overwrite the original macro. The original strict reviewer failure is retained
alongside the supplemental terminal review.

Offline comparison of the first case found an exact unchanged file prefix:
the exit removed 5,632 trailing bytes, corresponding to eleven sectors marked
free in the compound-file allocation table. Every stream payload and every
directory entry was identical. This distinguishes the observed container-size
change from a source or reference change; it neither establishes the complete
cause of truncation nor replaces byte-exact preservation acceptance.

The stress matrix runs three coverage copies successively within each owned
Excel, Word and PowerPoint session. Each cycle independently verifies measured
counters, original source/revision/disk, document and VBE-project detachment,
and a distinct run or retained copy. Word/PowerPoint event assertions apply to
every copy. In `completion-native33`, Excel passed all three cycles and its
original normal exit. Word passed its first copy but failed before the next
scope's native input because the fixture reused an immutable event receipt;
its original normal exit was verified. PowerPoint passed its first copy, then
its second copy's consent hit the same receipt-name defect after dispatch.
That uncertain case retained the host without automatic Close/Quit or release.
The lot therefore remains one passed case and two failed cases.

Separate recovery disabled macros on that exact second copy once and verified
the recovery worker's creation-handle exit. Both owned presentations were then
closed and one recovery `Quit` returned, but the host remained alive. Closing
its exact empty VBE once failed with `0x80004005`; neither action was replayed.
The original byte baseline remained intact before these actions. With explicit
maintainer authorization, the exact owned PowerPoint generation was forcibly
closed; its saved original bytes were verified unchanged. A separate limited
worker restored the scoped HKCU COM registration and all six nonsecret settings,
with verified readback and an original creation-handle exit of zero. All six
owned campaign/recovery tasks were removed after their workers exited. Forced
cleanup is not native stress acceptance; the original failed lot remains intact.
The next fixture revision gives each independent consent/event
scope distinct immutable receipts, preserving the original first-scope names.
In the fresh `completion-native34` lot, Word and PowerPoint passed all three
copy cycles, including independent consent/event receipts, original preservation,
project detachment and original-handle normal exit. Excel failed during its
second cycle: the document collection contained only the original workbook, but
the VBE still listed the closed copy's project. Excel also exited normally. The
coordinator retained its failed aggregate verdict; registration and six settings
were restored, all frozen inputs were unchanged and the exact task was removed.
The receipt correction is therefore natively accepted for the declared Word and
PowerPoint scenarios, while repeated Excel copy detachment remains open.

The next candidate explicitly owns native preparation interfaces until positive
callback and design-mode completion. Refused undispatched attempts release their
acquisitions; uncertain dispatched attempts retain them. Pane preparation,
focus validation, native diagnostics and command lookup balance each getter
acquisition without finalizing borrowed editor/project objects. Preparation cannot
be reused after release or retention, or disposed from another thread.

The complete fresh `completion-native35` lot passed on that candidate in all
three hosts. Each owned Excel, Word and PowerPoint session independently ran
three copy measurements, preserved original source/revision/saved bytes and
counters, and verified document and VBE-project detachment after every cycle.
Word and PowerPoint also verified each copy's native Open/BeforeClose events
and independent consent/event receipts. All original host handles exited
normally; the original coordinator handle exited with zero. Scoped COM and all
six settings were restored, the exact task was removed, and independent review
verified all frozen inputs unchanged. The Excel ghost-project failure was not
reproduced in this declared lot; the original failing lot remains unchanged.
This accepts the declared repeated-copy scenarios, not the entire VBE product,
native appearance, signatures, every document type or SOLIDWORKS storage.
The separate `completion-native36` lot retained the complete remaining Office
matrix on the same product. Excel, PowerPoint, Access and Publisher assistant
scenarios and Access/Publisher registered explorer scenarios passed. Word failed
before sending because its initial native-control lookup exceeded the unchanged
observation deadline. Outlook also failed before sending: the exact owned popup
was still hidden immediately after its single expansion. Its explorer scenario
was not run. All hosts were released, registration and the six settings restored,
and the original coordinator exit and exact task removal verified. These failed
and unexecuted cases remain unchanged. The separately prepared
`completion-native37` follow-up passed both Outlook scenarios, including installed
explorer callbacks, streaming, cancellation and a subsequent Send. Word's model
lookup fell from 34 seconds to 89 milliseconds, but the scenario still failed
before Send: the original mode popup remained hidden after its one expansion.
The follow-up retained that failure and independently verified normal fixture
cleanup, scoped registration restoration and restoration of the six settings.
Its test assembly differs; product bytes are identical. Outlook's lifecycle
receipt does not publish the numeric original process handle, a retained evidence
limitation rather than inferred additional proof.
The separately frozen `completion-native38` Word follow-up passed the exact
visible mode-popup selection and independent selection readback. It then failed
before Send: after one composer write, the independently read value did not match
the request. No SetValue or Send was replayed. Panel discovery also took roughly
38–40 seconds. Word's sole Quit returned, but its original five-second exit
observation did not prove termination; later process absence is not normal-exit
proof. Recovery from the same limited Default actor independently restored the
scoped COM registration and all six preferences, verified both original worker
exits and removed both owned tasks. The original scenario and restoration
failures remain preserved. Recovery does not qualify the assistant workflow.
The next helper changes restrict panel lookup to every eligible WinForms Window
class, preserve the exact AutomationId/HWND uniqueness checks, bind composer
readback to its original UIA runtime identity and record only lengths, hashes and
API error metadata on divergent writes. The affected managed gate passed. In
the separately frozen `completion-native39` Word bank, all three one-shot
composer writes were independently read back exactly on the original editor.
Visible streaming, Stop acknowledgement, the next Send, the unprompted native
module marker and source/reference preservation passed. The original isolated
observer exit was verified. Panel discovery fell to 3.7 seconds for the
transcript and 16 seconds for the composer; those UI Automation timings still
require improvement and are not chat-rendering benchmarks. The entire native
case remains FAIL: Word's single Quit returned, but its original five-second
normal-exit observation failed. The optional shutdown trace was also absent;
that secondary diagnostic failure does not replace the original exit failure.
After later host absence, the original limited Default actor restored COM and
all six preferences; both exact tasks were removed. Frozen inputs remained
unchanged. Functional evidence now covers the twelve planned Office workflows
on identical product bytes with explicitly differing helper identities, but
Word lifecycle acceptance and release qualification remain open. Prior failed
runs are not promoted and the composer or exit oracle is not relaxed.
The subsequent `completion-native40` comparison passed its three-copy Word
control and original normal exit without post-Quit collection. The assistant
again passed the four visible chat oracles, but its final second `read_module`
exceeded the response deadline. Source/reference preservation was not proved,
the original Quit was never entered, and the complete assistant case remains
FAIL. A bounded thread snapshot of this exact retained process observed at
least four recursive context-timer callbacks on the VBE STA while the bridge
waited in `Control.Invoke`. Hiding the native chat HWND alone did not settle
the request. Separately closing that original chat control once destroyed it
and made the serial bridge listener available in 99 ms without repeating the
read. One first normal recovery Quit then produced exit code zero on a separate
recovery handle and preserved the saved disposable bytes. The original limited
Default actor restored COM and all six preferences; all six owned tasks were
removed. Recovery does not replace the failed original preservation/exit scope.

The production fix suspends the polling timer across COM reads and rejects
reentrant ticks before event-source acquisition. Each deferred read revalidates
busy, scope-loading, visibility and runtime/disposal state; the guard is released
after errors and the timer is not restarted after disposal. Its managed gate is
identified above. The separately frozen `completion-native41` bank prepares all
twelve Office assistant/explorer workflows on that changed candidate, including
three distinct coverage copies in each Excel/Word/PowerPoint explorer case. It
disables explicit pre/post-Quit testhost collection and retains the original
Word shutdown deadline. That bank ended with one PASS, one FAIL and ten NOT_RUN
rows. The three-copy Word control passed and its original handle observed normal
exit within the unchanged bound. The Word assistant passed all four visible
oracles, its original observer exit and exact source/reference preservation;
the previously blocked final read did not recur. Its single Quit returned, but
normal process exit was not observed within five seconds. Later host absence
is not a replacement exit proof. A separate original-actor recovery restored
COM and all six preferences and removed both tasks; the failed row and ten
unexecuted rows remain unchanged. Independent terminal review verified the
frozen inputs before the next correction.

The shutdown load trace still showed component/reference source resolution for
the already closed document before its dirty scope was reconciled. The next
polling revision refreshes a dirty scope before source acquisition, defers all
source reads after an unavailable catalogue, and reconciles events pumped by
project-source acquisition before reading component sources. Its compilation and
complete affected regression gate passed as identified above.

The complete `completion-native42` matrix then recorded ten PASS, one FAIL and
one NOT_RUN on product MVID `1eb93d77-b4c7-4b91-978a-bab2a8331625`. Both assistant
and explorer workflows passed in Excel, Word, PowerPoint, Access and Publisher.
The Word assistant preserved source/references, observed the original observer's
exit, and its original host handle observed exit code zero in 2,496 ms within the
unchanged five-second window. Explicit pre/post-Quit collection remained disabled.
Outlook streamed while busy, then failed a bounded native transcript read
(Win32 1460); its subsequent Quit raised RPC failure `0x800706BE`, so its normal
exit remains unproved despite later absence. Outlook explorer was not run. The
original coordinator exited with code one; COM and all six preferences were
restored and its exact task removed. The independent terminal review verified
all ten passing scopes and frozen inputs before further observer edits. The
supplemental reviewer handles Excel's distinct receipt format without changing
the frozen original reviewer. The original failed Outlook row is retained.

The following observer revision treats a timed-out read of a still-owned row as
an unavailable complete snapshot within the original oracle deadline. It returns
no partial text, rejects foreign ownership and other live-row errors, and issues
no new Send/Stop action. Its complete regression matrix passed in the affected
gate identified above, including fourteen timeout, expiry, recycling and foreign
ownership scenarios. The production assembly hash matches build67 exactly.
The separate `completion-native43` bank then passed Outlook's registered explorer
workflow, including original project restoration, verified normal exit and OTM
cleanup. Its chat case failed the unchanged 125-second visible-streaming-while-
busy oracle. A final read observed the assistant already idle; no second Send
or Stop was issued. Guarded project restoration and normal original-host exit
passed, so the explorer case could proceed. The earlier RPC failure did not recur;
its cause remains unestablished. Both original case verdicts remain retained.
The original coordinator exited with code one; COM, all six preferences and the
exact task cleanup were verified. Independent review checked the frozen inputs
and the passing explorer evidence. Together with native42's ten successful
cases, eleven of the twelve declared workflows have passing evidence on the
exact same product hash. This does not make either failed aggregate green or
qualify the outstanding Outlook chat workflow or the complete Office suite.

The diagnostic `completion-native44` run retained that chat failure and its
unchanged deadline. A read of the actual owned transcript found a task-cancellation
error, although the scenario issued no Stop. Original Outlook shutdown, project
restoration, COM registration, all six preferences and exact task cleanup were
verified. No native acceptance follows from successful cleanup.

The separate HTTP-only `completion-transport73` batch on the same production
assembly passed the synthetic tool roundtrip but failed cancellation-after-text:
the long request supplied public reasoning fragments without any text before
the 120-second deadline. The local wire capture contains only fixture-owned
synthetic input/output; its independent review verifies the original runner exit
and the absence of native tool dispatch. This reproduces the failure without
Outlook and provides no proof of an Outlook-specific cause. The active Ollama
server is version `0.32.0` in WSL; old Windows desktop logs were not used as current
server evidence. Its matching journal records cancellation when the client
closed the connection. The complete Office scope and SOLIDWORKS remain open.

The completed `completion-sampling84` comparison retained both failed long-response
cases on build82. With temperature 0/top-p 0.8, the model produced no content before
the deadline; with temperature 1/top-p 0.95, it produced no content and ended with
`length`. Both profiles completed their separate synthetic tool roundtrips. Changing
sampling alone did not resolve the long-response case.
On the current build88 candidate, `completion-sampling90` passed both unchanged
HTTP-only scenarios with the explicit reasoning-off selection: real-text streaming,
cancellation, the next request, and the synthetic tool/result roundtrip. The retained
sampling remained temperature 1/top-p 0.95, and the original 120-second deadline and
content/cancellation oracles were unchanged. The independent review checks the
exact model, DLL and `reasoning_effort: "none"` request payloads and no native tool
dispatch. This qualifies only that explicit HTTP profile, not default-thinking or
the Office UI. The full twelve-case `completion-native91` bank completed on this
candidate with an immutable matrix and original native oracles. Five explorer cases
passed (Word, PowerPoint, Access, Publisher and Outlook); Excel's explorer case and
all six assistant cases failed. Excel's third copy had left the workbook collection
but remained in the VBE collection with an unreadable host path. Word and Excel
failed before numbered text was visible while busy; later assistant cases failed
while awaiting the model catalogue. These failures remain failures.
The independent bounded-read review checked all 1,548 frozen inputs, original
normal process exits, scoped COM restoration, six scalar settings, the owned
model-specific reasoning preference and exact temporary-task removal. It did not
accept the full matrix. Later read-only provider diagnosis found Ubuntu stopped;
its retained kernel journal also recorded GPU allocation failures followed by a
`llama-server` abort at 22:05:57 CEST. This does not establish one cause for every
assistant failure. Windows resource-exhaustion events also recorded system commit
at 99.84% of its limit at that time. Nine orphaned Windows model workers from earlier
Q028 campaigns were identified independently and stopped with explicit maintainer
authorization. Their original handles confirmed exit; available commit increased
by 29.81 GiB. Neither Office nor the active WSL provider was terminated. Current
qualification helpers already contain new backend children in a private Windows
job; the historical parent-only cleanup is not the current implementation.
The next candidate balances unrelated-project acquisitions in
the shared selector, defers passive chat polling during native inspection, and
explicitly selects/read-backs the declared reasoning choice in native scenarios.
The complete Unit gate for this historical candidate passed, with the conditional
live-provider scenario skipped. Its product MVID is
`650a7862-a34a-4f97-9571-f1be520a47b3`, SHA-256
`14138DA0A90EC507E4CA9B7880D76A9BFFAC510A286B68785DE20CA88B051529`,
from source `213231485c3bba0251357357a09392b750b8c24f` plus the working-tree fixes.
The original VSTest exit and exact test assembly were independently reviewed in
`completion-managed93-connected`. An initial sandbox dispatch failed before
executing any tests because the test-host socket was refused; its zero-test result
remains retained. This gate is not a new line/branch coverage measurement. Current
native acceptance remains pending; no previous candidate's native passes carry
forward to it.

On that same build92 candidate, `completion-sampling94` passed both original HTTP
scenarios. The real-desktop `completion-native95` matrix then recorded two passes,
three failures and seven unexecuted cases. Word's explorer passed its three copy
cycles. Excel's assistant passed visible streaming while busy, one Stop, the next
reply and reading a marker held only in a real disposable module; outgoing request
metadata independently confirmed the exact model and reasoning-off choice.
Word's assistant failed before Send because its mode dropdown did not remain
visible after its single expansion. Excel still retained the third closed copy in
the VBE collection. PowerPoint's second copy returned an uncertain native callback
outcome and halted the remaining matrix. The maintainer authorized stopping only
that owned PowerPoint PID; this forced recovery does not qualify normal shutdown.

Independent terminal review verified all 1,549 frozen inputs and retained every
original verdict. Recovery under the original limited interactive actor restored
the scoped COM snapshot, six scalar preferences and the model-specific reasoning
preference; both exact temporary tasks were removed. An earlier tool-process
attempt observed different registry/settings visibility and refused to overwrite
that state. The provider metadata observer closed normally; Ollama was preserved.
The complete Office matrix and SOLIDWORKS acceptance remain open.

The subsequent build96 balances the shared project inventory's COM acquisitions
and releases the Excel copy's owned project acquisition before closing its
workbook. Native callback timeout messages now distinguish the observed project
mode from the missing correlated verdict; deadlines and no-replay checks are
unchanged. The Office fixture selects real WinForms choices through one native
selection and one parent notification, without opening a popup or taking focus.
Its complete Unit gate passed **6,704 tests, no failures and one conditional
live-provider skip** (`completion-managed97`). The product MVID is
`b59e8627-bf5c-488b-8d17-4768734385b9`, SHA-256
`B853B4E8610C26D974F0794D326B9EC043CF1029CEC9662AACBA96A9DED773D8`,
from the same source revision plus working-tree fixes. Original VSTest exit,
assembly identity and all new regression results were independently reviewed.
These results do not measure line/branch coverage. Both real HTTP scenarios
passed in `completion-sampling98`. The complete twelve-case real-desktop
`completion-native99` bank recorded eight passes and four failures, with no
unexecuted cases. All six native explorers passed; Excel, Word and PowerPoint
each preserved the original and detached three independently measured copies.
The Excel and Access chats passed streaming while busy, Stop, the next reply and
an unprompted native-module marker. All owned hosts exited normally; scoped COM
registration, six scalar preferences and the reasoning preference were restored,
the original worker exit was observed, and the exact task was removed. Independent
review verified all 1,551 frozen inputs before subsequent fixture edits.

Word's chat failed after the scope selection notification. The fixture's combined
identity/enabled guard also rejected the temporary disabled state used by
asynchronous scope loading. Its corrected actions still require enabled original
controls, while post-notification readback keeps the ownership/identity checks.
Publisher, Outlook and PowerPoint failed the expected numbered-stream response:
their final visible replies respectively discussed the request, refused the
unrelated task or echoed it. These failures do not establish a rendering fault.
They remain failed results for the original requests and model profile.

The subsequent fixture-only gate passed **187 tests without failures or skips**
(`completion-managed103-fixture`), including temporary disable after notification,
disable before each mutation and revised display/recovery requests within VBAi's
actual context. The production DLL and companions are byte-identical to build96;
only four test sources changed. `completion-native104-chat` exited before native
dispatch: its preflight incorrectly associated the original full Unit TRX with
the newer fixture assembly. No native scenario, preference change or registration
change started. The fresh `completion-native105-chat` plan checks each managed
gate against its own exact original assembly and retains all functional oracles.
Its six chats recorded three passes (Excel, Access and Publisher) and three
failures (Word, Outlook and PowerPoint). Every host passed visible streaming while
busy, one Stop and the completed next reply. The three remaining failures concern
the unprompted native-module marker: the model returned unfounded refusals rather
than requesting the read. These are failed model workflows, not evidence that
rendering or native module access failed. All hosts exited normally; scoped COM,
six preferences and the reasoning preference were restored, the original worker
exit was verified and the exact temporary task removed. Independent review
verified the 1,551 frozen inputs before further edits. Full Office and SOLIDWORKS
acceptance remains open; no installer or release acceptance is implied.

The synthetic HTTP context matrix `completion-host-context107` then exercised the
production developer instructions and core schemas with all six document types.
Model-default reasoning passed every row. Reasoning-off recorded four failures:
three first-response refusals and one catalog gateway call not handled by the
initial strict fixture. That last outcome is a fixture limitation, not a proven
invalid model call; the original failed verdict is retained. The fixture now
accepts the normal gateway only for the same exact `read_module` and arguments.
These HTTP checks perform no native dispatch.

The fresh model-default native bank `completion-native108-chat` passed every Word
chat stage, including the random native marker, and preserved source/references.
Word's original five-second exit oracle failed after its one returned `Quit`, so
the other five cases were not run. Later passive observation under the original
limited Default actor found no live hosts; it restored scoped COM and all owned
preferences without force or another Word command. Both exact tasks were removed.
This recovery does not qualify the original normal-exit deadline. Independent
review checked all 1,555 frozen inputs before subsequent ownership edits.

The next source correction balances native component inventories and module
reads, including their acquired project, collection, enumerator, components and
code modules. Detached metadata/code preserve the existing response fields and
SHA-256 semantics. This addresses demonstrable unbalanced acquisition paths;
causality for Word's delayed exit remains unproved until candidate-specific native
verification. The build110 candidate's complete `TestCategory=Unit` gate passed
6,728 tests with no failures and one conditional skip in `completion-managed111`.
The original TRX and runner exit were reviewed, including all 18 added module
ownership scenarios. The candidate MVID is
`a1a7f00a-edc9-418b-a1ba-fa86c0d7cab6`; its assembly SHA-256 is
`C94B2F7E418149A59EC989D97B52665B629F44ABE1F208F9C9BD8DC71BC1EF48`.
This complete managed gate does not qualify native host behavior or supply a new
line/branch coverage measurement.

The same-candidate `completion-http112` batch passed all six synthetic host
context reads and its streaming/tool round trip. Its text-only cancellation
scenario failed: with model-default reasoning, the request for 1,000 numbered
objects reached the provider's `length` terminal after 3,769 JSON chunks without
a text fragment. The production reader rejected the incomplete response; the
fixture never entered its fragment-triggered cancellation or recovery request.
The original batch remains seven passes and one failure. It does not qualify
cancellation during a reasoning-only stream or native chat rendering.

The corrected fixture-only `completion-http120` batch subsequently passed all
eight scenarios on the identical product assembly: cancellation on actual text
or published in-progress reasoning, a subsequent completed request, a tool
round trip and six synthetic host-context reads. The context fixture accepts
bounded distinct read-only preflights before one exact module read; the original
`completion-http117` Word failure was a valid `status` request rejected by the
older fixture. `completion-http119` could not connect to testhost and executed
no tests. Both original reports remain retained. These HTTP passes qualify the
declared model-default profile, not native host behavior.

The independently prepared `solidworks-storage-main109` bank started its exact
SOLIDWORKS 2025 instance through Visual Studio (`33.1.1`, PID 93776), created and
saved a disposable macro, then failed to settle its one ordinary VBE close.
Five remaining scenarios were not run. The original matrix is retained; a
supplemental terminal review checked all 1,596 frozen inputs, original worker
exits, read-only observer exits, scoped COM restoration and exact task removal.
The debugger utility exited normally, but the original host-exit receipt is
missing. No force or second native action was used, and the host's disappearance
does not establish its exit cause or qualify normal shutdown. The same-candidate
Office bank is independent and retains all its own functional/lifecycle oracles.

The complete `completion-native113` Office bank then passed all 12 planned cases
on build110: installed explorer callbacks and genuine Ollama chat in Excel,
Word, PowerPoint, Access, Publisher and Outlook. Each chat observed visible
numbered streaming while busy, one Stop and acknowledgement, a completed next
Send, an unprompted random marker read from its native module and unchanged
source/references. Word, Excel and PowerPoint each completed three distinct
coverage copies with independently checked procedure-entry counters, preserved
original source/bytes and detached copies. All owned hosts exited normally;
Word retained its original five-second deadline without testhost GC. The
independent terminal review checked all 1,551 frozen inputs, original worker
exit, scoped COM and six-field preference restoration, model reasoning preference
restoration and exact task removal. The owned HTTP observer closed normally.
These passes apply to this matrix, candidate and explicit model-default profile;
they do not erase the separate HTTP cancellation or SOLIDWORKS close failures,
establish complete host compatibility, or supply line/branch coverage.

The managed gate additionally verifies every event/consent receipt phase across
three independent scopes and rejects overwriting an earlier receipt, unknown
phases, path-bearing names, nonpositive sequences and relative evidence roots.

Wider coverage-copy detachment stress cases, SOLIDWORKS disk behavior after exit
and native acceptance of the corrected control lookup remain open. A complete
SOLIDWORKS matrix was prepared to separate initial native writer behavior from
opening a previously closed synthetic file; both retain strict raw-byte equality
and the earlier failed results. On the same product, `solidworks-storage-main114`
passed the saved blank/direct-exit control with unchanged bytes, source and
references. Its explorer control was blocked before opening the macro or
dispatching the native test. Separate normal exits of the exact owned SOLIDWORKS
and utility IDE processes were observed through their original handles;
registration restoration and removal of all owned tasks were verified. This
does not qualify the explorer case or resolve the failed VBE close. The narrower
welcome-screen admission gate `completion-welcome-gate122` passed 121 host-free
tests, including the Windows PowerShell compiler boundary; unknown modality
still blocks native commands. The earlier empty test filter and compiler failure
remain retained.

The new disposable explorer copy in `solidworks-storage-main123` passed its
installed native explorer assertions and preserved source/references before
exit. Both original host and utility IDE handles exited normally with code zero;
scoped COM restoration and task removal were verified. Its overall preservation
verdict remains **FAIL**: the SWP changed from 27,136 to 28,160 bytes after exit.
Offline comparison found identical decompressed module source and identical
decoded protection fields, but changes in host metadata, the embedded macro
path, project version and compiled streams. This does not establish the cause
of serialization or qualify exact file preservation. The native run encountered
no visible welcome dialog at file-command admission, so the new UI Automation
modality branch has managed evidence only. No ordinary VBE close was replayed.

The `solidworks-storage-main7` bank terminated
with one failed case and five unexecuted cases after one VBE close request: close
settlement was not proved. A separate read-only observer on
`WinSta0\Default` found no visible original VBE or SOLIDWORKS frame while the exact
owned process remained alive. The exact utility debugger intercepted an unhandled
exception on the CLR finalizer thread; the stack included PresentationFramework,
and its output recorded `0x8007000E` before a fatal native abort. One debugger
continuation stopped at the same fatal exception. Detaching only that debugger
allowed Windows to process the crash; the original host handle recorded
`0xC0000409`, not a normal exit. The original Visual Studio and worker exits,
verified scoped COM restoration and removal of all eleven owned temporary tasks
are retained in `solidworks-storage-main7/root-terminal-review.json`.
The post-crash macro was preserved separately; raw bytes differ from the baseline.
No close request was replayed and no forced termination was performed. Allocation
failure causality and a product fix remain unproved. Neither root disappearance
nor successful recovery qualifies VBE close or storage. No new line/branch coverage
or rendering acceptance is claimed.

### Earlier frozen audit candidates

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

## Git and transcript follow-up

The 2026-10-08 follow-up uses source base
`5eb46e9c0a5542a64b946825cc9133a4ae19e7bb` plus the Git reference-error and
incremental transcript changes. Its measured product has MVID
`cc95b84d-ee20-4624-a861-b1f64aa7271b` and SHA-256
`F7BAB60DFF1700D58EDCEF461F6679C43391A3F9700C6479ACA2EE62F6B6D55F`.

The combined Git/chat gate recorded **353 passed, no failures and one conditional
skip**. The skip requires the separate real-provider opt-in. After improving the
native transcript observer's failure diagnostics, its focused guards/profile/
oracle gate recorded **47 passed, no failures or skips**. A sandboxed attempt
failed to connect VSTest to its test host before producing test results; the
subsequent gate is a separate run. No new complete line/branch measurement was
collected. An earlier chat run had three fixture failures, corrected by registering
the realized history views; its original failed report is retained.

The earlier `final-workload-results.json` performance workloads used actual
native RichEdit controls outside Office. They
cover growing and rewritten text, selection/format preservation, viewport
changes, retained history and expanded/collapsed activity groups. Timings are
machine-specific historical observations, not acceptance thresholds or current
candidate results. The build193/Unit194 measurement is recorded above:

| Workload | Observed cost | Remaining limit |
| --- | --- | --- |
| 131,072 characters, 64 updates, 340-pixel width | Incremental publication: 146 ms total, 4.0 ms p95; complete replacement: 611 ms total, 16.8 ms p95 | A control benchmark does not measure provider latency or hosted WPF composition. |
| 100 completed activity siblings, expanded, 32 updates | Historical candidate: 6,723 ms total, 758 ms p95, 881 ms maximum | Superseded for the tested detached workload by the recheck above; real-host rendering remains a separate qualification. |

The retained local evidence includes `final-managed.trx`,
`native-observer-gate.trx` and `final-workload-results.json` under
`artifacts/excel-audit-followup-20261008/`. Git tests now distinguish a silent
missing revision from repository/reference read errors and prove that errors
stop admission before mutation. They do not establish the cause of the earlier
intermittent stale-revision result.

The real-desktop Ollama bank with `qwen2.5:3b` failed in both hosts. Excel observed
streaming, one Stop, cancellation and a completed next Send, then failed on a
bounded native transcript read. Word did not satisfy the visible-streaming
oracle. Excel's original-handle shutdown was verified. Word's cleanup did not
observe exit within its deadline; the process was absent at a later observation.
Scoped registration and preferences were restored. These partial results do not qualify the candidate's complete
provider workflow, and neither failure is attributed to model quality alone.

The independent `qwen3.5:9b` bank used the same product and the updated observer.
Excel lost the fixed Send-control identity after its first Send; the observer
retained that uncertain host without replaying an action. Word timed out while
waiting for the model catalogue, before sending a prompt. Both tests failed;
they do not provide a valid comparison of the two models' response quality.
Scoped registration and preferences were restored. Word's fixture issued Quit
but did not observe exit within its deadline. Separate recovery of Excel verified the original process and sole
disposable workbook before issuing one Close and one Quit; normal process exit
was not observed within the recovery deadline. This is cleanup evidence, not
successful native qualification. Later observation found neither owned process
running; this does not reconstruct their original-handle exit receipts or turn
either failed cleanup deadline into a pass.

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

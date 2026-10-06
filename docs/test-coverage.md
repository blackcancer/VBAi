# Recorded validation

[Documentation](README.md)

This is the evidence index for VBAi. Results identify the tested candidate and
operation; rebuilding or merging a candidate requires new applicable checks.
Use [compatibility](compatibility.md) to choose a host and
[testing](../tests/README.md) to reproduce the relevant test family.

## How to read results

| Result | Meaning |
| --- | --- |
| PASS / qualified scope | The declared oracle passed for the identified candidate and environment. |
| FAIL | An assertion or required lifecycle check failed. A later pass preserves this verdict. |
| BLOCKED | A prerequisite prevented acceptance. |
| NOT_RUN / skipped | The operation has no execution proof in that campaign. |

Managed coverage measures the instrumented managed assembly. JavaScript,
native renderer code, provider services and host processes have separate checks.
The qualification sections below are the retained current decisions; detailed
failed attempts and superseded counters are available in Git history.

## Q028 Ollama and embedded Office assistant qualification (2026-10-06)

**Qualified for the complete prepared Q028 matrix on the real
`WinSta0\Default` desktop.** Frozen source
`85486c00eb28ac148f00d7e8a4b73e4bcc9c5173`, product MVID
`7f766edc-1a6a-40a2-bf55-fbfc187fcf93`, SHA-256
`DC8094BC06F25784045FC85E1CE8968C38D48161260574F8C8AAB7C69AB0FCB8`.
The isolated solution build passed. Native acceptance belongs to those bytes;
a later merge or rebuild does not inherit it automatically.

The selected backend is Ollama **0.34.4**, `qwen2.5:7b-instruct`, manifest
SHA-256 `845DBDA0EA48ED749CAAFD9E6037047AA19ACFCFD82E704D7CA97D631A0B697E`,
CPU, context 8192, one parallel request, temperature 0 and top-p 0.8.
The existing model blobs were freshly size/hash checked without download or
substitution. Each selected Office executable is x64 **16.0.20430.20092**.

| Exact-candidate validation | Result |
| --- | --- |
| Focused managed transport, chat, privacy, host path, ownership and lifecycle gate | 714 passed, 0 failed, 0 skipped |
| Real provider prerequisites | 3 passed: exact synthetic tool roundtrip, cancellation/recovery, shown detached streaming chat |
| Real embedded assistant | 6 passed: Excel, Word, PowerPoint, Access, Publisher and classic Outlook |
| Ordered VSTest matrix | 723 passed, 0 failed, 0 skipped; every bank invoked once |
| Independent wire auditor self-tests | 16 passed; synthetic refusal checks, separate from VSTest totals |
| Frozen independent campaign review | `OFFLINE_NATIVE_WIRE_PASS` for every embedded host |
| Original native lifecycle | All six owned hosts closed normally; no forced Office termination or uncertain native replay |
| Coordinator and resource release | Original main worker exit code 0; temporary settings and HKCU Registry64 COM registration restored and verified; launcher task removed |
| Owned Ollama backend lifecycle | Private kernel job has zero remaining members after stopping its synthetic backend and calculation workers |
| Documentation validation | 40 maintained Markdown files, 329 local links, 0 errors |

Every embedded bank proves actual installed-candidate/owner identity, assistant
text while Stop is active, one Stop and visible cancellation, a complete next
reply, and an unprompted native marker returned by exactly one `read_module`.
The independent wire review binds the exact conversation selector, tool
arguments, native result/source hash and final answer. All project source and
references remain unchanged. The temporary Outlook fixture restores its original
empty project and leaves its initially absent OTM absent; no mail or VBA runs.

The campaign corrected Outlook's fake absolute `VBProject.FileName`: when no
verified existing `.otm` path is available, the resolver publishes an explicit
null host path and uses the temporary project identity rather than a path derived
from a process-dependent current directory. During unsettled turns the observer
retains original native window ownership instead of rediscovering an unavailable
UIA parent. It drops settled managed UIA references before shutdown. Word also
uses its scoped testhost collection diagnostic; the unchanged five-second
original-handle exit bound passes. This does not prove native RCW release or
identify the cause of every earlier exit delay.

Stopping only the isolated `ollama.exe` leaked multi-gigabyte `llama-server.exe`
workers in earlier batches. The corrected runner attaches the fresh backend to
an unnamed Windows job before model requests and verifies empty membership at
settled shutdown. Recovery closed only the exact descendants proven by this
session's backend receipts; older unknown workers and personal backends were
preserved. Automatic kill-on-close is disabled so native uncertainty retains
diagnostic state.

Retained proof root: `artifacts/q28p1/native8/`, including `q028-plan.json`,
all original TRX files, actual embedded HTTP/SSE wire, host/native shutdown
receipts, `independent-review.json`, `backend-job-exit.json`, original coordinator
handle exit and `qualification-acceptance.json`. Earlier `native2` through
`native7` outcomes remain failed or incomplete, including the first failed
independent Outlook selector audit and the pre-Office allocation failure. Their
source/MVID differs and their actions were not replayed to change their verdict.

This closes the selected Q028 provider/Office operation matrix. It does not
qualify other models/devices, all Office operations, SOLIDWORKS, Visio, Project,
production Outlook projects or full release acceptance. The historical
`No text response.` / complete-empty backend cause remains unresolved; the
passing selected profile is not a universal reliability correction. No coverage
percentage was measured by this campaign and no failed broad aggregate is
relabelled green.

## Q024 Word and Q027 UserForm scoped qualification (2026-10-06)

**Qualified for the complete prepared Q024/Q027 operation matrix on the real
`WinSta0\Default` desktop.** The frozen source is
`c593d6af08582d4facbc9d3bd2532550ea35cef9`, product MVID
`0f1e0112-39dd-4f8e-b3a3-44d45954bb9a`, SHA-256
`336F6F7CA0635216D4C2339CE0D8DC12A320B9A286146413760950F8CED41E66`.
The isolated solution build completed without warnings or errors. Qualification
belongs to these bytes; a later documentation commit, merge or rebuild does not
inherit native acceptance automatically.

| Exact-candidate validation | Result |
| --- | --- |
| Complete default managed suite | 5,617 passed, 0 failed, 234 skipped; both original workers exited normally and private desktops were released |
| Synthetic UI matrix | All nine prepared actions passed |
| Q024 real Word workflows | Three passed: canonical same-name document isolation/stale SaveAs refusal; installed owner capture/checkpoint/compare; actual Chat-Git entry and known modal closure |
| Q024 host lifecycle | All three original Word processes exited normally |
| Q027 persisted native baselines | All twelve prepared layouts passed |
| Q027 installed owner checkpoint import | All twelve prepared layouts passed |
| Q027 one-save and fresh-process persistence after owner import | All twelve prepared layouts passed |
| Q027 installed owner local checkpoint/backup/interruption/rollback/reopen | All twelve prepared layouts passed |
| Q027 invalid-resource preflight | All three Missing, Empty and SignatureCorrupt cases refused before mutation, with unchanged sentinel resources, native state, disk and recovery state |
| Q027 authenticated exact remote | One private fixture push/fetch/owner-import/save/fresh-read-only-reopen scenario passed; production macros and repository main were untouched |
| Q027 aggregate | All 52 prepared scenarios passed; each of the six banks invoked once; original coordinator exit code 0 |
| Q027 independent lifecycle audit | All 78 original Excel normal exits proved within the unchanged 10,000 ms bound; no replay or forced termination |
| Q027 owner evidence | 24 installed import/persistence raw reports and 40 recovery/remote owner steps: 37 admitted mutations and three proved prewrite refusals |
| Q027 actual designer pixel review | All 80 indexed captures inspected in ten contact sheets; full originals 43 and 46 also inspected |
| Temporary registration and campaign release | Exact original HKCU Registry64 COM state restored and verified; owned completed task removed; no Word or Excel process remained |
| Pure verifier regressions | 13 aggregation cases, eight reflection cases, five owner-role positives and ten negatives passed; both historical singleton failures reproduced |
| Coverage collection | Not run; no line or branch percentage claimed |

The prepared layouts are LabelButton, TextBox, ComboBox, ListBox, CheckBox,
OptionButton, ToggleButton, ScrollBar, SpinButton, TabStrip, Image and
FrameMultiPage. Strict transport/readback checks retain source and raw FRM/FRX
resources, complete collections and parent relationships, native properties and
font descriptors, unchanged snapshot stability, meaningful-change detection,
backup/interruption/rollback, and saved/fresh-process persistence. Frame font
8.27 is not rounded. All native mutations retain the installed owning process,
original host generation and native STA checks.

Evidence roots are `artifacts/q27g3/` (managed prerequisite),
`artifacts/q24p3/` (Word) and `artifacts/q27p4/` (Excel). The final
`q27p4/qualification.json` binds the original frozen plan, independent strict
audit, actual `designer-review/visual-review.json`, Word receipt, exact registry
restoration and owned task removal. The strict audit independently verifies
physical resources and owner intent/mutation/terminal chains; a cleanup trace
or an observer response alone is not host-exit proof.

The final `q27p4/completion-audit.json` checks the declared requirements against
the physical receipts and exact data rows. A separate read-only limited
interactive observer confirms the original HKCU fingerprint, task absence and
free Office hosts in `completion-release-state.json`; its completed observation
task is also removed. The terminal caller's registry view differs, as already
shown by the pre-native previews. Its mismatched comparison is retained as a
context-specific failure, not classified as a failed interactive restoration.

The earlier `q27p3` aggregate remains failed: its native banks passed but its
original coordinator exited with an aggregation error. The fresh `q27p4`
coordinator completed normally. Its original pre-native strict auditor then
failed on PowerShell reflection boxing; the first amended auditor failed on
singleton-array collapse. Both failures are retained. Separate, sealed
**post-native verifier amendments** unbox reflection arguments and retain role
arrays; they change neither native oracles nor frozen native inputs and perform
no native replay. The successful strict auditor is version 7. Its amendment
manifest explicitly records `PreparedBeforeNative=false`, `NativeReplay=false`
and `NativeOraclesChanged=false`. The final owner-role proof uses the sealed
test's complete bytes; the earlier sealed proof had fewer negative cases and
remains unchanged.

Pixel review establishes the visible prepared controls and containers. Some
reopened tall forms have unused lower grid clipped by the MDI viewport; tested
controls remain visible. Installed owner construction captures precede import:
**owner restoration is proved by raw resources and native readbacks, not by
post-import visual evidence**. Pixels do not measure exact font sizes or hidden
collections. Chat panel appearance, native RCW release, Word templates/arbitrary
SaveAs or cancellation paths, other hosts and release-wide acceptance are not
inferred. The historical checkpoints below retain their original outcomes.

## Office adapter persistence (Q012, 2026-10-04)

Access and Publisher's existing-document adapter contract is accepted on frozen
product MVID `2106fd95-fb1d-4b3e-b971-0ba39e78494b`, SHA-256
`852DC8414D0E667ACAF9A3424978683F14E4EA8DED2ABF2780DBAE9851AF1CA8`.
Plans identify base `60b8b7f2d81bcc3ca04900481f31a2ea956c1f5a` and the frozen
pending-source manifest; base commit alone does not identify these DLL bytes.

| Validation | Result |
| --- | --- |
| Final focused managed gate | 766 passed, 0 failed, 0 skipped |
| Required native adapter cases | 20 accepted original individual cases |
| Owned host lifecycle | 35 original/fresh generations exited normally, exit code 0 |
| Complete managed suite and coverage on this product | NOT_RUN |

Accepted operations cover active-module and module/class source Save, retained
UserForm state, reference addition/removal, Description and explicit native
General HelpContextID/ANSI-representable HelpFile persistence. A Save is followed
by independent fresh-process readback. Publisher uses an audited serialized seed.
NewDocument/first SaveAs, positive unsupported-Unicode persistence, form rendering
and event execution remain outside this result. Unicode is accepted only as a
verified pre-write refusal. Legacy HelpFile/HelpContextID COM writes are refused;
the explicit General workflow has its own guards.

Original case receipts and fixture identities are retained under
`artifacts/q012-final-access/`, `q012-access-remaining/`,
`q012-publisher-remaining/` and `q012-publisher-pending-detached/`.
Acceptance combines individual original outcomes, not a relabelled failed
aggregate. Settings, registration and owned desktop release are verified.

## SOLIDWORKS native core and assistant (Q014, 2026-10-04)

The corrected frozen product MVID
`ddf638b2-30d2-40a5-8ad3-9d49f303ff7c`, SHA-256
`69BDE5B2D1CB55CA23597540F1CA108231F342FAF79A79B3F160E31FAC11AFD5`
is accepted for selected SOLIDWORKS 2019 SP5 (27.5.0) and 2025 SP1.1 (33.1.1).

The native banks verify connected load, modules/classes, a synthetic UserForm and
Label, stale refusals, BAS/CLS/FRM/FRX transport, compilation, breakpoint/run/step/
continue, product Save and native Edit Macro reopening of identical saved copies.
Scoped actual captures verify Monaco/designer transitions and pane closure with
source preservation. Separate local assistant evidence verifies native reading,
permission/stale refusals, streaming, cancellation and recovery; this does not
qualify an authenticated external provider in SOLIDWORKS.

The original host, owned IDE and coordinator exit normally, and temporary
COM/settings state is restored. The final decision is retained in
`artifacts/q014-final-decision-20261004/qualification.json`; original banks are
`q014-final-native-2019-20261004`, `q014-recovery-barrier-2019-20261004` and
`q014-recovery-barrier-2025-20261004`. Private-desktop acceptance is specific to
those operations and layouts; no full release or new coverage result is inferred.

## Q-020 and Q-030 native macro candidate, 2026-10-05

The isolated DLL was built from base commit
`4a5b05e0f9efc20fda3ee61f93c45d45cd32cef1` plus frozen creation, publication,
General-reader, picture-persistence and post-save verification changes. Its
MVID is `08689325-53fe-4d8d-86a5-3c02aa5c4a81`, SHA-256
`3A1B0A92803C737156EE381CE66070C22CA5A932715528E3B622C2DA8101F1B7`.
Build completed without warnings or errors. The scoped managed bank passes
**579 tests**, with zero failures or skips, in
`native-macros-managed-11/native-macros-managed-11.trx`, SHA-256
`3C057E3C67C840672E9D1B5848EDD15299687FC4252CCEAC7CFFDD1F0F160737`.
It covers the changed routes and admission, identity, revision, mode, owning
thread, metadata, transport, privacy and uncertainty guards, including real
Windows OLE picture persistence. It is not the full managed suite or a coverage
measurement. The earlier post-save false-to-true Saved transition has a retained
failing regression on the previous candidate and a passing focused correction.

| Native scope | Actual evidence and result |
| --- | --- |
| SOLIDWORKS 2019 SP5, revision 27.5.0 | Scoped phase acceptance: first-generation bank13, full original A readback and designer review in continuation18, full original B readback after explicit designer opening and designer review in continuation22. Both workflows preserve code/attributes/form/picture/references/General. First host exits normally; fresh host's settled typed ExitApp is correlated with exit zero on its original retained creation handle. Prior failed aggregate banks and the raw failed exit receipt remain unchanged. |
| SOLIDWORKS 2025 SP1.1, revision 33.1.1 | Full bank18: all 16 ordered stages PASS, including first-save creation/publication, source preservation, refusal of unsafe generic Open, fresh-host original A/B readback, and normal exit of both original host handles. Three actual native designer PNGs reviewed. |
| Cleanup | Actual per-user COM baselines restored exactly after each campaign; original worker exits observed and exact owned scheduled tasks removed. No production macro or user project is executed. |

The frozen candidate is identical across both hosts. Record source-file hashes
and acceptance receipt hashes in [the machine-readable evidence index](qualification/q020-q030.json).
Native acceptance requires explicit designer initialization before cold inspection.
The observed closed-designer E_FAIL remains a limit. Default form metadata,
resource content and recognized MSForms enum descriptors are compared with strict
negative checks; raw snapshots and revisions are not rewritten. The installed
.NET Framework WinForms implementation generates enum names from an ITypeInfo
pointer, also shown in the [official WinForms source](https://github.com/dotnet/winforms/blob/v3.1.0/src/System.Windows.Forms/src/System/Windows/Forms/ComponentModel/COM2Interop/COM2TypeInfoProcessor.cs#L868-L890).
The portable canaries cover complete snapshot guards, original-handle startup
identity, cross-generation enum descriptors, one-time designer materialization,
and transient/persistent startup readiness. They make no native acceptance claim.

Failed evidence remains recorded: first-generation bank13's second launch,
cold14's pre-open helper lookup, cold15's enum comparator, continuation18's closed
designer inspection, diagnostic19's result collection, and 2025-17's transient
startup dialog. None is relabeled green by later scoped evidence. See the
[authoritative acceptance limits](release-qualification.md#current-acceptance-summary).
No result qualifies arbitrary UserForms, signatures, macro runtime execution,
the complete embedded assistant UI, generic standalone Open, Type101 promotion,
or v1.0.0 release acceptance.

## Native options and debugger boundaries

Q026 is accepted for the selected Excel Size/options matrix on source
`08d7420`, MVID `7156af5b-941c-4452-9e78-1381cb69af0d`, Office x64
16.0.20430.20092. Exact Tabs/revision restoration, native mutation, original-handle
exit and the controlled metadata-only drift comparison pass under the maintainer's
approved acceptance criterion. The original September trigger remains unexplained.
The machine-readable decision is
[q026-qualified-candidate.json](../tools/tests/q026-qualified-candidate.json).

The Q006 Excel campaign on frozen MVID
`d2c3601b-893d-4e84-b9e3-c172da7e2437`, SHA-256
`C4D095D9427379AC8F2A82D047C8780637A0E815173241976F2C86664E777644`
accepts its declared supported scalar pages, ParamArray/Variant arrays, protection
and product Save followed by independent reopen of module/class/form/Label state.
The extended single-page terminal-trace oracle failed at the diagnostic event
cap; other scalar/runtime/host scopes and historical break-mode crashes remain
unqualified. Original campaigns are under `artifacts/q006-20261002-*`.

## Coverage status

There is **no current complete line/branch measurement** for the Q028 candidate.
Its focused managed gate is functional evidence and does not satisfy the full
coverage target. Native qualification also does not measure managed coverage.

For historical comparison only, source
`8f2315d04162f55b0956618f96b59d294a3fb681`, MVID
`d8f31d57-8612-465e-871c-93a62f2b3eae`, SHA-256
`C900BA09D92DA7CF50CC09033C63F5226DD04C18426CC386B30AF86EB0BA0941`
passed its instrumented managed run on 2026-10-01: 2,324 passed, 0 failed,
90 conditional skips. That collector measured 33,562/33,755 lines (99.43%) and
33,975/34,487 branch outcomes (98.52%). It excludes external hosts, JavaScript and
native C++; these values are not the current branch's coverage.

Its original collector is
`artifacts/qualification-v1/followup-20260930/managed-v6-path-diagnostic/d25f9f82-6146-4466-8c24-de444ccf25de/coverage.cobertura.xml`.
The target remains complete line and branch coverage, with meaningful mirrored
tests and no exclusions added to improve a percentage.

## Evidence retention

Local `artifacts/` proof sets are not distributed as source or automatically
uploaded. Preserve their candidate manifests, individual TRX outcomes, original
process receipts and recovery state when reviewing a qualification. A missing
local proof set must be reported as missing rather than reconstructed as a pass.
Historical investigation text remains in the documentation's Git history.
The repository does not infer a single globally qualified release by combining
results from different frozen binaries.

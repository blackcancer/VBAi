# Version 1.0.0 qualification

**Decision: not qualified for release yet.** This is the current acceptance gate,
not a claim of universal Office compatibility. Test totals and measured coverage
are maintained only in [recorded validation](test-coverage.md).

## Current candidate checkpoint (2026-10-02)

The active qualification focus is Q-024; Q-027 is deferred at the maintainer's
request and retains its recorded open gates. Source `b9dc446`, product MVID
`824eb4bf-abc7-44c8-aa75-5a51e030d04f`, passes its complete default managed
gate with clean source and exact hashes. Its single Word trial reads references
successfully through the owning-thread bridge, then the test helper rejects the
native unbroken `Normal` template reference's empty GUID/0.0 identity. No menu
intent is emitted. Original Close/Quit and exit code 0 are observed, and
registration restoration is verified. Production capture/checkpoint/compare and
the separate chat entry remain NOT_RUN. The corrected template-aware helper and
complete reference-revision readback require a fresh native trial. See
[recorded validation](test-coverage.md#word-template-reference-preparation-refusal-2026-10-02).

Source `b6f13fb`, product MVID `988010f8-e080-4caa-a54d-11ab51d58016`, passes
the complete default managed gate. Its single retained-copy Excel observation
reads `Application.VBE` successfully while the inert initial workbook is open,
with events disabled and AutomationSecurity 3. After the sole planned replacement
and exact Open/path/count/hash checks, the getter fails with `0x800A03EC` on the
same owner thread. No form capture, import or font transfer is reached. The cause
remains unknown; startup bridge connection is not current getter acceptance.
Normal original exit, unchanged source/copy bytes and registry restoration are
verified. Q-027 and its complete root-font/Frame/control/recovery/persistence
scope remain open. See
[recorded validation](test-coverage.md#retained-vbe-lifecycle-observation-and-managed-gate-2026-10-02).

The tests-only lifecycle source `f524848`, product MVID
`4114e0c5-375c-4639-bbee-aae21302288a`, fails its complete default managed gate
on an existing native-options ComboBox value assertion. Its embedded Git fixture
mirror passes, but the before/after-copy native observation is NOT_RUN. No Apply
or font delivery occurs; the installed DLL, registration and trust settings remain
unchanged. The preceding named native `Application.VBE` refusal remains the
authoritative retained-copy result. Q-027 stays open. See
[recorded validation](test-coverage.md#retained-vbe-lifetime-candidate-managed-control-failure-2026-10-02).

Source `9ba099e`, product MVID `deee6a02-2c77-47f0-8b7b-78b3553f40fc`,
includes merged PR 19 and passes the complete default managed gate. Its single
retained-copy Excel diagnostic identifies `Application.VBE` as the getter failing
with `0x800A03EC` after successful Open/path/count/hash verification.
`VBE.MainWindow`, baseline capture, manifest, menu, import and font transfer are
NOT_RUN. The original owned host exits normally, source/copy bytes stay unchanged
and temporary registration is restored with verified readback. The installed DLL
and trust settings are unchanged. The cause remains unknown; Q-027 and its full
matrix, default root font, Frame 8.27, recovery and persistence remain open. See
[recorded validation](test-coverage.md#named-vbe-getter-observation-after-pr-19-integration-2026-10-02).

The named-getter diagnostic source `2060217`, product MVID
`6eab1262-f577-4650-8f6c-c83bdb065d01`, fails its complete default managed gate
on a Git recovery-marker mirror assertion. The expected preview-created directory
is absent after an accepted `InvalidOperationException`; its actual origin remains
unknown. The Git path is unchanged and passed in the preceding complete gate,
but a timing or environment cause is not established. Registration previews are
read-only; Apply and a native trial for this MVID are NOT_RUN. The installed DLL
and host trust are unchanged. Stronger test diagnostics are prepared separately;
Q-027 remains open. See
[recorded validation](test-coverage.md#named-vbe-getter-candidate-managed-gate-failure-2026-10-02).

The gated PR integration source `3ff51bd` is exercised in an exact isolated
checkout using the frozen product MVID `d5b638d6-56b4-4cd1-9d39-6b1fb9c5722e`.
The retained workbook copy now passes Open/path/count/hash checks. Excel
`16.0.20430.20092` then refuses a source line containing `Application.VBE` and
`VBE.MainWindow` with HRESULT `0x800A03EC`; the precise getter and cause remain
unknown. Baseline capture, manifest, menu, import and transfer are NOT_RUN.
Original normal exit, unchanged workbook bytes and registry restoration are
independently verified. A tests-only named-getter diagnostic has focused mirror
regressions; its subsequent complete gate fails on the Git mirror above, so
its native observation remains NOT_RUN. It preserves ForceDisable and does not reopen or retry. Q-027 stays open. See
[recorded validation](test-coverage.md#retained-userform-copy-vbe-access-refusal-2026-10-02).

The retained-workbook diagnostic source `6d3d802`, MVID
`46bd53b3-7416-4530-bd86-bcbc251ca041`, passes its complete default managed gate.
Its one Excel trial stops after opening the synthetic copy, when the fixture
hash reader conflicts with Excel's file sharing. Baseline capture, menu, import
and font transfer are NOT_RUN. Original normal exit, unchanged source/copy
bytes and restored registration are independently verified. The tests-only
reader correction passes the complete default managed gate in the PR 19
integration source `3ff51bd`, MVID `d5b638d6-56b4-4cd1-9d39-6b1fb9c5722e`.
PR 19 is merged into main, and the local qualification branch is synchronized.
The installed DLL is unchanged; its subsequent native discriminator stops before baseline capture as recorded above.
Q-027 stays open. See
[PR integration validation](test-coverage.md#pr-19-integration-managed-gate-2026-10-02) and
[recorded validation](test-coverage.md#retained-userform-copy-preparation-failure-2026-10-02).

The latest gated diagnostic source `a5649c3`, MVID
`736dae63-1922-44b8-8f0f-65740b87e76d`, SHA-256
`423CC1AD64D0B1A4E25E0EB8CCC8C821B56E7AB03CFF21C79F498CBA45C95A4F`,
passes its complete default managed gate. A separately declared Arial 9.00
baseline survives save/reopen and actual owner-dispatched LabelButton import.
The initial and later independent snapshots are exact, with matching native
properties and font. The diagnostic refuses the already exact initial state
before any deferred transfer, so the native scenario remains failed. This is
scoped exact import evidence, not ordinary-workflow or restoration acceptance.
Default Tahoma explicit-font fidelity, Frame 8.27, recovery and post-import
persistence remain open. The owned host exits normally, disk bytes stay
unchanged and registration is restored. See
[recorded validation](test-coverage.md#explicit-root-font-import-and-diagnostic-gate-refusal-2026-10-02).

The preceding gated diagnostic source `269b187`, MVID
`e78f15bc-f918-425e-9d1f-a8e700da3caf`, SHA-256
`627EDE3B8038580E107791600CCA85E080FFD88F0FF6F2EA6A437C4DBC656CCA`,
passes its complete default managed gate. Its fresh owner-UI LabelButton trial
stops before manifest publication or Git menu emission: the saved/reopened
reference lacks the required root font binding. The declared post-capture
font transfer is NOT_RUN, rather than a failed native setter. The owned Excel
exits normally and temporary registration is restored. The baseline variation's
cause remains unproven; Q-027 and its full native scope remain open. See
[recorded validation](test-coverage.md#post-capture-font-diagnostic-baseline-refusal-2026-10-02).

The preceding gated observation source `be9dcd8`, MVID
`ab35c2ed-1c3d-47b4-9490-b26e0b8bc737`, SHA-256
`B53848A4838BF20CD6077A2B73A3371F624B38593ED6356E370021B325B77DC2`,
passes its complete default managed gate. Fresh owner-UI LabelButton trials
observe the root binding already absent before diagnostic font getters and
exact child-value readback without restored persistence. In the separate
distinct-name trial, the child value reads Arial while both attached font
objects still serialize Tahoma. Both imports fail exact comparison, with only
the FRX differing. Both owned hosts exit normally, saved workbook bytes remain
unchanged and temporary registration is restored. The child collection route
has not updated the observed attached descriptor; the internal cause remains
unproven. Q-027, the complete layout matrix, recovery, transfer and post-import
reopen remain open. See
[recorded validation](test-coverage.md#owner-thread-root-font-observation-and-distinct-child-name-2026-10-02).

The earlier gated restoration source `7745e1c`, MVID
`e5df96d3-56c3-4089-a8dd-d2b8785f2c5b`, SHA-256
`93D21AD271ECD87A656047BBB533173E1E35CF4D5917076D5DB8D91C1FC71C93`,
passes its complete default managed gate. Its first fresh owner-UI LabelButton
import passes the scalar-child phase without a reported setter error, but
fails exact snapshot comparison. Independent snapshot and frozen-parser
inspection confirm that only the FRX differs and the root StdFont binding
still disappears. Setter readback and the cause of this persistence failure
remain unknown. The original owned Excel exits normally, saved workbook bytes
stay unchanged and temporary registration is restored. Q-027 stays open; no
other layout, recovery or post-import reopen is qualified by this canary. See
[recorded validation](test-coverage.md#root-font-child-value-import-canary-2026-10-02).

The earlier fully gated diagnostic source `e3d94bc`, MVID
`8071360b-1c57-45fd-9ff9-5b575dee2576`, SHA-256
`62A2209C1D3E659CBC6CE4A96CB0CB2768AFB73F31E096C19E65BDBE4A1832DB`,
passes its complete managed gate. A fresh owner-dispatched LabelButton import
identifies `VBIDE.Property.Object.set` as the failing operation, with inner
`NotSupportedException` and managed HRESULT `0x80131515`. The original native
HRESULT remains unknown. The owned Excel exits normally and registration is
restored. A separate external-STA read-only trial observes all eight Font.Value
children and get/put declarations, with exact snapshot, StdFont, source, Saved
state and disk-byte preservation, normal original exit and restored registration.
It does not exercise setters or import. The following implementation changes
root restoration to prevalidated scalar child Value writes; actual owner-STA
import and reopen acceptance are still required. Neither diagnostic replaces
the full failed matrix. See
[recorded validation](test-coverage.md#root-font-operation-and-read-only-child-observation-2026-10-02).

The latest complete native UserForm matrix uses source `de5c619`, MVID
`22fe3345-5015-4949-860f-23cdf23449e3`, SHA-256
`4D3859E07F9479BA6E465F6F7346B66F80B2C91EF7D58458EBDE453B14F1C637`.
Its complete default managed gate passes. The root Font property route remains
unqualified: every declared native UserForm checkpoint-import case reaches a
terminal property/method error. The retained UI message does not identify the
precise COM member. First-case offline inspection still finds the root font
descriptor missing after import. All owned Excel processes exit normally and
the temporary future-activation registration is restored from its guarded
snapshot. Installed files remain unchanged. Q-027 stays open; counters and
candidate-specific evidence are in
[recorded validation](test-coverage.md#root-userform-font-property-route-2026-10-02).

The installed candidate is product source `608c002`, with documentation
checkpoint `d6b79b7`, MVID `d2c3601b-893d-4e84-b9e3-c172da7e2437`, SHA-256
`C4D095D9427379AC8F2A82D047C8780637A0E815173241976F2C86664E777644`.
Its complete default managed gate passes with unchanged frozen binaries using
the normal Windows temporary directory. The failed repository-local TEMP run
and its deferred Access continuation failure remain separately recorded;
the latter's original cause is not established.

An owned, visible Excel read-only surface trial confirms the installed MVID and
normal original-process exit. The saved/reopened root form exposes
`VBComponent.Properties.Item("Font").Object`, with the same COM identity as
`Designer.Font`. Corrected offline comparison preserves the root font descriptor
and all meaningful exported resources after the reads. Raw string padding and
CFB slack differ; the initial diagnostic comparison omitted production's line
normalization. No import or font setter is exercised by this trial, so Q-027
remains open. Exact counters, identities and evidence are in
[recorded validation](test-coverage.md#installed-candidate-and-read-only-userform-surface-2026-10-02).

### Earlier Q-027 import candidate

The earlier Q-027 follow-up candidate is source `420b3da`, MVID
`49d7c21c-f394-4d45-b55f-fd71f235e684`, SHA-256
`E5D8234F1C817CFB585227E32B7C925EE7E1BE101202B2C7F95FEC20A4C9EF64`.
It corrects the native accessibility ListBox selection cache and bounds owner
names before constructing font restoration paths. Its predecessor `198e0ec`
added an owning-thread, guarded exact-font restoration path without relaxing
FRX comparison. Focused managed regressions pass, but the complete managed
gate has a disk-full signature-size failure. The corrected selection fixture
reaches actual owner-dispatched LabelButton import: sources match but the root
FRX font descriptor is omitted, so exact comparison correctly refuses success.
Original normal exit is verified for that case. A subsequent seed attachment
failure blocks the other layouts. Q-027 remains open. At that checkpoint the
installed DLL had an unknown-provenance identity, separately recorded in
[recorded validation](test-coverage.md#q-027-userform-follow-up-2026-10-02).
That installation has since been backed up and replaced by the known candidate
above; the earlier native import results are not reruns against that candidate.

### Earlier 2026-10-01 operation checkpoint

Product source `4a323196`, MVID `25facee2-6b10-40fa-880c-cb9328333713`,
SHA-256 `474E9C106E109C693AC47587BD13BF7239E617F31EA7F210B1208E4EA4678F7D`
has a completed default managed gate and operation-specific native results in
`artifacts/native-qualification-20261001`. Tests-only fixture corrections are
integrated separately. At that checkpoint the installed DLL was unchanged and
both temporary COM registration contexts were restored from their guarded
snapshots. See
[recorded validation](test-coverage.md#native-qualification-refresh-2026-10-01)
for exact counters, assembly identities, failed preparation and terminal runs.

| Gate | Current candidate observation | Remaining acceptance |
| --- | --- | --- |
| Q-006 / Q-020 | Excel ParamArray/Variant arrays, protection reopen, breakpoint, options and the declared scalar page pass with normal owned-host exit. | Historical crashes and other runtime local types remain unresolved. |
| Q-011 | Word and PowerPoint adapter-only source/class/form save and fresh-process reopen pass. The final contract verifies distinct old/new PIDs and normal exit of all four processes with a 15000 ms observation bound; earlier same-process passes and failed fresh-process trials remain separate. Read-only getter retries are bounded without replaying Save/Close/Quit. | This qualifies the stated adapter operation only. |
| Q-012 | Earlier Access/Publisher reference and selected metadata scopes pass. Offline inspection confirms that the failed Publisher HelpFile document stores the exact path in every checked help-path record, despite altered native readback. A later batch again passes Publisher Description. | A common 15000 ms exit bound does not qualify Access Description: its original process remains alive. A later Publisher HelpContextID setter fails, and cleanup refuses Quit for its unsaved project. Other metadata cases are skipped. Earlier Access help exit failures and Publisher HelpFile readback remain open. Separate forced cleanup does not qualify normal exit. |
| Q-024 / Q-027 | Actual owner-dispatched Git capture/local checkpoint/compare passes with native state, saved-file hash and normal exit. Independent no-import reopen reproduces the initial geometry/list/image differences. Saved/reopened baselines remain exact in the accepted scopes. Stream inspection then isolates missing font descriptors; expanded native readback finds Frame.Font.Size changing from 8.27 to 8.25. | Strict FRX comparison still refuses import. External setters, owner assignment and guarded owner-STA font commands do not restore the tested snapshots. Persisted-font Load also fails the reached root-form scopes; its Frame trial stops at the active-form guard before import. Remote transfer, recovery and post-import reopen are not qualified. |
| Q-026 | Native Format font/category/color changes and complete restoration pass; an empty size catalogue is refused with unchanged state. | Font-size mutation and the historical incomplete trace remain unqualified. |
| Signature | Native unsigned state is observed; the disk verifier reports unavailable Office SIP. The single signing request is cancelled on a mismatched certificate name, followed by normal exit and removal of the owned certificate/key. The official Microsoft x64 SIP payload and a reversible registration worker are prepared but not executed. | Signature persistence, fresh reopen and cryptographic/trust verification remain failed or not reached. Temporary administrator registration awaits the maintainer decision; no trust policy is changed. |
| Q-028 | Earlier automatic-device/CPU batches fail model allocation. After authorized SOLIDWORKS closure and verified extra commit memory, the selected CPU profile passes headless tools, cancellation/recovery, shown detached chat and native Excel read-only inspection on this frozen candidate. | Acceptance applies to this explicit CPU/model/sampling profile and in-process native-tool dispatch. The embedded-host assistant, other device profiles and the historical intermittent cause remain unqualified. |
| Q-014 / Q-030 | No selected disposable SOLIDWORKS bridge is available for this campaign. The maintainer-authorized normal window close frees resources for the next provider trial; later process absence does not prove exit within the original observation deadline. | Current SOLIDWORKS save/reopen and native Edit Macro are NOT_RUN. Unsafe VBProjects.Open remains unavailable. |

No new coverage percentage or release-wide native acceptance is inferred. The
VBA test explorer feature is outside this native qualification campaign.

That UserForm diagnosis preserves exact raw recovery files and comparison
rules. The lost root and nested Frame font descriptors are meaningful data, not
CFB allocation or documented padding. The guarded Font.Size refusal is retained;
no rounding tolerance is widened to accept the changed size. Independent OLE
font controls distinguish persistence before property getters from later native
readback, but do not establish the VBE import cause. Subsequent guarded
restoration and list-cache corrections have not reached native import acceptance.
Exact source/binary identities and all terminal diagnostic results are
in [recorded validation](test-coverage.md#native-qualification-refresh-2026-10-01).

## Report scope

Qualification checkpoint: 2026-10-01, following the 2026-09-29/30 campaign.
The register below contains all 30 findings, Q-001 through Q-030. Each status
applies only to the stated operation, host and tested candidate; CLOSED does not
qualify an entire application. OPEN and PARTIAL entries remain release gates
where required by the agreed scope.

The remaining completion gates are Q-006, Q-012, Q-014, Q-020, Q-024,
Q-026, Q-027, Q-028 and Q-030. The other findings have the scoped corrections and
validation described below; they do not constitute complete current-candidate
qualification. The previously installed product at this checkpoint was source
`8f2315d`, MVID
`d8f31d57-8612-465e-871c-93a62f2b3eae`, SHA-256
`C900BA09D92DA7CF50CC09033C63F5226DD04C18426CC386B30AF86EB0BA0941`.
`deployment-v6.json` retains exact installation and previous-payload/registration
backup evidence. Its complete instrumented suite passes with unchanged product
hash; Q-015 closes only for this exact source/MVID/hash. Native/provider opt-ins
and the complete coverage target remain separate scopes. The preceding
Monaco-status candidate `d7a1c75` / `6a74af33` and earlier frozen v5 `2e75161` / `f9a36c85` have separate completed managed passes. The original
v5 harness failure remains failed with explained packaging/Git path defects;
none of those results qualifies a later product binary.

The earlier uninstalled combined candidate is source `b60996c`, MVID
`e79c6288-d384-475c-b8bc-276d7caaaf00`, SHA-256
`9ADBCFB96B1F2F4E4FA3066EA26F3CA2E0A2EC7996C85105070A3765BEFC8587`.
Its complete default managed gate and separate JavaScript, synthetic native,
managed-loader and detached Designer checks pass with unchanged product bytes
and source. Q-015 is accepted for that exact managed scope; this is not an
installed or native-qualified release. Exact counters and coverage are recorded
only in [recorded validation](test-coverage.md).

Tests-only follow-up `a28be69` prepares the actual embedded Git menu/window
workflow against the unchanged frozen product; its complete default managed gate
passes. Native execution remains NOT_RUN, and it does not qualify publish,
import, recovery or disk reopen. Follow-up `ce3d9da` permits a strictly selected
literal-loopback port for detached Ollama qualification. Its activated full
suite fails the headless tool-argument assertion and reproduces the chat's empty
response. Exact captured SSE bytes already contain no text or tools before
assembly in VBAi; the chat displays the empty-response fallback with ready
status. The current response boundary is proven, but backend/model internals
and the historical failure's cause remain unproven. No assertion is relaxed,
request replayed to obtain a pass, or product correction claimed. Q-028 stays
open; complete results and preparation failures are in recorded validation.

Subsequent direct backend controls are terminal: the exact original request and
its chosen-token variant deliver a recognized core-tool call; the nonstreaming
variant delivers a call missing a required argument; the no-tools variant
streams text. No tool is dispatched. An offline correction explains the separate
diagnostic runner's PowerShell singleton-array failure without altering its
original failed record or replaying requests. The exact owned helper is closed
after independent terminal verification. These new generations do not explain
the original empty response or qualify VBAi's provider integration. Q-028 stays
open; byte-level evidence and scope are in recorded validation.

After synchronizing main `2f28018`, an independent sampling comparison reproduces
a new complete-empty response and captures its generated malformed tool JSON.
The exact backend receives no explicit temperature from VBAi and applies 1;
lower temperatures change that observed outcome but still produce incorrect
tool choices or missing arguments. The captured prompt fits the allocated
context without truncation. This explains the new controlled response boundary,
not the earlier responses without generated-token evidence. No product settings,
parser or native dispatch policy changes are made, and no later result is used
to erase an earlier failure. Q-028 remains open; exact scope, counts and evidence
are in [recorded validation](test-coverage.md).

A new detached Ollama aggregate uses product source main `2f28018`, MVID
`044522a1-31cc-494c-98e6-46dee00af787`, SHA-256
`1A5B036B899A88D1E4B5315E3082C0689BA810B2BADB87F92A24926E1D251B4F`,
and tests-only source `f15c48a`. Cancellation/recovery and the strict synthetic
tool roundtrip pass, but the visible chat receives another complete-empty SSE
response and fails before its stop/next-send phases. The aggregate remains
failed. The artifact driver's MVID read also prevents coverage instrumentation;
its lock is verified on a disposable offline copy and a separate driver
correction is prepared. Original payload hashes and installed DLL remain
unchanged. These observations neither qualify the new native product nor
close Q-028. See recorded validation for exact counters and evidence.

Configured candidate `799ccef` adds optional persisted Ollama sampling and a
shared qualification selector for the headless, detached UI and native Excel
scenarios. The selected `qwen2.5:7b-instruct` profile explicitly sends temperature
0 and top-p 0.8, using an isolated context of 8192 and one parallel request.
Its four live scenarios pass, including a random unprompted marker read through
the real Excel tools with unchanged source and normal owned-host shutdown.
The aggregate still fails the language-catalogue completeness check. Completing
the new translations and restoring in-memory sampling after a failed settings
save require a subsequent complete gate. The installed product, embedded-host
assistant and other model/provider combinations remain separate scopes. Earlier
failed trials and their limits are retained in recorded validation.

Candidate `bbb6e6f` completes the language catalogues and restores shared sampling
values when settings persistence or theme application fails. Its complete managed
gate passes, with product MVID `f6f01687-c0a4-4b1c-acab-82cc1dfab2af` and SHA-256
`35E94E67D30819E32790854E317C55D5736DC74EA7EB2461F8301E56F2E6B9D1`.
The four strict real-provider scenarios pass on the same selected 7B configuration;
captured requests verify explicit sampling and the native Excel source remains
unchanged. Its owned Excel process exits normally and the backend is cleaned up.
Q-015 is accepted for this exact managed candidate. Q-028 is PARTIAL for the
selected detached chat, headless and external native-tool scope; the embedded-host
assistant and historical intermittent cause remain unqualified. The installed
product is unchanged. Exact results and scope are in recorded validation.

The Git menu, chat and LLM tools now share a compatible lookup of the native
document key and the former uppercase key. A single existing binding is reused
without moving or rewriting caches; two bindings or uncertain metadata refuse
automatic selection. Managed regressions and the complete candidate gate pass.
This corrects a source-proven entry-point mismatch, not the historical native
export failures. Q-024/Q-027 still require the actual embedded owner-thread Git
workflow; existing external-test-STA captures do not establish that acceptance.

The retained current-v6 Format trial commits and verifies the font change, then
refuses the requested size because the observed size catalogue is empty. Its
captures do not identify the Win32/UIA branch, native handle/style or intermediate
counts, so neither lazy list initialization nor a provider defect is proven.
Later read-only evidence observes the Options window closed, without qualifying
baseline font restoration. The original failed trial and pending recovery remain
unchanged; no size write or uncertain action is replayed. Q-026 stays open.

Current-v6 native evidence includes the explicitly authorized read-only
owner-STA path/effective-token diagnostic in owned Excel. Both synthetic GUID
directories and files are visible with attributes matching the testhost; normal
exit is verified. This neither reproduces nor explains the earlier export
path-not-found observation, and Q-027 remains open. Subsequent paired owner-STA
observations and single exports pass below LocalAppData and TEMP in explicitly
bootstrapped Excel, with normal exits. Fresh COM-activation trials on this same
product still fail below LocalAppData/VBAi/GitTemporary. Launch-context
differences do not establish a cause or complete Git capture acceptance.

Current-v6 Access HelpContextID returns 321 through all three native/managed
getters after a bridge setter and one verified Save, and also in the separate
external CLR setter case. Both initial hosts fail normal exit; fresh-disk reopen
is NOT_RUN. Authorized force cleanup preserves stable database copies but does
not qualify persistence. External HelpFile setters return through both production
CLR and raw IDispatch PUT, yet both fresh-disk reopen values are altered through
all three getters. The later raw HelpContextID PUT also returns successfully with
live value 321 and verified Save, but fails bounded exit and disk reopen remains
NOT_RUN. Its retained copy and authorized force cleanup do not qualify normal
exit. No conversion heuristic or causal
product fix is established. The initial current Publisher batch fails its active-project
startup guard despite exact disposable-document/persistence identity, before
baseline edits or product Save. Its independently guarded, authorized window
close later exits normally without force termination; the startup failure remains
failed. External Publisher VBE inventory observations are UNVERIFIED, as recorded
in an additive correction. Exact terminal outcomes and the correction of copied
getter-driver labels are in recorded validation. Q-006 and Q-012 remain open for
these scopes.

Current-v6 ordered Excel scalar trials now pass the unsupported-declaration page,
one Long value and the complete declared scalar page, with correlated owner-STA
phase evidence, unchanged identity/source/selection/mode and normal owned exits.
This supersedes the current-v6 NOT_RUN scalar status, while historical stalls and
crash causes remain open. The initial native Monaco attempts fail in fixture preparation
(a malformed selection request, then a decorated-tab caption mismatch); normal
cleanup does not qualify those attempts. After correcting the fixture, the
actual installed embedded editor passes the closed-project/live-status scenario
with exact source/identity preservation, normal exit and a reviewed native capture.
Exact scope, source revisions and terminal records are in recorded validation.

The broader current-v6 Excel campaign remains failed: the array/ParamArray
procedure results return before abnormal cleanup exits, project-protection
cleanup lacks a normal exit observation, extended-option delivery is uncertain,
and the breakpoint control is not ready. A separate editor wrapper records a
combase access violation and a distinct abnormal final exit; its subsequent
cases are NOT_RUN. Initial Publisher metadata/reference cases fail a common selector assertion after fresh
reopen, before final readback verification. Normal Publisher exits do not convert
those failures into accepted metadata/reference persistence. A later semantic
identity guard accepts the legitimate selector change and qualifies current
Publisher module/class/form, Description and reference addition/removal with
exact fresh-disk readback and normal exits. HelpFile still changes after reopen.
HelpContextID returns an error despite changing the live value; the unsaved
state is preserved before separately authorized discard and normal cleanup,
without setter replay or subsequent Save. Counts and exact
candidate/event scopes remain in recorded validation; lifecycle and remaining
adapter/UI gates stay open.

On preceding v5, the stable Access guard passes scoped module/class, Description
and reference addition save/reopen. HelpFile fresh-disk readback remains altered
through descriptor, CLR binder and raw IDispatch getters despite an intact
VARIANT canary; HelpContextID runtime metadata confirms I4 but its setter fails.
The read-only getter probe does not add a save or mutation and no heuristic fix
is applied. Reference-removal cleanup and Publisher preparation also remain
failed/unqualified. These are not current-v6 host acceptance.

Preceding-v5 SOLIDWORKS 2025 loading, disposable module/class/form preparation,
compile and verified Save have native evidence. One synthetic run succeeds,
but its strict whole-file preservation trial fails; a separate native reload
verifies persisted source and label without converting that failure into a pass.
The owned instance closes normally after live-source/form backup and one ExitApp.
Complete UI/debugger acceptance, strict SWP preservation and the preceding 2019
abnormal termination remain unresolved. Current-v6 SOLIDWORKS acceptance is NOT_RUN.

The native results below identify their own candidates. Historical
SOLIDWORKS acceptance uses MVID `aaf3a555-76d4-4b18-ae09-1e7b3e085934`;
the earlier Office and SOLIDWORKS 2019 follow-up used `7b5f11d8` from `0ddb0dd`.
Earlier Office candidates remain historical evidence. Source changes made after
each build are not covered by its results. The full qualification branch includes the
implementation changes, tests and this register. Its publication checks are
recorded separately in [recorded validation](test-coverage.md); they do not
replace the operation-specific native acceptance or close the remaining gates.

### Fresh SOLIDWORKS 2019 follow-up

The maintainer authorized autonomous application launch for this campaign.
An isolated Visual Studio utility profile opened SOLIDWORKS 2019 SP5, PID
47344, revision 27.5.0, and the exact PID ROT plus loaded add-in MVID were
independently verified. Native Edit Macro opened only a copied disposable
fixture; module/class hashes, the form label and unchanged disk bytes passed
readback. The designer resize and exact placement restoration passed.

A separately host-created disposable macro passed module/class/form edits,
stale-source refusal, compilation and one product save. Its synthetic marker
ran once and the project unloaded. The strict execution trial still failed
because SWP bytes changed; a distinct, guarded native reload verified retained
source hashes and the label against the observed post-execution file hash.
The cause of the binary change remains unproven. This is not whole-file
preservation or standalone-project qualification.

That historical instance has now closed normally on explicit maintainer
authorization, with exit code zero and no force termination. Before closing,
the exact live source and form were preserved; a fixture whose live whitespace
differed and Saved flag was false received one verified product Save without
source writes or macro execution. The preservation and normal-close records
do not prove fresh-disk reopen or acceptance of the subsequently installed
`d5e25e25` product.

The historical return-to-code trial activated the native module but showed an
empty Monaco shell. Preceding installed candidate `d5e25e25` loaded in a fresh owned
2019 SP5 instance, PID 51376. Native Edit Macro and independent module/class/form
readback preserve the copied file bytes. Designer and returned-code workspace
resize/restoration pass, and reviewed captures show the selected module and
class in Monaco. The class source then differs by an `on` prefix; the maintainer
reports possible keyboard input diverted when the window gained focus. Its
cause is not proven. The live source is retained without overwrite, save or
execution, so strict unchanged-source acceptance remains unqualified.
Its authorized ExitApp request stalled at native debugger heap corruption
`0xc0000374`. Unresolved frame observations include ntdll/ucrtbase/mfc140u/sldappu
but do not establish the originating defect. An authorized forced termination
timed out while the debugger still held the target; the owned debugger was
subsequently stopped after exact utility-solution and sole-target validation.
Final debugger design mode and no debug target are not normal ExitApp evidence;
the exit code is NOT_OBSERVED. The earlier transient PID absence is not the final
cleanup oracle.
Preceding-v5 SOLIDWORKS 2025 has separate partial native evidence in recorded
validation. Its selected class renders in Monaco and code workspace resize and
restoration pass, but a stale closed-project warning is visible. Strict binary
preservation remains failed despite successful post-runtime content readback;
the subsequent owned 2025 cleanup exits normally. Complete debugger and embedded
assistant acceptance remain pending. Q-014 stays open.

The preceding v4 Office campaign aborted during owned Word form preparation.
Its bridge timed out after a form-property request, and the old fixture entered
native cleanup despite uncertain delivery. A noninvasive, nonsuspending stack
observation finds the Word STA in FM20 overlay-window/visibility handling;
this does not establish the cause. Source now records request intent before
dispatch and retains native ownership after uncertain delivery, refusing further
requests, save/reopen and Close/Quit. The retained testhost and Office instances
have now received explicitly authorized forced cleanup; this is not normal
shutdown or a qualification pass. The terminal TRX does not contain a result
for the in-flight Word scenario, and unreached scenarios remain NOT_RUN.
Access save verification
also reports a changed CurrentProject COM identity, while later read-only Saved
observations are true. Removing that guard without establishing a stable Access
identity would weaken dispatch safety. A fresh read-only Access trial establishes
that repeated CurrentProject getters produce distinct retained IUnknown wrappers
while the database path and mapped/selected VBProject identity remain stable.
Its normal exit qualifies the diagnostic only. Source `2e75161` now uses stable
application/PID, database path and mapped/selected VBProject identity while
retaining VBE, source, selection, mode/protection and metadata/reference guards.
Preceding installed-v5 adapter-only save/reopen accepts module/class edits,
Description and Scripting-reference addition by GUID and file, with normal exits.
HelpFile changes to garbled Unicode after disk reopen despite correct live
readback; reference removal cannot reach fresh reopen because initial Quit does
not complete. Its disposable database was preserved before one authorized forced
cleanup without Quit replay. These outcomes do not qualify either operation.
Project HelpContextID setters also remain unresolved in Access and Publisher;
preceding-v5 Publisher metadata/reference acceptance has not been established.
Exact evidence is in recorded validation.

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
| Q-006 | P1 gate | Host tests could report success after uncertain saves, lost forms, stale loaded assemblies or forced termination. | Assert verified persistence, reopen without helper saving, retain form expectation, verify MVID and normal exit. | OPEN. Historical `7b5f11d8` scalar inspection stalled in break mode and cleanup after COM Reset was followed by a combase.dll access violation. Preceding `d5e25e25` passes the complete declared scalar page; the earlier fixture startup failure remains failed preparation evidence. Current `8f2315d` / installed `d8f31d57`, tests `304ee02`, also passes the ordered unsupported page, one Long and full declared scalar page: Long, String and Boolean values, unsupported array/Variant/object refusal, terminal native phase evidence, unchanged identity/source/selection/mode and normal exits. These scoped passes do not explain the historical stall/crash or qualify other runtime local types. Office forced cleanup and SOLIDWORKS abnormal termination remain failed lifecycle evidence. Exact results are in recorded validation. |
| Q-007 | P2 | External COM temporaries in host fixtures could outlive Quit and obscure shutdown results. | Explicitly release owned collections/windows/commands and require normal exit; distinguish forced termination from an independent crash. | Corrected Office and Monaco fixtures passed normal-exit checks; no forced termination counted as success. |
| Q-008 | P2 | An obsolete UI test called a removed Markdown rendering API. | Exercise the current native Markdown view, streaming, transcript virtualization, settings and Git views with isolated state. | Harness corrected; detached UI pass and captures recorded. |
| Q-009 | P2 | NuGet runtime license/notice payloads were absent from the build output. | Deliver exact upstream texts with provenance and verify output hashes. | Payload added and final Debug/Release delivery hashes verified. |
| Q-010 | P3 | WebView2 creates persistent per-PID profiles without a retention policy. | Define ownership and safe cleanup only after browser processes exit; preserve active/private state. | CLOSED for the newly owned editor-profile lifecycle on candidate `9924660b`: each editor has a unique environment folder; retirement and its matching BrowserProcessExited notification are both required before asynchronous cleanup. Detached real WebView2 acceptance verifies one editor can close and remove its profile while a second remains usable. Unknown legacy profiles, missing exit evidence, links and locked files are retained; no automatic legacy purge or crash recovery deletion is claimed. See recorded validation. |
| Q-011 | P1 gate | Word/PowerPoint adapter acceptance required project-access prerequisites. | Run verified adapter-only save/reopen under a maintainer-approved host configuration. | CLOSED for adapter-only save/reopen on candidate `7b5f11d8`: Word and PowerPoint preserve the pending module/class sources and form/Label, with no post-adapter helper Save, verified disk readback and normal host exit. This scope does not explain the older Word form-creation failure or qualify every host operation. Exact evidence is in test-coverage.md. |
| Q-012 | P1 gate | Access/Publisher host-document save adapters lacked native acceptance. | Implement and qualify an adapter, or explicitly narrow the release contract for this operation. | PARTIAL. Historical `7b5f11d8` Publisher save/readback passed; Access disk sources persisted but original responses remained uncertain. Preceding `d5e25e25` deferred verification then refused transient CurrentProject wrapper identity in the interrupted Office campaign. The read-only diagnostic establishes stable database path and mapped/selected VBProject despite distinct retained wrapper IUnknowns. Product `2e75161` / preceding installed `f9a36c85` corrects this guard without weakening owning-PID, VBE, mode, protection, selection or source/metadata checks. Fresh adapter-only module/class, Description and Scripting-reference addition by GUID/file pass exact save/reopen and normal exits. Preceding-v5 HelpFile is altered after reopen through descriptor, CLR binder and raw IDispatch, with intact raw VARIANT canary; HelpContextID runtime type information confirms I4 but its setter fails. Reference removal fails initial host exit and disk reopen is NOT_RUN. Current `d8f31d57` Access HelpContextID bridge/external CLR setters return 321 after a verified Save, but Quit nonexit prevents disk reopen. The fresh raw IDispatch PUT case on owned Access PID 38712 returns HRESULT 0 with VT_I4 321 and intact VARIANT canaries; all getters and one verified Save succeed, but its single Quit does not produce process exit, so fresh-disk reopen is NOT_RUN. A stable database copy precedes separately authorized forced exit -1; no Quit, setter or Save is replayed, and normal cleanup/persistence remain unqualified. External Access HelpFile production CLR and raw PUT both return normally but fresh-disk readback remains altered. Current Publisher startup and semantic reopen guards correct the fixture's no-code-pane and name-to-path selector assumptions; the original preparation/assertion failures remain failed. Adapter-only module/class/form, Description, and Scripting-reference addition by GUID/file and removal now pass exact fresh-disk source/metadata/reference readback and normal exit 0. Publisher HelpFile disk storage is correct in both CP1252 PROJECTHELPFILEPATH fields and the PROJECT text stream, yet the exposed fresh getter is altered; loading, BSTR construction or marshaling cause remains unproven. Publisher HelpContextID returns error 0x9CFD3148 despite terminal live value 321 and Saved=false. Live sources/identity/properties are preserved before one separately authorized owned discard and normal exit 0, without setter replay or subsequent Save; the failed mutation and disk persistence are not qualified. No heuristic or causal product fix is inferred. Access/Publisher metadata, lifecycle and remaining adapter scopes stay open; exact terminal evidence is in recorded validation. |
| Q-013 | P1 gate | Classic Outlook initially had no configured profile. | Qualify a read-only scenario in an explicitly configured classic profile without modifying mail or production VBA. | CLOSED for read-only startup/metadata: `outlook-accepted/native.trx`, MVID `ce19a20c-9708-4c17-b998-f3415b8e6303`; exact PID, project inventory, scoped debug state and environment passed, normal exit 0. No account configured by automation, no mail read/sent or VBA mutation. |
| Q-014 | P1 gate | SOLIDWORKS versions require independent native acceptance. | Qualify load, UI, disposable module/form operations, compile/debug, persistence and cleanup in each explicitly selected version. | PARTIAL. Historical aaf3 existing-Type100 save and saved-copy module/class/form readback passed independently in 2019 SP5 and 2025, with normal exits (Q-021); historical designer-resize failure remains recorded. Preceding `d5e25e25` passes 2019 load, copied native Edit Macro/source/form readback, Monaco return-to-code and designer/code resize. Later live class-source drift remains unproven and unmodified. Its ExitApp stalled at native heap corruption 0xc0000374; authorized forced/debugger cleanup is not normal shutdown, and exit code is NOT_OBSERVED. Preceding `f9a36c85` passes 2025 load, disposable module/class/form preparation, compile and verified Save. One synthetic run and a separate native reload verify marker/source/label, but the original whole-file-preservation trial fails. Monaco code rendering and resize/restoration pass with a stale closed-project warning still visible. Owned 2025 cleanup exits normally after backup and one ExitApp; complete lifecycle/debugger/assistant acceptance remains open. |
| Q-015 | P2 gate | The initial instrumented suite timed out on post-step observation; its direct relationship to Q-005 is not proven. | Repeat the complete suite on the corrected source and retain failures/skips honestly. | CLOSED for current source `8f2315d` / installed `d8f31d57`: the complete instrumented suite passes with unchanged product hash and independently verified individual TRX outcomes. Preceding Monaco-status source `d7a1c75` / `6a74af33` and earlier product `2e75161` / `f9a36c85` with tests `f0874e6` retain separate completed passes. The original v5 complete run remains failed with explained fixture-packaging/Git-path defects. Exact counters, below-target managed coverage and terminal evidence are in recorded validation. Native/provider opt-ins remain separate gates; a later product binary requires its own full run, and historical observation-timeout causes remain unproven. |
| Q-016 | P2 | Shift+Tab accepted a composer suggestion instead of allowing backward keyboard navigation. | Leave backward navigation unhandled while preserving plain-Tab suggestion acceptance. | Fixed; focused regression passed. |
| Q-017 | P2 | The French welcome card is clipped in the narrow native chat panel. | All welcome actions remain visible or reachable by normal scrolling, including after resize. | Fixed by measuring the Designer table at its available width; failing-before/passing-after geometry checks and inspected native recapture confirm all actions are visible. |
| Q-018 | P2 | Monaco tests assumed immediate reference refresh, stopped only one of two synchronization timers and accepted an unrelated pending-edit refusal as a save cancellation. | Respect the bounded reference cache; isolate draft synchronization and prove the native save callback actually runs; verify exact native/readback text. | Stronger fixtures passed the final native batch, including exact VBE canonical readback and renderer reconciliation. |
| Q-019 | P2 | Git for Windows rejects an absolute `--git-dir` beyond its internal path buffer, even when `core.longpaths=true`; the full suite exposed this with a deep cache path. | Keep the verified working directory and pass the repository path relatively; prove real local Git operations at the failing path length without global configuration changes. | Fixed with relative --git-dir and process-local core.longpaths; real Git long-cache commit/readback and the originally failing binding scenario passed focused validation. The final complete managed rerun passed. |
| Q-020 | P1 | In SOLIDWORKS 2019, creating a standalone project succeeds but its unsaved `FileName` getter throws, breaking collection readback and reporting an uncertain result. The same assumption prevents initial standalone save detection. | Recognize the narrowly identified unsaved standalone state; preserve refusal for unrelated getter errors and saved projects; validate first save and reload. | Corrected creation, first SaveAs, synthetic module/class/form, compilation, save and close completed on native 2019/353d. Reload did not complete: the first run was interrupted by the user; the second host exited unexpectedly during Open (Q-030). The original getter HRESULT remains NOT_OBSERVED. No macro ran. |
| Q-021 | P1 gate | The native Save command for an existing SOLIDWORKS Type100 SWP can return before the host's Saved flag becomes true, causing premature uncertain results. | Invoke Save once, yield to the owning UI thread, preserve identity/selection/path/source/metadata guards, and independently read persisted module/class/form content in both versions. | CLOSED for the tested existing-SWP Type100 path on `aaf3a555-76d4-4b18-ae09-1e7b3e085934`. Bounded asynchronous verification returned Verified=true/Uncertain=false in 2019 SP5 (PID 47384, revision 27.5.0) and 2025 (PID 1236, revision 33.1.1). Exact saved byte copies opened once through native Edit Macro retained module/class hashes and the form label: `solidworks/type100-2019/async-save-02/saved-copy-content-verification.json` and `solidworks/type100-2025/async-save-01/saved-copy-content-verification.json`. Pending saves are blocked across sessions on the owning thread; timeout or changed state remains uncertain without retry. Type100 SaveAs, signatures and reopening the original file after a full application restart are not qualified. Earlier 1541 uncertain results remain historical failures. The separate 2019 async-save-01 runtime marker succeeded, but its subsequent SWP hash change failed the unload scenario; that result is not reload proof. |
| Q-022 | P1 | Word VBProject.FileName throws before Save and can identify temporary VBA storage afterward; equating it to Document.FullName rejected a correctly matched document. | Use Word's identity-matched Document.FullName while preserving PID, COM identity, format, source, saved flags and file-byte guards; keep PowerPoint project-path checks strict. | CLOSED for the reproduced defect: native diagnostic proved DirectoryNotFoundException/0x80070003 then ~WRL0002.tmp, with unchanged document path/source and saved flags true. Red/green regressions and `office-accepted/native.trx` on ce19a20c prove adapter save/reopen; Q-024 remains separate. |
| Q-023 | P2 | The Outlook qualification fixture omitted the required Project field for debug_state despite successful add-in loading. | Resolve the unique project explicitly and retain normal-exit and metadata assertions. | CLOSED: corrected fixture passed `outlook-accepted/native.trx` on ce19a20c, with exact project metadata and normal owned-process exit 0. |
| Q-024 | P1 gate | Word VBA backing paths are unsuitable for persisted document identity; complete Git capture also fails during native export into GitTemporary. | Resolve the unique Word document by PID/IUnknown and canonical FullName; preserve raw FileName, refuse stale bindings and qualify capture separately. | PATH CONTRACT VERIFIED on 353d: `artifacts/qualification-v1/word-git-native/353-native-03/hosts/Word/2cb46193741e415ab5773c1539745096/qualification.json` proves two projects named Project resolved by their own DOCM paths, native selection, stale SaveAs binding refusal and unchanged source in the other document. Word PID 50144 exited normally. COMPLETE GIT REMAINS OPEN: production Capture fails with 0x800AC35C, while bridge and external-STA exports to separate artifact files succeed with identical bytes. The test remains failed; actual chat/Git UI opening is not established by the bridge selection evidence. No implicit scope/grant migration or Normal modification. Source `b9dc446` reaches successful independent bridge export and reference reading, then the test helper rejects the native Normal template reference before menu intent. Original Close/Quit exits 0 and registration is restored. The corrected reference helper still needs a fresh trial; no historical export repair is established. The owning-thread menu and chat acceptance gates remain open; see recorded validation. |
| Q-025 | P1 gate | Native Excel form fitting could corrupt memory or prevent shutdown. The installed x64 WinForms Com2PropertyDescriptor.SetValue allocates a 16-byte VARIANT buffer, while its marshaler writes 24 bytes; scroll fitting used this setter. | Avoid that native descriptor setter while retaining validation, conversion, readback and existing control restrictions. Qualify both fitting and the generic scalar dispatch with normal host exit. | CORRECTED on candidate 353d for the tested paths. `scalar-excel-native/native.trx` passes complete fitting, arrays and persistence; the independent scroll-only trial exited normally. `native-scalar/native-353d/qualification.json` verifies project Description, module Name with unchanged source, and Label BackColor through native getters; its reviewed VBE capture shows the yellow Label and renamed module, followed by exit 0. Earlier crashes, dumps and forced cleanup remain failed evidence; array/persistence failures are not automatically attributed to this defect. Other controls and complete Git workflows are not thereby qualified. |
| Q-026 | P1 gate | Activated Excel formatting-options scenario refused a stale revision and its restoration also failed. | Identify the revision drift and rerun the instrumented mutation/restoration scenario without relaxing revision checks. | OPEN; cause not proven. The original font and normal-text foreground were restored through the native bridge; complete options revision matched the retained baseline and Excel exited normally. Test diagnostics now retain failed requests, before/after observations, and every distinct restoration failure. Evidence: `artifacts/qualification-v1/excel-options-recovery-restore-02/report.json`. The instrumented ce19 rerun (`excel-options-instrumented/options.trx`) passed the full scenario, complete baseline restoration and normal exit. The earlier drift was not reproduced, so its cause is not claimed fixed; successful-run snapshot attachments were not retained by that VSTest invocation. The follow-up fixture now records every command and successful before/after state in an optional durable directory and restores each palette using its observed category explicitly. The current 096b2e2b native rerun failed while selecting a Code Colors category before the first font mutation. Durable intermediate readbacks matched the complete baseline, but the final read failed; complete restoration is therefore unqualified. This differs from the earlier stale-revision failure and does not prove its cause or correction. Follow-up selection diagnostics now retain the exact requested category, observed UIA category and native list index on a mismatch; a regression verifies the single selection/notification is not replayed. Native cause and full mutation/restoration acceptance remain unproven. Current v6 command-ledger analysis locates the extended-options failure in a read before the intended toggle; its redundant On-to-On finally write and normal owned exit do not qualify the missing mutation or Docking iteration. Successful retained response prefixes omit OptionsVersion, so their hashes are not complete restoration proof. No cause is established. See recorded validation. |
| Q-027 | P1 gate | UserForm export to the production GitTemporary parent fails natively, and raw FRX serialization varies in timestamp/padding bytes. Opaque FRX contents also need bounded preflight checks. | Qualify native capture/import and recovery, preserve real FRM/FRX through GitHub, and reject missing or manifestly invalid companions before mutation. | OPEN. Candidate `ce19a20c`: positive round-trip attempts stopped at fixture dimensions (corrected), then `Capture` / `Export` (`0x800AC373`) before push/import. Differential probes export the same form successfully under TEMP and E: but fail under GitTemporary, including installed-bridge export. Equal-ACL plain/EFS siblings both work; EFS alone and GUID segment length are not sufficient causes. Consecutive unedited exports differ only in identified CFB timestamp/MS-OFORMS padding fields. Owned probes exit normally. Separately, `userform-git-transport-20260929225354/transport.json` proves production Git commit/push/fresh-fetch with exact FRM/FRX hashes on retained branch `qualification-userform-20260929225354-93ed53dc`, main unchanged; this form-only transport is not native project import qualification. Checkpoint/backup remain intact; no permanent loss is established and no corrupt snapshot was published. Separate candidate `353ddf2a` bridge import from the exact Git-fetched FRM/FRX, native controls/both Caption properties, save/reopen and owned-process shutdown are verified in `userform-fetched-import-20260929231432`; source/import/reopen captures were reviewed. VBIDE adds exactly one leading CRLF to code (content preserved, exact source text differs). This does not qualify the blocked Git coordinator or raw FRX comparison. Follow-up candidate `7b2423c3` reproduces unequal unedited FRX captures and pre-import refusal in the historical jvc environment; production Capture succeeds under GitTemporary there. Those results do not establish a successful export on the current MOTHER environment. The new bounded OLE/CFB preflight accepts the native exports and refuses a nonempty signature-corrupted companion before native mutation. Follow-up candidate `82942b5d` now compares logical CFB contents and only structurally identified MS-OFORMS padding; raw FRX transport remains exact. Its owned Excel trial passes repeated captures, guarded form import with exact FRM readback, production local-Git checkpoint restore, verified backup, explicit rollback/final restore, helper save/reopen and normal exit. VBIDE adds one leading code line and omits the exported EOF terminator from CodeModule.Lines; the correction removes only the proven extra line, preserving full final snapshot verification. Unsupported form layouts retain exact logical stream comparison. Fresh 7b5f11d8 probes on MOTHER fail below LocalAppData, LocalAppData/VBAi and GitTemporary but pass below TEMP. Observed testhost/Excel token fields are identical, and generic EFS/volume controls pass. The crossed child-DACL trial on historical 7b5 succeeds under TEMP with the LocalAppData DACL but still fails under LocalAppData with the TEMP DACL; existing parents and EFS remain unchanged, and both owned hosts exit normally. Child-DACL differences alone do not explain the failure. A subsequent exact `d5e25e25` Excel export trace, after a measured disposable-helper CDB preflight, records paired native VBE CreateFileA and managed-parent open failures with path/name-not-found statuses while the same GUID child exists and root synthetic write/read succeeds. The single export fails; debugger detach and host normal exit do not qualify export. An earlier trace setup failed before attachment/export due to UTF-8 JSON decoding. Cause remains unproven, and neither trace is current-v6 native acceptance. Current-candidate GitHub transfer, embedded bridge/UI qualification, additional controls and historical export failures remain open. Current `d8f31d57` read-only owner-STA diagnostics see both synthetic LocalAppData/TEMP GUID directories and files with attributes matching the testhost, with normal exits. The primary-token fallback follows ERROR_NO_TOKEN with matching user SID/integrity/AuthenticationId; observations in different processes do not establish the effective context throughout export. Separate fresh COM-activation exports still fail below LocalAppData, LocalAppData/VBAi and GitTemporary while TEMP succeeds. Independent explicit /x /automation bootstrap trials pair the owner-STA diagnostic immediately with one export below LocalAppData or TEMP; both pass with retained FRM/FRX hashes, source and Label properties, and normal exit 0. No permissions, attributes, trust policy or token are changed and no failed export is replayed. These distinct launch contexts are scoped export acceptance, not a controlled causal explanation of the COM-activation failures. Complete Git capture/import/recovery and current-provider transfer remain open. See recorded validation. The current-v6 complete prepared explicit-launch layout matrix reaches production GitTemporary capture for every layout but fails the first unchanged-snapshot comparison. Source/FRM/manifest hashes are identical in the retained Label/Button pair; FRX differs. All owned hosts exit normally. Git transport/import/recovery and Save/reopen phases are NOT_RUN after that guard; no normalizer exception is introduced. |
| Q-028 | P1 gate | The real Ollama chat UI scenario intermittently failed to observe streamed text while busy. The ce19 diagnostic ended with the visible fallback `No text response.`, without an observed text fragment or refused tool call. | Preserve the streaming, cancellation and next-send assertions and retain exact synthetic wire evidence. Explain the earlier empty response before claiming a reliability correction. | PARTIAL. Candidate bbb6e6f passes the complete managed gate and four strict real-provider cases on the selected qwen2.5:7b-instruct configuration (temperature 0, top-p 0.8, context 8192). Exact synthetic arguments, detached UI streaming/Stop/next-send, cancellation recovery and one read-only native Excel call pass; source is unchanged and owned hosts/helpers close. Embedded-host acceptance and the historical intermittent cause remain unresolved. Earlier evidence is preserved: `ollama-ui-wire-353d/ollama-ui.trx` passes the unchanged strict scenario on 353d: visible text while busy, cancellation and the subsequent visible complete reply. Wire capture records a refused synthetic tool call and subsequent text; no LLM production fix was applied. The passive wrapper preserves transport configuration but changes timing, so this success does not explain the ce19 failures in `activated-ollama-ui` and `ollama-ui-diagnostic`. Source follow-up adds bounded content-free stream metadata (chunk counters, filtered terminal reason and complete-empty/text/tools/error outcomes) to future synthetic UI diagnostics. It does not explain the historical empty response, replay a request or change provider parsing. Current frozen b60996c with test source ce3d9da reproduces complete-empty in the detached UI: captured SSE already contains empty deltas, stop and DONE with no text or tools; diagnostics and visible fallback agree. The headless tool-roundtrip also fails its scalar argument assertion without a captured body; cancellation and the subsequent request pass separately. That earlier activated aggregate remains failed, and its backend/model cause is not established. Detached UI with simulated VBE only; no embedded-host qualification is inferred. |
| Q-029 | P2 | Empty-project assistant state re-enables Send after a prompt edit because `UpdateBudgetControls` omitted the scope predicate. | Keep Send disabled after scope loss, including draft and busy-state changes; preserve active-turn Stop and normal sending with a selected scope. | Reproduced by a real detached control event in `empty-scope-red/red.trx`. Source correction centralizes the predicate while preserving cancellation. The expanded detached real-control regression passes scope loss/loading, draft, busy, pause and stop-request transitions; native embedded-host/UIA acceptance remains separate. `EnsureCurrentScope` still guards dispatch; no privacy bypass is demonstrated. |
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

## Latest scoped follow-up

Q-026 remains open. Fresh current-v6 disposable checkbox trials now verify
Editor Format and Docking mutation, closed-dialog readback, guarded restoration
and normal owned-host exit. The final trial retains complete baseline/restored
options revisions in durable phase summaries. Earlier successful trials lack
durable summaries, and one copied driver selected an older test build; the
additive provenance correction preserves those original records. These outcomes
do not explain the original pre-write connection EOF or qualify every category.
The later full Format trial confirms a distinct defect: the empty-size refusal
returns before its owned Options dialog is confirmed closed. The dialog is
visible in the immediate observation and absent in a later read-only observation.
The guard retains the exact Excel
without another dispatch, restoration or cleanup. Remaining Format phases are
NOT_RUN; this refusal defect does not establish the earlier EOF's cause.

Q-027 remains open. The uninstalled combined candidate now compares all retained
layout pairs consistently in an offline production-snapshot analysis, including
Frame/MultiPage through bounded atomic storage-graph validation. This does
not turn the failed installed-v6 native matrix into a pass or execute its later
Git, import, recovery and persistence phases. The installed DLL is unchanged.
Its complete default managed gate passes. Native import/recovery, persistence
and current-provider transport remain unexecuted for that candidate. Scalar
failure-phase diagnostics preserve uncertain setter outcomes without replay;
they do not establish a causal fix for the Access/Publisher metadata failures.
The later uninstalled Options cancellation candidate fails its complete managed
gate on synthetic native-probe ownership fixtures. A focused pass does not
qualify that assembly, and it is not deployed.
Those fixture failures are explained and corrected without weakening ownership
guards. The next combined candidate verifies exact closure for reads and Accept,
preserves primary/cancellation errors and passes its complete default managed
gate. It remains uninstalled: the original native Format failure and earlier
revision/EOF causes remain open. Private CFB guard contracts complete managed
coverage of that preflight, without claiming native import/recovery acceptance.
Exact identities, counts and artifact paths are recorded in
[test coverage](test-coverage.md).

The latest combined candidate additionally refuses uncertain recovery marker
metadata and existing invalid entries before overwriting recovery refs or
starting rollback. Completion deletes a confirmed regular marker once and
verifies absence; original and completion errors remain distinct. Managed/local
regressions reproduce the old directory-marker defects and pass on the fix.
These safeguards are not a causal correction for historical native export,
persistence or lifecycle failures. The installed DLL remains unchanged.

A source/retained-evidence review confirms that the Word exact-path Git fixture
executes `VbaGitProject.Capture` on the external test STA, as its original report
explicitly states. The prepared sequential UserForm GitHub fixture also runs
its coordinator on the test STA. These scopes cannot qualify the embedded
owning-VBE-thread Git path. Both Word diagnostic exports succeeding through
different routes/destinations do not establish a thread/path cause for the
historical capture failures. Complete owner-dispatched production Git acceptance
is still required for Q-024/Q-027; no native operation was replayed by this review.

## Evidence location

Machine-local TRX, logs, screenshots, manifests and detailed delegated reviews
are under `artifacts/qualification-v1/`; detached UI captures are under
`artifacts/ui-review/screenshots/`. These ignored files are not public downloads.
The maintained [compatibility matrix](compatibility.md) records operation-specific
host outcomes; [testing](../tests/README.md) supplies the standard commands.

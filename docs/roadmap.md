# Roadmap

[Documentation](README.md)

The roadmap describes direction, not promised delivery dates. Current behavior
belongs in the guides and [compatibility](compatibility.md); verified results belong
in [recorded validation](test-coverage.md).

## Shared VBE experience

Consolidate the modern editor, live-project agent workflows, code review/recovery,
UserForm tooling and source versioning as one coherent workspace. Improve real
workflow reliability, discoverability, accessibility and performance without
weakening revision or permission checks.

## Product delivery plan

This plan covers the 36 chat/workspace improvements and four additions: an
editable system prompt, skills, MCP connections and the SolidAi bridge. F01-F36 retain the original
feature numbering; F37-F40 identify the additions. All entries below are planned,
not implemented or qualified by this document. The initial source review for this plan used `5eb46e9c` on 2026-10-08.
Recheck the current source and acceptance evidence before implementing each work
item; existing foundations do not establish completion of a feature.

The existing product already has caret-positioned #/@ completion, links for known
references, collapsible activities, conversation search and forks, queues, diffs,
guarded recovery, editor tools and a test explorer. Extend these foundations;
do not replace them merely to introduce a new presentation. See
[chat](chat-ui.md), [editor](modern-editor.md), [testing](vba-testing.md) and
[Git](github-integration.md) for implemented behavior.

The proposed version labels are delivery targets, not announced releases or dates.
Release a milestone only after its acceptance gate; split it into smaller releases
if native-host or provider qualification requires more time. Effort is relative:
S = localized change, M = several components, L = architectural or native work.
These are not calendar estimates and must be refined when issues are scoped.

| Milestone | Proposed target | Outcome and feature IDs | Prerequisites | Exit gate |
| --- | --- | --- | --- | --- |
| M0 | Before feature releases | Shared reference, instruction, extension and activity contracts; persistence migration plan; UX prototypes | Current-source review | Contracts cover legacy sessions, unsaved projects, queues, forks and provider differences; net48 integration approach demonstrated |
| M1 | 1.1 | Reliable composition and navigation: F01-F09, F25 | M0 | Typed, pasted and selected references behave consistently; no lost drafts; correct target in VBE/Monaco |
| M2 | 1.2 | Personalized assistant: F20, F22, F36, F37, F38 | M0, M1 context presentation | Editable instructions and selected skills reach every supported provider path with visible provenance and unchanged runtime guards |
| M3 | 1.3 | Readable, reviewable conversations: F10-F19 | M1; M0 activity contract | Long streamed conversations remain navigable; approvals and partial outcomes are understandable; accessibility scenarios pass |
| M4 | 1.4 | Connected assistant: F21, F23, F24, F39, F40 | M2 configuration/instructions, M3 approvals/activities | Local and remote MCP scopes are independently qualified; context accounting, images and continuation preserve data boundaries |
| M5 | 1.5 | Integrated VBA workspace: F26-F35 | M1 navigation, M3 review; F20 for task templates | A disposable-project workflow links source, tests, forms, debugging, Git and recovery without confusing their states |

M2 and M3 may be developed on separate branches after M1; M4 requires both.
M5 can proceed independently of MCP once its prerequisites are available. M0 is
a bounded foundation, not permission for a broad rewrite of the existing chat.

## Feature backlog and acceptance

### Composition and references

| ID | Deliverable | Milestone / effort | Minimum acceptance scenario |
| --- | --- | --- | --- |
| F01 | Styled, interactive #/@ mentions inside the composer | M1 / L | Mention identity survives edit, undo, copy/paste and draft restoration; removing a mention updates its context consistently |
| F02 | Rich completion with signatures, grouping and relevant ranking | M1 / M | Keyboard and mouse select the same symbol; duplicate names show their project/module; large indexes remain responsive |
| F03 | Resolve manually typed and pasted tags | M1 / M | An unambiguous token gets the same context as a picked token; ambiguous, missing and unauthorized targets are visibly distinguished |
| F04 | Consistent reference links in replies, activities and relevant inline code | M1 / M | Identical visible tokens in different messages retain their original target; unavailable targets report an actionable state; streaming integration completes with F12 |
| F05 | Project navigation that preserves the draft | M1 / S | Clicking a project reference does not replace the composer; explicit insertion occurs at the caret |
| F06 | Source peek and choice of VBE or modern editor | M1 / M | Preview is read-only; opening either editor selects the intended authorized source and reports navigation failure |
| F07 | Explicit refresh of stale references | M1 / M | A changed source offers comparison and refresh; the send path still rejects stale revisions |
| F08 | Context inspector with origin, ranges, freshness, access and size | M1 / M | Users can remove or narrow context and inspect exactly those explicit attachments; total provider-context accounting is identified separately under F23 |
| F09 | Expanded composer, send-key preference and independent Stop action | M1 / M | Long drafts persist; completion consumes Enter before send; Stop remains available while drafting a queued message |
| F25 | Unified chat, test and editor navigation | M1 / M | A test or diagnostic opens the correct Monaco module/range or explicit VBE fallback; navigation back restores the previous location |

Implement F05 and the resolver/identity model for F03-F04 first. Then add F02,
F06-F08 and F25. Choose the rich-input implementation for F01 only after checking
WPF/WinForms hosting, IME, spell checking, undo and accessibility; F09 shares that
input surface. Keep clipboard text portable rather than dependent on private tags.

### Reasoning, activity and review

| ID | Deliverable | Milestone / effort | Minimum acceptance scenario |
| --- | --- | --- | --- |
| F10 | Persistent compact/summary/detail reasoning preferences | M3 / M | User collapse choices survive refresh; providers without public reasoning show no fabricated reasoning |
| F11 | Filterable activity timeline with target links | M3 / M | Reads, edits, verification and approval waits are distinguishable; filtering or collapse does not hide failed or uncertain outcomes |
| F12 | Progressive Markdown rendering | M3 / L | Partial fences, lists and links settle correctly without losing selection or scroll position; F04 links work during streaming |
| F13 | Visible actions on code blocks | M3 / M | Copy, open source and attach a selected excerpt act on the displayed block and correct project; no implicit application of generated code |
| F14 | Structured tables and expanded long-output views | M3 / M | Columns remain readable at narrow widths; copy preserves rows; truncation is explicit and does not create misleading completeness |
| F15 | Human-readable approvals with available diffs | M3 / L | Exact operation and target are shown; approve/decline paths revalidate live revision and scope before dispatch |
| F16 | Structured completion and interruption summaries | M3 / M | Applied, compiled, tested, saved, pending and uncertain states reflect recorded results; resume does not replay uncertain mutations |
| F17 | Reading density, zoom and keyboard/accessibility improvements | M3 / M | Focus, accessible names, zoom and reading position work across transcript virtualization and supported DPI settings |
| F18 | In-conversation search with navigable matches | M3 / M | Search reaches older virtualized entries and collapsed sections, shows counts and highlights the selected occurrence |
| F19 | Editable prompts through forks and branch comparison | M3 / L | The parent remains intact; branches retain provenance and privacy scope without acquiring ownership of parent rollback actions |

Implement the activity/result model behind F11/F16 before its presentation.
F15 supplies the common approval surface later used by F39. F12 must retain the
existing streaming batching and virtualization behavior. Accessibility applies to
every milestone; F17 adds the broader user-facing controls and qualification.

### Personalization and connected context

| ID | Deliverable | Milestone / effort | Minimum acceptance scenario |
| --- | --- | --- | --- |
| F20 | Reusable task templates with parameters | M2 / M | A template prepares an editable draft and explicit context; it does not send or execute on selection |
| F21 | Image/screenshot attachments with preview and crop | M4 / L | User sees what will be transmitted; supported adapters receive the image; unsupported/unknown capability is handled explicitly without silently dropping it |
| F22 | Structured, selectively attached project notes | M2 / M | Existing notes migrate without loss; conventions, decisions and open questions can be attached independently; temporary projects remain in memory |
| F23 | Provider usage and context-budget display | M4 / L | Reported usage and estimates are labeled separately; explicit attachments are not presented as total context; unavailable prices/usage remain unknown |
| F24 | Editable summaries for long-conversation continuation | M4 / L | Users inspect the summary and source links before a new conversation starts; historical claims do not become current source state or transferable action permissions |
| F36 | Guided setup and actionable prerequisite diagnostics | M2 / M | Users can resolve provider, model, project and component issues; availability and observed qualification remain distinct; later skills/MCP/image diagnostics extend this view |
| F37 | Editable system prompt with scope, profiles and reset | M2 / L | New, resumed, queued and forked turns use the recorded effective instructions; save/cancel/reset and failed persistence behave consistently across adapters |
| F38 | Skill manager and explicit skill activation | M2 / L | Import, inspect, enable, update and disable a local skill; record the selected version; missing resources and conflicting instructions are visible |
| F39 | MCP connection manager and guarded integration | M4 / L | Add/test/enable/disable local and remote servers; tools, selected resources and prompts preserve consent, scope, lifecycle and per-provider capability limits |
| F40 | Same-host SolidAi-to-VBAi MCP tool bridge with one LLM orchestrator | M4 / L | SolidAi discovers and calls guarded VBE tools in the verified VBAi instance; no VBAi provider call is made. Preserve project/revision checks, approvals, correlated outcomes and cooperative cancellation |

Implement F37 before F38 and F20 so instructions are composed once. F22 uses the
same context provenance but remains user data, not an implicit instruction layer.
F23 accounts for templates, skills, MCP schemas/results and provider history where
observable. F24 depends on F18 source navigation and the F37 snapshot contract.

### Integrated workspace

| ID | Deliverable | Milestone / effort | Minimum acceptance scenario |
| --- | --- | --- | --- |
| F26 | Editor action palette for existing guarded tools | M5 / M | Supported rename, symbol, bookmark and navigation actions expose preview and applicable limits without requiring a chat request |
| F27 | Clear draft/synchronized/saved/committed/published states | M5 / M | Native edits, failed save, local commit and failed push produce distinct observed states; unknown persistence is never displayed as saved |
| F28 | Analyze a test failure in chat | M5 / M | A chosen failure prepares source, assertion and run revision; review and explicit rerun form one traceable workflow |
| F29 | Persistent test runs and revision comparisons | M5 / M | Reports survive restart for saved projects; comparisons identify stale source and differing test sets; retention is manageable |
| F30 | Revision-bound coverage navigation in Monaco | M5 / M | Procedure-entry results map to the measured source; stale/incomplete reports are visible; no statement/branch percentage is inferred |
| F31 | Usages and caller-impact view | M5 / L | Resolved uses navigate correctly and can be attached to chat; dynamic or ambiguous calls remain explicitly unresolved |
| F32 | Visual review of UserForm changes | M5 / L | Layout/tab-order changes show a preview tied to the current form version; control/property/event navigation and partial recovery are explicit |
| F33 | Git conflict resolution by block in Monaco | M5 / L | Ancestor/local/incoming/result are reviewable; accepted blocks still pass existing merge and native-import guards |
| F34 | Unified recovery center | M5 / L | Drafts, checkpoints and backups retain origin and scope; comparison precedes restoration; partial or conflicting recovery is accurately reported |
| F35 | Selectable debugging snapshot for chat | M5 / M | Snapshot records time, location and inspected values; shared-context consent is preserved; reading and expression evaluation remain distinct |

F27 supplies the state model for F34. F28 precedes the combined F29/F30 workflow;
F29 stores source identities used by comparisons and coverage. F31 requires a
bounded binding/ambiguity design before expanding existing syntax indexing.

## Prompt, skills and MCP design gates

### F37: editable system prompt

The current instruction source is
[LlmVbeContext](../src/VBAi/Llm/Chat/LlmVbeContext.cs). Chat send, HTTP resume,
session reconstruction and forks consume it; the Codex adapter already records
an applied-instruction hash. This is a foundation to extend, not an existing
end-user prompt editor.

The proposed editor exposes the actual customizable assistant instructions,
initialized from a versioned VBAi default, with profiles, import/export, validation,
a comparison against defaults and Reset. Users can edit role, style, workflow and
coding conventions. Display the effective composition and the origin of each
layer rather than offering only an opaque appended text box.

Define three configuration scopes: global profile, project override and
conversation override. A narrower explicit selection overrides the inherited
customizable profile; an empty field must distinguish inheritance from intentional
empty customization. Selected skill instructions are composed in a documented,
stable order. Project notes and attached content stay separately identified.

Keep the VBE host contract, encoding facts and access rules in a visible,
non-editable application layer. Enforcement remains in runtime code: editing a
prompt cannot grant access, bypass revisions or enable a disabled tool. The goal
is meaningful customization without making prose the security boundary.

Capture effective text, profile/skill versions and a hash at the turn boundary.
Queued messages retain their captured configuration; changed or revoked access is
revalidated at dispatch. Edits apply to future turns, never half of a running
turn. A fork records its inherited configuration. A provider that cannot safely
update instructions must offer an explicit new conversation with reviewed context,
not silently continue with old instructions or replay actions. Resuming a paused
turn retains its original configuration; changing it requires a new turn.

Use versioned, backward-compatible persistence. Existing sessions inherit the
built-in profile until explicitly customized; temporary project/session overrides
remain in memory. Preserve detached Save/Cancel behavior and secret-free exports.
An application upgrade must not overwrite custom profiles; offer a default diff.

### F38: skills

The proposed VBAi skill format starts with a local folder containing `SKILL.md`
and declared supporting resources. Define and version the supported subset;
compatibility with another product's skill ecosystem is not implied.

1. Deliver create/import, preview, edit, enable/disable and remove for local skills.
   Display name, description, source, version/hash, scope and required capabilities.
2. Add explicit activation from the composer and task library. Show active skills
   before sending and their contribution to context. Names must not collide with
   the existing #/@ reference syntax or built-in slash commands.
3. Resolve resources only inside approved package roots, with bounded sizes and
   defined link/path handling. An update produces a reviewable diff and a new
   version; a selected turn retains its original snapshot.
4. Start with instruction/resource skills calling already permitted VBAi tools.
   Script execution, installation hooks and automatic downloads are outside the
   first skill release. Later automation requires a separate execution contract.
5. Show conflicts or unavailable dependencies instead of silently ignoring them.
   Skill text cannot grant permissions or configure a server by itself. A later
   skill may declare an MCP dependency, but its connection needs explicit setup.

Use the same instruction/context composer across providers rather than relying
on a particular provider's skill loader. The current
[Copilot adapter](../src/VBAi/Llm/Providers/CopilotClient.cs) explicitly disables
native skills and external MCP configuration; changing one flag is insufficient.

### F39: MCP connections

VBAi should act as an MCP client for user-selected services. Its existing named
pipe remains a separate diagnostic interface; this milestone does not turn it
into a public MCP server. Reuse VBAi's tool catalog and activity UI through a
provider-neutral extension dispatcher, with separate server/tool identities and
validated schema mapping. Provider-owned external execution must not create a
second path around VBAi permissions.

| Slice | Deliverable | Acceptance gate |
| --- | --- | --- |
| 39a | Compatibility spike and connection records | Pin supported protocol/SDK versions; demonstrate net48 compatibility or an explicitly packaged companion process; verify licensing, deployment and failure isolation |
| 39b | Local stdio servers and tools | Explicit command/arguments/working directory, test connection, catalog inspection and per-project enablement; bounded startup/shutdown and no unowned process termination |
| 39c | Remote Streamable HTTP servers | Endpoint/authentication setup, applicable OAuth or explicit credentials, expiry/revocation and connection diagnostics; no credential forwarding to unrelated origins |
| 39d | Resource picker and server prompt templates | Preview selected resources before attachment; server prompts prepare an editable request; no automatic crawling or permission elevation |
| 39e | Lifecycle and provider qualification | Disconnect, cancellation, schema changes, duplicate/late results, oversized output, host exit and restart recovery are exercised for each supported route |

These slices may ship as clearly labeled increments; F39 is complete only after
the declared local and remote scopes and selected tools/resources/prompts pass.
MCP defines stdio and Streamable HTTP transports; use the official
[transport specification](https://github.com/modelcontextprotocol/modelcontextprotocol/blob/main/docs/specification/2026-07-28/basic/transports/index.mdx)
and [C# SDK transport documentation](https://csharp.sdk.modelcontextprotocol.io/v2/concepts/transports/transports.html)
as design inputs. Recheck and pin actual supported versions during 39a; do not mix
different protocol revisions or assume the SDK can load into this net48 add-in.

Store credentials separately with the existing user-protected secret mechanism.
Export redacted connection definitions. Local servers receive only the intended
environment and run as owned background processes; connecting one is an explicit
execution decision. An external process is not a sandbox and may have the user's
OS permissions. Surface that scope when configuring it.

Separate permission to transmit selected data from permission to call an external
tool. Classify tool effects through a reviewed local policy; server-provided
annotations alone do not establish trust. Discussion/Plan allow only admitted
inspection; mutating or unknown-effect calls require the applicable Agent policy
and approval. Any permission to an external service is separate from VBE access.
Results and resource content are untrusted data, not authorization.

Bound calls to their project, session, turn and connection generation. Show server,
tool, target, result and any truncation in the activity timeline. Cancel/timeout
does not prove remote rollback. Reconnect may restore the catalog, but must not
automatically replay an uncertain mutation. Settings changes or disabling a
server revoke future admission without misreporting already dispatched effects.

## F40: SolidAi access to guarded VBAi tools through MCP

Status: planned. Added on 13 October 2026; revised the same day after the
maintainer requested one LLM orchestrator to avoid duplicate inference between
SolidAi and VBAi. The preferred route is:

```text
SolidAi conversation and provider -> MCP client -> VBAi tool adapter
    -> guarded VBE dispatcher -> observed result -> SolidAi conversation
```

SolidAi owns the conversation and model calls. VBAi exposes its VBE capabilities
and executes admitted operations on the owning host thread. This route must not
submit a natural-language task to VBAi's chat, start its provider, request model
sampling or duplicate SolidAi's conversation. VBAi's standalone assistant remains
available for users who open it directly. MCP is the tool transport; it does not
make tool schemas, selected source or returned results token-free. Discover
capabilities progressively and return bounded results with explicit truncation.

The current `VBAi.<PID>` named pipe is a diagnostic foundation, not an implemented
MCP server. Its status identifies the host and assembly, but raw bridge commands
do not automatically enforce the chat's mode, project grants and approvals.
An MCP facade must call a shared guarded tool boundary rather than simply expose
those raw commands. Do not use raw `run_procedure` as an external agent shortcut.

Deliver the bridge as one coherent lot:

1. Pin a supported MCP revision and runtime/SDK. Prove net48 compatibility or
   package a local companion with bounded stdio startup and shutdown. Advertise
   only implemented capabilities; do not infer compatibility from registry entries.
2. Bind the client to the intended host PID, process generation, connected VBE,
   protocol version and tool catalog. Verify the named-pipe peer before sending
   data. Different simultaneous hosts must not share project identities.
3. Establish an explicit external-tool context with origin, conversation/turn/call
   IDs, target project/module identity, revision, privacy grants, effective mode
   and approval policy. SolidAi's CAD permissions do not grant VBA access.
4. Reuse the VBE tool catalog, owner-thread dispatch and existing guards for every
   call, including catalog gateways. Revalidate live identity, revision and grants
   before native effects. Discussion/Plan remain read-only. Keep approval and
   execution receipts available to the initiating conversation.
5. Return correlated progress, results and refused, canceled or uncertain states.
   Transport receipt and `Delivered` are not proof of successful macro execution.
   Report editing, compilation, testing and saving separately. Cancellation stops
   future admission; it cannot certify interruption of synchronous COM work or
   reset the global VBE debugger. Never replay an uncertain mutation after a lost
   reply or reconnection.

Acceptance must prove zero VBAi provider calls during the SolidAi route, accurate
scope/approval enforcement, small initial discovery, bounded output and trustworthy
outcome correlation. Exercise missing/disconnected VBAi, wrong PID/generation,
unsupported protocol, busy/loading hosts, stale revisions, denied access,
duplicate calls, cancellation before/after dispatch and lost replies as one bank.
Qualify disposable projects on SOLIDWORKS 2019 and 2025 with frozen binaries;
2020-2024 require their own host evidence. Protocol tests, actual guarded VBE calls
and native host qualification remain separate. Task delegation to a second VBAi
LLM is outside this default route.

## Implementation sequence and release checks

Start with these independently reviewable changes:

The [chat context and dispatch contracts](chat-context-contracts.md) define the
M0 reference, turn, instruction, extension, activity and persistence requirements.
Their compatibility prototypes and product integration remain unqualified.

1. M0: define reference identity, turn configuration snapshots, extension dispatch
   and activity outcomes; prototype rich input; settle the net48 MCP approach.
2. M1: fix draft replacement and surfaced navigation errors (F05/F06), then unify
   typed/pasted references and per-message link identity (F03/F04).
3. M1: add completion, context refresh/inspection and unified navigation
   (F02/F07/F08/F25); finish rich input and composition controls (F01/F09).
4. M2: implement prompt profiles and all provider injection/resume paths (F37),
   then local skills, templates, notes and guided setup (F38/F20/F22/F36).
5. M3: implement activity/outcome contracts and approvals (F10/F11/F15/F16), then
   rendering, search, branches and reading controls (F12-F14/F17-F19).
6. M4: integrate MCP slices 39a-39e and context/usage accounting, then image input
   and reviewed continuation summaries (F21/F23/F24). Implement slices in their
   listed order; revisit the M0 compatibility spike before selecting dependencies.
   F40 is a separate coherent bridge lot with its own guarded dispatch and
   same-host acceptance gate.
7. M5: implement navigation-backed actions, state/test workflows and then forms,
   conflict resolution and recovery. Keep native changes small enough to qualify.

Each feature issue should name its ID, source baseline, dependencies, included and
excluded behavior, migration, UI states, localization, acceptance scenarios and
rollback/fallback. No owner or duration is assigned until implementation work is
scheduled. A large item should become several reviewable PRs, not one release-sized
change. Use `planned`, `in progress`, `implemented`, `qualified` and `released`
as distinct delivery states; the table above currently remains `planned`.

For each affected milestone:

- Run focused managed, UI or JavaScript tests for the changed behavior, including
  existing draft, queue, scope, privacy and recovery regressions.
- Exercise provider adapters independently: new turn, continuation, resume,
  cancellation, missing capabilities and configuration changes. A simulated
  provider pass does not establish authenticated-provider acceptance.
- Check narrow/wide panels, supported DPI, keyboard navigation, accessibility,
  localization, long histories, streaming and unavailable services using realistic
  synthetic data. Set measurable latency/memory targets before performance work.
- Verify storage migration from existing data, failed writes, restart, concurrent
  hosts, transient projects and interrupted operations; never silently discard
  user profiles, history or recoverable drafts.
- Qualify changed native operations on disposable projects with an explicitly
  selected host and frozen binary identity. Record PASS/FAIL/BLOCKED/NOT_RUN;
  SOLIDWORKS remains preloaded and explicitly selected. Follow [testing](../tests/README.md).
- Update the authoritative user/architecture/privacy/provider guides when behavior
  lands, run `python tools/docs/check_docs.py`, and record evidence once in
  [test coverage](test-coverage.md). Feature implementation is not release acceptance.

Unresolved choices to settle during the relevant design gate are the rich composer
control, instruction override/migration details, skill schema and dependency model,
MCP protocol/SDK/runtime and authentication scope, and provider-specific image and
usage support. They do not remove any feature from the backlog.

## Host compatibility

Continue qualifying the shared VBE layer in additional applications. Add or repair
host-specific operations where VBIDE is insufficient, especially document identity,
saving and execution. Extend the qualified existing-document save and native macro workflows to
additional document types and operations. Broaden debugger, control, language,
DPI and real-desktop evidence on the current integrated candidate.

A host matrix should describe observed operations and prerequisites, not restrict
the intended ecosystem to the first applications used for testing.

## Quality and security

Extend the [VBA test explorer](vba-testing.md) beyond its annotated discovery,
shared COM result channel, serial batches and human/LLM reports. Qualify native
callback execution and installed tool-window workflows separately in Excel,
Word, PowerPoint, Access, Publisher, Outlook and SOLIDWORKS. Visio and Project are
outside the current requested qualification scope.

Procedure-entry coverage is implemented on explicit Excel, Word and PowerPoint document copies.
Qualify native copying, preserved original source, mapped probes, independently
calculated counters, host-specific open/close event effects, recovery and incomplete measurements before claiming
production readiness. Broader host-copy adapters and statement/branch coverage
remain future work. UI/source implementation and known host names do not replace
operation-specific evidence. Word saved-file/source matching and owned macro
activation, and Word/PowerPoint open/close handler effects, require native
qualification. Production acceptance remains pending.

Maintain regression coverage for concurrent edits, interrupted operations, privacy
boundaries, transport failures and recovery. Re-measure the current instrumented
scope after implementation changes. Broaden authenticated-provider and real-host
checks separately from simulated protocol tests.

Keep native appearance experimental until its lifecycle, recovery and rendering
are qualified across the intended environments. Preserve usable fallback behavior.

## Distribution

The 1.0.0 setup and uninstaller are unsigned. Qualify installation, occupied-host
handling, upgrade, uninstall and recovery on the packaged candidate. Complete
[code-signing onboarding](code-signing.md) before supplying installers through the
[automatic updater](updates.md), which continues to require trusted signatures.

Before a public release, include the declared [project licenses](../LICENSING.md),
complete the payload's third-party notice review and publish only genuine
compatibility and validation information.

## Community

Keep documentation concise and maintained, provide reproducible issue templates,
and accept voluntary support without exclusive functionality or a service
commitment. A donation channel can be added once the maintainer provides its
official destination; it is not a product subscription roadmap.

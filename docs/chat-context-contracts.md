# Chat context and dispatch contracts

[Documentation](README.md) · [Architecture](architecture.md) · [Roadmap](roadmap.md)

This document defines the M0 contracts needed by the chat roadmap. It separates
existing behavior from the proposed integration. It is a design specification,
not a claim that the M0 or M1 release gates have passed.

## Current integration

The shared VBE layer owns project access and native operations. The chat reads
explicit references through `VbeChatReferences`, builds attachments in
`ChatWindow.Workflow`, and stores turn references with conversation entries.
Queued messages capture reference metadata independently of subsequent edits.
Transcript links resolve against the originating turn, not the latest discovery
index. Project navigation uses a module picker and preserves the draft.
Selected code references retain their module hash. Attachment preparation rejects
a changed or missing hash for both whole modules and procedures; it does not
silently refresh the reference. The send path captures detached references and
draft attachments once, prepares their text, and uses the captured references for
the conversation entry. A preparation failure preserves the draft.

Several gaps prevent a consistent composition contract:

- Manually typed and pasted mentions do not become references through the same
  path as selected suggestions.
- Discovery uses a project's display name even though the native resolver needs
  an exact host path when project names collide.
- Explicit refresh still requires removing the stale reference and selecting it
  again; a comparison and refresh interface remains to be implemented.
- The complete detached turn contract below also includes instructions, provider
  configuration and policy; detached references alone do not establish that gate.

Existing snapshot and navigation regressions are foundations for the changes
below. They do not qualify the remaining composition, refresh or host scenarios.

## Reference identity

A reference has three separate concerns:

| Concern | Required information | Rule |
| --- | --- | --- |
| Destination | Host session, exact project selector, component and symbol identity | A visible label is never sufficient to choose between conflicting destinations. |
| Presentation | Project label, component label, token and origin in the draft | Editing a label cannot silently retarget a captured destination. |
| Source snapshot | Module hash, range and declaration metadata | The attached text and stored reference must describe the same source revision. |

Use the exact `HostPath` returned by project discovery when available. Preserve
the displayed project name separately. Existing `Project` request fields retain
their meaning; this work does not rename protocol fields or COM identities.

For an unsaved project, a name is usable only when the owning live session can
resolve it uniquely. A duplicate name must remain ambiguous. A persisted
reference to a departed unsaved project becomes unavailable; opening another
project with the same name must not rebind it. Saving or renaming a live project
requires a verified identity transition before updating its selector.

## One resolution path

Selected, typed and pasted mentions must use one resolver. Its result records the
original text range, destination, captured revision and resolution state.

| State | Send behavior | User action |
| --- | --- | --- |
| Resolved | Attach the validated snapshot. | Inspect or remove context. |
| Ambiguous | Preserve the draft; require an explicit destination. | Choose a project or symbol. |
| Missing | Preserve the draft and identify the missing mention. | Correct or remove it. |
| Unauthorized | Preserve the draft; do not read its source. | Change access through the existing policy flow. |
| Stale | Preserve the draft; do not substitute newer source. | Compare and explicitly refresh. |
| Unavailable | Preserve the draft; do not adopt another host or project. | Reconnect to the original scope or replace the reference. |

Only explicit reference syntax is resolved. Ordinary email addresses, prose,
code comments and Markdown examples must not become unintended attachments.
Token parsing must account for punctuation, Unicode names and names containing
spaces. Caret and selection positions use the same UTF-16 offsets as the composer.
The final syntax and escaping rules remain a rich-input prototype decision.

Resolution must not execute a macro, alter source, refresh a stale revision or
change approval policy. Native reads still require live project authorization.
Indexes may provide suggestions; they cannot override a refused live read.

## Turn capture and refresh

Prepare one detached turn snapshot before dispatch. It contains the question,
resolved references, attachment text and metadata, selected memory, provider,
model, reasoning setting and execution policy. Request construction, transcript
entry creation and retry recovery all consume that same snapshot.

Preparation is atomic from the user's perspective: a refused reference leaves
the composer, attachments and queue intact. No user entry is added before
preparation succeeds. Changing the composer after dispatch cannot change that
turn's attachments or transcript destinations.

Refreshing context is an explicit operation. It reads authorized current source,
shows the previous and current revisions, and replaces the selected snapshot
only after the user accepts it. Sending must validate the accepted snapshot
again because source can change after preview. Procedure ranges and declaration
identity must be rediscovered rather than reused after a structural edit.

Queues retain their captured context. Editing a queued message produces a new
draft without mutating its originating turn. Forks copy snapshots; later refresh
in one fork cannot change another fork. Navigating a transcript link uses the
stored destination and reports unavailable or changed source explicitly.

## Instruction and provider configuration

Instruction state is separate from project read/write approval. Capture the
effective instruction text, origin and content hash for each turn. Host-specific
capabilities, source encoding and active VBE scope are observed context, not
permanent assumptions about Excel or SOLIDWORKS.

For a persistent provider conversation, compare the effective instruction hash
with the last acknowledged hash. Changed instructions must reach the provider
before the next user message. Record the new acknowledged hash only after the
provider accepts them. An uncertain update cannot be reported as applied or
silently replaced by old instructions.

Stateless providers still receive the required instructions with each request.
Provider caching and billing are transport-specific; suppressing required text
is not a portable token-saving mechanism. Model and reasoning selection remains
in the conversation UI, and unsupported settings produce an explicit state.
Credentials never belong in turn snapshots, transcripts, diagnostics or model
context.

## Extension dispatch and activity outcomes

Skills, local tools and future MCP connections remain separate opt-in extension
sources. The provider cannot grant them additional native access. Every native
dispatch, including catalog gateways, retains the existing project identity,
revision, mode, approval, privacy and owning-thread checks.

An activity records its turn, operation identity, source, lifecycle and observed
outcome. Distinguish preparation, admission, dispatch, completion, refusal,
cancellation and uncertain outcome. Cancellation is not proof that an already
dispatched operation had no effects. An uncertain mutation is never replayed
automatically, and recovery results retain the original failure.

Revoking an extension prevents future admission. It cannot erase previously
dispatched effects or mark them rolled back. Displayed activity summaries describe
observable steps and provider-supplied summaries; they must not fabricate private
reasoning.

## Persistence and compatibility

Keep existing session fields readable. New reference metadata requires an
additive migration with an explicit format/version discriminator. Missing
metadata means unknown provenance, not a current validated revision.

Legacy transcript references remain readable under their recorded identity and
range. Legacy drafts without a revision require an explicit current capture
before they can acquire freshness guarantees. Saving the migrated state must
preserve provider sessions, queues, drafts, forks, bookmarks and unsaved-project
history management. Migration must be idempotent and retain recoverable originals.

The COM add-in continues to target .NET Framework 4.8. No live service, provider
connection or native discovery belongs in a WinForms Designer constructor. Rich
input and extension transport choices require separate compatibility prototypes;
this specification does not claim that those prototypes exist.

Owning-thread marshaling and chat responsiveness are separate requirements.
`BeginInvoke` or a synchronization context cannot keep a chat on the host's
blocked UI thread responsive. The compatibility prototype must include a busy
host operation that does not pump UI messages. A cooperative macro using
`DoEvents` provides a narrower observation and cannot qualify that requirement.
Any execution isolation must keep COM on its owner and transfer explicit data
snapshots rather than live COM objects across its boundary.

## Acceptance matrix

Implement complete scenario groups before executing their tests. Use mirrored
managed tests for deterministic contracts and real owned hosts for native scope
and UI behavior. Evidence is specific to the compiled candidate.

| Group | Required scenarios |
| --- | --- |
| Identity | Duplicate saved names, duplicate unsaved names, departed unsaved project, save/rename transition, two host sessions with identical labels. |
| Resolution | Selected/typed/pasted equivalence, repeated mentions, conflicting destinations, punctuation, Unicode and spaces, ordinary prose and code examples. |
| Preparation | Atomic refusal, unchanged draft/attachments/queue, authorized reads only, consistent reference/text hashes, no entry added on refusal. |
| Freshness | Changed module, changed procedure, moved declaration, explicit compare/refresh, change between refresh and send, legacy missing revision. |
| Continuity | Edit/undo/copy/paste, draft restoration, queue editing, fork isolation, transcript links after discovery changes, host disconnection. |
| Providers | Persistent and stateless transports, changed/unchanged instructions, rejected and uncertain instruction updates, unsupported model/effort, credential exclusion. |
| Dispatch | Every gateway preserves native guards, revocation before admission, cancellation after dispatch, uncertain mutation without replay, partial recovery. |
| UI and compatibility | WPF/WinForms hosting, IME and selection, keyboard/mouse equivalence, accessibility, narrow layouts, owning-thread behavior and busy-host responsiveness. |

M0 remains incomplete until the persistence migration and compatibility prototypes
are demonstrated. M1 remains incomplete until these contracts are integrated into
the actual composer, inspector, refresh and navigation flows and their applicable
native scenarios are qualified. See [test coverage](test-coverage.md) for current
evidence rather than interpreting this matrix as executed tests.

# Conversations and agent workflows

[Documentation](README.md)

A conversation belongs to a selected VBA project. Saved projects use their document
path as the durable scope; unsaved projects have a temporary scope. Select the
intended project before attaching context or resuming work.

## Context and modes

Use `#` to find projects/modules, `@` to find procedures, and the selection action
to attach code. Use `/` to choose an available command such as explanation,
correction, refactoring, documentation or planning. Selecting a command can prepare
a mode; verify the visible mode before sending.

Explicit references, selections and optional project notes appear in the context
preview. Code revisions are checked before transmission; stale attachments must
be refreshed. The displayed size describes explicit context, not all system
instructions, tool schemas or provider history.

| Mode | Intended use | Enforcement |
| --- | --- | --- |
| Discussion | Understand code and ask questions. | Inspection and permitted compilation; editing/execution tools are blocked. |
| Plan | Prepare an approach before making changes. | The same mutation/execution restriction applies. |
| Agent | Perform an intervention. | Project permissions, revisions, mode checks and approval policy still apply. |

Only the linked project is allowed by default. Additional read access and shared
VBE context are separate permissions. See [privacy](privacy.md) before enabling them.

## Editing approval

**Read-only** refuses mutations. **Ask each time** uses the configured confirmation
path. **Automatic** permits eligible actions without a dialog for each one; it does
not disable scope, revision or other tool guards.

New settings currently start with Automatic; migration of older settings without
an explicit policy uses Ask each time. Choose deliberately before the first agent
request. Some native evaluation/execution tools require Automatic and are not
available merely because the mode is Agent.

Inspection is not synonymous with no local UI effect: navigation can select a
pane and compilation can open a diagnostic dialog. The permission categories
concern the tool contract, not an operating-system sandbox.

## Sending, stopping and queued messages

Enter sends and Shift+Enter inserts a line. A visible suggestion takes precedence
when accepting it with Enter or Tab.

During an active intervention, an empty composer offers **Stop**. With text in the
composer, sending queues that message and its captured context without stopping
the current response. Queued messages can be edited, removed or prioritized with
**Send now**. Prioritizing requests interruption and waits for cleanup before
starting the selected message.

The queue is scoped to the session. Project permissions and code revisions are
revalidated at dispatch; stale context leaves the message queued. An error, manual
stop or workflow pause does not silently drain the remaining queue. Reopening a
saved session displays queued messages without executing them automatically.

Stopping a response does not reverse a COM action already started. An action with
no recorded terminal result is uncertain: inspect the live project before retrying.

## Progressive tools and pauses

The model starts with a small core and discovers code, forms, debug, Git or
environment tools as needed. Discovery does not grant permissions. The
[tool reference](reference/vbe-tools.md) explains the common gateway.

For the HTTP workflow, eight consecutive tool-bearing responses without new
successful results trigger a safety pause. A separate ceiling of 64 responses
bounds one segment. Results and call IDs are retained; a pause is not an exception
that discards completed work. Progress means a new tool/arguments/result tuple,
not proof that the user's business objective is complete.

Use the paused-turn resume action with the recorded provider/model/effort/mode.
The project and permissions are checked again. Recorded tool-call IDs are not
executed twice. A new text message starts a new request instead. Codex and SDK-backed
providers retain their own internal orchestration; these HTTP limits are not a
universal provider limit.

## Review, recovery and verification

The transcript groups agent activity without discarding individual outcomes.
Code changes provide diffs and supported undo actions for a hunk, a change or a
whole intervention. Current code is read before recovery; ambiguous or conflicting
changes are refused rather than overwritten.

A multi-module recovery is not an atomic COM transaction. Partial recovery must
be reported, and a code rollback cannot undo files, host data or other external
effects caused by running VBA. UserForm designer recovery has its own boundaries.

Compilation can be requested manually or after an intervention. The result and
available source location are shown in the conversation. A suggested correction
is still a message to send; compilation does not run all macros or tests.

## Sessions and local history

The history button replaces the transcript with the document history across the
full width of the central panel. The same button returns to the conversation;
selecting a saved session also returns there. Opening history preserves the
unsent draft and the current transcript.

Search, rename, pin, archive, restore, delete or export conversations from the chat history.
Deletion requires confirmation and waits for pending local writes. It removes the
local entry; it does not request deletion from the provider.
Forking a conversation copies relevant context but does not create a second owner
of earlier rollback actions or reuse another session's Codex thread.

`%APPDATA%\VBAi\chat.db` stores sessions, drafts, protocol history, queues, explicit
context and recovery snapshots. Provider settings live separately. Project notes
are local and are attached only when explicitly selected. A storage failure is
reported; do not assume that an unsaved history will survive a crash.

For a new unsaved document, sessions and project notes stay in memory. Saving the
exact live project for the first time promotes its history to the canonical saved
path when storage is available. Abandoning the document does not leave a persisted
orphan session. Temporary in-memory entries remain manageable while the scope is
alive; they are not durable across host restarts.

[Privacy](privacy.md) covers permission changes, legacy-session migration and the
limits of deleting or revoking already transmitted context.

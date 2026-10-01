# Privacy and safety

VBAi processes live VBA inside the host application. Cloud or CLI-backed AI use
can transmit code and context to the selected provider. Decide what may leave the
machine before enabling an integration; follow the rules applicable to your work.

## Project permissions

A conversation is bound to one project. Both project-targeted reads and writes
are checked before execution. Another project being open in the same VBE process
is not permission to send its code to the model.

The chat's project-access controls can grant additional **read-only** access to
specific projects. Writes remain limited to the bound project. Ambiguous project
names do not broaden access; saved absolute paths provide stronger identity.

A separate permission grants **shared VBE and clipboard context**. Some debugger
windows, navigation surfaces, collections and clipboard content cannot reliably
be attributed to one project. These tools are refused by default. Granting this
permission may disclose unfiltered content from other projects or the system
clipboard; it does not authorize writing to those projects.

External-file reads follow another route: the user supplies the path and confirms
transmission. The model cannot grant itself these permissions through a tool call.

## Context and conversations

Provider input can include your messages, resolved code attachments, permitted
tool results, optional project notes and prior conversation context. Inspect the
preview and permissions, but do not mistake explicit-context size for the whole
provider request or remote conversation history.

Changing project-access permissions starts a fresh conversation in the same scope.
Existing sessions retain their original permissions; a fork that carries history
also carries its applicable permissions. Revoking access cannot recall content
already received by a provider.

Older sessions predating the privacy policy reset provider-facing history and
remote continuation when opened. The local transcript remains visible. A fork
must not silently reintroduce pre-migration context into a new provider thread.

## Local data

The conversation history provides **Delete conversation** for the selected local
session, with confirmation. Deletion is unavailable during an active response or
while a document scope loads. Pending local writes finish before deletion, and a
revision changed by another host must be reopened first. Deleting the active
conversation opens another unarchived conversation or creates a new one.
This removes the SQLite history entry and attempts to remove its ordinary writer
recovery snapshot. A retained recovery copy is reported separately from deletion.
It does not erase provider conversations, CLI state, document memory, exports or
backups, and it is not a secure wipe of the database file. Recovery copies retained
after an earlier shutdown timeout remain available for manual review.

| Data | Location and boundary |
| --- | --- |
| Settings and saved provider keys | `%APPDATA%\VBAi\settings.json`; saved keys use current-user DPAPI. |
| Conversations, context, notes and recovery snapshots | `%APPDATA%\VBAi\chat.db`; local SQLite history, separate from provider credential storage. |
| CLI-provider state | `%LOCALAPPDATA%\VBAi\Providers\Codex` and `...\Copilot`; child-process homes are isolated from ordinary CLI homes. |
| Editor drafts | `%LOCALAPPDATA%\VBAi\EditorDrafts`; current-user DPAPI protection. |
| Git data and exports | `%LOCALAPPDATA%\VBAi\Git` and `...\GitTemporary`; may contain complete source snapshots. |
| VBA test support backups | `%LOCALAPPDATA%\VBAi\TestSupportBackups`; retained original/proposed source from reviewed support edits. |
| VBA coverage copies and plans | `%LOCALAPPDATA%\VBAi\CoverageRuns`; Excel/Word/PowerPoint document copies, `coverage-plan.json` with original/instrumented full source, hashes and original-source mapping. |
| Problem reports | `%LOCALAPPDATA%\VBAi\CrashReports`; local Markdown copies are retained. |
| Update state | `%LOCALAPPDATA%\VBAi\Updates`; catalog, jobs, host registrations and downloaded payloads. |

Do not assume that the history database, Git objects, coverage copies/plans,
support backups, exports or reports are
covered by the credential-encryption guarantee. Windows profile access, backups
and device security still matter. DPAPI is not protection against all software
running as the same user. Close the relevant processes and back up needed data
before manual deletion or migration.

Deleting a local session is not a provider-side deletion request. Each provider
has its own retention and account settings. Sending `store=false` to OpenAI does
not impose a policy on other providers or guarantee zero retention.

## Network activity beyond chat

Git/GitHub operations can publish sources when requested. Update checks contact
GitHub; prerequisite installation can download a verified Microsoft bootstrapper.
CLI providers may make their own service requests. Choosing a local model does
not make every other enabled feature offline automatically.

Problem reports are previewed and saved locally before sending. Technical reports
exclude raw exception messages, logs, credentials, conversations and VBA by default,
but user-entered descriptions and attachments can contain sensitive data. A GitHub
submission may fall back to classic Outlook; see [troubleshooting](troubleshooting.md).
Do not use this public issue path for undisclosed security findings.

## Execution is a separate risk

The [VBA test explorer](vba-testing.md) keeps results in session memory and can
export human/JSON reports to an explicitly selected local path. Reports can
contain private assertion messages and source identity. Reviewed support-module
edits preserve original/proposed source in `%LOCALAPPDATA%\VBAi\TestSupportBackups`;
these retained recovery files are not uploaded or automatically deleted.

Measured procedure coverage creates an explicitly authorized separate Excel,
Word or PowerPoint document copy. It instruments that copy rather than the original
source and does not save the original document. The copy runs in the same host
application with its privileges and environment; it does not isolate external
effects or other open documents. Excel temporarily suppresses workbook events
while saving, opening and closing its copy and restores the prior setting. PowerPoint
has no equivalent boundary: application-level open and before-close handlers
may execute, affect external systems or prevent copy closure.
Word copies the saved file and refuses unsaved VBA source/reference mismatches;
an unsaved document is refused before copying and checked again after copying
and opening. AutoOpen, document open/close and
application handlers and the Normal template are not disabled, and trust settings
are preserved. These handlers can execute with the host's privileges; a copy is
not an external-system sandbox.
The retained document can contain private workbook, document or presentation data as well as VBA. The
coverage plan contains full original and instrumented source, not just anonymous
hit counters. There is no automatic cleanup or encryption of these artifacts.
Treat manual deletion and any later sharing as separate decisions.

Human and compact reports contain the same canonical measured coverage, source
locations, exclusions and diagnostics. Coverage previews expose procedure identity
and capability; a preview is not an execution measurement or permission to run.

Approval and revision checks are not an OS sandbox. VBA execution can affect host
data, files and external systems with the host's privileges. Code undo cannot
reverse those effects. Stop does not guarantee that a COM action already begun
was canceled. Work on backups and inspect uncertain outcomes before repeating a
mutation. Use [private reporting](../SECURITY.md) for a suspected boundary failure.

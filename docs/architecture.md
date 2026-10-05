# Architecture

VBAi is an in-process COM add-in. It connects a shared VBE automation layer to a
modern editor, an assistant and source-control workflows. Host-specific behavior
is an additional compatibility concern, not the definition of the product.

## Solution structure

`VBAi.sln` contains the managed add-in, the independent updater and test/diagnostic
projects. Common build properties target .NET Framework 4.8, C# 7.3 and x64.
The native renderer is built by the add-in's build target using the C++ toolchain.

| Area | Responsibility |
| --- | --- |
| `src/VBAi/Host` | COM connection lifecycle, menus, native tool-window hosting and load diagnostics. |
| `src/VBAi/Vbe` | Live project resolution and VBIDE operations for code, forms, references, windows and debugging. |
| `src/VBAi/Bridge` | Named-pipe diagnostics and request dispatch in the host process. |
| `src/VBAi/Editor` | Monaco hosting, document revisions, language services, synchronization and draft recovery. |
| `src/VBAi/Llm/Chat` | Sessions, context permissions, workflow state and guarded tool dispatch. |
| `src/VBAi/Llm/Providers` | Provider protocols, authentication integration and catalogs. |
| `src/VBAi/Llm/Settings` | Saved configuration and settings views. |
| `src/VBAi/Git` | Exported-source repositories, bindings, imports, checkpoints and GitHub views. |
| `src/VBAi/Testing` | Annotated tests, versioned assertion/result support, guarded serial execution, native explorer, procedure coverage on Excel/Word/PowerPoint copies and human/LLM reports. |
| `src/VBAi/Ui` and `Localization` | Shared presentation, native appearance, problem reports and translations. |
| `src/VBAi/Updates` and `src/VBAi.Updater` | Update coordination and an out-of-process application step. |
| `src/VBAi.Native` | Native renderer and hook lifecycle. |
| `assets/editor` | Version-locked editor source and generated distribution. |
| `tests` and `tools` | Tests, owned-host fixtures, build helpers and explicit diagnostics. |

The add-in remains one managed assembly. Folder separation is not an assertion
that every layer has already been extracted behind a public extension interface.
The updater deliberately does not reference the loaded add-in assembly.

## Shared VBE services and host compatibility

`VbeSession` resolves the live target and routes operations. Standard VBIDE objects
provide projects, components, code modules, references and code panes. Native
command and accessibility services provide operations not exposed as complete
public VBIDE collections.

For native Access and Publisher host projects (VBIDE Type 100),
`set_project_property` refuses HelpFile and HelpContextID before either COM setter.
Observed partial failures do not justify another setter or an automatic UI
fallback. `project_properties` retains the COM descriptor's `ReadOnly` flag and
raw value, while its existing `SetterStatus` reports `HostLegacyWriteUnsupported`
and `Display` identifies the explicit General alternative. Clients must refresh
the opaque project version and choose that asynchronous, approved workflow;
existing command names and request fields remain unchanged. COM HelpFile readback
and `open_project_help` are not repaired by this change. Other hosts, project types
and properties retain their existing routes.

The explicit asynchronous `read_project_general` and `set_project_general`
commands use the original VBE General dialog on its owning STA. They are separate
from the existing COM metadata commands and never run as recovery after a failed
COM setter. Reads cancel the original dialog; writes target one HelpFile or
HelpContextID field and request one OK after exact field readback. The native
`OptionsVersion` protects all General fields independently of the existing COM
project revision. Original project identity, design mode, protection, approval
and privacy are revalidated across dispatch and final publication. A pending or
uncertain operation blocks further session operations, including direct native
bridge routes. Bridge admission is held on the owning STA until worker dispatch
settles, preventing General from entering between admission and a native call.
Managed status remains available. Persistence is qualified
separately by adapter Save and independent reopen. Native host acceptance is
recorded in the compatibility and qualification pages, not inferred from this design.

A document save, application-level procedure invocation or standalone project
persistence can require a host-specific path. These adapters must identify the
actual document and report an unsupported operation or failed prerequisite rather
than simulate success. The [compatibility guide](compatibility.md) records outcomes
per operation instead of maintaining a misleading all-or-nothing host whitelist.

SOLIDWORKS macro creation has two explicit asynchronous routes.
`create_solidworks_macro` requires a fresh absolute `.swp` destination and invokes
the selected host's native New Macro command once. The owning STA handles its
addressed save dialog and verifies the resulting Type100 project, native
`ThisLibrary`, file and original project collection. A successful creation result
does not prove persistence after an independent reopen.

`publish_solidworks_macro` preserves an unprotected Type101 draft and creates a
separate native macro at the explicit destination. It reports both identities,
transports standard modules, classes and UserForms with their resource companions,
and verifies code, hidden attributes, complete designer state, pictures, references
and Description before and after one native save. Staging exports remain available
on failure. Nondefault help or conditional-compilation metadata and project
references are refused; digital signatures are not transported or qualified.
Publication never promotes the original project object in place.

Live MSForms descriptor type names can include a process-local COM type-info
pointer. Opaque form revisions remain live concurrency tokens; they are not
cross-process persistence hashes. Qualification compares recognized enum
families, ordered choices and typed values without rewriting the raw protocol
objects or suppressing other descriptor errors. Native cold inspection explicitly
opens the designer before enumerating form properties; the observed closed-form
inspection limit is recorded in the qualification guide.

Publication compares persisted designer content separately from each live
project's concurrency revision. The canonical source and destination FRM headers
must agree; only the validated companion FRX filename is normalized. Missing
root `Visible`, `HelpContextID`, `ShowModal`, `WhatsThisButton` and `WhatsThisHelp`
properties use their documented defaults only after their exported settings have
been checked. Control properties, ordinary font members, pictures, unknown root
properties and read errors remain strict. Live undo, redo and clipboard
capabilities are excluded only from this transport comparison; raw revisions
still include them. These boundaries follow the
[VBFrame specification](https://learn.microsoft.com/en-us/openspecs/office_file_formats/ms-ovba/76ecef26-5e1c-4729-8074-c76584e25b6b)
and [Microsoft Forms property reference](https://learn.microsoft.com/en-us/office/vba/language/reference/user-interface-help/properties-microsoft-forms).

An imported UserForm may receive surplus initial empty CRLF lines. After a
verification export proves its other content and hidden attributes, publication
can remove only that exact surplus prefix, once, from the guarded imported
component. Exact code readback and a distinct verification export must succeed
before Save. Other code differences are refused, and an uncertain call is never
replayed. The source draft and its raw revision remain unchanged.

After a verified publication Save, the destination content check permits only an
error-free Boolean `Saved` transition from false to true, or an unchanged value.
This is the documented edited-since-save flag; it is not persisted component
content. All other component descriptors, including `HasOpenDesigner`, remain
strict. A distinct guarded post-save export verifies code, hidden attributes and
persisted form content. Source fingerprints and pre-save raw component versions
remain unchanged; they still include `Saved` and live designer revision data.
A remaining mismatch records bounded immutable observations in owned staging,
keeps the successful Save result as a partial outcome and quarantines the
operation. It does not repeat Save. See the
[VBIDE Saved property contract](https://learn.microsoft.com/en-us/office/vba/language/reference/visual-basic-add-in-model/properties-visual-basic-add-in-model#saved).

MSForms exposes `_Font_Reserved` as a write-only alias. Designer inspection skips
its getter only when complete COM type metadata proves the exact reserved DISPID
and setter-only accessors. The ordinary `Font` property and its members remain
part of the snapshot. Missing or ambiguous metadata, and actual font or resource
read errors, still prevent publication.

Publication reads General metadata through the existing addressed native Project
Properties command and closes the dialog with Cancel. Conditional compilation
arguments are not an exposed VBIDE project COM property. A scoped reader verifies
all General fields before staging, before and after native creation, and before
and after save. It retains the original source identity and restores the prior
active project only after the read settles; an uncertain modal call stops the
operation without another attempt. The reader remains on the owning STA and
preserves approval, privacy, revision and cross-session quarantine checks.

UserForm picture revisions compare OLE persistence data, type and dimensions,
using SHA-256 rather than a process-local GDI handle. Inspection calls
`IPersistStream.Save(false)` without clearing the picture's dirty flag; it bounds
the declared and actual stream sizes to 64 MiB and releases only its owned stream.
An unreadable size/content contract is refused. Empty OLE pictures have no content
digest. Windows may adapt a WMF handle during persistence inspection, so a stable
handle is not the content oracle. Import and independent host reopening remain
separate qualification gates for preservation of those data.

Generic Type101 Save/SaveAs is refused inside SOLIDWORKS because its output has not
passed native reopening. Existing native Type100 macro save remains a separate
route. Unsafe `VBProjects.Open` remains disabled; reopening uses the host's native
Edit Macro workflow. Both new mutation routes revalidate identity, revision, mode,
approval and privacy on the owning STA, including the final cached authorization
boundary. A pending or uncertain native mutation quarantines further operations
on that STA until host restart; it is not retried or rolled back automatically.
These implemented contracts require candidate-specific native qualification.

Adding a host means qualifying shared behavior first, then implementing and testing
only the missing application-specific operations. Keep knowledge of a particular
application out of provider protocols and reusable editor code where possible.

## Threading and mutation lifecycle

The test explorer uses the shared owning-thread project/revision guards. A
registered x64 `VBAi.TestRuntime` COM callback binds only to an already authorized
pending attempt on the VBA thread. Versioned project-local support publishes a
bound verdict; a delivery acknowledgement or Immediate echo is not a test result.
Ordinary native runs use the registered callback by default. Excel, Word and PowerPoint
have returned-value fallback transports, also used for copy-based coverage.
Recognized host names describe capability routing, not observed qualification.

Procedure-entry coverage uses an explicit Excel, Word or PowerPoint document copy, validates separate
project/path/source identity and instruments only that copy. Its denominator,
original-source probes, exclusions and diagnostics feed both report formats.
Unknown or incomplete measurements cannot become a complete project percentage.
Statement and branch coverage are unavailable. Copy/plan artifacts retain private
full source; the original document is not saved by measurement. Copies run in
the same host application with its privileges. Excel restores temporary event
suppression around copy opening/closing; PowerPoint application-level open and
before-close handlers may execute. Word copies the saved document file and
refuses source/reference mismatches, leaving AutoOpen, document/application
handlers and the Normal template subject to the host's existing policies. Its returned-value path activates and
revalidates the exact owned document before document-qualified invocation.
Copying is not an external-system sandbox. See the
[test contract](vba-testing-design.md), [usage guide](vba-testing.md) and
[qualification evidence](test-coverage.md).

COM/VBE, WinForms and WebView2 calls stay on their owning UI/STA threads. Background
work prepares diffs, indexes snapshots, performs Git/network I/O and persists
eligible data without treating COM objects as thread-safe values.

Before an edit, resolve identity, verify permissions/mode/protection and compare
the expected revision. Apply the operation, re-read the result and record recovery
information. Concurrent changes invalidate stale plans. Native operations may be
partially applied before failure; COM does not supply a general transaction for
these workflows. A timeout is an uncertain outcome, not an automatic retry signal.

The editor uses versioned changes and reconciliation against native snapshots.
A draft, a synchronized VBA module and a document saved on disk are distinct states.

## Assistant boundary

Provider transports normalize conversation and tool events. A small core discovers
additional tool families or uses the guarded invocation gateway. All paths must
retain project access, shared-context consent, chat mode, approval and revision
checks. [Tool reference](reference/vbe-tools.md) describes the schema authority.

Conversation history, queued messages and workflow pauses are persisted separately
from provider credentials. The permissions on a conversation are not permissions
on the host process or the diagnostic bridge. See [privacy](privacy.md).

## Local bridge

The host exposes `VBAi.<PID>` as a named pipe with an access rule for the current
Windows user. It accepts one UTF-8 JSON request per connection, terminated by LF
(optionally CRLF), and returns a serialized response. It is a diagnostic interface,
not an MCP server or the same permission boundary as the chat.

The receiver bounds incoming data at **10 MiB** and **10 seconds total after
connection**. A slow sender does not reset the deadline. Invalid UTF-8, missing
line termination, excessive size and incomplete requests are refused; a failed
client must not permanently occupy the single worker. The reception deadline is
not an execution timeout for a VBE command.

Use `tools/Invoke-VBAi.ps1` against an explicitly selected PID. Keep automation on
disposable projects and never equate a successful protocol response with a verified
runtime effect or successful disk save.

## Identity, storage and updates

Current names are `VBAi.dll`, namespace `VBAi`, `VBAi.AddIn`,
`VBAi.ChatToolWindow`, `VBAi.TestRuntime` and `VBAi.<PID>`. The rename retained COM GUIDs and includes
known legacy registration/data migration. Do not rename persisted Git refs or
protocol fields for cosmetic consistency.

[Privacy](privacy.md) inventories persistent data. [The update contract](updates.md)
explains why a separate process waits for loaded hosts before applying a future
signed installer. No installer is implemented by this documentation work.

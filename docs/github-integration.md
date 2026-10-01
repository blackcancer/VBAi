# Git and GitHub

VBAi versions **exported VBA sources**, not the entire application document. The
same workflow is available to a compatible VBE project regardless of its host,
subject to reliable document identity and import/export capabilities.

## Requirements and connection

Use Git for Windows, a configured Git identity and an HTTPS credential manager
such as Git Credential Manager. The GitHub account selected in VBAi settings is
independent of the AI provider, including GitHub Copilot. Credentials stay with the
credential manager; VBAi does not put access tokens in source manifests.

Open VBAi's GitHub view for the intended project. The document must be saved and
accessible, the project unlocked and the VBE in design mode for source mutations.
Link an existing GitHub HTTPS repository and branch. Linking inspects the remote
branch; it does not publish or import VBA. GitHub Enterprise and SSH URLs are not
part of the documented connection path.

## Keep the operations separate

| Operation | Effect |
| --- | --- |
| Compare | Inspect exported source changes and file diffs. |
| Commit | Export the live VBA and record changes locally. |
| Push | Publish local commits to the selected remote branch. |
| Fetch | Refresh remote information without changing live VBA. |
| Pull and import | Validate incoming sources and import them with a prior backup. |
| Restore VBA | Attempt guarded recovery of the pre-import source state. |

Do not confuse a successful commit or import with saving the document on disk.
After importing, compile/check the result and save through the application.
The Git workflow does not automatically run macros or save the host document.

GitHub repository and pull-request views also exist in the UI. Their availability
depends on credentials and repository permissions; not every UI operation is
exposed as an LLM tool. Remote counts reflect the last fetch.

## Checkpoints, branches and merges

A checkpoint captures live VBA and can be restored without rewriting commits.
Restoration first checkpoints displaced work. Keep `.frm` and `.frx` files together.

Uncommitted changes block branch switching. Switching imports the selected
branch's VBA with a backup. Divergent pulls are refused rather than automatically
rebased or force-pushed.

Three-way merge preparation requires Git 2.38 or later. Preparing a merge does not
immediately change HEAD or VBA. Review conflicts and choose a complete side for
binary form resources, or provide the full resolved text for text sources. Completing
a valid merge creates a two-parent commit and imports the result with a checkpoint.
Aborting discards the preparation, not unrelated work.

## Versioned source format

Only the managed `vba/` subtree is replaced by source commits. Other repository
files are preserved. Do not place unmanaged files inside that subtree.

| Entry | Meaning |
| --- | --- |
| `manifest.json` | Format version, component names/types, resources and reference identities. |
| `.bas` | Native standard-module export including attributes. |
| `.cls` | Native class-module export including attributes. |
| `.frm` and `.frx` | Form definition and associated binary resources. |
| `.vba` | Visible code for an existing host-owned document module. |

Git text is normalized to UTF-8/LF. Imports convert to the VBE's Windows code page
and fail before mutation if conversion is impossible. Binary resources stay binary.
Document contents, workbook data, conversations and provider settings are not part
of this VBA source export. Secrets embedded in VBA code are still code: review
sources before publication.

## Import safety

Document modules must already exist with matching names and are updated in place.
VBAi does not create application objects such as sheets to satisfy a source manifest.
References must match; importing does not install missing COM libraries automatically.

For a UserForm's `OleObjectBlob`, preflight checks the native LB/08 resource
envelope and the bounded compound-storage allocation graph before import. It
rejects truncated containers, invalid sector references, cycles and overlapping
allocations, without activating an OLE object or altering FRX bytes. The storage
checks follow Microsoft's [MS-CFB header](https://learn.microsoft.com/en-us/openspecs/windows_protocols/ms-cfb/05060311-bfce-4b12-874d-71fd4ce63aea)
and [directory format](https://learn.microsoft.com/en-us/openspecs/windows_protocols/ms-cfb/60fe8611-66c3-496b-b70d-a504c94c9ace).
These are structural checks, not a complete validation of embedded MS-OFORMS
properties or third-party controls.

Comparison uses logical CFB storage/stream contents rather than physical sector
placement, unused allocation bytes or directory timestamps. Names, storage CLSIDs,
state bits and all stream contents remain significant. Invalid UTF-16 directory
names are refused instead of replacing characters during comparison. For a
complete form graph identified by the UserForm storage CLSID and recognized
stream grammars, comparison also ignores padding
identified by Microsoft's [MS-OFORMS site-data grammar](https://learn.microsoft.com/en-us/openspecs/office_file_formats/ms-oforms/f65e0b17-6383-4570-b030-7b868f2c07d5)
and [alignment rules](https://learn.microsoft.com/en-us/openspecs/office_file_formats/ms-oforms/622ed335-0723-4491-b271-e4767d7453e3).
Recognized variants include Label, CommandButton, TextBox, ListBox, ComboBox,
CheckBox, OptionButton, ToggleButton, SpinButton, ScrollBar, TabStrip and Image,
plus nested Frame, MultiPage and Page storage graphs. Font readers recognize
bounded TextProps and StdFont variants; selected picture-bearing controls retain
their complete supported picture envelopes without decoding the payload.
Recognition remains restricted by each reader's property masks, versions,
lengths and storage relationships. This is not support for every property
combination or third-party control.

Unsupported controls, fonts, picture layouts, container relationships or
extensions retain exact logical stream comparison for the **whole graph**;
VBAi does not normalize recognized descendants of an unsupported parent.
Any opaque resource declaration retains exact FRX comparison: its offset alone
cannot establish its extent or exclude overlap with an OLE envelope. Validation
and comparison share resource-reference extraction, including multiline values.
Status, revision guards, import selection and readback use these same comparison
rules. Git blobs, checkpoints and native import files retain the original FRX
bytes. When native UserForm import adds exactly one leading empty code line,
VBAi removes it only after complete visible-code matching and project/component
identity revalidation. Intentional whitespace and hidden export attributes remain
protected by the final exact FRM readback. Unexpected code is not rewritten and
an uncertain COM mutation is not retried. Complete native form roundtrips remain tracked in
[qualification](release-qualification.md).

Before mutation, the workflow records a backup and recovery marker. Incoming sources
are validated and read back. A COM error after partial application does not trigger
an automatic repeat. An unresolved recovery state blocks further synchronization.
If partial-state readback also fails, automatic recovery can be refused while the
backup remains available for manual recovery. VBAi reports both the original
import error and the failure to record the resulting state; it does not claim a
verified after-state or retry the import.

Only a definitely missing recovery marker permits normal synchronization.
Metadata access and I/O errors propagate; an existing entry is pending even if it
is a directory. Recovery refuses directories and filesystem links before native
import. Preparing a new recovery independently checks the marker before replacing
the backup or clearing the recorded after-state. These checks do not make external
filesystem changes atomic.

Completion requests deletion of a confirmed regular marker once and verifies its
absence. Failed deletion, unreadable metadata or a replacement entry prevents a
success report without another deletion or import attempt. If an import is refused
before mutation and marker cleanup also fails, both errors are retained. A
legitimate rollback of a completed import can still start with no pending marker.

Private refs under `refs/codex/*`, including checkpoints and backup state, are not
published by normal pushes. Their legacy spelling is a storage contract, not a
reason to rename them in a documentation change.

## Local storage

Bindings and bare repositories live under `%LOCALAPPDATA%\VBAi\Git`, keyed by the
saved document path. No sidecar is added next to the document. Native COM exports
and imports use `%LOCALAPPDATA%\VBAi\GitTemporary`; a crash can leave temporary files.

The VBE menu, chat and agent tools resolve the same binding from the native saved
document path. Lookup also checks the former uppercase path key to preserve a
single existing binding created by earlier chat/tool versions. It does not move,
copy, merge or rewrite either cache, its commits or its private recovery refs.
When neither binding exists, a new link uses the native path key. If both keys
have bindings, or a binding entry is unreadable, a directory or a reparse point,
lookup refuses automatic selection. Resolve the ambiguity while preserving both
repositories; VBAi does not choose one history over the other.

A binding is local to the machine and document path. Moving a document requires
relinking it. **Do not purge the Git directory as disposable cache:** it can contain
unpushed commits and recovery snapshots.

Agent tools use the existing project binding and a current `ExpectedState`. They
cannot substitute an arbitrary repository/cache/shell command. Review the editing
policy and request remote publication explicitly. See [privacy](privacy.md) and
[recorded validation](test-coverage.md) for boundaries and test scope.

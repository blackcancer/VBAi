# Git and GitHub

[Documentation](README.md)

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

## A first project, step by step

Think of the integration as three copies with different jobs:

1. **Open VBA project:** the live project in the host application and VBE. This is
   the source VBAi reads for Compare and Commit, and the destination for Pull.
2. **Local Git history:** a private bare repository in the current Windows user's
   `%LOCALAPPDATA%\VBAi\Git` cache. It stores commits, branch state, checkpoints,
   and import recovery data. It has no checkout folder to edit in Explorer.
3. **GitHub branch:** the remote copy addressed by the selected HTTPS repository
   and branch. Push sends commits here; Fetch reads its latest commit.

For example, suppose `Budget.xlsm` contains a standard module `ModuleCalcul`, a
class `ClassBudget`, a UserForm `FrmBudget`, and the workbook document module
`ThisWorkbook`. A committed VBA source tree can look like this:

```text
Budget-tools/                 # Other repository files are preserved
  README.md
  vba/
    manifest.json             # Component names/types, resource flags, references
    ModuleCalcul.bas          # Standard module
    ClassBudget.cls           # Class module
    FrmBudget.frm             # UserForm definition and visible code
    FrmBudget.frx             # Exact companion form-resource bytes, when present
    ThisWorkbook.vba          # Visible code for the existing host-owned module
```

The `.vba` entry records code only; it does not create or replace `ThisWorkbook`,
worksheets, or other objects owned by Excel. VBAi imports document-module code
into the already existing component with the same name. It also requires the
project's reference identities to match the manifest. Resolve missing or different
references through the host's References dialog before importing. A Git commit
does not contain workbook cells, sheets, formulas, other host document data, or
the VBA project's digital signature.

### Publish the current project for the first time

1. Save `Budget.xlsm` in Excel, open the intended project in the VBE, and make
   sure the project is unlocked and in design mode. Confirm the selected document
   label in VBAi before linking.
2. In the GitHub view, enter the HTTPS URL and branch, for example `main`, and
   connect. This creates or opens the local bare cache and fetches that branch.
   It does not import source or publish anything. Compare shows the live project
   against the last synchronized baseline. On a new link with no local baseline,
   the live files appear as additions. If the selected remote branch already has
   a VBA source package, import it first when that version should be the starting
   point; an initial commit refuses to overwrite a different remote VBA snapshot.
3. Review the changed-file list. Check the `.bas`, `.cls`, `.frm`, matching `.frx`,
   and document-module `.vba` files you intend to publish. Inspect the source for
   credentials, personal data, or other content that should not be public.
4. Enter a descriptive commit message and choose **Commit**. VBAi captures the
   live project and records its VBA snapshot as a local Git commit. This changes
   local history and its baseline only; it does not save the workbook or contact
   GitHub.
5. Choose **Push** to publish the selected local branch. Push fetches the branch
   again and proceeds only if the remote can fast-forward to the local commit.
   A rejected push leaves local history intact. Fetch and review the remote change
   before deciding how to reconcile it; VBAi does not force-push or silently
   merge divergent histories.

Immediately after the first commit, the live project matches the new local
baseline. Once the remote state is known, the local branch is one commit ahead;
after a successful push, the selected remote branch points to the same commit and
the outgoing count returns to zero. If the remote already had a commit but no
`vba/` package, the new commit preserves the other repository files and adds the
managed `vba/` subtree.

If the connected GitHub branch already contains a valid `vba/manifest.json` and
VBA source, use **Pull** when the open project should receive that version. Pull
requires the live project to match its synchronized baseline when one exists. On
a first link without a baseline, compare and review the incoming change summary
carefully before confirming the import. VBAi then verifies that the remote commit
is still the one selected for import and that the update is a fast-forward. A
branch with no VBA source package cannot be imported as a VBA project.

Before changing the live project, Pull creates a checkpoint and a separate private
backup of the current VBA snapshot, then records a recovery marker. It validates
the complete target, applies source to existing components, and captures the live
project again. Only after that readback succeeds does it advance the local branch
and baseline to the incoming commit. Review and compile the result in the host,
then save the host document there. VBAi does not automatically run macros or save
the workbook.

If an import reports an error, do not repeat Pull immediately. Check the Git view
for a pending recovery state and use **Restore VBA** only after reviewing the
current project. Automatic restore proceeds only when the live source still
matches the recorded post-import state; otherwise it refuses to overwrite newer
edits and keeps the backup for manual recovery. If the after-state could not be
recorded, VBAi cannot prove that automatic restore is safe. Inspect the live
project and retained backup before deciding what to recover.

### Save, Commit, Push, and Pull

These actions affect different copies:

| Action | Changes | Does not do |
| --- | --- | --- |
| Save in Excel or the host | Persists the host document on disk. | Create a Git commit or publish it. |
| Commit in VBAi | Captures current VBA into local Git history and advances the local baseline. | Save the host document or contact GitHub. |
| Push in VBAi | Sends the selected local branch commit to GitHub. | Read unsaved editor changes or mutate the live VBA project. |
| Fetch in VBAi | Updates the local record of the remote branch. | Change local commits or import into the host. |
| Pull in VBAi | Imports a reviewed fast-forward source commit into the open VBA project, with a prior backup. | Save the host document to disk. |

After editing code in the VBE, save the host document as appropriate, compare,
commit, and then push. After Pull, review and validate in the host and save again
to persist that imported VBA in the document file. A local commit and a saved
workbook are separate records; either can be newer than the other.

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
Import does not preserve or recreate the VBA project's digital signature. If the
host invalidates or requires a signature after source changes, follow the host's
signing policy and sign the project again before distribution.

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

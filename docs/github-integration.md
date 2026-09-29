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

Before mutation, the workflow records a backup and recovery marker. Incoming sources
are validated and read back. A COM error after partial application does not trigger
an automatic repeat. An unresolved recovery state blocks further synchronization.
If partial-state readback also fails, automatic recovery can be refused while the
backup remains available for manual recovery.

Private refs under `refs/codex/*`, including checkpoints and backup state, are not
published by normal pushes. Their legacy spelling is a storage contract, not a
reason to rename them in a documentation change.

## Local storage

Bindings and bare repositories live under `%LOCALAPPDATA%\VBAi\Git`, keyed by the
saved document path. No sidecar is added next to the document. Native COM exports
and imports use `%LOCALAPPDATA%\VBAi\GitTemporary`; a crash can leave temporary files.

A binding is local to the machine and document path. Moving a document requires
relinking it. **Do not purge the Git directory as disposable cache:** it can contain
unpushed commits and recovery snapshots.

Agent tools use the existing project binding and a current `ExpectedState`. They
cannot substitute an arbitrary repository/cache/shell command. Review the editing
policy and request remote publication explicitly. See [privacy](privacy.md) and
[recorded validation](test-coverage.md) for boundaries and test scope.

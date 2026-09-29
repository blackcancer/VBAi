# Modern editor

VBAi embeds Monaco in the VBE document area. The native VBA compiler, debugger and
UserForm designer remain the execution and design back ends; Monaco does not
replace the VBA runtime.

## Workspace and language assistance

The editor opens with the add-in and follows the VBE document area. Double-clicking
a code module opens its tab. Native code panes remain associated with the tabs;
UserForm designers and the Object Browser retain their native surfaces.

Completion, signature help, hover and declaration navigation use project code and
referenced COM type libraries. Open drafts participate in the index. Current
services handle typed receivers, property/call chains, nested `With` blocks,
several default-member cases and locally inferred `Set ... = New ...` types.
Reference changes invalidate relevant results without requiring a code edit.

The editor also supplies block completion, indentation/formatting, automatic
parenthesis closing and VBA aliases such as `Left`/`Left$`. These services do not
instantiate application objects to inspect their members.

This is not a full replacement for compiler semantics. Arbitrary late-bound
`Object`/`Variant` expressions, ambiguous declarations and conditional compilation
remain limited; unavailable COM metadata cannot supply completion. Old-revision
responses are discarded.

## Synchronization is not saving

Background synchronization applies eligible drafts to the live VBA project. It
does **not** save the containing document to disk.

**Ctrl+S** first synchronizes, then requests the native Save command for the tab's
project. Resolve host prompts and check the reported result. This native route is
not the same as the `save_host_document` application adapter. See
[compatibility](compatibility.md) for adapter-specific results.

For important work, save in the host and reopen a disposable copy to verify
persistence. Git source commits are another separate operation.

Each write checks component identity, mode, protection and the expected prior
revision. Conflicts offer comparison, reload and explicit resolution. Further
native edits after a comparison invalidate an overwrite plan.

In break mode, a supported single-line procedure-body correction can use
`ReplaceLine` without resetting execution. Declaration changes, insertion and
removal remain drafts until design mode. Incompatible Windows-code-page characters
are refused rather than silently corrupted.

## AI actions

Explain, Fix and Refactor prepare a message for the document's conversation and
attach the selection, or the module when there is no selection. The action itself
does not send a provider request. Review the context and mode before sending.
A changed draft invalidates an older attachment.

## Compilation and debugging

| Action | Shortcut in Monaco |
| --- | --- |
| Compile project | Ctrl+Shift+B |
| Request breakpoint toggle | F9 or the gutter action |
| Step into | F8 |
| Step over | Shift+F8 |
| Step out | Ctrl+Shift+F8 |
| Show next statement | Command palette/context menu |

Compilation synchronizes first and reports the first available native diagnostic
at the location supplied by VBE. Editing invalidates diagnostic markers. Compilation
and synchronization do not themselves run a macro.

Stepping requires the appropriate native break state. The execution marker follows
native VBE observations, including Show Next Statement; it is not an independent
read of the runtime instruction pointer.

VBIDE exposes no public complete breakpoint collection. Local gutter markers
represent toggle requests, **not a verified inventory of installed breakpoints**.
Check the native state when precision matters. A reload may lose native breakpoints
and Undo history, which cannot be reconstructed exactly by this integration.

The assistant's `debug` tools can inspect visible native panes and, with the
required project, mode and approval checks, evaluate a selected expression or
inspect a bounded set of simple local scalars while paused. These native reads
may change selection or evaluate VBA and do not create an independent debugger
state database. See [tool discovery](reference/vbe-tools.md) and the
[recorded Excel checks](test-coverage.md) for their different validation limits.

## Attributes and recovery

Hidden VBA attributes require more care than replacing visible text. Supported
edits use native export, metadata verification and guarded reload. Multi-line
attributed declarations may require a controlled component replacement: export,
import under a temporary name, verify code/metadata/Designer resources, re-read the
original, then replace it. Failure triggers a restoration attempt and preserves
the original export if recovery cannot be confirmed.

Document modules are not replaced by importing a new component: their identity
belongs to the host document. Ambiguous attribute associations, unsupported
third-party controls and unverifiable Designer properties block replacement.
Original `.frx` resources are retained where required.

Drafts are DPAPI-protected under `%LOCALAPPDATA%\VBAi\EditorDrafts`. Cleanup can remove
older versions after 30 days but retains the latest draft per module and files
belonging to a live process; reparse directories are ignored. An unsaved project's
identity is temporary. A crash before the last edit is received/persisted can still
lose that edit; drafts do not replace document backups.

## Appearance and editor assets

Monaco resources, workers and available translations are bundled rather than
loaded from a CDN. WebView2 blocks external navigation, permissions, downloads,
new windows and host-object exposure in this editor surface. Its installed runtime
is a separate prerequisite from the bundled loader/assets.

The add-in's own appearance is separate from the **experimental native VBE dark
theme**. Native palette recovery must not be deleted to silence an error. See
[troubleshooting](troubleshooting.md#appearance-and-native-palette-recovery).

# UserForms and native designer data

VBAi works with the host's native VBA UserForm designer. It can inspect forms and
containers, add supported controls, change typed properties, adjust geometry,
work with events and perform bounded duplication/recovery operations. These
capabilities do not imply universal support for every ActiveX control.

## Identity and revisions

Resolve the project, form and control from the live tree. Nested controls use
canonical container/control paths rather than a caption alone. Layout and mutation
operations require the relevant form/tree revision and the permitted VBE mode.
Re-read after editing; an obsolete tree cannot safely identify the current target.

Property discovery is not a writable-property guarantee. A COM property can be
readable but read-only, inaccessible in the current mode, dependent on a control's
state or different across hosts. Preserve types and treat errors as explicit results.

## Layout, events and resources

Use preview/inspection before changing a form layout. MultiPage/Page, Frame and
TabStrip containers have distinct semantics; removing a page can remove its child
controls. Event-procedure creation changes code as well as designer state and
must retain the corresponding revision checks.

Images, fonts, list data and other properties can require special handling.
Duplication is bounded by supported properties/control types; do not silently
replace unsupported third-party controls with a different built-in control.

Native form exports pair a `.frm` definition with `.frx` binary resources when
present. Preserve the pair through Git operations and recovery. A controlled
component replacement must preserve and verify the original designer data; host-owned
document modules must not be replaced as ordinary imported components.

A supported designer cut can offer a recovery action. This is distinct from text
undo and does not establish an unlimited designer transaction history. Preserve
exports when recovery cannot be confirmed.

## Reference inventories

The following existing CSV files are retained as **historical Excel-host property
inventories**, not universal writeability or compatibility specifications:

- [Control properties](excel-control-properties.csv).
- [UserForm properties](excel-userform-properties.csv).

Use live discovery and operation-specific tests for the actual host. The
[compatibility guide](../compatibility.md) includes later Office form scenarios;
these do not qualify every control, property combination or runtime effect.

See [modern-editor attribute recovery](../modern-editor.md),
[Git import safety](../github-integration.md) and [tool discovery](vbe-tools.md).

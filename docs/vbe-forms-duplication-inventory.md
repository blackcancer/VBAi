# VbeForms duplication test inventory

Status: **NOT_RUN**. `VbeFormsFrameDuplicationTests` uses in-memory VBIDE/MSForms
fakes and is included in the MSTest project. No build, test or coverage run was
performed for this lot.

## Scenarios written

- `FrameCopyPlan` and `FrameSimpleCopyPlan`: read-only eligible Label plan,
  TextBox allowed only in the simple plan, proposed paths and profiles,
  collision/length issues, stale tree, noncanonical path, invalid name and
  wrong native control type.
- `DuplicateEmptyFrame`: geometry and Caption success, changed tree version,
  refusal of children/stale version/unsupported geometry, and removal of a
  created control after a simulated setter failure.
- `DuplicateFrameWithLabels`: supported Label fields and child path success,
  refusal of empty or unsupported children before mutation, and rollback after
  a simulated Frame setter failure.
- `DuplicateCommandButton`: narrow property copy with events explicitly
  omitted, wrong-type refusal and rollback after a setter failure.
- `DuplicateTextBox`, `DuplicateCheckBox`, `DuplicateToggleButton`,
  `DuplicateOptionButton`, and `DuplicateComboBox`: supported value and shell
  profiles, refusal of unsafe source values or stale paths, and setter-failure
  rollback. The fake preserves the deliberate omission of selection, group,
  items and bindings.
- `DuplicateFrameWithSimpleChildren`: Label and TextBox success, refusal of a
  Frame without TextBox or of nontext TextBox.Value, and removal of both a
  created child and Frame after a simulated child setter failure.
- `FrameProfileCopyPlan` and `DuplicateFrameProfiled`: the six positive child
  profiles, refusal of an unprofiled child or invalid CheckBox.Value, full
  child-and-Frame rollback, and an explicit incomplete-rollback diagnostic.

## Remaining limits

- The fakes do not prove MSForms COM `Controls.Add`, design mode persistence,
  event procedure behavior, visual layout, or real parent identity. These
  require the automated Excel integration host and targeted SOLIDWORKS checks.
- The fake can reject `Remove` to assert the incomplete rollback diagnostic;
  it does not prove that a real COM failure leaves the same control hierarchy.
- Duplicate-name, geometry, and race branches remain uneven across individual
  profiles; the grouped coverage report must identify the actual gaps. A
  green result for this lot cannot imply global 100% coverage.
- The grouped VSTest/Cobertura run after integration must report the actual
  coverage. No source exclusion was added.

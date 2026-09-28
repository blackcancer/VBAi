# VbeForms core branch coverage lot

> Archive conservée le 28 septembre 2026. Ce document contient des observations et des décisions de sa période de rédaction ; ses états « à faire » et ses anciens chiffres ne constituent pas le bilan actuel. Voir [la documentation actuelle](../../README.md) et [les travaux restants](../../roadmap.md).

Status: VSTest **NOT_RUN** until the grouped suite is executed. The tests use an
in-memory VBIDE/MSForms fake. The isolated Release build succeeded with zero
warnings and zero errors; compilation does not prove COM runtime behavior.

## Scenarios in this lot

- `AddPageOrTab` and `RemovePageOrTab`: canonical Page and Tab paths, indexed
  insertion, name and version refusals, wrong parent, rollback after a native
  Add failure, and the explicit diagnostic when rollback cannot be verified.
- `SetNodeProperty`: scalar conversion and fresh tree versions, unknown or
  managed object members, and preflight refusal of the five setters recorded
  as unsafe after host crashes or unavailable native setters.
- `FrameProfileCopyPlan`: empty Frame, proposed child collision, bad Frame
  geometry, unsupported font, nontext TextBox value, invalid ComboBox width,
  stale tree and the 128 direct-child limit.

## Remaining host boundaries

- The fake's Page/Tab collections are deterministic. Real VBIDE collection
  indexing, designer persistence, native rollback and visual tabs still need
  a disposable Excel host check.
- Managed fake objects cannot exercise COM-only `SetNodeProperty` object member
  setters. Image/OLE picture paths require a real compatible designer object.
- Form tree identity across distinct COM wrappers and `ZOrder` visual effect
  remain host-dependent. The grouped Cobertura result determines actual line
  and branch coverage; no exclusions were added.

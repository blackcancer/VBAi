# VbeSession dispatch coverage follow-up

Status: new tests compiled successfully; grouped VSTest and coverage were not
run for this follow-up batch.

## Added deterministic scenarios

- Project dispatch: properties/version change, persistence state, signature
  status, and signature persistence report their host limits outside Excel.
- Mutation dispatch: disabled project rename, missing save tokens, and refusal
  of Save/SaveAs in an unsupported host.
- Signature dialog dispatch: missing expected design mode, a running project,
  and an inactive project are refused before any native dialog command.
- Reference dispatch: negative version numbers, absent identity, stale inventory,
  case-insensitive version tokens, and add/remove lifecycle.

## Remaining acceptance boundary

This fake VBE cannot prove Excel workbook persistence, certificate store
eligibility, a native signature dialog, COM object identity in the VBE UI, or
reference effects in a real project. The wider dispatcher routes to editor,
debugger, forms, and code navigation classes; this file does not claim that
all routes, lines, or branches are covered. No coverage exclusion was added.

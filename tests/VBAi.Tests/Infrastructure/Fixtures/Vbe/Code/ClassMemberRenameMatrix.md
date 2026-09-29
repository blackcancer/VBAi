# Class member rename matrix

## Supported and verified by prepared fixtures
- Private class Sub: exact declaration, direct Call, direct statement, Me-qualified Call, inline Then/Else, physical continuation and casing.
- Private class Function: declaration, recursive calls, Me-qualified calls, direct expression calls and return-name assignment.
- Private property family: Get/Let/Set declarations together, direct and Me-qualified reads/writes, Get return name.
- Comments, strings, labels, named argument names, types and unrelated procedures in other components remain unchanged.
- Complete project snapshot, target SHA, project selector, design mode, metadata/catalogue/source races and per-module undo guards.

## Refused before mutation
- Missing/duplicate/incomplete project snapshots, non-class target, stale SHA, bad identifier or reserved name.
- Public/Friend members and implicit Public declarations; private class instancing and all typed consumers are not proven by this contract.
- Wrong physical declaration/ProcKind; unmatched/nested procedures; duplicate accessor kind, mixed property/Sub/Function or mixed accessor visibility.
- Callback conventional names/underscores; Implements, Attribute, conditional compilation, implicit types, event/callback/dynamic invocation, With binding, brackets.
- Target/new-name variable/parameter/type/enum/external shadows, existing replacement member, project/component collision.
- Unknown qualified receiver (including another instance, external component, bang, leading-dot With), chains before Me, suffix-based/implicit target types.
- Unbound Sub expression and assignment; Function assignment outside its own return body; target token outside a procedure.
- Read failures, write failures and mismatching readback; failed native mutation is reported uncertain with no retry/rollback.

No compile or host runtime acceptance follows from lexical fixtures. Mutation uses the existing complete-project snapshot and per-module history contract.
## Complementary branch matrix (prepared before execution)
- Prepare guards: null request, blank project, oversized catalogue, null component/name/source, query callback/project collision, empty other component.
- Binding guards: replacement token without declaration, suffixed member, outside-procedure token, Sub expression/assignment, Me Function assignment, chained Me receiver.
- Member reader: standalone End, nested/missing End, Static modifier, modifier-only statement, incomplete Sub/Property signatures.
- Exclusions: terminal token, terminal label, spaced/tabbed named argument, non-label colon after expression, As/New/GoSub/Resume references.
- Service: null apply request, missing/non-class target, native wrong type/empty version, oversized export, complete attributes ending at EOF, missing/wrong-name class attributes, source null and bad BEGIN/END.
- No-op preview/apply verifies zero mutation and zero history entries.
- Redundant empty-family guard removed with authorization: exact chosen-member selection already refuses empty families.
- Per-instance export path factory: null rejected; already occupied file preserved and no export performed; default remains fresh GUID.
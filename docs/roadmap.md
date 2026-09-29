# Roadmap

The roadmap describes direction, not promised delivery dates. Current behavior
belongs in the guides and [compatibility](compatibility.md); verified results belong
in [recorded validation](test-coverage.md).

## Shared VBE experience

Consolidate the modern editor, live-project agent workflows, code review/recovery,
UserForm tooling and source versioning as one coherent workspace. Improve real
workflow reliability, discoverability, accessibility and performance without
weakening revision or permission checks.

## Host compatibility

Continue qualifying the shared VBE layer in additional applications. Add or repair
host-specific operations where VBIDE is insufficient, especially document identity,
saving and execution. Close the known Office save-adapter gaps, qualify standalone
macro persistence and broaden debugger, control, language and DPI evidence.

A host matrix should describe observed operations and prerequisites, not restrict
the intended ecosystem to the first applications used for testing.

## Quality and security

Maintain regression coverage for concurrent edits, interrupted operations, privacy
boundaries, transport failures and recovery. Re-measure the current instrumented
scope after implementation changes. Broaden authenticated-provider and real-host
checks separately from simulated protocol tests.

Keep native appearance experimental until its lifecycle, recovery and rendering
are qualified across the intended environments. Preserve usable fallback behavior.

## Distribution — later milestone

Develop and qualify the standalone installer after the current add-in work. Reuse
the [existing update contract](updates.md) for signed payloads, occupied-host handling,
upgrade, uninstall and rollback. This is not part of the documentation refactor.

Before a public release, choose the project license explicitly, complete the
payload's third-party notice review and publish only genuine compatibility and
validation information.

## Community

Keep documentation concise and maintained, provide reproducible issue templates,
and accept voluntary support without exclusive functionality or a service
commitment. A donation channel can be added once the maintainer provides its
official destination; it is not a product subscription roadmap.

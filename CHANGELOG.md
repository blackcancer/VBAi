# Changelog

This file summarizes user-visible changes. Development work is grouped under
**Unreleased** until it is associated with an actual release tag. Commit-by-commit
investigations and test transcripts belong in Git history and test artifacts.

## Unreleased

### Added

- AI conversations with Discussion, Plan and Agent modes, explicit context,
  provider selection, project permissions and reviewable code changes.
- Monaco-based editing with VBA and COM language assistance, native compilation
  diagnostics, synchronization, draft recovery and guarded attribute preservation.
- UserForm inspection and editing, native VBE tools, Git source versioning,
  checkpoints, branches, merges and GitHub integration.
- Progressive tool discovery, queued messages and resumable HTTP tool-workflow
  pauses with recorded call results.
- Opt-in native dark-theme work and the external update-program foundation.
- Host qualification scenarios for Word, PowerPoint, Access and Publisher,
  alongside existing VBE/Excel and SOLIDWORKS checks.

### Changed

- Renamed the solution, managed assembly, namespaces, ProgIDs, named pipe and
  repository to VBAi while retaining the existing COM GUIDs.
- Added migration of known legacy registration and user-data locations, without
  overwriting files already present at the destination.
- Reorganized maintained documentation in English around the shared VBE platform
  and its host-specific compatibility layers.
- Consolidated overlapping guides and retired historical investigations from the
  active documentation tree; their content remains in Git history.

### Fixed

- Cross-project read permissions and shared-context handling in conversations.
- Unbounded reception of incomplete local bridge requests.
- The former fixed eight-response HTTP workflow failure, replaced by progress-aware
  bounded segments and a resumable pause.
- Several editor synchronization, window ordering, palette recovery and
  conversation presentation issues; consult commit history for individual fixes.

### Known limitations

A standalone installer is still planned. Qualification remains operation- and
host-specific. The recorded Office batch did not qualify VBAi's Word/PowerPoint
save paths; its PowerPoint handle failure was corrected in code afterward but has
not been retested in that host. Access and Publisher still lack save adapters.
The current revision does not have a new instrumented coverage result after the
latest implementation changes.

See [compatibility](docs/compatibility.md), [recorded validation](docs/test-coverage.md)
and [the roadmap](docs/roadmap.md) for the current boundaries.

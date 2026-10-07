# Changelog

This file summarizes user-visible changes. Development work is grouped under
**Unreleased** until it is associated with an actual release tag. Commit-by-commit
investigations and test transcripts belong in Git history and test artifacts.

## 1.0.0 — 2026-10-07

### Added

- Windows x64 setup and uninstall integration, per-user registration, installation
  identity, compiled offline help and preserved user data. This release is unsigned;
  automatic updates continue to require trusted Authenticode signatures.

- A draft VBA test explorer with annotated discovery, guarded single/batch runs,
  fixtures and assertions, themed result icons, human/JSON reports, LLM tools and
  procedure coverage on Excel/Word/PowerPoint copies. Registered native execution
  and production acceptance remain incomplete; see the VBA testing guide.
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

- Outlook project identity no longer treats a process-dependent fake filename
  as persisted storage. The Q028 runner now contains its owned Ollama calculation
  workers and verifies their shutdown.

- Workspace-hosted editor closure now preserves the embedded editor for native
  WM_CLOSE requests. Detached VBE/ActiveX focus sites no longer interrupt normal
  host shutdown; the corrected candidate is qualified for the recorded Q-014
  operations in SOLIDWORKS 2019 SP5 and 2025 SP1.1.
- Cross-project read permissions and shared-context handling in conversations.
- Unbounded reception of incomplete local bridge requests.
- The former fixed eight-response HTTP workflow failure, replaced by progress-aware
  bounded segments and a resumable pause.
- Several editor synchronization, window ordering, palette recovery and
  conversation presentation issues; consult commit history for individual fixes.

### Known limitations

A standalone installer is planned. Qualification is operation-, host- and
candidate-specific; current results include selected Office persistence,
UserForm/Git workflows, SOLIDWORKS native macro workflows and embedded Ollama
assistant scenarios. A new integrated build requires its own applicable checks.
The latest changes have no new complete instrumented line/branch measurement.

See [compatibility](docs/compatibility.md), [recorded validation](docs/test-coverage.md)
and [the roadmap](docs/roadmap.md) for the current boundaries.

# Contributing to VBAi

Thank you for helping improve VBAi. Useful contributions include reproducible bug
reports, host-compatibility evidence, documentation corrections, translations,
regression tests and focused implementation changes.

Read the [code of conduct](CODE_OF_CONDUCT.md) and use
[private security reporting](SECURITY.md) for vulnerabilities rather than an issue.

## Before you start

Search existing issues and pull requests. For a substantial change, describe the
problem and proposed scope before implementing a broad redesign. Keep unrelated
refactoring separate from a bug fix.

VBAi targets the shared VBE ecosystem. Prefer host-neutral behavior at the VBIDE
boundary; isolate application-specific behavior in the compatibility layer. Do
not advertise an entire host as supported because one operation succeeds.

The repository does not yet declare a project license. Discuss external code
contributions with the maintainer before submitting them; do not assume a license
or add third-party material with unclear redistribution terms.

## Set up a development environment

Follow [source-build setup](docs/installation.md) and the
[development guide](docs/development.md). Build to an isolated output while an
application has the installed DLL loaded. Do not overwrite a running installation.

Use a branch for the change. Do not force-push shared branches, remove another
contributor's work or change COM identities as part of an unrelated task.

## Implementation expectations

- Validate project identity, permissions, revision and execution mode before a
  mutation. A newly exposed tool must not bypass these checks.
- Keep COM and UI operations on their owning threads. Treat timeouts and partial
  native mutations as uncertain outcomes, not permission to repeat an operation.
- Preserve recoverable exports, drafts and checkpoints when an operation fails.
  Never claim a successful rollback without checking the resulting state.
- Keep fixed WinForms controls editable in their Designer. Constructors used by
  the Designer must not open a host, a database, a browser or a network connection.
- Update documentation and localization with the behavior, not as an afterthought.

Follow the existing `.editorconfig`, build settings and naming conventions. Avoid
mass-formatting or translating source comments in a documentation-only change.

## Tests and evidence

Add a regression test that fails without the fix. Place focused tests in the
production-file mirror; use scenarios and integration fixtures for cross-cutting
work. [Testing](tests/README.md) describes commands and explicit host opt-ins.

Report the commands actually run, the tested revision, failures and skipped tests.
Keep unit coverage, JavaScript checks, native C++ checks and real-host evidence
separate. Never add production exclusions or empty mirror files to improve a
coverage claim. Coverage is not proof of correct behavior.

Use disposable documents and only the processes owned by the test. Existing
SOLIDWORKS sessions require explicit selection and must not be launched or closed
by an autonomous test. Do not enable all live-host or authenticated-provider flags
on a working machine indiscriminately.

## Documentation changes

Write maintained Markdown in English. Keep one authoritative page per subject and
use relative links. The README explains the product; testing evidence belongs in
[recorded validation](docs/test-coverage.md), not in promotional badges or repeated
historical counters. Follow the [documentation rules](docs/development.md#documentation).

Before submitting, run `python tools/docs/check_docs.py` from the repository root.

## Pull requests

Explain the user-visible outcome, scope, tests and remaining limitations. Include
sanitized before/after screenshots for UI work and exact application/VBE versions
for host-specific changes. Never upload production documents, credentials, private
chat histories or unreviewed logs.

The maintainer may request changes or decline work that expands the maintenance
burden beyond the project scope. Contributions and donations do not create a
support contract or a guaranteed delivery date.

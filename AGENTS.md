# Repository guidance

VBAi is a Windows x64 COM add-in for applications that host the Visual Basic Editor.
The shared VBE/VBIDE layer is the product foundation; application adapters add
host-specific operations. Read [the architecture](docs/architecture.md),
[contribution rules](CONTRIBUTING.md) and [testing guide](tests/README.md).

## Scope and boundaries

Keep a change focused. Preserve unrelated work and inspect the current branch
before editing. Do not change repository visibility, release a package, push code
from a user's macro, choose a software license or configure a donation recipient
without an explicit maintainer decision.

The maintainer authorized a public, unsigned 1.0.0 installer on 2026-10-07.
Do not claim signed distribution or installer lifecycle qualification without
current evidence. Keep the automatic updater's Authenticode trust checks intact.

Do not rename COM GUIDs, ProgIDs, public protocol fields, environment variables or
persisted formats without a migration plan. Use the current VBAi identity; legacy
names belong only in explicit migration logic and historical references.

## Build and test

From a Windows development shell, use an isolated output when a host has loaded
the installed DLL:

```powershell
dotnet build VBAi.sln -c Debug -p:BuildOutputRoot="$PWD/artifacts/build"
dotnet test tests/VBAi.Tests/VBAi.Tests.csproj -c Debug --no-build -p:BuildOutputRoot="$PWD/artifacts/build"
```

Follow [testing](tests/README.md) for native-host opt-ins. Never claim to have run
Windows/COM tests in a non-Windows environment. Report failures, skipped scenarios,
which assembly was loaded, and the exact scope of any coverage measurement.

Use only disposable projects and owned host processes. Never execute a user's
production macro as a test. SOLIDWORKS must be preloaded and explicitly selected;
do not launch or terminate it autonomously. Do not bypass host trust policies or
introduce global keyboard shortcuts/SendKeys for VBE automation.

## Implementation invariants

Keep COM and UI work on the correct owning thread. Preserve project identity,
revision, mode, approval and privacy checks across every dispatch path, including
catalog gateways. Do not retry a native mutation just because its outcome is
uncertain. Preserve backups and report partial recovery accurately.

Keep fixed WinForms layouts in their Designer and keep Designer construction free
of live services. Add focused tests in the matching production-file mirror; do not
create empty tests or coverage exclusions to satisfy a metric.

## Documentation

Maintained Markdown is English. The README is a product entry point, not an agent
work log. Consolidate overlapping pages; keep investigation history in commits,
issues or pull requests. Update the authoritative page when behavior changes.

Keep architecture targets, implemented behavior and observed host qualification
separate. Record test counts once in `docs/test-coverage.md`, tied to their actual
source revision. Do not carry forward old percentages as current results.

Run the documentation checks after moving or editing Markdown:

```sh
python tools/docs/check_docs.py
```

Preserve third-party notices, non-Markdown reference data and links used by the
application. Do not fabricate screenshots, badges, contact details, releases or
payment links. New documentation must describe the current source, not obsolete
audit findings or intended behavior that is not implemented.

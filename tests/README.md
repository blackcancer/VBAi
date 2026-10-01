# Testing VBAi

The test suite covers the shared VBE core, provider protocols, storage, editor
services and selected host integrations. A passing local suite is not a claim
that every feature works in every application that embeds the VBE.

See [development setup](../docs/development.md), [recorded results](../docs/test-coverage.md)
and the [compatibility matrix](../docs/compatibility.md). Run commands from the
repository root in a Windows development environment.

## Build and run

Use a separate output directory while an application has the installed DLL loaded:

```powershell
dotnet build VBAi.sln -c Debug -p:BuildOutputRoot="$PWD/artifacts/build"
dotnet test tests/VBAi.Tests/VBAi.Tests.csproj -c Debug --no-build -p:BuildOutputRoot="$PWD/artifacts/build"
```

An isolated build does not replace the registered add-in. Native tests must verify
which assembly the target process actually loaded. Do not overwrite a loaded DLL
or close unrelated applications to make a build succeed.

For a category-specific pass, append `--filter "TestCategory=Unit"`. Some local tests
create real Windows controls, WebView2 instances, temporary Git repositories and
simulated provider processes. They require the relevant runtimes and an interactive
Windows desktop; they are not all platform-independent pure unit tests.

## Test organization

| Location | Purpose |
| --- | --- |
| `tests/VBAi.Tests/Unit/` | Contracts and branch behavior, with source-file mirrors. |
| `tests/VBAi.Tests/Integration/` | Storage, processes, runtime controls and host boundaries. |
| `tests/VBAi.Tests/Scenarios/` | Cross-component workflows and recovery scenarios. |
| `tests/VBAi.Tests/Infrastructure/` | Shared doubles, fixtures and host helpers. |
| `tests/VBAi.Git.Smoke/` | Standalone Git diagnostics; shared scenarios also run in VSTest. |
| `tests/VBAi.Providers.Smoke/` | Provider diagnostics and a simulated CLI process. |
| `tests/native/` | Native-renderer lifecycle and managed-loader checks. |

Tests reference the production assembly rather than recompiling its sources.
Mirrors follow production paths with a `.Tests.cs` suffix, including partial
classes. Cross-cutting scenarios should complement those mirrors rather than
create a second implementation of production logic.

```powershell
powershell.exe -NoProfile -File tools/tests/Test-TestLayout.ps1 -ReportPath artifacts/test-layout/mirror-inventory.json
```

Prompt files under `Infrastructure/Fixtures/Prompts/` and the test-environment skill
under `tests/Infrastructure/.agents/` are test inputs, not user documentation.
Preserve their contents unless deliberately changing the corresponding fixture.

## Native host tests are opt-in

Use disposable documents and identify the intended process/project before any
write. Record application version, architecture, language, DPI and loaded VBAi
build. Restore temporary settings and verify that unrelated documents remain
unchanged. Skip unavailable hosts honestly instead of treating a skip as a pass.

| Opt-in variable | Scope |
| --- | --- |
| `VBAi_RUN_EXCEL_TESTS=1` | Excel integration tests (`TestCategory=Excel`). |
| `VBAi_RUN_OFFICE_TESTS=1` | The additional Office host qualification fixtures. |
| `VBAi_RUN_OUTLOOK_TESTS=1` | Outlook tests; a usable profile is also required. |
| `VBAI_EDITOR_EXCEL_TEST=1` | Monaco/Excel roundtrip (`TestCategory=MonacoExcel`). |
| `VBAI_EDITOR_LANGUAGE_EXCEL_TEST=1` | Language-service qualification (`FullyQualifiedName~MonacoLanguageExcelTests`). |
| `VBAI_NATIVE_PALETTE_EXCEL_TEST=1` | Native palette (`TestCategory=NativePaletteExcel`); observe the fixture's other-host exclusions. |
| `VBAi_SOLIDWORKS_PID` | PID of a user-preloaded SOLIDWORKS/VBE instance (`TestCategory=SolidWorks`). |

Set only the variables needed for the intended run and remove them afterward.
The SOLIDWORKS workflow must not create or kill an application instance on the
user's behalf. A host fixture can use its own native save helper; that result does
not automatically qualify VBAi's `save_host_document` adapter.

For example, after confirming that Excel tests are safe to run:

```powershell
$env:VBAi_RUN_EXCEL_TESTS = '1'
try {
    dotnet test tests/VBAi.Tests/VBAi.Tests.csproj -c Debug --no-build -p:BuildOutputRoot="$PWD/artifacts/build" --filter "TestCategory=Excel"
} finally {
    Remove-Item Env:VBAi_RUN_EXCEL_TESTS -ErrorAction SilentlyContinue
}
```

### VBA test explorer native execution

The `VbaTestExcel` category uses the Excel opt-in above and creates its own
disposable macro-enabled workbook in a newly owned Excel process. Run this
category alone to qualify the generated VBA wrapper, assertions, fixtures,
results, stale-source refusal and the fixture-owned Excel coverage-copy path:

```powershell
$env:VBAi_RUN_EXCEL_TESTS = '1'
try {
    dotnet test tests/VBAi.Tests/VBAi.Tests.csproj -c Debug --no-build -p:BuildOutputRoot="$PWD/artifacts/build" --filter "TestCategory=VbaTestExcel"
} finally {
    Remove-Item Env:VBAi_RUN_EXCEL_TESTS -ErrorAction SilentlyContinue
}
```

The compiled service runs in the external STA test host with a fixture-owned
Excel adapter. Its synthetic production module contains two eligible procedures,
only one entered by the selected test bodies: the expected hand-calculated
procedure coverage is 1/2 = 50%. This is an assertion in the fixture, not a claim
that its current native run passed. It verifies original-source preservation and
mapped measurement when execution completes. It does not measure statements or
branches.

This external fixture does not qualify installed in-process callback activation,
native tool-window docking or other VBE hosts. The fixture records the assembly MVID
and Excel version, verifies normal owned-process exit and preserves report
examples under the test output's `test-explorer-native` directory. It does not
change Trust Center settings or execute an existing user macro.

For installed Excel qualification,
[Test-RegisteredVbaTestExplorer.ps1](../tools/testing-explorer/Test-RegisteredVbaTestExplorer.ps1)
checks the candidate assembly/MVID, registered add-in bytes and x64 callback
registration before creating a host. Its required arguments are
`CandidateAssemblyPath`, `ExpectedMvid` and `OutputDirectory`; use Windows
PowerShell 5.1 x64 in STA mode. Without `ExecuteOwnedFixture`, it only performs
preflight and reports that execution was not started. The explicit switch creates
an owned disposable Excel fixture. The script neither installs/registers the
candidate nor changes Office trust settings; deployment requires a separate
reviewed decision.

Use `-TestSubsystemOnly -ExecuteOwnedFixture` to qualify the VBA testing
subsystem: discovery/support installation, the registered native batch, human and
JSON reports, measured procedure coverage on a disposable copy, exact native
explorer actions, report consistency and stale-revision refusal. This option
skips the general Excel `Workbook_BeforeSave` positive save control and explicitly
records `NOT_TESTED_BY_SCOPE`. The synthetic event remains an eligible coverage
procedure; its counters and entry probe are retained as observations only, without
claiming that an independently qualified handler was suppressed. Original source,
probe mappings, coverage totals/percentage and exact `EnableEvents` restoration
remain required assertions. The fixture is saved, then the event-counter baseline
is read after the completed batch with owned PID/project/path/source checks.

The legacy default additionally requires the positive save control.
`-ExecutionAndUiOnly` skips that control **and measured coverage**; it cannot
qualify coverage. These two scope switches are mutually exclusive and are rejected
before any host activation if combined. None of these scenarios qualifies unrelated
VBAi features or the LLM permission/approval boundary; run the focused testing-tool
boundary regressions separately.

Shared callback code recognizes known VBE host names, but recognition is not
native qualification. The requested production scope includes Excel, Word,
PowerPoint, Access, Publisher, Outlook and a user-preloaded, explicitly selected
SOLIDWORKS instance; Visio and Project are excluded. Identify each loaded assembly,
actual callback transport, host version and tested operation. Do not infer that a
returned-value Excel fixture qualified the shared COM callback in another host.
Record failures and skipped scenarios in [validation](../docs/test-coverage.md),
not as successful host coverage.

Excel, Word and PowerPoint returned-value and coverage-copy adapters exist in the
source; that fact does not qualify those adapters or the shared callback natively.
The fixture and installed qualification script above remain Excel-specific.
PowerPoint qualification must independently verify its PIA argument-array
invocation, macro-enabled presentation copy and application-level event behavior.
Word qualification must verify saved DOCM/DOTM/DOC/DOT copying, refusal of unsaved
VBA source/reference mismatches, exact owned activation and document-qualified
positional invocation without retries after uncertain completion.

`VbaTestOffice` uses `VBAi_RUN_OFFICE_TESTS=1` for registered in-process Office
execution. Its Word/PowerPoint rows include a separate synthetic production module
with two eligible functions, only one called by the selected test: the expected
procedure measurement is 1/2 = 50%. This is a fixture assertion, not an observed
passing result. It checks canonical IDs/revision/probes, human/compact reports,
unchanged original source, file and counters, closed copies and retained artifacts.
The fixture may coexist with existing hosts only after identifying exactly one
new PID against its initial process inventory; it never reuses a user host.

`VbaTestExcelLarge` requires `VBAi_RUN_EXCEL_LARGE_TESTS=1`. It builds a fresh
catalogue spanning dispatcher leaf and route boundaries, verifies compilation of
the complete support module through a returned VBA verdict, and executes only
explicit boundary and failure selections. It uses the external STA Excel adapter;
it does not qualify the registered callback or explorer UI. Sources, references,
revisions and both reports are retained in its private temporary evidence folder.

`VbaTestOutlook` requires `VBAi_RUN_OUTLOOK_TESTS=1`, a usable existing profile,
no running Outlook process and an initially absent `VbaProject.OTM`. It creates
only an unsaved disposable inspector. It never sends or saves a mail item,
changes a profile, replaces an existing OTM or changes trust settings. An opaque
FileName on a new unpersisted project is selected by its unique project name;
the fixed per-user OTM path is used only for recovery. The fixture verifies its
initial blank source/reference baseline, removes only its unchanged owned modules,
and deletes a newly created OTM only after verified normal exit and retained backup.
Unknown native completion prevents both inspector closure and Quit. All failures
and retained recovery files are reported.

`VbaTestSolidWorks` requires `VBAi_RUN_SOLIDWORKS_TEST_EXPLORER=1`,
`VBAi_SOLIDWORKS_PID` and the absolute `VBAi_SOLIDWORKS_TEST_MACRO` path. The user
must preload that process and open the selected disposable `.swp` manually through
**Tools > Macro > Edit**. Save a blank macro containing only empty standard modules
or `Option Explicit`; remove the default generated macro body before qualification.
The fixture refuses other source, a different loaded candidate or an unsaved macro.
It never opens or closes a macro, activates application COM, saves, launches or
terminates SOLIDWORKS. Its own test/support modules are removed only after verified
settled completion and unchanged ownership hashes. The macro remains open; its
disk bytes remain unchanged, and the in-memory project may remain marked unsaved.
Other projects' sources are not inspected. This scenario covers ordinary tests,
reports, stale refusal and native window identity; it does not measure coverage.

Coverage qualification must verify Excel/Word/PowerPoint clone identity, unchanged
original live/disk source, original-source probes, explicit exclusions,
unsafe-syntax refusal and incomplete/unknown measurements. Verify restored Excel
event settings, Word AutoOpen/document/application/Normal-template effects and
PowerPoint open/before-close handler effects, including
cancelled closure and retained copies. These copies are not external-system
sandboxes. Retained document copies and coverage plans can contain full private source and data. Keep evidence
outside maintained documentation and follow [privacy](../docs/privacy.md).

## JavaScript and native renderer

The editor's JavaScript tests use Node's test runner:

```powershell
node --test tools/tests/Test-MonacoLanguage.mjs tests/VBAi.Tests/Infrastructure/Fixtures/Editor/MonacoEditing.Scenarios.mjs
```

See [native renderer testing](native/README.md) for the C++ self-test and loader.
These checks are distinct from real VBE rendering, host integration and measured
native-code coverage.

## Provider qualification

`VBAi_RUN_OLLAMA_TESTS=1` enables `TestCategory=Ollama` against the loopback
server through the production HTTP client. `VBAi_TEST_OLLAMA_MODEL` selects an
already-installed model (default `qwen2.5:3b`). These synthetic scenarios cover
streamed text, a harmless tool roundtrip, cancellation and a subsequent request;
they do not execute native VBE tools or read saved provider settings. Remove the
opt-in variables after the run.

`VBAi_RUN_OLLAMA_UI_TESTS=1` enables `TestCategory=OllamaUi` with the same
model selector. It shows the real chat controls and checks send, rendered
streaming, Stop and a subsequent completed response through the production
loopback HTTP client. Settings and history are isolated; the VBE project is
simulated and native tools are refused. This qualifies a detached chat workflow,
not Office or SOLIDWORKS integration. Run it separately from native host UI tests
to avoid competing for focus, then remove its opt-in variable.

`VBAi_RUN_OLLAMA_EXCEL_TESTS=1` enables `TestCategory=OllamaExcel` using the
fixed loopback endpoint and already-installed `qwen2.5:3b` model. It creates a disposable Excel module containing a
random marker absent from the prompt, dispatches the model's `read_module` call
through the real project-bound VBE tools, and verifies the final answer and
unchanged source and normal exit of its owned Excel process. No macro runs.
This covers production provider/tools/session code dispatching through native
Excel COM from the test process, not the installed bridge or embedded assistant UI.
Run it separately from other native host tests
and remove the opt-in afterwards.

Existing-account checks require both `VBAi_CONNECTED_PROVIDER_TESTS=1` and
`VBAi_CONNECTED_SOURCE_TESTS=1`. The Git read check additionally requires
`VBAi_TEST_GITHUB_MANIFEST` identifying an explicitly authorized synthetic private
repository (URL, ID, verified ownership/private state and expected main commit).
The Codex check requires `VBAi_TEST_SOLIDWORKS_MANIFEST` with `OwnedDisposable`,
absolute `Path`, `FileSha256`, `Module`, `ModuleSha256`, `Marker`, `Pid` and `Mvid`.
It accepts only the manifested disposable SWP under the qualification artifacts,
uses an already-connected account and never signs in or copies authentication state.

Native UserForm GitHub qualification is a separate explicit scope:
`VBAi_RUN_USERFORM_GITHUB_TESTS=1`, `VBAi_RUN_EXCEL_TESTS=1`, the Git manifest and
an absolute `VBAi_TEST_USERFORM_GIT_OUTPUT`. The fixture is constrained to the
maintainer-authorized retained qualification repository and creates a dedicated
branch containing synthetic exports. It verifies controls, code and FRX bytes,
retains backups and captures the owned designer windows for visual review.
`VBAi_RUN_USERFORM_CORRUPTION_TESTS=1` separately enables the local-only malformed
FRX diagnostic; it must not publish corrupt content. Run either scenario only
while it owns the desktop. Normal host shutdown is part of acceptance; a verified
transfer does not excuse a subsequent crash. These fixtures are not enabled by
ordinary connected-account opt-ins.

Local provider tests use simulated HTTP or CLI transports. Do not use personal
credentials, paid API calls or private project data without explicit permission.
A successful model catalog lookup alone does not qualify streamed responses,
tool execution, cancellation or recovery.

For a live qualification, use synthetic content and a harmless tool first. Record
the provider, model, CLI/protocol version and precisely what was exercised. Keep
secrets out of fixtures, logs, screenshots and reports. See the
[provider guide](../docs/providers.md).

## Coverage and evidence

A coverage run can use the installed collector:

```powershell
dotnet test tests/VBAi.Tests/VBAi.Tests.csproj -c Debug --no-build -p:BuildOutputRoot="$PWD/artifacts/build" --collect:"XPlat Code Coverage" --results-directory artifacts/coverage
```

Inspect the resulting report's assembly scope and filters. State numerator,
denominator, tested commit, exclusions and skipped host tests. Coverlet results
for the managed add-in do not measure C++, JavaScript or every native COM path.
Do not exclude production code or weaken assertions to manufacture a target.

Keep machine-local artifacts outside the maintained guide tree. Publish a concise,
versioned summary in [recorded validation](../docs/test-coverage.md), separating
unit/runtime tests, native host observations, Designer checks and live-provider
runs. An old 100% result does not describe a later build.

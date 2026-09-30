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

For an owned-host scalar-inspection investigation, set `VBAi_VBE_INSPECTION_TRACE`
to an absolute local JSONL file path in the host's environment before launching it.
The parent directory must already exist. Tracing is disabled when absent or invalid;
unavailable logging never changes command execution. Evidence is bounded to 128
events per inspection and 1 MiB per file. Rows contain correlation, fixed phase,
PID, thread/apartment, elapsed time and exception type; they contain no project
path, code, expression or inspected value. A client timeout does not cancel or
authorize replay of a pending native command. Retain the host and phase evidence
when the bridge stops responding.

The two `ExcelScalarDiagnostics` scenarios use a dedicated explicit bootstrap,
not the ordinary Excel COM-activation fixture. Set `VBAi_RUN_EXCEL_TESTS=1`,
`VBAi_VBE_INSPECTION_TRACE` to an absolute local JSONL path with an existing parent
directory, and `VBAi_EXCEL_RESULTS` to a durable local evidence directory. The
bootstrap resolves the installed x64 Excel executable from HKLM App Paths and
launches it once with `/x /automation` and an owned, empty, macro-free `.xlsx`
seed. Microsoft documents that
[/automation suppresses automatically opened files and auto-run macros](https://learn.microsoft.com/en-us/troubleshoot/microsoft-365-apps/excel/files-open-automatically).
The trace variable is supplied explicitly in the child process environment;
setting it only in the test process does not prove that COM activation inherited
it. No machine/user environment setting is changed.

NativeOM attachment accepts only the launched PID's `EXCEL7` document. Before
any workbook/VBE mutation, the fixture checks the retained process handle,
executable/start identity, application HWND/PID and an inventory containing only
the exact seed workbook. It never falls back to COM activation or an active
application. Startup, loaded add-in MVID and shutdown evidence remain in the
fixture directory even on success. If attachment or startup is uncertain, the
process is retained without Close, Quit, termination or another launch. Execute
`InstalledBridgeSkipsUnsupportedScalarPageWithoutQuickWatch` first, then
`InstalledBridgeReadsOneLongScalarWithNativePhaseEvidence`, using separate exact
method filters; a category batch does not guarantee their order. Both require
nonempty correlated host phase evidence and the installed product MVID must
match the assembly referenced by the tests.

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
The formatting-options scenario can additionally use an absolute
`VBAi_TEST_FORMAT_OPTIONS_OUTPUT` directory to retain request/response and
before/after evidence independently of VSTest attachment retention, including
successful runs. It creates a unique subdirectory and never retries a failed
mutation to gather evidence.
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

Native **local Git** UserForm layout qualification uses
`VBAi_RUN_USERFORM_LOCAL_GIT_TESTS=1` and `VBAi_RUN_EXCEL_TESTS=1`, with optional
absolute `VBAi_TEST_USERFORM_LOCAL_GIT_OUTPUT` for durable evidence. Filter on
`FullyQualifiedName~NativeUserFormLocalGitTests`. Each layout owns a separate
disposable Excel process and tests unchanged captures, a native property change,
exact local Git transport, checkpoint import, measured explicit rollback and
helper save/reopen with macros disabled. It covers Label/Button, TextBox,
ComboBox, ListBox, CheckBox, OptionButton, ToggleButton, ScrollBar, SpinButton,
TabStrip, an Image with a picture, and nested Frame/MultiPage. Unsupported
comparison grammars are retained unchanged and may fail these scenarios; no
normalizer exception or coverage claim is added. No remote Git operation occurs.
Run only while owning the desktop; retain failure artifacts and review the
source/reopened designer captures separately. Normal host exit is required.

The Image layout reads its actual content digest through the production
`form_tree` descriptor inside Excel. It does not marshal a process-local OLE
picture to the test process. That readback records the descriptor type and PNG
content SHA-256, rather than an external picture handle or HIMETRIC dimensions.

To diagnose native UserForm export failures separately, enable
`VBAi_RUN_USERFORM_EXPORT_PROBES=1` and `VBAi_RUN_EXCEL_TESTS=1`, then filter on
`FullyQualifiedName~NativeUserFormExportProbeTests`. The six disposable trials
compare external STA and production host-bridge dispatch to fixture temporary,
production `GitTemporary` and evidence directories. Each trial uses a new
directory and the same form filename, attempts at most one export, records exact
PID, loaded assembly, document/component identity and raw files, and requires
normal owned-host exit. Failed exports are not retried or redirected. Partial
files remain available for diagnosis. Use the same optional absolute
`VBAi_TEST_USERFORM_LOCAL_GIT_OUTPUT` to retain reports. These probes do not
qualify Git capture comparison, imports or recovery.

The separate `ControlledSiblingEfsExportPreservesIdentityAndRetainsRawEvidence`
method runs two synthetic host-bridge trials with the same opt-ins. Each creates
a fresh unencrypted parent under the evidence root and two ASCII-named sibling
directories, encrypts only the empty `efs001` sibling, verifies both attributes,
then exports once to the selected sibling. Setup failures remain failures before
export. It records parent/temp attributes and retains both siblings; it never
decrypts a directory or modifies the real VBAi profile. This isolates EFS from
the production cache path without relocating any user's macro data.

`ControlledVolumeAndAncestorEfsExportRetainsOwnedIdentity` adds four synthetic
host-bridge trials: C/E volumes crossed with encryption on a fresh parent before
creating its children, or only on the empty export leaf. C trials use a unique
owned directory under the existing C TEMP root; E trials require an explicit
evidence root on E. The new parent must initially be unencrypted, and inherited
versus directly applied EFS is verified before the single export. All siblings,
raw files and reports remain available. This changes no existing directory's
encryption, decrypts nothing and does not qualify the actual GitTemporary cache.

`InheritedStorageAncestorExportRetainsExactOwnedIdentity` runs four host-bridge
trials in fresh children of the existing LocalAppData, LocalAppData/VBAi,
GitTemporary and TEMP directories. It changes no parent's ACL or encryption.
Each child records inherited attributes and ACL, verifies a new synthetic text
file can be written/read, and queries public EFS certificate hashes while keeping
metadata access denial explicit. Each process then attempts exactly one native
form export. Reports, raw and partial files are retained; no private key is read.
The same four trials record the testhost and exact owned Excel process tokens:
public user SID, session, integrity SID, elevation type, restricted-token flag and
AppContainer flag. These observations use query-only handles, change no token or
permission, and retain access failures as unverified metadata.

Designer screenshots reobserve the COM window after bounded UI settlement. A
zero designer HWND permits capture of the owned VBE root only after verifying
the exact active project and designer COM identities, captions, type, visibility
and root PID. Each screenshot has a `.png.json` sidecar recording observations,
capture scope and failures. A root capture covers the entire VBE with the exact
designer active and must be reviewed with that scope; it is not a form-only crop.

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

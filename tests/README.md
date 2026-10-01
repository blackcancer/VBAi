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

The Q-027 path-visibility diagnostic is a temporary, explicit host opt-in. Run
only `ExcelPathVisibilityTests.OwnedOwnerStaObservesTwoTesthostCreatedGuidDirectoriesAndEffectiveToken`
with `VBAi_RUN_EXCEL_PATH_VISIBILITY=1`, `VBAi_RUN_EXCEL_TESTS=1` and an absolute
durable `VBAi_EXCEL_RESULTS` directory. The owned `/x /automation` bootstrap
passes `VBAi_TEST_PATH_VISIBILITY_MANIFEST` explicitly to the child. The manifest
contains only two direct GUID children of LocalAppData and TEMP. Fixed synthetic
files and the manifest remain under read leases during observation and are
retained as evidence. No personal file is read, no macro/export/source edit runs,
and no ACL, attribute, trust policy or token is modified.

The internal, parameter-free `diagnostic_path_visibility` bridge command is
disabled unless the host received this manifest before add-in connection. It is
absent from the LLM catalogue. It checks captured owner PID/native TID/STA and
records managed existence/attribute results, explicit exception types/HRESULTs,
native attributes and immediately captured LastError (meaningful only on native
failure). Effective token evidence queries the thread with OpenAsSelf=true and
falls back to the primary token only for ERROR_NO_TOKEN (1008); metadata is
limited to SID, integrity, restricted/AppContainer state, type/impersonation and
AuthenticationId/TokenId. Compare testhost before/after with the actual owner STA;
do not label the testhost's STA as Excel's. Visibility differences are recorded,
not assumed equal or attributed to a cause. Pending delivery/startup retains the
host and read leases without replay or native cleanup. This diagnostic changes
the product candidate; qualification of an earlier binary does not cover it.

After deploying and verifying the matching candidate, run exactly one native
observation (the command below is a recipe, not recorded execution):

```powershell
$env:VBAi_RUN_EXCEL_TESTS = '1'
$env:VBAi_RUN_EXCEL_PATH_VISIBILITY = '1'
$env:VBAi_EXCEL_RESULTS = "$PWD/artifacts/path-visibility-native"
dotnet test tests/VBAi.Tests/VBAi.Tests.csproj -c Debug --no-build -p:BuildOutputRoot="$PWD/artifacts/build" --filter 'FullyQualifiedName=VBAi.Tests.Integration.ExcelPathVisibilityTests.OwnedOwnerStaObservesTwoTesthostCreatedGuidDirectoriesAndEffectiveToken' --logger 'trx;LogFileName=path-visibility-native.trx' --results-directory "$PWD/artifacts/path-visibility-native"
```

For an owned-host scalar-inspection investigation, set `VBAi_VBE_INSPECTION_TRACE`
to an absolute local JSONL file path in the host's environment before launching it.
The parent directory must already exist. Tracing is disabled when absent or invalid;
unavailable logging never changes command execution. Evidence is bounded to 128
events per inspection and 1 MiB per file. Rows contain correlation, fixed phase,
PID, thread/apartment, elapsed time and exception type; they contain no project
path, code, expression or inspected value. A client timeout does not cancel or
authorize replay of a pending native command. Retain the host and phase evidence
when the bridge stops responding.

The `ExcelScalarDiagnostics` scenarios use a dedicated explicit bootstrap,
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
`InstalledBridgeReadsOneLongScalarWithNativePhaseEvidence`, and then
`InstalledBridgeReadsFullScalarPageWithThreeNativeObservers`, using separate exact
method filters; a category batch does not guarantee their order. Each requires
nonempty correlated host phase evidence and the installed product MVID must
match the assembly referenced by the tests.

The full page issues one `inspect_local_scalars` request with Offset=0/Limit=6.
It checks the exact Long/String/Boolean probe values, declaration order/type/
position, and array/Variant/object rows skipped without a value or error. Native
phase evidence must contain one correlation and terminal inspection, with a
complete Command229/observer lifecycle for each eligible scalar. Selection,
source and break mode are checked before and after; normal owned Close/Quit and
exit are required. These synthetic runtime assertions do not establish complete
Locals enumeration or explain a prior crash. Recorded acceptance and loaded
candidate identity are in [validation results](../docs/test-coverage.md).

For a tests-only follow-up against a frozen installed candidate, use a separate
output containing hash-verified copies of its product/dependency binaries and
build the test project with `BuildProjectReferences=false`. Verify the product
hash again in the final test output; do not rebuild the product implicitly or
overwrite a test assembly whose native campaign is still active. The exact full
page filter is:

```text
FullyQualifiedName=VBAi.Tests.Integration.ExcelLocalScalarInspectionTests.InstalledBridgeReadsFullScalarPageWithThreeNativeObservers
```

Office adapter-only diagnostics distinguish transient Access CurrentProject
wrapper identity from the database path and selected VBProject identity. Keep
sampled objects alive across identity comparisons and balance their temporary
IUnknown references; do not weaken native dispatch guards or invoke Save to
diagnose an identity getter. A read-only identity trial with normal exit does
not qualify adapter persistence. Adapter/property/reference acceptance requires
its own verified save and fresh-disk reopen without post-save helper saving.

For Access save/reopen, preserve the exact live and reopened metadata strings,
not only a successful Save response or Saved flag. Reference-removal acceptance
requires the initial owned process to exit normally before a fresh process
reopens the database. A returned Quit is insufficient; preserve the disposable
database and report disk readback as NOT_RUN if exit is unverified. Do not replay
Quit or native mutations to obtain a passing result.

Office fixtures retain the original process handle, PID/start/image identity and
shutdown failure when Quit returns but exit is not observed within the fixture
deadline. `shutdown-lifecycle.json` records the single Quit outcome, bounded exit
observation and handle disposition; the final qualification report retains this
ledger too. An unverified shutdown refuses further native requests, saving,
Close/Quit attempts and reopening. Later PID absence, manual termination or an
eventual exit does not overwrite the original failed gate or establish an exit
code that was not read. Releasing a client COM reference is not proof that every
server reference disappeared. These retention checks do not diagnose a host's
failure to exit.

### Read-only Access metadata getter probe

Set both `VBAi_RUN_OFFICE_TESTS=1` and
`VBAi_RUN_OFFICE_METADATA_GETTER_PROBE=1` to enable this diagnostic. The
`FreshAccessMetadataGetterContractsReadOnly` test starts a new owned disposable
database through the existing Access fixture. Database creation and its initial
fixture save remain prerequisites; the probe adds no setter, save, macro or help
invocation. The baseline requires three successful, exactly equal getters and
runtime getter/setter contracts for both DISPIDs. It does not qualify metadata
persistence.

The diagnostic reads `HelpFile` and `HelpContextID` on the exact mapped, selected
VBProject through `PropertyDescriptor.GetValue`, CLR `InvokeMember` with
`GetProperty`, and raw `IDispatch.Invoke` with `DISPATCH_PROPERTYGET`. The raw
result uses the x64 24-byte VARIANT ABI, a checked boundary canary and OLE cleanup.
Runtime `GetTypeInfo` records the getter/setter VARTYPEs for DISPIDs 116 and 117
and the containing library identity. Reads run on the external fixture STA;
COM marshaling dispatches them to the Office object's apartment. These are not
in-process bridge getter observations.

Each `metadata-getters-<phase>.json` retains candidate MVID and matching assembly
file hashes, PID/start identity,
project path/IUnknown identity, mode/protection/Saved state, getter outcomes and
bounded exact BSTR bytes. No ANSI repair or value normalization is performed.
A pending-read marker identifies an unfinished call. Normal owned process exit
must be verified separately through the fixture's lifecycle evidence.

With the same diagnostic opt-in, the existing Access HelpFile/HelpContextID
adapter scenarios also record reads before and after their existing mutation,
after their existing adapter save, and immediately after fresh-disk reopen.
These hooks add no mutation or save. Getter disagreements remain observations;
the original exact metadata assertions still decide adapter acceptance. An
`OBSERVED` report does not mean its getters agree or persistence passed.

Prepare a new durable evidence directory and use a test output referencing the
exact installed candidate as described above. The following filter selects the
complete diagnostic batch; no ordering between tests is assumed:

```powershell
$env:VBAi_RUN_OFFICE_TESTS = "1"
$env:VBAi_RUN_OFFICE_METADATA_GETTER_PROBE = "1"
$env:VBAi_OFFICE_RESULTS = "$PWD/artifacts/metadata-getter-evidence"
dotnet test tests/VBAi.Tests/VBAi.Tests.csproj -c Debug --no-build -p:BuildOutputRoot="$PWD/artifacts/build" --filter "FullyQualifiedName=VBAi.Tests.Integration.OfficeMetadataGetterProbeTests.FreshAccessMetadataGetterContractsReadOnly|FullyQualifiedName=VBAi.Tests.Integration.OfficeAdapterOnlyMetadataQualificationTests.Access16HelpFilePathAdapterSaveReopen|FullyQualifiedName=VBAi.Tests.Integration.OfficeAdapterOnlyMetadataQualificationTests.Access16HelpContextIdAdapterSaveReopen" --results-directory "$PWD/artifacts/metadata-getter-evidence" --logger "trx;LogFileName=metadata-getters.trx"
```

### One-shot Access metadata setter comparison

The additional `VBAi_RUN_OFFICE_METADATA_SETTER_PROBE=1` opt-in enables four
`OfficeMetadataSetterProbe` cases: HelpFile/HelpContextID, each through the
existing production CLR `SetNative` implementation or raw `IDispatch` PROPERTYPUT.
Every case creates a separate owned Access database, prepares the existing
synthetic module/class baseline and attempts one metadata setter. A returned
failure is followed only by read-only getter observations before the original
error is rethrown. Independent read/cleanup failures are aggregated; uncertain
bridge delivery refuses further native reads or mutations. There is no setter
replay, rollback claim or post-failure save.

Both new setter paths run from the same external fixture STA, use the exact
selected project identity and current production revision, and require design
mode, no project protection and the verified candidate MVID/assembly-file hashes.
Their context differs from the in-process bridge path; include the existing
adapter HelpFile/HelpContextID scenarios as separate controls. A difference
between a bridge case and an external case alone does not establish a binder bug.

The raw setter uses exact BSTR/I4 inputs without coercion, architecture-correct
24-byte argument/result buffers, checked canaries and OLE cleanup. It passes one
named `DISPID_PROPERTYPUT` argument, `DISPATCH_PROPERTYPUT` only and the invariant
locale used by the CLR comparison. Microsoft documents that
[PROPERTYPUT requires the named argument and ignores the result](https://learn.microsoft.com/en-us/windows/win32/api/oaidl/nf-oaidl-idispatch-invoke).
`metadata-setter.json` retains the original native HRESULT and input/result
buffer observations. Successful calls proceed through the unchanged adapter
save, normal process exit and fresh-disk exact metadata assertions. Byte-packing
is retained as observed; these diagnostics add no ANSI recovery heuristic.

After preparing a matching frozen-candidate test output and a new evidence
directory, run the four cases and both bridge controls together:

```powershell
$env:VBAi_RUN_OFFICE_TESTS = "1"
$env:VBAi_RUN_OFFICE_METADATA_GETTER_PROBE = "1"
$env:VBAi_RUN_OFFICE_METADATA_SETTER_PROBE = "1"
$env:VBAi_OFFICE_RESULTS = "$PWD/artifacts/metadata-setter-evidence"
dotnet test tests/VBAi.Tests/VBAi.Tests.csproj -c Debug --no-build -p:BuildOutputRoot="$PWD/artifacts/build" --filter "TestCategory=OfficeMetadataSetterProbe|FullyQualifiedName=VBAi.Tests.Integration.OfficeAdapterOnlyMetadataQualificationTests.Access16HelpFilePathAdapterSaveReopen|FullyQualifiedName=VBAi.Tests.Integration.OfficeAdapterOnlyMetadataQualificationTests.Access16HelpContextIdAdapterSaveReopen" --results-directory "$PWD/artifacts/metadata-setter-evidence" --logger "trx;LogFileName=metadata-setters.trx"
```

The native project-properties dialog exposes these Help fields on its
[General tab](https://learn.microsoft.com/en-us/office/vba/language/reference/user-interface-help/project-properties-dialog-box).
The existing guarded `queue_project_properties_dialog` command opens that
dialog; it does not edit its controls. UI field discovery, a separately owned
single-change trial and normal save/reopen proof are prerequisites for evaluating
that alternative. No native dialog edit is included in this setter batch.

### Controlled native export tracing

`tools/probes/Trace-NativeUserFormExport.ps1` defaults to a prepare-only plan.
Its explicit `-Execute` path requires the absolute pending report, expected
MVID/SHA-256 and a measured `-DebuggerPreflightReport` from
`tools/probes/Test-NativeTracePreflight.ps1`. That preflight uses only a new
disposable non-Office helper and must prove paired native tracing, attachment,
target survival, normal shutdown and debugger detach with the exact CDB,
JsProvider and trace-script hashes. It does not qualify native Office export.

Trace execution validates the owned Excel PID/start identity and existing
absolute GUID export child, refuses a preexisting trace output or attached
debugger, and arms the export permission marker only after attachment is proven.
The bounded capture records paired syscall arguments/statuses and its detach
lifecycle; it does not issue exports or change ACLs, EFS, tokens or trust. The
fixture issues one export only. If setup fails before arming, preserve zero-export
evidence; if a native export fails, retain its original response and do not replay
it. Record detach and owned-host shutdown independently from export acceptance.

Pending/preflight JSON is read explicitly as UTF-8, including accented repository
paths. CDB command files intentionally use the active Windows ANSI code page
without a BOM; debugger logs use Unicode. These are separate encoding contracts.
A native path/name-not-found status despite a successful synthetic root write
localizes the observed failure but does not prove an EFS, ACL or token cause.
Exact candidates and measured results belong in recorded validation.

If a native campaign aborts, retain the final TRX and reconcile individual
outcomes with its raw counters and the planned scenario inventory. An in-flight
scenario absent from the TRX and unreached cases are NOT_RUN; NotExecuted is not
a successful fixture skip. Explicitly authorized forced cleanup is never normal
host shutdown acceptance. Record native debugger ownership, any termination
timeout and whether an exit code was actually observed; transient PID absence
does not establish completed cleanup.

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

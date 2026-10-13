# Testing VBAi

Run commands from the repository root on Windows. Start with a host-free managed
scope; select native hosts or real providers explicitly when the change requires
them. [Recorded validation](../docs/test-coverage.md) contains results, while this
guide describes how to run checks.

## Build and run

Use a separate output when an application has loaded the registered DLL:

```powershell
dotnet build VBAi.sln -c Debug -p:BuildOutputRoot="$PWD/artifacts/build"
dotnet test tests/VBAi.Tests/VBAi.Tests.csproj -c Debug --no-build -p:BuildOutputRoot="$PWD/artifacts/build"
```

Use Visual Studio MSBuild from a development shell if the .NET SDK cannot locate
the Windows SDK/native resources. Required runtimes and toolchains are documented
in [source-build setup](../docs/installation.md).
For a focused gate, append `--filter "TestCategory=Unit"` or a declared class
filter. Even local tests can create actual controls, WebView2 instances, synthetic
CLI processes and temporary Git repositories; use a suitable Windows desktop.
An isolated build does not replace the registered add-in.

`VBAi.Tests` uses [its runsettings file](VBAi.Tests/VBAi.Tests.runsettings) to run
the Windows UI/COM bank in a dedicated x64 testhost process without a child
AppDomain. This avoids requesting MSTest's child-domain unload for this native
interop bank. Fixtures still disconnect owned providers, drain admitted work and
observe their original threads and processes; the complete gate requires normal
testhost exit and unchanged inputs. This configuration does not enable native-host
opt-ins or skip tests. VSTest supports the setting through
[`DisableAppDomain`](https://github.com/microsoft/vstest/blob/main/docs/configure.md).
The project supplies it to `dotnet test` and Visual Studio. For a direct DLL run,
pass `/Settings:<absolute-path-to-VBAi.Tests.runsettings>` explicitly. Test assets
are resolved from the physical test assembly rather than the runner's directory.
Ordinary test execution uses MTA. UI tests must select their dedicated STA or
WinForms fixture attribute so a control-created synchronization context cannot
leak onto the shared asynchronous-test thread.

Packaged PowerShell controller tests run local scripts in child processes. A
restricted default execution policy can reject them before their guard assertions
run. Use an approved process-only policy for that test session when necessary;
no machine or current-user policy change is required. For example, set and restore
`PSExecutionPolicyPreference` around the selected test command in a dedicated
PowerShell process. Recorded gates state the effective scope and retain original
policy failures separately. See Microsoft's
[execution-policy scopes](https://learn.microsoft.com/en-us/powershell/module/microsoft.powershell.core/about/about_execution_policies?view=powershell-5.1).
## Choose the check

| Change | Applicable checks |
| --- | --- |
| Managed behavior | Source-file mirror tests and affected integration/scenario tests |
| Monaco source/assets | JavaScript scenarios and WebView contracts |
| Fixed forms | Designer construction/serialization plus relevant rendering tests |
| COM/VBE/native behavior | [Owned native-host qualification](native-hosts.md) |
| Provider/authentication/streaming | [Provider qualification](providers.md) |
| Native renderer | [Native build, loader and lifecycle checks](native/README.md) |
| Markdown | Documentation structure/link checks below |

Implement the complete scenario matrix and oracles before running a native batch.
Run once for the frozen candidate, then diagnose failures with retained evidence.
Do not repeatedly open hosts during development or replay uncertain mutations.

## Test organization

| Location | Responsibility |
| --- | --- |
| `tests/VBAi.Tests/Unit/` | Production-file mirrors, including partial classes |
| `tests/VBAi.Tests/Integration/` | Storage, processes, controls and live boundaries |
| `tests/VBAi.Tests/Scenarios/` | Cross-component workflows and recovery |
| `tests/VBAi.Tests/Infrastructure/` | Shared doubles, fixtures and host helpers |
| `tests/VBAi.Git.Smoke/`, `tests/VBAi.Providers.Smoke/` | Standalone diagnostics and simulated provider processes |
| `tests/native/` | Native-renderer and managed-loader checks |

Tests reference production assemblies. Mirrored `.Tests.cs` files follow the
production directory/name and cover behavior, not a copied implementation.
Cross-cutting scenarios complement the mirrors. Check layout with:

```powershell
powershell.exe -NoProfile -File tools/tests/Test-TestLayout.ps1 -ReportPath artifacts/test-layout/mirror-inventory.json
```

Markdown under `Infrastructure/Fixtures/` and the environment skill under
`tests/Infrastructure/.agents/` are test inputs. Preserve them when editing guides.

## Native host tests are opt-in

Use only disposable projects and owned process identities. SOLIDWORKS requires
explicit selection/authorization under repository rules. Native execution never
uses SendKeys or global shortcuts and never changes host trust policy.
See [native-hosts.md](native-hosts.md) for flags, original-handle lifecycle,
candidate registration and diagnostic retention.

## JavaScript, UI and documentation

```powershell
node --test tools/tests/Test-MonacoLanguage.mjs tests/VBAi.Tests/Infrastructure/Fixtures/Editor/MonacoEditing.Scenarios.mjs
python -m unittest discover -s tools/docs -p "test_*.py"
python tools/docs/check_docs.py
```

Edit `assets/editor/src/` and rebuild the intended distribution using
`tools/Build-MonacoAssets.ps1`. [Development](../docs/development.md) describes
asset provenance, localization and Designer rules.
The [tools index](../tools/README.md) distinguishes offline, standalone UI,
registered-host and explicit rendering commands. Choose the appropriate desktop;
a detached screenshot cannot qualify native hosting.

## Coverage and evidence

```powershell
dotnet test tests/VBAi.Tests/VBAi.Tests.csproj -c Debug --no-build -p:BuildOutputRoot="$PWD/artifacts/build" --collect:"XPlat Code Coverage" --results-directory artifacts/coverage
```

State the source revision, loaded MVID/SHA, collector include/exclude rules,
executed tests, failures and skips. Keep line and branch metrics separate from
VBA procedure-entry measurements, native-host qualification and synthetic UI.
Native/provider opt-ins must be declared; do not enable every available flag.
Skips and failed aggregates remain visible. No empty tests or metric-driven
exclusions are accepted.

When a gate includes test-process lifetime, retain the VSTest manager and host
logs as well as the original runner's exit receipt. A successful runner exit and
passing TRX do not prove normal test-host termination. A forced-cleanup warning,
an unhandled host exception or an interrupted run leaves that lifecycle gate
unqualified. Process absence proves cleanup, not the manner of exit.

VSTest 18.7 reads `VSTEST_TESTHOST_SHUTDOWN_TIMEOUT` in its manager process; its
default budget is 100 ms. For an explicitly declared shutdown investigation,
set a bounded budget in the dedicated launcher before starting VSTest and retain
that setting with the candidate plan. Do not change machine-wide settings or
silently repeat a native mutation. See the
[VSTest shutdown implementation](https://github.com/microsoft/vstest/blob/v18.7.0/src/Microsoft.TestPlatform.CrossPlatEngine/Client/ProxyOperationManager.cs#L298).
Increasing the observation budget does not establish successful host shutdown;
the recorded normal-exit evidence is still required.

Reports, frozen plans, screenshots and binary manifests belong under `artifacts/`.
Only reviewed synthetic evidence should be published. Record test counts once in
[recorded validation](../docs/test-coverage.md), tied to the candidate; do not
copy old percentages into feature guides or the README.

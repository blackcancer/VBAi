# Tests VSTest

`CodexVBE.Tests` is an x64 .NET Framework 4.8 MSTest project in `CodexVBE.sln`. Test Explorer discovers the cases by category. Tests reference the production DLL; `InternalsVisibleTo` exposes internal code to this test assembly without copying production source.

Run unit tests by default from the repository root:

```powershell
dotnet test CodexVBE.sln --filter TestCategory=Unit
```

Run unit tests plus the local SQLite integration test with coverage:

```powershell
dotnet test tests/CodexVBE.Tests/CodexVBE.Tests.csproj --filter "TestCategory=Unit|TestCategory=LocalIntegration" --collect:"XPlat Code Coverage"
```

The latest complete measured run on 2026-09-27 passed **315 tests**, with **2 host tests skipped** and no failures. The `CodexVBE` package in Cobertura reports **73.68% of production lines** and **70.50% of production branches**. The isolated build also places the auxiliary `ProviderTests` executable in the test output; its separate package must not be included when reporting add-in coverage. No production files or methods were excluded. The two skipped host cases still need a separate live run.

The `Excel` category is opt-in runtime integration. Its smoke test creates a separate Excel process and temporary workbook, reads and saves the workbook, opens the VBE through Excel's native `CommandBars.ExecuteMso("VisualBasic")` command, and queries the add-in pipe `CodexVBE.<PID>`. It verifies the VBE environment, the exact temporary project path, and `CodexVBE.AddIn` with `Connect=true`. The test records existing Excel PIDs and closes only its own workbook and process.

```powershell
$env:CODEXVBE_RUN_EXCEL_TESTS = '1'
dotnet test tests/CodexVBE.Tests/CodexVBE.Tests.csproj --filter TestCategory=Excel
Remove-Item Env:CODEXVBE_RUN_EXCEL_TESTS
```

The actual opt-in run on 2026-09-27 passed **1/1 Excel test**. Direct external access to `excel.VBE` was separately attempted and refused by Excel's AccessVBOM trust setting; the passing test instead validates VBE, add-in, and project via the native command and the production bridge. It does not modify that user-level setting.

The standalone console harness executables in `tools/tests/Providers` and `tools/tests/Git` are separate from VSTest; selected source files are linked as VSTest integration cases. The PowerShell debug scripts are linked under **Manual/Debug** in the test project so they appear with the tests in Visual Studio. They remain opt-in host probes, not MSTest cases or part of the coverage run. The production project compiles only `src/**/*.cs`; no debug test class is shipped in the add-in. Excel is the first host for automated COM integration.

The `SolidWorks` category is strictly opt-in. It requires `CODEXVBE_SOLIDWORKS_PID` to name an **existing** `SLDWORKS` process whose VBE and CodexVBE add-in are already loaded. The test connects to that PID's bridge, verifies a VBE project and `CodexVBE.AddIn` with `Connect=true`, and does not start, close, or modify SOLIDWORKS. To run it after preloading the host:

```powershell
$env:CODEXVBE_SOLIDWORKS_PID = '<existing SLDWORKS PID>'
dotnet test tests/CodexVBE.Tests/CodexVBE.Tests.csproj --filter TestCategory=SolidWorks
Remove-Item Env:CODEXVBE_SOLIDWORKS_PID
```

With SOLIDWORKS closed and no PID supplied, VSTest reported **1 ignored** SolidWorks test on 2026-09-27. The runtime check is `NOT_RUN`. The existing SOLIDWORKS PowerShell probes remain separate and were not executed by this VSTest case.

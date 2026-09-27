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

On 2026-09-27, this measured run passed **16/16** tests: 15 Unit and one LocalIntegration. Cobertura reported **375/12,997 production lines (2.88%)** and **257/14,220 branches (1.80%)**. No production files or methods were excluded. The suite covers provider protocol and streaming behavior, history and export, code preview and rollback, and SQLite persistence. This is the full DLL denominator, not a claim of full product coverage.

The `Excel` category is opt-in runtime integration. Its smoke test creates a separate Excel process and temporary workbook, reads and saves the workbook, opens the VBE through Excel's native `CommandBars.ExecuteMso("VisualBasic")` command, and queries the add-in pipe `CodexVBE.<PID>`. It verifies the VBE environment, the exact temporary project path, and `CodexVBE.AddIn` with `Connect=true`. The test records existing Excel PIDs and closes only its own workbook and process.

```powershell
$env:CODEXVBE_RUN_EXCEL_TESTS = '1'
dotnet test tests/CodexVBE.Tests/CodexVBE.Tests.csproj --filter TestCategory=Excel
Remove-Item Env:CODEXVBE_RUN_EXCEL_TESTS
```

The actual opt-in run on 2026-09-27 passed **1/1 Excel test**. Direct external access to `excel.VBE` was separately attempted and refused by Excel's AccessVBOM trust setting; the passing test instead validates VBE, add-in, and project via the native command and the production bridge. It does not modify that user-level setting.

The console harnesses in `tools/tests/Providers` and `tools/tests/Git` are separate from VSTest and are not included in this coverage. PowerShell scripts under `tools/tests` are also separate; they are not VSTest cases. Excel is the first host for automated COM integration.

The `SolidWorks` category is strictly opt-in. It requires `CODEXVBE_SOLIDWORKS_PID` to name an **existing** `SLDWORKS` process whose VBE and CodexVBE add-in are already loaded. The test connects to that PID's bridge, verifies a VBE project and `CodexVBE.AddIn` with `Connect=true`, and does not start, close, or modify SOLIDWORKS. To run it after preloading the host:

```powershell
$env:CODEXVBE_SOLIDWORKS_PID = '<existing SLDWORKS PID>'
dotnet test tests/CodexVBE.Tests/CodexVBE.Tests.csproj --filter TestCategory=SolidWorks
Remove-Item Env:CODEXVBE_SOLIDWORKS_PID
```

With SOLIDWORKS closed and no PID supplied, VSTest reported **1 ignored** SolidWorks test on 2026-09-27. The runtime check is `NOT_RUN`. The existing SOLIDWORKS PowerShell probes remain separate and were not executed by this VSTest case.

using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Runtime.CompilerServices;
using System.Runtime.ExceptionServices;
using System.Web.Script.Serialization;
using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace VBAi.Tests.Integration
{
    /// <summary>Démarre et nettoie une instance Excel isolée pour les tests d’intégration du pont VBE.</summary>
    internal sealed partial class ExcelVbeFixture : IDisposable
    {
        /// <summary>Récupère l’identifiant du processus propriétaire d’une fenêtre Win32.</summary>
        /// <param name="window">Handle de la fenêtre à examiner.</param>
        /// <param name="processId">Reçoit l’identifiant du processus propriétaire.</param>
        /// <returns>L’identifiant du thread qui possède la fenêtre.</returns>
        [DllImport("user32.dll")]
        private static extern uint GetWindowThreadProcessId(IntPtr window, out uint processId);

        /// <summary>Instance COM Excel démarrée par la fixture.</summary>
        private object application;
        /// <summary>Indicates that temporary scenario roots have left the non-inlined execution frame.</summary>
        private bool scenarioFrameReturned;
        /// <summary>Collection COM des classeurs Excel.</summary>
        private object workbooks;
        /// <summary>Classeur temporaire créé pour isoler les commandes VBE.</summary>
        private object workbook;
        /// <summary>Indique si cette fixture a créé le processus et peut le fermer.</summary>
        private bool owned;
        // Retain a process handle before shutdown so even a fast crash remains observable.
        private Process ownedProcess;
        // This fixture owns and releases the optional embedded-project identity lease.
        private IntPtr embeddedGitProjectIdentity;

        /// <summary>Crée une fixture avant son initialisation par <see cref="Start"/>.</summary>
        private ExcelVbeFixture() { }

        /// <summary>Identifiant du processus Excel isolé.</summary>
        /// <value>Identifiant du processus créé pour le test.</value>
        internal int ProcessId { get; private set; }
        /// <summary>Répertoire temporaire réservé aux fichiers du test.</summary>
        /// <value>Chemin racine des fichiers temporaires de la fixture.</value>
        internal string Root { get; private set; }

        /// <summary>Retains exact owned-process shutdown observations after temporary-file cleanup.</summary>
        internal IDictionary<string, object> ShutdownDiagnostics { get; private set; }

        /// <summary>Suspends Close/Quit when a diagnostic debugger has not been proven detached.</summary>
        internal bool PreserveForDiagnosticRecovery { get; set; }

        /// <summary>Démarre Excel de façon isolée et vérifie la disponibilité du pont VBE.</summary>
        /// <returns>La fixture prête à envoyer des commandes au pont.</returns>
        /// <exception cref="AssertInconclusiveException">Les tests Excel sont désactivés, Excel est absent ou une session existante a été détectée.</exception>
        internal static ExcelVbeFixture Start()
        {
            if (Environment.GetEnvironmentVariable("VBAi_RUN_EXCEL_TESTS") != "1")
                Assert.Inconclusive("Excel automation is opt-in. Set VBAi_RUN_EXCEL_TESTS=1.");
            string desktop = Environment.GetEnvironmentVariable("VBAi_TEST_DESKTOP_NAME");
            if (!string.IsNullOrWhiteSpace(desktop))
            {
                IsolatedTestDesktop.RequireCurrent(desktop);
                var privateFixture = StartOwnedWithTrace(ExcelOwnedBootstrapPlan.RequireLocalAbsolutePath(
                    Environment.GetEnvironmentVariable(VbeInspectionTrace.EnvironmentName)));
                // Generic scenarios require the same unsaved-workbook precondition as COM
                // activation. Retire only the verified macro-free seed, after loaded MVID checks.
                privateFixture.PreserveForDiagnosticRecovery = true;
                try
                {
                    ((dynamic)privateFixture.workbook).Close(false);
                    Release(privateFixture.workbook); privateFixture.workbook = null;
                    privateFixture.workbook = ((dynamic)privateFixture.workbooks).Add();
                    Assert.AreEqual(1, Convert.ToInt32(((dynamic)privateFixture.workbooks).Count));
                    Assert.IsTrue(string.IsNullOrEmpty(Convert.ToString(((dynamic)privateFixture.workbook).Path)));
                    privateFixture.WriteEvidence("private-unsaved-workbook.json", new {
                        Desktop = desktop, privateFixture.ProcessId, Workbook = Convert.ToString(((dynamic)privateFixture.workbook).Name),
                        SavedPath = Convert.ToString(((dynamic)privateFixture.workbook).Path), HelperSaveInvoked = false,
                        SeedClosedWithoutSaving = true, Utc = DateTime.UtcNow.ToString("o") });
                    privateFixture.PreserveForDiagnosticRecovery = false;
                    return privateFixture;
                }
                catch
                {
                    lock (retainedBootstraps) retainedBootstraps.Add(privateFixture);
                    throw; // Unknown Close/Add outcomes never authorize replay or cleanup.
                }
            }
            var excelType = Type.GetTypeFromProgID("Excel.Application");
            if (excelType == null) Assert.Inconclusive("Excel.Application is unavailable.");
            var existing = Process.GetProcessesByName("EXCEL");
            int[] existingIds = existing.Select(process => process.Id).ToArray();
            foreach (var process in existing) process.Dispose();
            string output = Environment.GetEnvironmentVariable("VBAi_EXCEL_RESULTS");
            bool durable = !string.IsNullOrWhiteSpace(output);
            if (durable && (!Path.IsPathRooted(output) || !string.Equals(Path.GetPathRoot(output), Path.GetPathRoot(Path.GetFullPath(output)), StringComparison.OrdinalIgnoreCase)))
                throw new ArgumentException("VBAi_EXCEL_RESULTS must specify an absolute output directory.");
            var fixture = new ExcelVbeFixture {
                retainEvidence = durable,
                Root = Path.Combine(durable ? Path.GetFullPath(output) : Path.Combine(Path.GetTempPath(), "VBAi-VSTest"), Guid.NewGuid().ToString("N")) };
            try
            {
                Directory.CreateDirectory(fixture.Root);
                fixture.RecordStartup("Created", existingIds);
                fixture.application = Activator.CreateInstance(excelType);
                dynamic excel = fixture.application;
                uint processId;
                GetWindowThreadProcessId(new IntPtr(Convert.ToInt64(excel.Hwnd)), out processId);
                fixture.ProcessId = (int)processId;
                fixture.owned = processId != 0 && !existingIds.Contains((int)processId);
                fixture.RecordStartup("ProcessIdentified", existingIds);
                if (!fixture.owned)
                    Assert.Inconclusive("Excel returned an existing session; no workbook was opened.");
                fixture.ownedProcess = Process.GetProcessById(fixture.ProcessId);
                IntPtr retainedHandle = fixture.ownedProcess.Handle;
                uint windowThread = GetWindowThreadProcessId(new IntPtr(Convert.ToInt64(excel.Hwnd)), out uint observedPid);
                string ownerDesktop = IsolatedTestDesktop.DesktopName(IsolatedTestDesktop.GetCurrentThreadId());
                string windowDesktop = IsolatedTestDesktop.DesktopName(windowThread);
                RecordComAttachedIdentity(fixture.startupEvidence, fixture.ProcessId, (int)observedPid,
                    retainedHandle.ToInt64(), fixture.ownedProcess.StartTime.ToUniversalTime().ToString("o"),
                    windowThread, ownerDesktop, windowDesktop, IsolatedTestDesktop.InputDesktopName());
                fixture.RecordStartup("OwnedComIdentityObserved", existingIds);
                excel.Visible = true;
                excel.DisplayAlerts = false;
                fixture.workbooks = excel.Workbooks;
                fixture.workbook = ((dynamic)fixture.workbooks).Add();
                fixture.RecordStartup("WorkbookCreated", existingIds);
                object commandBars = null;
                try { commandBars = excel.CommandBars; ((dynamic)commandBars).ExecuteMso("VisualBasic"); }
                finally { Release(commandBars); }
                var status = fixture.Command("status");
                fixture.RecordStartup("BridgeObserved", existingIds, status);
                Assert.IsNotNull(status, "The isolated Excel VBE has no VBAi bridge.");
                Assert.AreEqual(true, status["Ok"]);
                var state = VbeBridgeClient.Object(status["Data"]);
                Assert.AreEqual(typeof(VBAi.VbeSession).Module.ModuleVersionId.ToString("D"), state["AssemblyModuleVersionId"],
                    "Excel loaded another build of the shared COM add-in; native tests must qualify this assembly.");
                fixture.RecordStartup("Ready", existingIds, status);
                return fixture;
            }
            catch (Exception startup)
            {
                var failures = new List<Exception> { startup };
                try { fixture.RecordStartup("Failed", existingIds, error: startup); }
                catch (Exception evidence) { failures.Add(evidence); }
                try { fixture.Dispose(); }
                catch (Exception cleanup) { failures.Add(cleanup); }
                if (failures.Count > 1) throw new AggregateException("Excel startup, evidence or cleanup failed; all errors are retained.", failures);
                throw;
            }
        }

        /// <summary>Preserves the scenario failure when cleanup independently fails.</summary>
        internal static void Run(Action<ExcelVbeFixture> scenario, Action<ExcelVbeFixture> shutdownVerified = null)
        {
            if (scenario == null) throw new ArgumentNullException(nameof(scenario));
            RunPreparedScenario(Start(), scenario, shutdownVerified);
        }

        /// <summary>Runs an already identified owned fixture; mirrors supply only managed fake application objects.</summary>
        internal static void RunPreparedScenario(ExcelVbeFixture fixture, Action<ExcelVbeFixture> scenario,
            Action<ExcelVbeFixture> shutdownVerified = null)
        {
            if (fixture == null) throw new ArgumentNullException(nameof(fixture));
            if (scenario == null) throw new ArgumentNullException(nameof(scenario));
            Exception failure = null;
            failure = ExecuteScenario(scenario, fixture);
            fixture.scenarioFrameReturned = true;
            try { fixture.Dispose(); shutdownVerified?.Invoke(fixture); }
            catch (Exception cleanup)
            {
                if (failure != null) throw new AggregateException("The Excel scenario and its shutdown both failed; both errors are retained.", failure, cleanup);
                throw;
            }
            if (failure != null) ExceptionDispatchInfo.Capture(failure).Throw();
        }

        /// <summary>Returns after the scenario frame has released its temporary dynamic COM roots.</summary>
        [MethodImpl(MethodImplOptions.NoInlining)]
        private static Exception ExecuteScenario(Action<ExcelVbeFixture> scenario, ExcelVbeFixture fixture)
        {
            try { scenario(fixture); return null; }
            catch (Exception error) { return error; }
        }

        /// <summary>Envoie une commande nommée au pont du processus Excel.</summary>
        /// <param name="name">Nom de la commande à exécuter.</param>
        /// <returns>La réponse reçue, ou <see langword="null"/> si le canal ne répond pas.</returns>
        internal IDictionary<string, object> Command(string name)
        {
            return RecordCommand(new { Command = name }, () => VbeBridgeClient.Read(ProcessId, name));
        }

        /// <summary>Envoie un objet de requête au pont du processus Excel.</summary>
        /// <param name="request">Requête sérialisable à envoyer.</param>
        /// <returns>La réponse reçue, ou <see langword="null"/> si le canal ne répond pas.</returns>
        internal IDictionary<string, object> Command(object request)
        {
            return RecordCommand(request, () => VbeBridgeClient.Read(ProcessId, request));
        }

        /// <summary>Construit un chemin dans le répertoire temporaire de la fixture.</summary>
        /// <param name="name">Nom ou chemin relatif du fichier.</param>
        /// <returns>Chemin absolu du fichier dans le répertoire temporaire.</returns>
        internal string File(string name) { return Path.Combine(Root, name); }

        /// <summary>Relit une cellule du classeur possédé, indépendamment des réponses du pont VBE.</summary>
        /// <param name="address">Adresse A1 de la cellule attendue.</param>
        /// <returns>Valeur native Excel après l'exécution de la macro de test.</returns>
        internal object ReadCell(string address)
        {
            object sheets = null, sheet = null, cell = null;
            try
            {
                sheets = ((dynamic)workbook).Worksheets;
                sheet = ((dynamic)sheets)[1];
                cell = ((dynamic)sheet).Range[address];
                return ((dynamic)cell).Value2;
            }
            finally { Release(cell); Release(sheet); Release(sheets); }
        }

        /// <summary>Ferme les ressources COM et fichiers temporaires appartenant à cette fixture.</summary>
        public void Dispose()
        {
            if (!ClaimShutdown()) return;
            try { DisposeOwnedHostCore(); CompleteShutdown(); }
            catch (Exception error) {
                Exception combined = CombineShutdownFailure(error);
                RetainFailedShutdown(combined);
                if (ReferenceEquals(combined, error)) throw;
                throw combined;
            }
        }

        /// <summary>Runs the only permitted cleanup sequence on the original owner thread.</summary>
        private void DisposeOwnedHostCore()
        {
            var diagnostics = new Dictionary<string, object> {
                ["ProcessId"] = ProcessId, ["FixtureRoot"] = Root, ["StartedUtc"] = DateTime.UtcNow.ToString("o"),
                ["AssemblyMvid"] = typeof(VbeSession).Module.ModuleVersionId.ToString("D"), ["ForcedTermination"] = false
            };
            CopyComAttachedIdentityToShutdown(startupEvidence, diagnostics);
            diagnostics["DurableEvidence"] = retainEvidence;
            diagnostics["CommandCount"] = commandSequence;
            diagnostics["CommandRecordsOmitted"] = Math.Max(0, commandSequence - MaximumCommandRecords);
            ShutdownDiagnostics = diagnostics;
            RecordShutdownThread(diagnostics, "CleanupStarted");
            if (owned && (ownedProcess == null || ProcessId <= 0 || ownedProcess.Id != ProcessId))
                throw new InvalidOperationException("No exact retained original Excel process is available; no native cleanup is authorized.");
            if (PreserveForDiagnosticRecovery)
            {
                diagnostics["CleanupSuspended"] = true;
                diagnostics["Reason"] = "Diagnostic debugger detachment was not proven; no Close/Quit or forced termination was attempted.";
                WriteShutdownDiagnostics(diagnostics);
                throw new InvalidOperationException("Exact owned Excel host left untouched for diagnostic recovery. PID: " + ProcessId + "; fixture: " + Root);
            }
            var watch = Stopwatch.StartNew();
            Exception closeFailure = null, quitFailure = null, evidenceFailure = null;
            Action writeDiagnostics = () => {
                try { WriteShutdownDiagnostics(diagnostics); }
                catch (Exception error) { if (evidenceFailure == null) evidenceFailure = error; RememberShutdownFailure(error); }
            };
            BeginDiagnosticCleanup();
            if (embeddedGitProjectIdentity != IntPtr.Zero)
            {
                Marshal.Release(embeddedGitProjectIdentity);
                embeddedGitProjectIdentity = IntPtr.Zero;
            }
            if (scenarioFrameReturned)
            {
                // Drain only after the non-inlined scenario frame returns, before
                // the one Close/Quit. Preserve the original 10-second exit oracle.
                RecordShutdownThread(diagnostics, "ScenarioReturned");
                diagnostics["ScenarioRcwDrainStarted"] = true;
                diagnostics["ScenarioRcwDrainCompleted"] = false;
                writeDiagnostics();
                var drain = Stopwatch.StartNew();
                CollectScenarioReferences();
                diagnostics["ScenarioRcwDrainCompleted"] = true;
                diagnostics["ScenarioRcwDrainElapsedMs"] = drain.ElapsedMilliseconds;
                writeDiagnostics();
            }
            if (owned && workbook != null)
                try { ((dynamic)workbook).Close(false); }
                catch (Exception error) { closeFailure = error; RememberShutdownFailure(error); }
            diagnostics["CloseElapsedMs"] = watch.ElapsedMilliseconds;
            if (owned && application != null)
                try { ((dynamic)application).Quit(); }
                catch (Exception error) { quitFailure = error; RememberShutdownFailure(error); }
            diagnostics["QuitElapsedMs"] = watch.ElapsedMilliseconds;
            diagnostics["CloseError"] = closeFailure?.ToString();
            diagnostics["QuitError"] = quitFailure?.ToString();
            Release(workbook);
            Release(workbooks);
            Release(application);
            workbook = workbooks = application = null;
            diagnostics["ReleaseElapsedMs"] = watch.ElapsedMilliseconds;
            if (teardownTrace != null)
            {
                try { teardownTrace.AfterCleanup(); }
                catch (Exception traceFailure)
                {
                    diagnostics["DiagnosticFailure"] = traceFailure.ToString();
                    writeDiagnostics();
                    lock (retainedBootstraps) retainedBootstraps.Add(this);
                    throw new AggregateException("Diagnostic cleanup failed; native cleanup is never replayed.",
                        new[] { traceFailure, closeFailure, quitFailure, evidenceFailure }.Where(error => error != null));
                }
            }
            var process = ownedProcess;
            if (process != null)
                {
                    RecordShutdownThread(diagnostics, "BeforeExitWait");
                    cleanupStage = "EXIT_WAIT"; exitWaitAttempted = true;
                    bool exited = WaitForOwnedExcelExit(process, 10000);
                    exitWaitReturned = true;
                    RecordShutdownThread(diagnostics, "AfterExitWait");
                    diagnostics["Exited"] = exited;
                    diagnostics["ElapsedMs"] = watch.ElapsedMilliseconds;
                    if (exited)
                    {
                        cleanupStage = "EXIT_CODE";
                        int code = ReadOwnedExcelExitCode(process);
                        exitCodeObserved = true;
                        diagnostics["ExitCode"] = code;
                        diagnostics["ExitCodeHex"] = "0x" + unchecked((uint)code).ToString("X8");
                    }
                    ObserveAddInShutdownTrace(diagnostics);
                    writeDiagnostics();
                    try
                    {
                        if (!exited)
                            Assert.Fail("Excel did not exit after Quit and COM release. PID: " + ProcessId + "; fixture: " + Root + ". The process was left running for diagnosis; shutdown.json preserves each phase.");
                        int code = (int)diagnostics["ExitCode"];
                        Assert.AreEqual(0, code, "Excel exited abnormally. PID: " + ProcessId +
                            "; exit code: 0x" + unchecked((uint)code).ToString("X8") + "; fixture: " + Root);
                        if (privateDesktopChild != null)
                        {
                            Assert.IsTrue(privateDesktopChild.Wait(0), "The original native Excel handle must independently observe exit.");
                            Assert.AreEqual(0u, privateDesktopChild.ExitCode());
                            privateDesktopChild.Dispose(); privateDesktopChild = null;
                        }
                    }
                    catch (Exception exitFailure)
                    {
                        if (evidenceFailure != null)
                            throw new AggregateException("Excel shutdown and its evidence write failed; both errors are retained.",
                                new[] { exitFailure, closeFailure, quitFailure, evidenceFailure }.Where(error => error != null));
                        throw;
                    }
                    if (evidenceFailure != null) throw evidenceFailure;
                    if (addInShutdownFailure != null) throw addInShutdownFailure;
                    cleanupStage = "PROCESS_RELEASE";
                    processReleaseEntered = true;
                    ReleaseOwnedExcelProcess(process);
                    processReleaseReturned = true;
                    ownedProcess = null;
                    diagnostics["ProcessHandleRetained"] = false;
                    writeDiagnostics();
                }
            else writeDiagnostics();
            if (privateDesktopChild != null)
            {
                Assert.IsTrue(privateDesktopChild.Wait(0), "The original private launch handle must observe exit.");
                Assert.AreEqual(0u, privateDesktopChild.ExitCode(), "The original private launch exited abnormally.");
                privateDesktopChild.Dispose();
                privateDesktopChild = null;
            }
            if (closeFailure != null || quitFailure != null || evidenceFailure != null || addInShutdownFailure != null)
                throw new AggregateException("Excel Close/Quit or cleanup observation reported errors; shutdown.json preserves diagnostics.",
                    new[] { closeFailure, quitFailure, evidenceFailure, addInShutdownFailure }.Where(error => error != null));
            if (retainEvidence) return;
            if (string.IsNullOrWhiteSpace(Root) || !Directory.Exists(Root)) return;
            if (retainEvidence) { WriteEvidence("shutdown.json", diagnostics); return; }
            try
            {
                foreach (string file in Directory.GetFiles(Root, "*", SearchOption.TopDirectoryOnly))
                    System.IO.File.Delete(file);
                Directory.Delete(Root, false);
            }
            catch (IOException) { }
            catch (UnauthorizedAccessException) { }
        }

        private void WriteShutdownDiagnostics(IDictionary<string, object> diagnostics)
        {
            if (WriteShutdownReceipt != null) { WriteShutdownReceipt(diagnostics); return; }
            if (string.IsNullOrWhiteSpace(Root) || !Directory.Exists(Root)) return;
            try { System.IO.File.WriteAllText(Path.Combine(Root, "shutdown.json"), new JavaScriptSerializer().Serialize(diagnostics)); }
            catch (IOException) { }
            catch (UnauthorizedAccessException) { }
        }

        /// <summary>Libère une référence COM sans propager une erreur de nettoyage.</summary>
        /// <param name="value">Objet COM à libérer, ou <see langword="null"/>.</param>
        private static void Release(object value)
        {
            try { if (value != null && Marshal.IsComObject(value)) Marshal.FinalReleaseComObject(value); }
            catch (COMException) { }
        }
    }
}

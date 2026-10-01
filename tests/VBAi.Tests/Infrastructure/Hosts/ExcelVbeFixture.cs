using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
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
        /// <summary>Collection COM des classeurs Excel.</summary>
        private object workbooks;
        /// <summary>Classeur temporaire créé pour isoler les commandes VBE.</summary>
        private object workbook;
        /// <summary>Indique si cette fixture a créé le processus et peut le fermer.</summary>
        private bool owned;
        // Retain a process handle before shutdown so even a fast crash remains observable.
        private Process ownedProcess;

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
                _ = fixture.ownedProcess.Handle;
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
            var fixture = Start();
            Exception failure = null;
            try { scenario(fixture); }
            catch (Exception error) { failure = error; }
            try { fixture.Dispose(); shutdownVerified?.Invoke(fixture); }
            catch (Exception cleanup)
            {
                if (failure != null) throw new AggregateException("The Excel scenario and its shutdown both failed; both errors are retained.", failure, cleanup);
                throw;
            }
            if (failure != null) ExceptionDispatchInfo.Capture(failure).Throw();
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
            var diagnostics = new Dictionary<string, object> {
                ["ProcessId"] = ProcessId, ["FixtureRoot"] = Root, ["StartedUtc"] = DateTime.UtcNow.ToString("o"),
                ["AssemblyMvid"] = typeof(VbeSession).Module.ModuleVersionId.ToString("D"), ["ForcedTermination"] = false
            };
            diagnostics["DurableEvidence"] = retainEvidence;
            diagnostics["CommandCount"] = commandSequence;
            diagnostics["CommandRecordsOmitted"] = Math.Max(0, commandSequence - MaximumCommandRecords);
            ShutdownDiagnostics = diagnostics;
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
                catch (Exception error) { evidenceFailure = error; }
            };
            BeginDiagnosticCleanup();
            if (embeddedGitProjectIdentity != IntPtr.Zero)
            {
                Marshal.Release(embeddedGitProjectIdentity);
                embeddedGitProjectIdentity = IntPtr.Zero;
            }
            if (owned && workbook != null)
                try { ((dynamic)workbook).Close(false); }
                catch (Exception error) { closeFailure = error; }
            diagnostics["CloseElapsedMs"] = watch.ElapsedMilliseconds;
            if (owned && application != null)
                try { ((dynamic)application).Quit(); }
                catch (Exception error) { quitFailure = error; }
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
            ownedProcess = null;
            if (process != null)
                using (process)
                {
                    bool exited = process.WaitForExit(10000);
                    diagnostics["Exited"] = exited;
                    diagnostics["ElapsedMs"] = watch.ElapsedMilliseconds;
                    if (exited)
                    {
                        diagnostics["ExitCode"] = process.ExitCode;
                        diagnostics["ExitCodeHex"] = "0x" + unchecked((uint)process.ExitCode).ToString("X8");
                    }
                    writeDiagnostics();
                    try
                    {
                        if (!exited)
                            Assert.Fail("Excel did not exit after Quit and COM release. PID: " + ProcessId + "; fixture: " + Root + ". The process was left running for diagnosis; shutdown.json preserves each phase.");
                        Assert.AreEqual(0, process.ExitCode, "Excel exited abnormally. PID: " + ProcessId +
                            "; exit code: 0x" + unchecked((uint)process.ExitCode).ToString("X8") + "; fixture: " + Root);
                    }
                    catch (Exception exitFailure)
                    {
                        if (evidenceFailure != null)
                            throw new AggregateException("Excel shutdown and its evidence write failed; both errors are retained.",
                                new[] { exitFailure, closeFailure, quitFailure, evidenceFailure }.Where(error => error != null));
                        throw;
                    }
                }
            else writeDiagnostics();
            if (closeFailure != null || quitFailure != null || evidenceFailure != null)
                throw new AggregateException("Excel Close/Quit reported errors; shutdown.json preserves diagnostics.",
                    new[] { closeFailure, quitFailure, evidenceFailure }.Where(error => error != null));
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

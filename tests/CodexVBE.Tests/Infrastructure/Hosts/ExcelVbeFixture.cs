using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace CodexVBE.Tests.Integration
{
    /// <summary>Démarre et nettoie une instance Excel isolée pour les tests d’intégration du pont VBE.</summary>
    internal sealed class ExcelVbeFixture : IDisposable
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

        /// <summary>Crée une fixture avant son initialisation par <see cref="Start"/>.</summary>
        private ExcelVbeFixture() { }

        /// <summary>Identifiant du processus Excel isolé.</summary>
        /// <value>Identifiant du processus créé pour le test.</value>
        internal int ProcessId { get; private set; }
        /// <summary>Répertoire temporaire réservé aux fichiers du test.</summary>
        /// <value>Chemin racine des fichiers temporaires de la fixture.</value>
        internal string Root { get; private set; }

        /// <summary>Démarre Excel de façon isolée et vérifie la disponibilité du pont VBE.</summary>
        /// <returns>La fixture prête à envoyer des commandes au pont.</returns>
        /// <exception cref="AssertInconclusiveException">Les tests Excel sont désactivés, Excel est absent ou une session existante a été détectée.</exception>
        internal static ExcelVbeFixture Start()
        {
            if (Environment.GetEnvironmentVariable("CODEXVBE_RUN_EXCEL_TESTS") != "1")
                Assert.Inconclusive("Excel automation is opt-in. Set CODEXVBE_RUN_EXCEL_TESTS=1.");
            var excelType = Type.GetTypeFromProgID("Excel.Application");
            if (excelType == null) Assert.Inconclusive("Excel.Application is unavailable.");
            var existing = Process.GetProcessesByName("EXCEL");
            int[] existingIds = existing.Select(process => process.Id).ToArray();
            foreach (var process in existing) process.Dispose();
            var fixture = new ExcelVbeFixture {
                Root = Path.Combine(Path.GetTempPath(), "CodexVBE-VSTest", Guid.NewGuid().ToString("N")) };
            try
            {
                Directory.CreateDirectory(fixture.Root);
                fixture.application = Activator.CreateInstance(excelType);
                dynamic excel = fixture.application;
                uint processId;
                GetWindowThreadProcessId(new IntPtr(Convert.ToInt64(excel.Hwnd)), out processId);
                fixture.ProcessId = (int)processId;
                fixture.owned = processId != 0 && !existingIds.Contains((int)processId);
                if (!fixture.owned)
                    Assert.Inconclusive("Excel returned an existing session; no workbook was opened.");
                excel.Visible = true;
                excel.DisplayAlerts = false;
                fixture.workbooks = excel.Workbooks;
                fixture.workbook = ((dynamic)fixture.workbooks).Add();
                excel.CommandBars.ExecuteMso("VisualBasic");
                var status = fixture.Command("status");
                Assert.IsNotNull(status, "The isolated Excel VBE has no CodexVBE bridge.");
                Assert.AreEqual(true, status["Ok"]);
                return fixture;
            }
            catch
            {
                fixture.Dispose();
                throw;
            }
        }

        /// <summary>Envoie une commande nommée au pont du processus Excel.</summary>
        /// <param name="name">Nom de la commande à exécuter.</param>
        /// <returns>La réponse reçue, ou <see langword="null"/> si le canal ne répond pas.</returns>
        internal IDictionary<string, object> Command(string name)
        {
            return VbeBridgeClient.Read(ProcessId, name);
        }

        /// <summary>Envoie un objet de requête au pont du processus Excel.</summary>
        /// <param name="request">Requête sérialisable à envoyer.</param>
        /// <returns>La réponse reçue, ou <see langword="null"/> si le canal ne répond pas.</returns>
        internal IDictionary<string, object> Command(object request)
        {
            return VbeBridgeClient.Read(ProcessId, request);
        }

        /// <summary>Construit un chemin dans le répertoire temporaire de la fixture.</summary>
        /// <param name="name">Nom ou chemin relatif du fichier.</param>
        /// <returns>Chemin absolu du fichier dans le répertoire temporaire.</returns>
        internal string File(string name) { return Path.Combine(Root, name); }

        /// <summary>Ferme les ressources COM et fichiers temporaires appartenant à cette fixture.</summary>
        public void Dispose()
        {
            if (owned && workbook != null)
                try { ((dynamic)workbook).Close(false); } catch { }
            if (owned && application != null)
                try { ((dynamic)application).Quit(); } catch { }
            Release(workbook);
            Release(workbooks);
            Release(application);
            workbook = workbooks = application = null;
            if (owned && ProcessId != 0)
            {
                try
                {
                    using (var process = Process.GetProcessById(ProcessId))
                        if (!process.WaitForExit(10000)) { process.Kill(); process.WaitForExit(10000); }
                }
                catch (ArgumentException) { }
            }
            if (string.IsNullOrWhiteSpace(Root) || !Directory.Exists(Root)) return;
            try
            {
                foreach (string file in Directory.GetFiles(Root, "*", SearchOption.TopDirectoryOnly))
                    System.IO.File.Delete(file);
                Directory.Delete(Root, false);
            }
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

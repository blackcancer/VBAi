using System;
using System.Diagnostics;
using System.IO;
using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace VBAi.Tests.Scenarios.Editor
{
    /// <summary>Exécute la matrice des fournisseurs JavaScript réels dans le projet VSTest.</summary>
    [TestClass, TestCategory("MonacoLanguage")]
    public sealed class MonacoLanguageScriptScenarios
    {
        /// <summary>Vérifie les complétions, survols, blocs et formats sans simulation de leur implémentation.</summary>
        [TestMethod]
        public void LanguageAndEditingProvidersPassTheirCompleteScenarioMatrix()
        {
            var directory = new DirectoryInfo(AppDomain.CurrentDomain.BaseDirectory);
            while (directory != null && !File.Exists(Path.Combine(directory.FullName, "assets", "editor", "src", "language.js"))) directory = directory.Parent;
            Assert.IsNotNull(directory, "Repository source assets are required for the JavaScript provider tests.");
            string root = directory.FullName;
            using (var process = Process.Start(new ProcessStartInfo("node", "--test \"" + Path.Combine(root, "tools", "tests", "Test-MonacoLanguage.mjs") + "\" \"" + Path.Combine(root, "tests", "VBAi.Tests", "Infrastructure", "Fixtures", "Editor", "MonacoEditing.Scenarios.mjs") + "\"")
            { UseShellExecute = false, CreateNoWindow = true, RedirectStandardOutput = true, RedirectStandardError = true, WorkingDirectory = root }))
            {
                var output = process.StandardOutput.ReadToEndAsync(); var error = process.StandardError.ReadToEndAsync();
                bool finished = process.WaitForExit(120000);
                if (!finished) process.Kill();
                Assert.IsTrue(finished, "JavaScript scenario matrix exceeded its time limit.");
                Assert.AreEqual(0, process.ExitCode, output.GetAwaiter().GetResult() + error.GetAwaiter().GetResult());
            }
        }
    }
}

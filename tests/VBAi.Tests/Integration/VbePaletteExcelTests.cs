using Microsoft.VisualStudio.TestTools.UnitTesting;
using System;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;

namespace VBAi.Tests.Integration
{
    /// <summary>Qualifies native palette conflict recovery through the real Excel Options dialog.</summary>
    [TestClass, TestCategory("NativePaletteExcel")]
    public sealed class VbePaletteExcelTests
    {
        /// <summary>Applies a theme over manual colors, archives recovery and restores the complete native baseline.</summary>
        [STATestMethod]
        public void ConflictingPaletteReconcilesArchivesAndRestoresThroughNativeExcelOptions()
        {
            ExcelScenarioLifetime.Run(RunScenario);
        }

        [System.Runtime.CompilerServices.MethodImpl(System.Runtime.CompilerServices.MethodImplOptions.NoInlining)]
        private static void RunScenario(ExcelScenarioLifetime lifetime)
        {
            if (Environment.GetEnvironmentVariable("VBAI_NATIVE_PALETTE_EXCEL_TEST") != "1") Assert.Inconclusive("Explicit disposable Excel palette test opt-in required.");
            if (Process.GetProcessesByName("EXCEL").Length != 0 || Process.GetProcessesByName("SLDWORKS").Length != 0)
                Assert.Inconclusive("Existing hosts must be preserved because native colors are shared preferences.");
            string root = Path.Combine(Path.GetTempPath(), "VBAi-native-palette-" + Guid.NewGuid().ToString("N"));
            Directory.CreateDirectory(root);
            string experiment = Environment.GetEnvironmentVariable("VBAi_NATIVE_DARK_EXPERIMENT");
            dynamic excel = null, book = null;
            VbeNativePaletteState.ColorRow[] before = null;
            bool restored = false;
            try
            {
                // The disposable host must not consume the production recovery snapshot.
                Environment.SetEnvironmentVariable("VBAi_NATIVE_DARK_EXPERIMENT", "1");
                excel = Activator.CreateInstance(Type.GetTypeFromProgID("Excel.Application"));
                lifetime.Capture((object)excel);
                excel.Visible = true; excel.DisplayAlerts = false; excel.EnableEvents = false;
                book = excel.Workbooks.Add(); excel.VBE.MainWindow.Visible = true;
                object vbe = (object)excel.VBE;
                before = VbeNativePaletteDialog.Visit(vbe, rows => null);
                var original = before.Select(row => new VbeNativePaletteState.ColorRow { Name = row.Name, Foreground = row.Foreground, Background = row.Background, Indicator = row.Indicator }).ToArray();
                var dark = VbeNativePaletteState.Dark(original);
                int changed = Array.FindIndex(before, row => row.Foreground != dark[Array.IndexOf(before, row)].Foreground);
                Assert.IsTrue(changed >= 0, "A native foreground must differ from the configured dark foreground.");
                original[changed].Foreground = Enumerable.Range(0, 17).First(value => value != before[changed].Foreground && value != dark[changed].Foreground);
                var previous = new VbeNativePaletteState { VbeVersion = (string)excel.VBE.Version, Original = original, Applied = VbeNativePaletteState.Dark(original) };
                string path = Path.Combine(root, "palette.json"); previous.SaveNew(path);
                string previousBytes = File.ReadAllText(path);
                VbeNativePalette.Change(vbe, true, path);
                var rebased = VbeNativePaletteState.Load(path, previous.VbeVersion);
                Assert.IsTrue(VbeNativePaletteState.Equal(before, rebased.Original), "Manual native colors must become the restored baseline.");
                Assert.AreEqual(previousBytes, File.ReadAllText(Directory.GetFiles(root, "*.previous-*").Single()));
                Assert.IsTrue(VbeNativePaletteState.Equal(rebased.Applied, VbeNativePaletteDialog.Visit(vbe, rows => null)));
                VbeNativePalette.Change(vbe, false, path);
                Assert.IsFalse(File.Exists(path));
                Assert.IsTrue(VbeNativePaletteState.Equal(before, VbeNativePaletteDialog.Visit(vbe, rows => null)));
                restored = true;
            }
            finally
            {
                try
                {
                    if (!restored && before != null && (object)excel != null)
                        VbeNativePaletteDialog.Visit((object)excel.VBE, rows => before);
                }
                finally
                {
                    if ((object)book != null) { book.Close(false); Marshal.FinalReleaseComObject((object)book); }
                    if ((object)excel != null)
                    {
                        if (lifetime.OwnsApplication) excel.Quit();
                        Marshal.FinalReleaseComObject((object)excel);
                    }
                    Environment.SetEnvironmentVariable("VBAi_NATIVE_DARK_EXPERIMENT", experiment);
                }
            }
        }
    }
}

using System;
using System.IO;
using System.Linq;
using CodexVBE;
using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace CodexVBE.Tests.Unit
{
    [TestClass, TestCategory("Unit")]
    public sealed class VbeNativePaletteTests
    {
        private static VbeNativePaletteState State()
        {
            var original = Enumerable.Range(0, 10).Select(index => new VbeNativePaletteState.ColorRow
            {
                Name = "Category " + index, Foreground = index, Background = 16 - index, Indicator = index + 1
            }).ToArray();
            return new VbeNativePaletteState { VbeVersion = "7.1", Original = original, Applied = VbeNativePaletteState.Dark(original) };
        }

        [TestMethod]
        public void ThemeRetainsIndicatorsAndRefusesManualColorChanges()
        {
            var state = State();
            CollectionAssert.AreEqual(state.Original.Select(row => row.Indicator).ToArray(), state.Applied.Select(row => row.Indicator).ToArray());
            state.RequireUnchanged(state.Original);
            state.RequireUnchanged(state.Applied);
            var customized = VbeNativePaletteState.Dark(state.Original);
            customized[5].Foreground = 12;
            Assert.ThrowsException<InvalidOperationException>(() => state.RequireUnchanged(customized));
            Assert.AreEqual(12, customized[5].Foreground);
            Assert.AreEqual(5, state.Original[5].Foreground);
        }

        [TestMethod]
        public void RecoveryRoundTripPreservesOriginalAndCannotOverwriteIt()
        {
            string directory = Path.Combine(Path.GetTempPath(), "CodexVBE-palette-" + Guid.NewGuid().ToString("N"));
            string path = Path.Combine(directory, "palette.json");
            try
            {
                var state = State();
                state.SaveNew(path);
                Assert.IsTrue(VbeNativePaletteState.Equal(state.Original, VbeNativePaletteState.Load(path, "7.1").Original));
                string contents = File.ReadAllText(path);
                Assert.ThrowsException<IOException>(() => State().SaveNew(path));
                Assert.AreEqual(contents, File.ReadAllText(path));
                Assert.ThrowsException<InvalidDataException>(() => VbeNativePaletteState.Load(path, "6.0"));
            }
            finally { if (Directory.Exists(directory)) Directory.Delete(directory, true); }
        }

        [TestMethod]
        public void InvalidRecoveryNeverSuppliesRestoreValues()
        {
            var state = State();
            state.Applied[0].Background = 9;
            Assert.ThrowsException<InvalidDataException>(() => state.Validate("7.1"));
            state = State();
            state.Original[0].Indicator = 17;
            Assert.ThrowsException<InvalidDataException>(() => state.Validate("7.1"));
        }
    }
}

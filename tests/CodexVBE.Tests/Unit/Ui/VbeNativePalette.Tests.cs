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

namespace CodexVBE.Tests.Unit
{
    [TestClass, TestCategory("Unit")]
    public sealed class VbeNativePaletteTransactionTests
    {
        [TestMethod]
        public void RestoreWithoutSnapshotDoesNotOpenOptionsAndReleasesLock()
        {
            using (var fixture = new NativePaletteFixture())
            {
                fixture.Change(false);
                Assert.AreEqual(0, fixture.Visits);
                Assert.IsFalse(File.Exists(fixture.PathName));
                using (File.Open(fixture.PathName + ".lock", FileMode.Open, FileAccess.ReadWrite, FileShare.None)) { }
            }
        }

        [TestMethod]
        public void ApplyRepeatAndRestorePreserveOriginalUntilVerifiedReopen()
        {
            using (var fixture = new NativePaletteFixture())
            {
                fixture.Change(true);
                var saved = VbeNativePaletteState.Load(fixture.PathName, "7.1");
                Assert.IsTrue(VbeNativePaletteState.Equal(NativePaletteFixture.Rows(), saved.Original));
                Assert.IsTrue(VbeNativePaletteState.Equal(saved.Applied, fixture.Current));
                string recovery = File.ReadAllText(fixture.PathName);
                fixture.Change(true);
                Assert.AreEqual(recovery, File.ReadAllText(fixture.PathName));
                Assert.AreEqual(1, fixture.Updates);
                fixture.Change(false);
                Assert.IsTrue(VbeNativePaletteState.Equal(NativePaletteFixture.Rows(), fixture.Current));
                Assert.IsFalse(File.Exists(fixture.PathName));
                Assert.AreEqual(6, fixture.Visits);
                Assert.AreEqual(2, fixture.Updates);
            }
        }

        [TestMethod]
        public void AlreadyOriginalRestoreDoesNotWriteControls()
        {
            using (var fixture = new NativePaletteFixture())
            {
                NativePaletteFixture.State().SaveNew(fixture.PathName);
                fixture.Change(false);
                Assert.AreEqual(0, fixture.Updates);
                Assert.AreEqual(2, fixture.Visits);
                Assert.IsFalse(File.Exists(fixture.PathName));
            }
        }

        [TestMethod]
        public void ChangedSettingsAndUncommittedUpdatesRetainRecoveryAndReleaseLock()
        {
            using (var fixture = new NativePaletteFixture())
            {
                fixture.RejectCommit = true;
                Assert.ThrowsException<InvalidOperationException>(() => fixture.Change(true));
                string originalRecovery = File.ReadAllText(fixture.PathName);
                Assert.IsTrue(VbeNativePaletteState.Equal(NativePaletteFixture.Rows(), fixture.Current));
                fixture.RejectCommit = false;
                fixture.Current[0].Foreground = 12;
                Assert.ThrowsException<InvalidOperationException>(() => fixture.Change(true));
                Assert.AreEqual(originalRecovery, File.ReadAllText(fixture.PathName));
                Assert.AreEqual(12, fixture.Current[0].Foreground);
                using (File.Open(fixture.PathName + ".lock", FileMode.Open, FileAccess.ReadWrite, FileShare.None)) { }
            }
        }

        [TestMethod]
        public void DialogAndConcurrentTransactionFailuresPreserveRecovery()
        {
            using (var fixture = new NativePaletteFixture())
            {
                NativePaletteFixture.State().SaveNew(fixture.PathName);
                string originalRecovery = File.ReadAllText(fixture.PathName);
                fixture.Failure = new InvalidOperationException("synthetic dialog unavailable");
                Assert.ThrowsException<InvalidOperationException>(() => fixture.Change(false));
                Assert.AreEqual(originalRecovery, File.ReadAllText(fixture.PathName));
                fixture.Failure = null;
                using (File.Open(fixture.PathName + ".lock", FileMode.Open, FileAccess.ReadWrite, FileShare.None))
                    Assert.ThrowsException<IOException>(() => fixture.Change(true));
                Assert.AreEqual(1, fixture.Visits);
                Assert.AreEqual(originalRecovery, File.ReadAllText(fixture.PathName));
            }
        }
    }
}

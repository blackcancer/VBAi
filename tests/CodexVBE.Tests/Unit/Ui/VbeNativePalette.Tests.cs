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
                fixture.Change(true);
                Assert.AreEqual(12, VbeNativePaletteState.Load(fixture.PathName, "7.1").Original[0].Foreground);
                Assert.AreEqual(originalRecovery, File.ReadAllText(Directory.GetFiles(fixture.DirectoryPath, "*.previous-*").Single()));
                fixture.Change(false);
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

namespace CodexVBE.Tests.Unit
{
    [TestClass, TestCategory("Unit"), DoNotParallelize]
    public sealed class VbeNativePaletteSchedulerTests
    {
        [TestMethod]
        public void ConstructorRejectsUnsafeVersionsAndComputesDefaultRecoveryPathWithoutOpeningSettings()
        {
            NativeThemeFixture.OnSta(() =>
            {
                foreach (string version in new[] { null, "", "unsafe/name" })
                    Assert.ThrowsException<InvalidOperationException>(() => new VbeNativePalette(new NativeThemeVbe { Version = version }, IntPtr.Zero));
                using (var service = new VbeNativePalette(new NativeThemeVbe(), IntPtr.Zero))
                {
                    string path = (string)typeof(VbeNativePalette).GetField("path", System.Reflection.BindingFlags.Instance | System.Reflection.BindingFlags.NonPublic).GetValue(service);
                    StringAssert.EndsWith(path, System.IO.Path.Combine("CodexVBE", "native-theme", "palette-7.1.json"));
                }
            });
        }

        [TestMethod]
        public void RequestsQueueOnlyWorkThatRemainsAndDisposedServiceRejectsFurtherRequests()
        {
            NativeThemeFixture.OnSta(() =>
            {
                using (var fixture = new NativePaletteSchedulerFixture())
                {
                    fixture.Service.Request(false); Assert.IsFalse(fixture.TimerEnabled);
                    System.IO.File.WriteAllText(fixture.Recovery.PathName, "synthetic recovery marker");
                    fixture.Service.Request(false); Assert.IsTrue(fixture.TimerEnabled);
                    fixture.Tick(); Assert.IsFalse(fixture.TimerEnabled); Assert.AreEqual(false, fixture.Field("applied"));
                    fixture.Service.Request(false); Assert.IsFalse(fixture.TimerEnabled);
                    fixture.Service.Request(true); Assert.IsTrue(fixture.TimerEnabled);
                    fixture.Tick(); Assert.AreEqual(true, fixture.Field("applied")); Assert.IsFalse(fixture.TimerEnabled);
                    fixture.Service.Request(true); Assert.IsFalse(fixture.TimerEnabled);
                    fixture.Busy = 1; fixture.Service.Request(true); Assert.IsTrue(fixture.TimerEnabled); fixture.Busy = 0;
                    fixture.Service.Dispose(); fixture.Service.Dispose(); fixture.Service.Request(false); fixture.Tick();
                    Assert.IsFalse(fixture.TimerEnabled); Assert.IsTrue((bool)fixture.Field("requested"));
                    Assert.AreEqual(2, fixture.Targets.Count);
                }
            });
        }

        [TestMethod]
        public void HiddenDisabledDisposedAndBusyOwnersDoNotRunPendingTransactions()
        {
            NativeThemeFixture.OnSta(() =>
            {
                using (var fixture = new NativePaletteSchedulerFixture())
                {
                    fixture.Service.Request(true);
                    fixture.Visible(false); fixture.Tick(); Assert.AreEqual(0, fixture.Targets.Count);
                    fixture.Visible(true); fixture.Enabled(false); fixture.Tick(); Assert.AreEqual(0, fixture.Targets.Count);
                    fixture.Enabled(true); fixture.Busy = 1; fixture.Tick(); Assert.AreEqual(0, fixture.Targets.Count);
                    Assert.IsTrue(fixture.TimerEnabled); fixture.Busy = 0;
                    fixture.Service.Dispose(); fixture.Tick(); Assert.AreEqual(0, fixture.Targets.Count);
                    Assert.AreEqual(0, fixture.Errors.Count); Assert.IsNull(fixture.Field("applied"));
                }
            });
        }

        [TestMethod]
        public void ChangedTargetRestartsTimerThenCommitsRestorationOnNextTick()
        {
            NativeThemeFixture.OnSta(() =>
            {
                using (var fixture = new NativePaletteSchedulerFixture())
                {
                    fixture.DuringChange = () => fixture.Service.Request(false);
                    fixture.Service.Request(true); fixture.Tick();
                    Assert.AreEqual(true, fixture.Field("applied")); Assert.IsTrue(fixture.TimerEnabled); Assert.AreEqual(0, fixture.Busy);
                    fixture.DuringChange = null; fixture.Tick();
                    CollectionAssert.AreEqual(new[] { true, false }, fixture.Targets);
                    Assert.AreEqual(false, fixture.Field("applied")); Assert.IsFalse(fixture.TimerEnabled); Assert.AreEqual(0, fixture.Busy);
                    Assert.IsTrue(fixture.Messages.Exists(m => m.Contains("palette applied and verified")));
                    Assert.IsTrue(fixture.Messages.Exists(m => m.Contains("palette restored and verified")));
                }
            });
        }

        [TestMethod]
        public void DisposingDuringTransactionPreventsTimerRestartForChangedRequest()
        {
            NativeThemeFixture.OnSta(() =>
            {
                using (var fixture = new NativePaletteSchedulerFixture())
                {
                    fixture.DuringChange = () => { fixture.Service.Request(false); fixture.Service.Dispose(); };
                    fixture.Service.Request(true); fixture.Tick();
                    Assert.IsFalse(fixture.TimerEnabled); Assert.IsTrue((bool)fixture.Field("disposed"));
                    Assert.AreEqual(true, fixture.Field("applied")); Assert.AreEqual(0, fixture.Busy); Assert.AreEqual(1, fixture.Targets.Count);
                }
            });
        }

        [TestMethod]
        public void NestedTickCannotEnterAnotherTransactionWhileOuterTransactionOwnsUpdateLock()
        {
            NativeThemeFixture.OnSta(() =>
            {
                using (var fixture = new NativePaletteSchedulerFixture())
                {
                    fixture.DuringChange = () => { fixture.Service.Request(true); fixture.Tick(); };
                    fixture.Service.Request(true); fixture.Tick();
                    Assert.AreEqual(1, fixture.Targets.Count); Assert.AreEqual(0, fixture.Busy); Assert.AreEqual(true, fixture.Field("applied"));
                    Assert.AreEqual(0, fixture.Errors.Count);
                }
            });
        }

        [TestMethod]
        public void TransactionFailureReportsExactErrorRetainsAppliedStateAndReleasesUpdateLock()
        {
            NativeThemeFixture.OnSta(() =>
            {
                using (var fixture = new NativePaletteSchedulerFixture())
                {
                    fixture.Service.Request(true); fixture.Tick();
                    var expected = new InvalidOperationException("synthetic native commit rejected"); fixture.Failure = expected;
                    fixture.Service.Request(false); fixture.Tick();
                    Assert.AreEqual(true, fixture.Field("applied")); Assert.AreEqual(0, fixture.Busy); Assert.IsFalse(fixture.TimerEnabled);
                    Assert.AreSame(expected, fixture.Errors[0]); Assert.AreEqual(1, fixture.Errors.Count);
                    Assert.IsTrue(fixture.Messages.Exists(m => m.Contains("Native editor palette failed: System.InvalidOperationException: synthetic native commit rejected")));
                }
            });
        }
    }
}

namespace CodexVBE.Tests.Unit
{
    [TestClass, TestCategory("Unit"), DoNotParallelize]
    public sealed class VbeNativePaletteHostBoundaryTests
    {
        [TestMethod]
        public void HostChangeWithoutSnapshotNeverOpensOptionsAndMissingOwnerIsRejectedBeforeSettings()
        {
            using (var fixture = new NativePaletteFixture())
            {
                var vbe = new NativeThemeVbe();
                VbeNativePalette.Change(vbe, false, fixture.PathName);
                VbeNativePalette.Change(vbe, false, fixture.PathName);
                Assert.IsFalse(System.IO.File.Exists(fixture.PathName));
                var actual = Assert.ThrowsException<InvalidOperationException>(() => VbeNativePalette.Change(vbe, true, fixture.PathName));
                StringAssert.Contains(actual.Message, "must be visible and have no modal dialog open");
                Assert.IsFalse(System.IO.File.Exists(fixture.PathName));
                using (System.IO.File.Open(fixture.PathName + ".lock", System.IO.FileMode.Open, System.IO.FileAccess.ReadWrite, System.IO.FileShare.None)) { }
            }
        }

        [TestMethod]
        public void DefaultFailurePresenterShowsTranslatedDiagnosticAndBaseErrorOnOwnedStaOnly()
        {
            NativeThemeFixture.OnSta(() =>
            {
                using (var fixture = new NativePaletteSchedulerFixture())
                {
                    string marker = "Synthetic-palette-error-" + Guid.NewGuid().ToString("N");
                    var error = new InvalidOperationException("outer", new InvalidOperationException(marker));
                    using (var service = new VbeNativePalette(fixture.Vbe, fixture.Window, fixture.Recovery.PathName,
                        (vbe, enabled, path) => { throw error; }))
                    {
                        service.Request(true);
                        string body = OwnedPaletteMessageBox.Dismiss(() => typeof(VbeNativePalette)
                            .GetMethod("ApplyPending", System.Reflection.BindingFlags.Instance | System.Reflection.BindingFlags.NonPublic)
                            .Invoke(service, new object[] { null, EventArgs.Empty }), marker);
                        StringAssert.Contains(body, UiText.Get("Native editor colors could not be updated. See the log for details."));
                        StringAssert.Contains(body, marker); Assert.AreEqual(0, fixture.Busy);
                    }
                }
            });
        }
    }
}

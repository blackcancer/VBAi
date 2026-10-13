using Microsoft.VisualStudio.TestTools.UnitTesting;
using System;
using System.Collections.Generic;
using System.IO;
using System.Reflection;
using VBAi.Tests.Integration;

namespace VBAi.Tests.Unit
{
    [TestClass, TestCategory("Unit")]
    public sealed class OfficeWordCoverageEventScopeTests
    {
        [DataTestMethod, DataRow(2), DataRow(3)]
        public void OneEnableAndRestorePreserveTheOriginalSuppressionAndReleaseOnlyOwnedCommands(int after)
        {
            var probe = new Probe(); var scope = new OfficeWordCoverageEventScope(probe.Boundary);
            scope.Enter(); Assert.IsTrue(scope.Active); Assert.IsFalse(scope.Uncertain);
            probe.Security = after; scope.Restore();
            Assert.IsFalse(scope.Active || scope.Uncertain);
            Assert.IsTrue(scope.Restored && scope.ReferencesReleased);
            CollectionAssert.AreEqual(new[] { 0, 1 }, probe.Writes.ToArray());
            CollectionAssert.AreEqual(new[] { "enable-intent", "enable-returned", "enable-checked", "restore-intent", "restore-returned", "restore-released" }, probe.Records.ToArray());
            Assert.AreEqual(1, probe.Releases);
            Assert.ThrowsException<InvalidOperationException>(() => scope.Enter());
            Assert.ThrowsException<InvalidOperationException>(() => scope.Restore());
            Assert.AreEqual(2, probe.Writes.Count);
        }

        [DataTestMethod, DataRow("RequireOwner"), DataRow("ReviewTemplates"), DataRow("ReadSecurity"), DataRow("AcquireCommands"), DataRow("ApplySuppression"), DataRow("Record"), DataRow("Release")]
        public void MissingDependencyCannotAcquireOrWrite(string field)
        {
            var probe = new Probe();
            typeof(OfficeWordCoverageEventScope.Boundary).GetField(field, BindingFlags.Instance | BindingFlags.NonPublic).SetValue(probe.Boundary, null);
            var scope = new OfficeWordCoverageEventScope(probe.Boundary);
            Assert.ThrowsException<ArgumentNullException>(() => scope.Enter());
            Assert.AreEqual(0, probe.Calls.Count); Assert.IsFalse(scope.Uncertain);
        }

        [TestMethod]
        public void NullBoundaryRefusesBeforeMutationAndCannotBeRepeated()
        {
            var scope = new OfficeWordCoverageEventScope(null);
            Assert.ThrowsException<ArgumentNullException>(() => scope.Enter());
            Assert.ThrowsException<InvalidOperationException>(() => scope.Enter());
            Assert.ThrowsException<InvalidOperationException>(() => scope.Restore());
        }

        [DataTestMethod, DataRow("owner:1"), DataRow("security:1"), DataRow("unsafe-security"), DataRow("templates"), DataRow("acquire"),
            DataRow("null-command"), DataRow("owner:2"), DataRow("security:2"), DataRow("security-changed"), DataRow("enable-intent")]
        public void KnownSetupRefusalsNeverChangeSuppressionAndReleaseAnyAcquiredCommand(string fault)
        {
            var probe = new Probe { Fault = fault }; var scope = new OfficeWordCoverageEventScope(probe.Boundary);
            Assert.IsNotNull(Capture(scope.Enter)); Assert.IsFalse(scope.Active || scope.Uncertain);
            Assert.AreEqual(0, probe.Writes.Count);
            Assert.AreEqual(fault == "owner:2" || fault == "security:2" || fault == "security-changed" || fault == "enable-intent" ? 1 : 0, probe.Releases);
            Assert.ThrowsException<InvalidOperationException>(() => scope.Enter());
        }

        [DataTestMethod, DataRow("write:1"), DataRow("enable-returned"), DataRow("templates:2"), DataRow("owner:3"),
            DataRow("security:3"), DataRow("post-security-changed"), DataRow("enable-checked")]
        public void UncertainSetupRetainsTheCommandWithoutRestoreOrReplay(string fault)
        {
            var probe = new Probe { Fault = fault }; var scope = new OfficeWordCoverageEventScope(probe.Boundary);
            Assert.IsNotNull(Capture(scope.Enter)); Assert.IsTrue(scope.Uncertain);
            Assert.AreEqual(1, probe.Writes.Count); Assert.AreEqual(0, probe.Releases);
            Assert.ThrowsException<InvalidOperationException>(() => scope.Restore());
            Assert.ThrowsException<InvalidOperationException>(() => scope.Enter());
        }

        [DataTestMethod, DataRow("owner:4"), DataRow("security:4"), DataRow("restore-unsafe"), DataRow("restore-intent"),
            DataRow("write:2"), DataRow("restore-returned"), DataRow("release"), DataRow("restore-released")]
        public void EveryRestorationFailureRetainsOwnershipAndCannotRepeatNativeWrites(string fault)
        {
            var probe = new Probe(); var scope = new OfficeWordCoverageEventScope(probe.Boundary); scope.Enter();
            probe.Fault = fault; if (fault == "restore-unsafe") probe.Security = 1;
            Assert.IsNotNull(Capture(scope.Restore)); Assert.IsTrue(scope.Uncertain);
            int writes = probe.Writes.Count, releases = probe.Releases;
            Assert.ThrowsException<InvalidOperationException>(() => scope.Restore());
            Assert.AreEqual(writes, probe.Writes.Count); Assert.AreEqual(releases, probe.Releases);
        }

        [TestMethod]
        public void SetupRefusalAndFailedReleasePreserveBothErrorsWithoutRetry()
        {
            var probe = new Probe { Fault = "owner:2", FailRelease = true }; var scope = new OfficeWordCoverageEventScope(probe.Boundary);
            var error = Assert.ThrowsException<AggregateException>(() => scope.Enter());
            Assert.AreEqual(2, error.InnerExceptions.Count); Assert.AreSame(probe.Error, error.InnerExceptions[0]);
            Assert.IsTrue(scope.Uncertain); Assert.AreEqual(1, probe.Releases); Assert.AreEqual(0, probe.Writes.Count);
        }

        [TestMethod]
        public void UncertainNativeCopyRetainsActiveScopeWithoutAutomaticallyChangingFlags()
        {
            var probe = new Probe(); var scope = new OfficeWordCoverageEventScope(probe.Boundary); scope.Enter();
            scope.Retain(); scope.Retain(); Assert.IsTrue(scope.Active && scope.Uncertain);
            Assert.ThrowsException<InvalidOperationException>(() => scope.Restore());
            Assert.AreEqual(1, probe.Writes.Count); Assert.AreEqual(0, probe.Releases);
        }

        private static Exception Capture(Action action) { try { action(); return null; } catch (Exception error) { return error; } }

        private sealed class Probe
        {
            internal readonly List<string> Calls = new List<string>(), Records = new List<string>();
            internal readonly List<int> Writes = new List<int>();
            internal readonly object Commands = new object();
            internal readonly IOException Error = new IOException("injected");
            internal readonly OfficeWordCoverageEventScope.Boundary Boundary;
            internal int Security = 3, Releases, owners, reads, reviews;
            internal string Fault;
            internal bool FailRelease;
            internal Probe()
            {
                Boundary = new OfficeWordCoverageEventScope.Boundary
                {
                    RequireOwner = () => Phase("owner:" + ++owners),
                    ReviewTemplates = () => { Phase("templates"); Phase("templates:" + ++reviews); },
                    ReadSecurity = () => { Phase("security:" + ++reads); return Fault == "unsafe-security" || (Fault == "security-changed" && reads == 2) || (Fault == "post-security-changed" && reads == 3) ? 1 : Security; },
                    AcquireCommands = () => { Phase("acquire"); return Fault == "null-command" ? null : Commands; },
                    ApplySuppression = (value, flag) => { Assert.AreSame(Commands, value); Writes.Add(flag); Phase("write:" + Writes.Count); },
                    Record = phase => { Records.Add(phase); Phase(phase); },
                    Release = value => { Assert.AreSame(Commands, value); Releases++; Phase("release"); if (FailRelease) throw new IOException("release"); }
                };
            }
            private void Phase(string name) { Calls.Add(name); if (Fault == name) throw Error; }
        }
    }
}

using Microsoft.VisualStudio.TestTools.UnitTesting;
using System;
using System.Collections.Generic;
using System.IO;
using System.Reflection;
using VBAi.Tests.Integration;

namespace VBAi.Tests.Unit
{
    /// <summary>Complete refusal, mutation and recovery matrix for a single macro-free Word control, without Office activation.</summary>
    [TestClass, TestCategory("Unit")]
    public sealed class OfficeWordLifecycleControlTests
    {
        private const string Original = @"C:\Owned\Original.docm", Copy = @"C:\Owned\Control.docx";

        [DataTestMethod, DataRow(false, false), DataRow(false, true), DataRow(true, false), DataRow(true, true)]
        public void EveryVisibilityAndSuppressionControlRestoresSafetyAndReleasesExactlyItsAcquisitions(bool visible, bool suppress)
        {
            var probe = new Probe(); var control = new OfficeWordLifecycleControl();
            control.Run(probe.Boundary, Original, Copy, visible, suppress);
            Assert.IsTrue(control.Entered); Assert.IsFalse(control.Uncertain);
            Assert.AreEqual(1, probe.Opens); Assert.AreEqual(1, probe.Closes);
            Assert.AreEqual(visible, probe.Visible);
            CollectionAssert.AreEqual(new[] { suppress ? 1 : 0, 1 }, probe.Suppression.ToArray());
            CollectionAssert.AreEqual(new[] { 3, 2 }, probe.SecurityWrites.ToArray());
            CollectionAssert.AreEqual(new[] { probe.Document, probe.Commands, probe.Application }, probe.Released.ToArray());
            Assert.AreEqual(2, probe.Security); Assert.IsFalse(probe.IsOpen);
            Assert.IsTrue(control.OpenEnteredWithForceDisable && control.OriginalSuppressionRestoreReturned && control.OriginalSecurityRestored);
            Assert.ThrowsException<InvalidOperationException>(() => control.Run(probe.Boundary, Original, Copy, visible, suppress));
            Assert.AreEqual(1, probe.Opens); Assert.AreEqual(1, probe.Closes);
        }

        [DataTestMethod, DataRow("RequireCopy"), DataRow("RequireAutoMacroScope"), DataRow("Acquire"), DataRow("RequireOwner"), DataRow("Paths"),
            DataRow("ReadSecurity"), DataRow("WriteSecurity"), DataRow("WordBasic"), DataRow("SuppressAutoMacros"),
            DataRow("Open"), DataRow("DocumentPath"), DataRow("HasVba"), DataRow("ReadOnly"), DataRow("Close"), DataRow("Release")]
        public void MissingDependencyRefusesBeforeAnyAcquisitionOrNativeAction(string dependency)
        {
            var probe = new Probe(); var control = new OfficeWordLifecycleControl();
            typeof(OfficeWordLifecycleControl.Boundary).GetField(dependency, BindingFlags.Instance | BindingFlags.NonPublic)
                .SetValue(probe.Boundary, null);
            Assert.ThrowsException<ArgumentNullException>(() => control.Run(probe.Boundary, Original, Copy, false, true));
            Assert.AreEqual(0, probe.Calls.Count); Assert.IsFalse(control.Uncertain);
        }

        [DataTestMethod, DataRow("original"), DataRow("copy"), DataRow("same"), DataRow("boundary")]
        public void InvalidInputsCannotEnterTheNativeBoundary(string fault)
        {
            var probe = new Probe(); var control = new OfficeWordLifecycleControl();
            try
            {
                control.Run(fault == "boundary" ? null : probe.Boundary,
                    fault == "original" ? null : Original, fault == "copy" ? null : fault == "same" ? Original : Copy, false, false);
                Assert.Fail("Invalid inputs were accepted.");
            }
            catch (ArgumentException) { }
            Assert.AreEqual(0, probe.Calls.Count); Assert.IsFalse(control.Uncertain);
        }

        [DataTestMethod, DataRow("copy:1"), DataRow("acquire"), DataRow("nullApplication"), DataRow("owner:1"),
            DataRow("paths:1"), DataRow("foreignOriginal"), DataRow("read:1"), DataRow("unsafeSecurity"),
            DataRow("commands"), DataRow("nullCommands")]
        public void KnownPreflightRefusalsReleaseAcquiredReferencesAndNeverOpenADocument(string fault)
        {
            var probe = new Probe { Fault = fault }; var control = new OfficeWordLifecycleControl();
            Assert.IsNotNull(Capture(() => control.Run(probe.Boundary, Original, Copy, false, true)));
            Assert.IsFalse(control.Uncertain);
            Assert.AreEqual(0, probe.Opens + probe.Closes + probe.SecurityWrites.Count);
            Assert.AreEqual(fault == "copy:1" || fault == "acquire" || fault == "nullApplication" ? 0 : 1, probe.Released.Count);
        }

        [DataTestMethod, DataRow("write:1"), DataRow("read:2"), DataRow("forceMismatch"), DataRow("suppress:1"),
            DataRow("owner:2"), DataRow("copy:2"), DataRow("read:3"), DataRow("recheckMismatch"), DataRow("open"), DataRow("nullDocument"), DataRow("path"),
            DataRow("wrongPath"), DataRow("hasVba"), DataRow("macros"), DataRow("readonly"), DataRow("writable"),
            DataRow("paths:2"), DataRow("foreignOpen"), DataRow("owner:3"), DataRow("close"), DataRow("paths:3"),
            DataRow("copyRemains"), DataRow("copy:3"), DataRow("suppress:2"), DataRow("write:2"), DataRow("read:4"),
            DataRow("restoreMismatch"), DataRow("release")]
        public void AnyUncertainMutationOrRecoveryRetainsReferencesAndNeverRetriesOrClosesAnotherDocument(string fault)
        {
            var probe = new Probe { Fault = fault }; var control = new OfficeWordLifecycleControl();
            Assert.IsNotNull(Capture(() => control.Run(probe.Boundary, Original, Copy, true, false)));
            Assert.IsTrue(control.Uncertain);
            Assert.AreEqual(fault == "release" ? 1 : 0, probe.Released.Count);
            Assert.IsTrue(probe.Opens <= 1 && probe.Closes <= 1);
            int entries = probe.Calls.Count;
            Assert.ThrowsException<InvalidOperationException>(() => control.Run(probe.Boundary, Original, Copy, false, true));
            Assert.AreEqual(entries, probe.Calls.Count);
            if (fault == "wrongPath" || fault == "macros" || fault == "writable" || fault == "foreignOpen")
                Assert.AreEqual(0, probe.Closes, "An unverified or foreign document must never be closed as cleanup.");
        }

        [TestMethod]
        public void PreflightAndReleaseFailuresAreBothPreservedWithoutSecondRelease()
        {
            var probe = new Probe { Fault = "owner:1", FailRelease = true }; var control = new OfficeWordLifecycleControl();
            var failure = Assert.ThrowsException<AggregateException>(() => control.Run(probe.Boundary, Original, Copy, false, true));
            Assert.AreEqual(2, failure.InnerExceptions.Count);
            Assert.AreSame(probe.Error, failure.InnerExceptions[0]);
            Assert.IsTrue(control.Uncertain); Assert.AreEqual(1, probe.Released.Count);
            Assert.AreEqual(0, probe.Opens + probe.Closes);
        }

        [TestMethod]
        public void ForeignTemplateRefusesBeforeSecurityWritesOrAnySuppressionChange()
        {
            var probe = new Probe { Fault = "templates" }; var control = new OfficeWordLifecycleControl();
            Assert.AreSame(probe.Error, Assert.ThrowsException<IOException>(() => control.Run(probe.Boundary, Original, Copy, false, false)));
            Assert.AreEqual(0, probe.Opens + probe.Closes + probe.SecurityWrites.Count + probe.Suppression.Count);
            Assert.AreEqual(1, probe.Released.Count); Assert.IsFalse(control.Uncertain);
        }

        [TestMethod]
        public void SuppressedControlDoesNotReadOrEnableAnyForeignTemplate()
        {
            var probe = new Probe { Fault = "templates" }; var control = new OfficeWordLifecycleControl();
            control.Run(probe.Boundary, Original, Copy, false, true);
            Assert.IsFalse(probe.Calls.Contains("templates"));
            CollectionAssert.AreEqual(new[] { 1, 1 }, probe.Suppression.ToArray());
        }

        [TestMethod]
        public void TemplateScopeChangingAfterTheFlagWriteRefusesOpenAndRemainsUncertainWithoutReplay()
        {
            var probe = new Probe(); int checks = 0;
            probe.Boundary.RequireAutoMacroScope = value => { if (++checks == 2) throw new IOException("scope changed"); };
            var control = new OfficeWordLifecycleControl();
            Assert.ThrowsException<IOException>(() => control.Run(probe.Boundary, Original, Copy, false, false));
            Assert.AreEqual(2, checks); Assert.AreEqual(0, probe.Opens + probe.Closes);
            Assert.IsTrue(control.Uncertain); Assert.AreEqual(0, probe.Released.Count);
            CollectionAssert.AreEqual(new[] { 0 }, probe.Suppression.ToArray());
            Assert.ThrowsException<InvalidOperationException>(() => control.Run(probe.Boundary, Original, Copy, false, false));
        }

        private static Exception Capture(Action action) { try { action(); return null; } catch (Exception error) { return error; } }

        private sealed class Probe
        {
            internal readonly object Application = new object(), Commands = new object(), Document = new object();
            internal readonly OfficeWordLifecycleControl.Boundary Boundary;
            internal readonly List<string> Calls = new List<string>();
            internal readonly List<object> Released = new List<object>();
            internal readonly List<int> SecurityWrites = new List<int>(), Suppression = new List<int>();
            internal readonly IOException Error = new IOException("The original native control failed.");
            internal string Fault;
            internal bool FailRelease, IsOpen, Visible;
            internal int Opens, Closes, Security = 2, owners, copies, reads, writes, paths;
            internal Probe()
            {
                Boundary = new OfficeWordLifecycleControl.Boundary
                {
                    Acquire = () => { Phase("acquire"); return Fault == "nullApplication" ? null : Application; },
                    RequireCopy = () => Phase("copy:" + ++copies),
                    RequireAutoMacroScope = value => { Assert.AreSame(Application, value); Phase("templates"); },
                    RequireOwner = value => { Assert.AreSame(Application, value); Phase("owner:" + ++owners); },
                    Paths = value =>
                    {
                        Assert.AreSame(Application, value); Phase("paths:" + ++paths);
                        if (paths == 1 && Fault == "foreignOriginal") return new[] { "foreign.docm" };
                        if (IsOpen && Fault == "foreignOpen") return new[] { Copy, "foreign.docm" };
                        if (!IsOpen && Closes != 0 && Fault == "copyRemains") return new[] { Original, Copy };
                        return IsOpen ? new[] { Original, Copy } : new[] { Original };
                    },
                    ReadSecurity = value =>
                    {
                        Phase("read:" + ++reads);
                        return Fault == "unsafeSecurity" && reads == 1 ? 1 : Fault == "forceMismatch" && reads == 2 ? 2 :
                            Fault == "recheckMismatch" && reads == 3 ? 2 : Fault == "restoreMismatch" && reads == 4 ? 3 : Security;
                    },
                    WriteSecurity = (value, security) => { SecurityWrites.Add(security); Phase("write:" + ++writes); Security = security; },
                    WordBasic = value => { Phase("commands"); return Fault == "nullCommands" ? null : Commands; },
                    SuppressAutoMacros = (value, state) =>
                    {
                        Assert.AreSame(Commands, value); Assert.AreEqual(3, Security, "Flag changes require ForceDisable.");
                        Suppression.Add(state); Phase("suppress:" + Suppression.Count);
                    },
                    Open = (value, path, visible) =>
                    {
                        Opens++; Phase("open"); Assert.AreEqual(Copy, path); Assert.AreEqual(3, Security);
                        Visible = visible; IsOpen = true; return Fault == "nullDocument" ? null : Document;
                    },
                    DocumentPath = value => { Assert.AreSame(Document, value); Phase("path"); return Fault == "wrongPath" ? Original : Copy; },
                    HasVba = value => { Phase("hasVba"); return Fault == "macros"; },
                    ReadOnly = value => { Phase("readonly"); return Fault != "writable"; },
                    Close = value => { Assert.AreSame(Document, value); Closes++; Phase("close"); IsOpen = false; },
                    Release = value => { Released.Add(value); Phase("release"); if (FailRelease) throw new IOException("release failed"); }
                };
            }
            private void Phase(string name) { Calls.Add(name); if (Fault == name) throw Error; }
        }
    }
}

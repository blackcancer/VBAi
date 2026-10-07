using Microsoft.VisualStudio.TestTools.UnitTesting;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using VBAi.Tests.Integration;

namespace VBAi.Tests.Unit
{
    [TestClass, TestCategory("Unit")]
    public sealed class OfficeVbeFixtureDesktopTests
    {
        [TestMethod]
        public void OfficeDesktopPairSelectsOnlyMatchingGeneratedNamesWithoutHostActivation()
        {
            string name = "VBAiTests_" + Guid.NewGuid().ToString("N");
            Assert.IsNull(OfficeVbeFixture.RequireOfficeDesktopPair(null, null));
            Assert.IsNull(OfficeVbeFixture.RequireOfficeDesktopPair("", ""));
            Assert.AreEqual(name, OfficeVbeFixture.RequireOfficeDesktopPair(null, name));
            Assert.AreEqual(name, OfficeVbeFixture.RequireOfficeDesktopPair(name, name));
            foreach (string configured in new[] { null, "", "Default", name.ToLowerInvariant(), "VBAiTests_" + Guid.NewGuid().ToString("N") })
                Assert.ThrowsException<InvalidOperationException>(() => OfficeVbeFixture.RequireOfficeDesktopPair(name, configured));
            Assert.ThrowsException<ArgumentException>(() => OfficeVbeFixture.RequireOfficeDesktopPair(null, "Default"));
        }

        /// <summary>Normal Publisher GUI mode is a changed bootstrap hypothesis; no automation/embedding switches are selected.</summary>
        [TestMethod]
        public void PrivateOfficeArgumentsUseNormalGuiWithoutAutomationEmbeddingOrFallback()
        {
            CollectionAssert.AreEqual(new string[0], OfficeVbeFixture.PrivateOfficeArguments("Access"));
            var first = OfficeVbeFixture.PrivateOfficeArguments("Publisher");
            CollectionAssert.AreEqual(new string[0], first);
            Assert.AreNotSame(first, OfficeVbeFixture.PrivateOfficeArguments("Publisher"));
            CollectionAssert.AreEqual(new string[0], OfficeVbeFixture.PrivateOfficeArguments("PowerPoint"));
            foreach (string host in new[] { "Word", "Excel", "", null })
                Assert.ThrowsException<InvalidOperationException>(() => OfficeVbeFixture.PrivateOfficeArguments(host));
        }

        [DataTestMethod]
        [DataRow("\"C:\\Office\\MSPUB.EXE\" /Automation", "/Automation")]
        [DataRow("C:\\Office\\MSPUB.EXE /Automation -Embedding", "/Automation -Embedding")]
        [DataRow("\"C:\\Office\\MSPUB.EXE\"", "")]
        public void RegisteredPublisherArgumentsAreObservedSeparatelyFromEmptySelectedArguments(string command, string expected)
        {
            Assert.AreEqual(expected, OfficeVbeFixture.RegisteredPublisherServerArguments(@"C:\Office\MSPUB.EXE", command));
            CollectionAssert.AreEqual(new string[0], OfficeVbeFixture.PrivateOfficeArguments("Publisher"));
        }

        [DataTestMethod]
        [DataRow("\"C:\\Other\\MSPUB.EXE\" /Automation")]
        [DataRow("C:\\Office\\MSPUB.EXEX /Automation")]
        [DataRow("")]
        [DataRow(null)]
        [DataRow("\"C:\\Office\\MSPUB.EXE\" /Automation\n-Embedding")]
        public void RegisteredPublisherArgumentsNeverInferAnotherImageOrMalformedCommand(string command)
        {
            Assert.ThrowsException<InvalidOperationException>(() => OfficeVbeFixture.RegisteredPublisherServerArguments(@"C:\Office\MSPUB.EXE", command));
        }
        /// <summary>Getter receipts are written before and after exactly one read, with no environment or native mutation.</summary>
        [DataTestMethod, DataRow("UserControl"), DataRow("Visible")]
        public void PrivateAccessAutomationGetterRecordsOneExactBooleanRead(string property)
        {
            var receipts = new List<IDictionary<string, object>>();
            int calls = 0;
            bool observed = OfficeVbeFixture.ObservePrivateAccessAutomationGetter(property, () =>
            {
                Assert.AreEqual(1, receipts.Count);
                Assert.AreEqual("PENDING", receipts[0]["State"]);
                calls++;
                return false;
            }, receipts.Add);
            Assert.IsFalse(observed);
            Assert.AreEqual(1, calls);
            Assert.AreEqual(2, receipts.Count);
            Assert.AreEqual("RETURNED", receipts[1]["State"]);
            Assert.AreEqual(false, receipts[1]["Value"]);
            Assert.IsTrue(receipts.All(row => Equals(row["MutationInvoked"], false)));
        }

        /// <summary>A transient-looking COM refusal is terminal once attachment is accepted.</summary>
        [TestMethod]
        public void PrivateAccessAutomationGetterFailurePreservesOriginalHResultWithoutReadReplay()
        {
            var receipts = new List<IDictionary<string, object>>();
            int calls = 0;
            var original = new COMException("Synthetic getter refusal", unchecked((int)0x80010001));
            var observed = Assert.ThrowsException<COMException>(() => OfficeVbeFixture.ObservePrivateAccessAutomationGetter(
                "UserControl", () => { calls++; throw original; }, receipts.Add));
            Assert.AreSame(original, observed);
            Assert.AreEqual(1, calls);
            Assert.AreEqual(2, receipts.Count);
            Assert.AreEqual("FAILED", receipts[1]["State"]);
            Assert.AreEqual("0x80010001", receipts[1]["HResult"]);
            Assert.AreEqual(false, receipts[1]["AutomaticRetry"]);
            Assert.AreEqual(false, receipts[1]["MutationInvoked"]);
        }

        /// <summary>Incomplete evidence or non-Boolean metadata cannot authorize a setter.</summary>
        [TestMethod]
        public void PrivateAccessAutomationGetterRefusesUnknownFieldsMalformedValuesAndEvidenceFailures()
        {
            int calls = 0;
            Assert.ThrowsException<ArgumentException>(() => OfficeVbeFixture.ObservePrivateAccessAutomationGetter(
                "Version", () => { calls++; return false; }, row => { }));
            Assert.AreEqual(0, calls);
            foreach (object value in new object[] { null, 0, "False" })
                Assert.ThrowsException<InvalidOperationException>(() => OfficeVbeFixture.ObservePrivateAccessAutomationGetter(
                    "Visible", () => value, row => { }));
            var evidenceFailure = new InvalidOperationException("Synthetic pending evidence failure");
            Assert.AreSame(evidenceFailure, Assert.ThrowsException<InvalidOperationException>(() =>
                OfficeVbeFixture.ObservePrivateAccessAutomationGetter("Visible", () => { calls++; return false; },
                    row => { throw evidenceFailure; })));
            Assert.AreEqual(0, calls);
            var original = new COMException("Synthetic getter failure", unchecked((int)0x800A0997));
            var aggregate = Assert.ThrowsException<AggregateException>(() => OfficeVbeFixture.ObservePrivateAccessAutomationGetter(
                "Visible", () => { throw original; }, row =>
                {
                    if (Equals(row["State"], "FAILED")) throw evidenceFailure;
                }));
            Assert.AreSame(original, aggregate.InnerExceptions[0]);
            Assert.AreSame(evidenceFailure, aggregate.InnerExceptions[1]);
        }

        /// <summary>A proven visible owned instance skips the setter in either launch mode.</summary>
        [DataTestMethod, DataRow(true), DataRow(false)]
        public void PrivateAccessVisibleStateIsOwnedNoOpAndConsumedOnce(bool userControl)
        {
            var decision = new OfficeVbeFixture.PrivateAccessVisibilityDecision(userControl, true);
            var receipts = new List<IDictionary<string, object>>();
            int guards = 0, setters = 0;
            decision.Apply(() => guards++, () => setters++, receipts.Add);
            Assert.AreEqual(1, guards);
            Assert.AreEqual(0, setters);
            Assert.AreEqual(2, receipts.Count);
            Assert.AreEqual("ExistingOwnedVisibleState", receipts[0]["PrivateAccessVisibility"]);
            Assert.AreEqual("NO_OP", receipts[1]["State"]);
            Assert.AreEqual(false, receipts[1]["MutationInvoked"]);
            Assert.AreEqual(0, receipts[1]["SetterEntries"]);
            Assert.ThrowsException<InvalidOperationException>(() =>
                decision.Apply(() => guards++, () => setters++, receipts.Add));
            Assert.AreEqual(1, guards);
            Assert.AreEqual(0, setters);
            Assert.AreEqual(2, receipts.Count);
        }

        /// <summary>A hidden automation instance enters the setter once after ownership and durable intent.</summary>
        [TestMethod]
        public void PrivateAccessHiddenAutomationStateSetsVisibleOnceAfterOwnership()
        {
            var decision = new OfficeVbeFixture.PrivateAccessVisibilityDecision(false, false);
            var receipts = new List<IDictionary<string, object>>();
            int guards = 0, setters = 0;
            decision.Apply(() => guards++, () =>
            {
                Assert.AreEqual(1, guards);
                Assert.AreEqual(1, receipts.Count);
                Assert.AreEqual("PENDING", receipts[0]["State"]);
                Assert.AreEqual(1, receipts[0]["MaximumSetterEntries"]);
                setters++;
            }, receipts.Add);
            Assert.AreEqual(1, setters);
            Assert.AreEqual("RETURNED", receipts[1]["State"]);
            Assert.ThrowsException<InvalidOperationException>(() =>
                decision.Apply(() => guards++, () => setters++, receipts.Add));
            Assert.AreEqual(1, setters);
            Assert.AreEqual(1, guards);
        }

        [TestMethod]
        public void PrivateAccessManuallyControlledHiddenStateRefusesAnyDecision()
        {
            Assert.ThrowsException<InvalidOperationException>(() =>
                new OfficeVbeFixture.PrivateAccessVisibilityDecision(true, false));
        }

        /// <summary>Ownership or intent failure consumes the decision before entering a native setter.</summary>
        [DataTestMethod, DataRow(true), DataRow(false)]
        public void PrivateAccessVisibilityGuardOrIntentFailureNeverEntersSetterOrReplays(bool ownerFails)
        {
            var decision = new OfficeVbeFixture.PrivateAccessVisibilityDecision(false, false);
            int setters = 0, guards = 0, receipts = 0;
            var original = new InvalidOperationException("Synthetic ownership or evidence refusal");
            Action guard = () => { guards++; if (ownerFails) throw original; };
            Action<IDictionary<string, object>> record = row => { receipts++; if (!ownerFails) throw original; };
            Assert.AreSame(original, Assert.ThrowsException<InvalidOperationException>(() =>
                decision.Apply(guard, () => setters++, record)));
            Assert.AreEqual(0, setters);
            Assert.AreEqual(1, guards);
            Assert.AreEqual(ownerFails ? 0 : 1, receipts);
            Assert.ThrowsException<InvalidOperationException>(() => decision.Apply(guard, () => setters++, record));
            Assert.AreEqual(0, setters);
            Assert.AreEqual(1, guards);
        }

        /// <summary>An uncertain setter failure preserves its HRESULT and the original host without replay.</summary>
        [TestMethod]
        public void PrivateAccessVisibilitySetterFailureIsUncertainAndCannotReplay()
        {
            var decision = new OfficeVbeFixture.PrivateAccessVisibilityDecision(false, false);
            var receipts = new List<IDictionary<string, object>>();
            var original = new COMException("Synthetic visibility refusal", unchecked((int)0x800A0997));
            int setters = 0;
            Action setter = () => { setters++; throw original; };
            Assert.AreSame(original, Assert.ThrowsException<COMException>(() => decision.Apply(() => { }, setter, receipts.Add)));
            Assert.AreEqual(1, setters);
            Assert.AreEqual("FAILED", receipts[1]["State"]);
            Assert.AreEqual("0x800A0997", receipts[1]["HResult"]);
            Assert.AreEqual(true, receipts[1]["SetterOutcomeUncertain"]);
            Assert.AreEqual(false, receipts[1]["AutomaticRetry"]);
            Assert.ThrowsException<InvalidOperationException>(() => decision.Apply(() => { }, setter, receipts.Add));
            Assert.AreEqual(1, setters);
        }

        /// <summary>Missing completion evidence never creates permission to repeat a completed or uncertain setter.</summary>
        [DataTestMethod, DataRow(true), DataRow(false)]
        public void PrivateAccessVisibilityCompletionEvidenceFailureCannotReplaySetter(bool setterFails)
        {
            var decision = new OfficeVbeFixture.PrivateAccessVisibilityDecision(false, false);
            var original = new COMException("Synthetic setter refusal", unchecked((int)0x800A0997));
            var evidence = new InvalidOperationException("Synthetic completion evidence refusal");
            int setters = 0;
            Action setter = () => { setters++; if (setterFails) throw original; };
            Action<IDictionary<string, object>> record = row =>
            {
                if (!Equals(row["State"], "PENDING")) throw evidence;
            };
            if (setterFails)
            {
                var aggregate = Assert.ThrowsException<AggregateException>(() => decision.Apply(() => { }, setter, record));
                Assert.AreSame(original, aggregate.InnerExceptions[0]);
                Assert.AreSame(evidence, aggregate.InnerExceptions[1]);
            }
            else
                Assert.AreSame(evidence, Assert.ThrowsException<InvalidOperationException>(() => decision.Apply(() => { }, setter, record)));
            Assert.ThrowsException<InvalidOperationException>(() => decision.Apply(() => { }, setter, record));
            Assert.AreEqual(1, setters);
        }

        [TestMethod]
        public void PrivateOfficeBootstrapAcceptsOnlyReviewedLocalAccessOrPublisherImages()
        {
            string desktop = "VBAiTests_" + Guid.NewGuid().ToString("N");
            Assert.AreEqual(@"C:\Program Files\Microsoft Office\root\Office16\MSACCESS.EXE",
                OfficeVbeFixture.RequirePrivateOfficeExecutable("Access", desktop, @"C:\Program Files\Microsoft Office\root\Office16\MSACCESS.EXE"));
            Assert.AreEqual(@"C:\Program Files\Microsoft Office\root\Office16\MSPUB.EXE",
                OfficeVbeFixture.RequirePrivateOfficeExecutable("Publisher", desktop, @"C:\Program Files\Microsoft Office\root\Office16\MSPUB.EXE"));
        }

        [TestMethod]
        public void PrivateOfficeBootstrapRefusesInputDesktopUnsupportedHostsAndAmbiguousExecutablePaths()
        {
            string desktop = "VBAiTests_" + Guid.NewGuid().ToString("N");
            foreach (string invalidDesktop in new[] { null, "Default", "WinSta0\\Default", "VBAiTests_bad" })
                Assert.ThrowsException<ArgumentException>(() => OfficeVbeFixture.RequirePrivateOfficeExecutable("Access", invalidDesktop, @"C:\MSACCESS.EXE"));
            Assert.AreEqual(@"C:\POWERPNT.EXE", OfficeVbeFixture.RequirePrivateOfficeExecutable("PowerPoint", desktop, @"C:\POWERPNT.EXE"));
            Assert.ThrowsException<ArgumentException>(() => OfficeVbeFixture.RequirePrivateOfficeExecutable("PowerPoint", desktop, @"C:\MSACCESS.EXE"));
            foreach (string host in new[] { "Word", "Excel", null })
                Assert.ThrowsException<InvalidOperationException>(() => OfficeVbeFixture.RequirePrivateOfficeExecutable(host, desktop, @"C:\MSACCESS.EXE"));
            foreach (string invalidPath in new[] { null, "", "MSACCESS.EXE", @"C:MSACCESS.EXE", @"\MSACCESS.EXE", @"\\server\share\MSACCESS.EXE", @"C:\MSPUB.EXE", @"C:\MSACCESS.EXE.cmd" })
                Assert.ThrowsException<ArgumentException>(() => OfficeVbeFixture.RequirePrivateOfficeExecutable("Access", desktop, invalidPath));
        }
    }
}

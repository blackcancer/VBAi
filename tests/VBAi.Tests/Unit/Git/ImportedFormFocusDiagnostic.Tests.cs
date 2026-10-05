using System;
using System.Collections.Generic;
using System.Runtime.ExceptionServices;
using System.Threading;
using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace VBAi.Tests.Unit
{
    [TestClass, TestCategory("Unit")]
    public sealed class ImportedFormFocusDiagnosticTests
    {
        private sealed class Probe
        {
            internal readonly Dictionary<string, int> Reads = new Dictionary<string, int>();
            internal string Failed;
            internal bool Enabled = true, PreviousMatches;
            internal T Read<T>(string name, T value)
            {
                Reads[name] = Reads.ContainsKey(name) ? Reads[name] + 1 : 1;
                if (Failed == name) throw new InvalidOperationException("secret native details must not enter observation");
                return value;
            }
            internal ImportedFormFocusDiagnostic.Snapshot Observe(bool active = true, bool same = false,
                bool previous = true, bool root = true) => ImportedFormFocusDiagnostic.Observe(active, same, previous, root, 1, 0,
                    () => Read("type", 2), () => Read("handle", 31L), () => Read("active-caption", "Other code pane"),
                    () => Read("expected-caption", "Imported form"), () => Read("enabled", Enabled),
                    () => Read("previous", PreviousMatches));
        }

        [DataTestMethod]
        [DataRow(false, false)] [DataRow(false, true)] [DataRow(true, false)] [DataRow(true, true)]
        public void EnabledAndDisabledRootsDoNotChangeObservedComIdentityAndEveryGetterIsReadOnce(bool enabled, bool same)
        {
            Sta(() => {
                var p = new Probe { Enabled = enabled, PreviousMatches = true }; var value = p.Observe(same: same);
                Assert.AreEqual(same ? "expected" : "different", value.ActiveIdentity);
                Assert.AreEqual(enabled.ToString(), value.RootEnabled); Assert.AreEqual("True", value.ActiveMatchesPrevious);
                Assert.AreEqual(same ? "1" : "2", value.ActiveType); Assert.AreEqual(same ? "0" : "31", value.ActiveHandle);
                Assert.AreEqual(same ? "Imported form" : "Other code pane", value.ActiveCaption);
                Assert.AreEqual(same ? 3 : 6, p.Reads.Count);
                foreach (var read in p.Reads) Assert.AreEqual(1, read.Value, read.Key);
            });
        }

        [DataTestMethod, DataRow(false), DataRow(true)]
        public void NullActiveWindowNeverInvokesItsGettersOrPreviousIdentityComparison(bool previous)
        {
            Sta(() => {
                var p = new Probe(); var value = p.Observe(active: false, previous: previous);
                Assert.AreEqual("null", value.ActiveIdentity); Assert.AreEqual(previous, value.PreviousPresent);
                Assert.AreEqual("not-observed-null-active", value.ActiveType);
                Assert.AreEqual("not-observed-null-active", value.ActiveHandle);
                Assert.AreEqual("not-observed-null-active", value.ActiveCaption);
                Assert.AreEqual("not-observed-null-active", value.ActiveMatchesPrevious);
                CollectionAssert.AreEquivalent(new[] { "expected-caption", "enabled" }, new List<string>(p.Reads.Keys));
            });
        }

        [DataTestMethod, DataRow(false), DataRow(true)]
        public void PriorViewIsObservedWithoutChangingTheActiveIdentity(bool matches)
        {
            Sta(() => {
                var p = new Probe { PreviousMatches = matches }; var value = p.Observe();
                Assert.AreEqual(matches.ToString(), value.ActiveMatchesPrevious); Assert.AreEqual("different", value.ActiveIdentity);
                Assert.AreEqual(1, p.Reads["previous"]);
            });
        }

        [TestMethod]
        public void MissingPriorViewAndZeroRootAreNotQueriedOrReplacedWithInventedFalseValues()
        {
            Sta(() => {
                var p = new Probe(); var value = p.Observe(previous: false, root: false);
                Assert.AreEqual("not-observed-no-previous", value.ActiveMatchesPrevious);
                Assert.AreEqual("not-observed-zero-root", value.RootEnabled);
                Assert.IsFalse(p.Reads.ContainsKey("previous")); Assert.IsFalse(p.Reads.ContainsKey("enabled"));
            });
        }

        [DataTestMethod]
        [DataRow("type")] [DataRow("handle")] [DataRow("active-caption")]
        [DataRow("expected-caption")] [DataRow("enabled")] [DataRow("previous")]
        public void UnavailableGetterIsNeverRetriedAndDoesNotSuppressOtherObservations(string failed)
        {
            Sta(() => {
                var p = new Probe { Failed = failed }; var value = p.Observe();
                string text = ImportedFormFocusDiagnostic.Format(value);
                StringAssert.Contains(text, "unavailable(InvalidOperationException:0x80131509)");
                Assert.IsFalse(text.Contains("secret native details")); Assert.AreEqual(6, p.Reads.Count);
                foreach (var read in p.Reads) Assert.AreEqual(1, read.Value, read.Key);
            });
        }

        [TestMethod]
        public void FailedCaptionOfTheSameComWindowIsNotReadAgainThroughItsActiveAlias()
        {
            Sta(() => {
                var p = new Probe { Failed = "expected-caption" }; var value = p.Observe(same: true);
                Assert.AreEqual(value.ExpectedCaption, value.ActiveCaption);
                Assert.AreEqual(1, p.Reads["expected-caption"]); Assert.IsFalse(p.Reads.ContainsKey("active-caption"));
                StringAssert.Contains(value.ActiveCaption, "unavailable(");
            });
        }

        [TestMethod]
        public void MissingDiagnosticGetterIsReportedWithoutNativeFallback()
        {
            Sta(() => {
                var value = ImportedFormFocusDiagnostic.Observe(true, false, true, true, 1, 0, null, null, null, null, null, null);
                Assert.AreEqual("unavailable-missing-getter", value.ActiveType);
                Assert.AreEqual("unavailable-missing-getter", value.ActiveHandle);
                Assert.AreEqual("unavailable-missing-getter", value.ActiveCaption);
                Assert.AreEqual("unavailable-missing-getter", value.ExpectedCaption);
                Assert.AreEqual("unavailable-missing-getter", value.RootEnabled);
                Assert.AreEqual("unavailable-missing-getter", value.ActiveMatchesPrevious);
            });
        }

        [DataTestMethod]
        [DataRow("null")] [DataRow("controls")] [DataRow("long")] [DataRow("surrogate")] [DataRow("normal")]
        public void CaptionFormattingIsBoundedSingleLineAndCannotForgeFields(string kind)
        {
            string input = kind == "null" ? null : kind == "controls" ? "a\r\n\t\0;=|\u2028\u2029b" :
                kind == "long" ? new string('x', 600) : kind == "surrogate" ? new string('x', 127) + "\ud83d\ude00" : "Owned form (UserForm)";
            string result = ImportedFormFocusDiagnostic.Bounded(input);
            Assert.IsTrue(result.Length <= ImportedFormFocusDiagnostic.FieldLimit);
            foreach (char c in result) Assert.IsFalse(char.IsControl(c) || c == ';' || c == '=' || c == '|' || c == '\u2028' || c == '\u2029');
            if (kind == "null") Assert.AreEqual("<null>", result);
            if (kind == "long") Assert.AreEqual(128, result.Length);
            if (kind == "surrogate") Assert.AreEqual(127, result.Length);
            if (kind == "normal") Assert.AreEqual(input, result);
            var snapshot = new ImportedFormFocusDiagnostic.Snapshot { ActiveCaption = input, ExpectedCaption = input };
            string text = ImportedFormFocusDiagnostic.Format(snapshot);
            Assert.IsTrue(text.Length < 600); Assert.IsFalse(text.Contains("\n"));
        }

        [TestMethod]
        public void WrongApartmentCannotInvokeAnyAdditionalComOrWindowGetter()
        {
            Run(ApartmentState.MTA, () => {
                var p = new Probe(); var value = p.Observe();
                Assert.AreEqual("not-observed-non-STA", value.State); Assert.AreEqual(0, p.Reads.Count);
            });
        }

        [DataTestMethod, DataRow(0), DataRow(12)]
        public void PassingSelectorPerformsNoDiagnosticReadsAndKeepsItsExistingOwnerChecks(int designer)
        {
            Sta(() => {
                int owners = 0;
                IntPtr result = ImportedFormMaterialization.SelectObservedTarget("before-first-render", new IntPtr(11),
                    new IntPtr(designer), true, true, true, true, 1, 42, _ => { owners++; return 42; },
                    () => throw new AssertFailedException("No failure owner formatting on PASS"),
                    () => throw new AssertFailedException("No new observation on PASS"));
                Assert.AreEqual(new IntPtr(designer == 0 ? 11 : designer), result); Assert.AreEqual(2, owners);
            });
        }

        [DataTestMethod, DataRow("ownership"), DataRow("focus"), DataRow("format")]
        public void DiagnosticFailurePreservesTheExactOriginalRefusalWithoutReplayingOwners(string phase)
        {
            Sta(() => {
                var original = new InvalidOperationException("Original ownership refusal"); int owners = 0, diagnostics = 0;
                var thrown = Assert.ThrowsException<InvalidOperationException>(() => ImportedFormMaterialization.SelectObservedTarget(
                    "before-first-render", new IntPtr(11), IntPtr.Zero, true, true, true, true, 1, 42,
                    _ => { owners++; throw original; }, () => { if (phase == "ownership") throw new Exception("diagnostic"); return "observed"; },
                    () => { diagnostics++; if (phase == "format") return ImportedFormFocusDiagnostic.Format(null); throw new Exception("diagnostic"); }));
                Assert.AreSame(original, thrown); Assert.AreEqual(1, owners); Assert.AreEqual(phase == "ownership" ? 0 : 1, diagnostics);
            });
        }

        [DataTestMethod]
        [DataRow("before-first-render")] [DataRow("immediately-before-PrintWindow")] [DataRow("after-PrintWindow")]
        public void EqualCaptionAndDesignerTypeNeverSubstituteForExactComIdentity(string stage)
        {
            Sta(() => {
                int owners = 0, observations = 0;
                var thrown = Assert.ThrowsException<InvalidOperationException>(() => ImportedFormMaterialization.SelectObservedTarget(
                    stage, new IntPtr(11), IntPtr.Zero, true, false, true, true, 1, 42,
                    _ => { owners++; return 42; }, () => "not-evaluated", () => {
                        observations++;
                        return ImportedFormFocusDiagnostic.Format(ImportedFormFocusDiagnostic.Observe(true, false, false, true, 1, 0,
                            () => 1, () => 0L, () => "Same caption", () => "Same caption", () => false, null));
                    }));
                Assert.AreEqual(0, owners); Assert.AreEqual(1, observations);
                StringAssert.Contains(thrown.Message, "Stage=" + stage + ";");
                StringAssert.Contains(thrown.Message, "DesignerMatches=False"); StringAssert.Contains(thrown.Message, "ActiveIdentity=different");
                StringAssert.Contains(thrown.Message, "RootEnabled=False"); Assert.IsInstanceOfType(thrown.InnerException, typeof(InvalidOperationException));
            });
        }

        [TestMethod]
        public void OtherOriginalFailuresAreNotConvertedOrObservedAgain()
        {
            Sta(() => {
                var original = new System.Runtime.InteropServices.COMException("Native owner query failed");
                var thrown = Assert.ThrowsException<System.Runtime.InteropServices.COMException>(() => ImportedFormMaterialization.SelectObservedTarget(
                    "before-first-render", new IntPtr(11), IntPtr.Zero, true, true, true, true, 1, 42,
                    _ => throw original, () => throw new AssertFailedException("No diagnostic"), () => throw new AssertFailedException("No diagnostic")));
                Assert.AreSame(original, thrown);
            });
        }

        private static void Sta(Action action) => Run(ApartmentState.STA, action);
        private static void Run(ApartmentState apartment, Action action)
        {
            Exception failure = null;
            var thread = new Thread(() => { try { action(); } catch (Exception error) { failure = error; } }) { IsBackground = true };
            thread.SetApartmentState(apartment); thread.Start();
            Assert.IsTrue(thread.Join(TimeSpan.FromSeconds(5)), "The pure diagnostic mirror did not return.");
            if (failure != null) ExceptionDispatchInfo.Capture(failure).Throw();
        }
    }
}

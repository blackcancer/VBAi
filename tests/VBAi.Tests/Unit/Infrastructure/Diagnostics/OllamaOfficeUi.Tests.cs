using Microsoft.VisualStudio.TestTools.UnitTesting;
using System;
using System.Collections.Generic;
using System.Web.Script.Serialization;
using VBAi.Tests.Integration;

namespace VBAi.Tests.Unit
{
    /// <summary>Native guards reject foreign, hidden and replaced assistant ancestry without re-reading stale UIA providers.</summary>
    [TestClass, TestCategory("Unit")]
    public sealed class OllamaOfficeUiTests
    {
        [TestMethod]
        public void ComboFocusAndExpansionRecheckTheOriginalOwnerBeforeEachAction()
        {
            var events = new System.Collections.Generic.List<string>();
            OllamaOfficeUi.ExpandOwnedComboOnce(() => events.Add("owner"), () => events.Add("focus"),
                () => events.Add("expand"), events.Add);
            CollectionAssert.AreEqual(new[] { "owner", "ComboFocusIntentOnce", "focus", "owner", "ComboExpandIntentOnce", "expand" }, events);
        }

        [DataTestMethod]
        [DataRow(1)]
        [DataRow(2)]
        public void ObserverOwnerLossPreventsTheNextComboAction(int refusedGuard)
        {
            int guards = 0, focus = 0, expand = 0;
            var error = new InvalidOperationException("original observer ended");
            Assert.AreSame(error, Assert.ThrowsException<InvalidOperationException>(() =>
                OllamaOfficeUi.ExpandOwnedComboOnce(() => { if (++guards == refusedGuard) throw error; },
                    () => focus++, () => expand++, unused => { })));
            Assert.AreEqual(refusedGuard - 1, focus);
            Assert.AreEqual(0, expand);
        }

        [DataTestMethod]
        [DataRow("focus")]
        [DataRow("expand")]
        public void UncertainComboActionIsRetainedWithoutRepeatingOrContinuing(string failed)
        {
            int focuses = 0, expands = 0;
            var error = new InvalidOperationException("native result uncertain");
            Assert.AreSame(error, Assert.ThrowsException<InvalidOperationException>(() =>
                OllamaOfficeUi.ExpandOwnedComboOnce(() => { },
                    () => { focuses++; if (failed == "focus") throw error; },
                    () => { expands++; if (failed == "expand") throw error; }, unused => { })));
            Assert.AreEqual(1, focuses);
            Assert.AreEqual(failed == "expand" ? 1 : 0, expands);
        }

        [DataTestMethod]
        [DataRow("ComboFocusIntentOnce")]
        [DataRow("ComboExpandIntentOnce")]
        public void FailedIntentPublicationPreventsItsNativeComboAction(string refused)
        {
            int focuses = 0, expands = 0;
            Assert.ThrowsException<InvalidOperationException>(() => OllamaOfficeUi.ExpandOwnedComboOnce(
                () => { }, () => focuses++, () => expands++, phase => { if (phase == refused) throw new InvalidOperationException(); }));
            Assert.AreEqual(refused == "ComboExpandIntentOnce" ? 1 : 0, focuses);
            Assert.AreEqual(0, expands);
        }

        [DataTestMethod]
        [DataRow("guard")]
        [DataRow("focus")]
        [DataRow("expand")]
        [DataRow("record")]
        public void IncompleteComboActionPlanRunsNoNativeAction(string missing)
        {
            Action action = () => Assert.Fail("No native action may start.");
            Assert.ThrowsException<ArgumentNullException>(() => OllamaOfficeUi.ExpandOwnedComboOnce(
                missing == "guard" ? null : action, missing == "focus" ? null : action,
                missing == "expand" ? null : action, missing == "record" ? null : new Action<string>(unused => action())));
        }

        [DataTestMethod]
        [DataRow("scopePicker")]
        [DataRow("providerPicker")]
        [DataRow("modelPicker")]
        [DataRow("effortPicker")]
        [DataRow("modePicker")]
        [DataRow("approvalPicker")]
        public void ComboLookupInspectsEveryNativeComboAndSkipsUnrelatedUiaProviders(string id)
        {
            Assert.IsTrue(OllamaOfficeUi.FixedControlClassCanMatch(id, "WindowsForms10.COMBOBOX.app.0.abc"));
            foreach (string cls in new[] { null, "ComboLBox", "WindowsForms10.BUTTON.app.0.abc", "WindowsForms10.Window.8.app.0.abc", "WindowsForms10.EDIT.app.0.abc" })
                Assert.IsFalse(OllamaOfficeUi.FixedControlClassCanMatch(id, cls));
        }

        [DataTestMethod]
        [DataRow("modelSummary")]
        [DataRow("send")]
        [DataRow("stop")]
        [DataRow("refreshModels")]
        public void ButtonLookupAvoidsComboAndEditProvidersWithoutAdoptingAnotherWindowClass(string id)
        {
            Assert.IsTrue(OllamaOfficeUi.FixedControlClassCanMatch(id, "WindowsForms10.BUTTON.app.0.abc"));
            Assert.IsFalse(OllamaOfficeUi.FixedControlClassCanMatch(id, "WindowsForms10.COMBOBOX.app.0.abc"));
            Assert.IsFalse(OllamaOfficeUi.FixedControlClassCanMatch(id, "WindowsForms10.EDIT.app.0.abc"));
            Assert.IsFalse(OllamaOfficeUi.FixedControlClassCanMatch(id, "Button"));
        }

        [TestMethod]
        public void OtherFixedControlsRetainTheCompleteWinFormsIdentitySearch()
        {
            foreach (string id in new[] { "futureControl" })
            {
                Assert.IsTrue(OllamaOfficeUi.FixedControlClassCanMatch(id, "WindowsForms10.Window.8.app.0.abc"));
                Assert.IsTrue(OllamaOfficeUi.FixedControlClassCanMatch(id, "WindowsForms10.EDIT.app.0.abc"));
                Assert.IsFalse(OllamaOfficeUi.FixedControlClassCanMatch(id, "NativeForeignWindow"));
            }
        }

        [DataTestMethod]
        [DataRow("promptHost")]
        [DataRow("transcriptPanel")]
        public void PanelLookupRetainsEveryWindowPeerAndSkipsNonPanelProviders(string id)
        {
            foreach (string cls in new[] { "WindowsForms10.Window.8.app.0.abc", "WindowsForms10.Window.b.app.0.xyz" })
                Assert.IsTrue(OllamaOfficeUi.FixedControlClassCanMatch(id, cls));
            foreach (string cls in new[] { null, "WindowsForms10.EDIT.app.0.abc", "WindowsForms10.BUTTON.app.0.abc",
                "WindowsForms10.COMBOBOX.app.0.abc", "NativeForeignWindow", "WindowsForms10.WindowReplacement" })
                Assert.IsFalse(OllamaOfficeUi.FixedControlClassCanMatch(id, cls));
        }

        [DataTestMethod]
        [DataRow(false)]
        [DataRow(true)]
        public void ComposerAcceptsOnlyExactIndependentReadbackIncludingAnApiErrorAfterApplication(bool apiError)
        {
            int writes = 0, reads = 0;
            var receipts = new List<object>();
            OllamaOfficeUi.WriteComposerOnce("request", () => { }, () => ++reads == 1 ? "" : "request",
                text => { Assert.AreEqual("request", text); writes++; if (apiError) throw new InvalidOperationException(); }, receipts.Add);
            Assert.AreEqual(1, writes); Assert.AreEqual(2, reads);
            var serializer = new JavaScriptSerializer();
            var receipt = serializer.Deserialize<Dictionary<string, object>>(serializer.Serialize(receipts[2]));
            Assert.AreEqual("ComposerExactReadback", receipt["Phase"]);
            Assert.AreEqual(apiError, receipt["ApiErrorAfterApplication"]);
        }

        [DataTestMethod]
        [DataRow("")]
        [DataRow("different")]
        [DataRow("request\r\n")]
        [DataRow(null)]
        public void DivergentComposerReadbackRecordsLengthAndHashesWithoutTextOrWriteReplay(string observed)
        {
            int writes = 0, reads = 0;
            var receipts = new List<object>();
            Assert.ThrowsException<InvalidOperationException>(() => OllamaOfficeUi.WriteComposerOnce("request",
                () => { }, () => ++reads == 1 ? "" : observed, text => writes++, receipts.Add));
            Assert.AreEqual(1, writes); Assert.AreEqual(2, reads);
            var serializer = new JavaScriptSerializer();
            string json = serializer.Serialize(receipts[1]);
            var receipt = serializer.Deserialize<Dictionary<string, object>>(json);
            Assert.AreEqual("ComposerIndependentReadback", receipt["Phase"]);
            Assert.AreEqual(false, receipt["Exact"]);
            Assert.AreEqual(7, receipt["PromptLength"]);
            Assert.AreEqual(observed == null ? (object)null : observed.Length, receipt["ObservedLength"]);
            Assert.AreEqual(observed == null, receipt["ObservedIsNull"]);
            Assert.AreEqual(EditorDocument.Hash("request"), receipt["PromptSha256"]);
            Assert.AreEqual(observed == null ? null : EditorDocument.Hash(observed), receipt["ObservedSha256"]);
            Assert.IsFalse(json.Contains("\"request\""));
            if (!string.IsNullOrEmpty(observed)) Assert.IsFalse(json.Contains(serializer.Serialize(observed)));
        }

        [DataTestMethod]
        [DataRow("draft")]
        [DataRow(" ")]
        [DataRow(null)]
        public void ExistingOrUnverifiableDraftIsNeverOverwritten(string existing)
        {
            int reads = 0;
            Assert.ThrowsException<InvalidOperationException>(() => OllamaOfficeUi.WriteComposerOnce("request", () => { },
                () => { reads++; return existing; }, unused => Assert.Fail("No write expected."),
                unused => Assert.Fail("No intent expected.")));
            Assert.AreEqual(1, reads);
        }

        [TestMethod]
        public void ComposerWriteErrorIsPreservedWhenIndependentReadbackDoesNotConfirmApplication()
        {
            int writes = 0;
            var error = new InvalidOperationException("native write failed");
            var observed = Assert.ThrowsException<InvalidOperationException>(() => OllamaOfficeUi.WriteComposerOnce("request",
                () => { }, () => "", text => { writes++; throw error; }, unused => { }));
            Assert.AreSame(error, observed.InnerException); Assert.AreEqual(1, writes);
        }

        [TestMethod]
        public void FailedIndependentComposerReadRetainsItsErrorAndDoesNotRepeatTheWrite()
        {
            int reads = 0, writes = 0;
            var error = new InvalidOperationException("original editor replaced");
            var receipts = new List<object>();
            Assert.AreSame(error, Assert.ThrowsException<InvalidOperationException>(() => OllamaOfficeUi.WriteComposerOnce("request",
                () => { }, () => ++reads == 1 ? "" : throw error, text => writes++, receipts.Add)));
            Assert.AreEqual(1, writes); Assert.AreEqual(2, reads);
            var serializer = new JavaScriptSerializer();
            var receipt = serializer.Deserialize<Dictionary<string, object>>(serializer.Serialize(receipts[1]));
            Assert.AreEqual("ComposerReadbackFailed", receipt["Phase"]);
            Assert.AreEqual(error.HResult, receipt["ErrorHResult"]);
        }

        [DataTestMethod]
        [DataRow(1)]
        [DataRow(2)]
        [DataRow(3)]
        [DataRow(4)]
        [DataRow(5)]
        [DataRow(6)]
        public void ComposerOwnerLossAtEveryReadAndWriteBoundaryPreventsLaterEffects(int lostGuard)
        {
            int guards = 0, reads = 0, writes = 0;
            var error = new InvalidOperationException("owner ended");
            Assert.AreSame(error, Assert.ThrowsException<InvalidOperationException>(() => OllamaOfficeUi.WriteComposerOnce("request",
                () => { if (++guards == lostGuard) throw error; }, () => ++reads == 1 ? "" : "request", text => writes++, unused => { })));
            Assert.AreEqual(lostGuard >= 5 ? 1 : 0, writes);
            Assert.AreEqual(lostGuard == 1 ? 0 : lostGuard >= 6 ? 2 : 1, reads);
        }

        [TestMethod]
        public void FailedComposerIntentPublicationDoesNotWrite()
        {
            var error = new InvalidOperationException("receipt unavailable");
            Assert.AreSame(error, Assert.ThrowsException<InvalidOperationException>(() => OllamaOfficeUi.WriteComposerOnce("request",
                () => { }, () => "", unused => Assert.Fail("No write expected."), unused => { throw error; })));
        }

        [DataTestMethod]
        [DataRow("prompt")]
        [DataRow("owner")]
        [DataRow("read")]
        [DataRow("write")]
        [DataRow("record")]
        public void MissingComposerActionPlanRunsNoEffects(string missing)
        {
            Action effect = () => Assert.Fail("No effect expected.");
            Assert.ThrowsException<ArgumentNullException>(() => OllamaOfficeUi.WriteComposerOnce(missing == "prompt" ? null : "request",
                missing == "owner" ? null : effect, missing == "read" ? null : new Func<string>(() => { effect(); return ""; }),
                missing == "write" ? null : new Action<string>(unused => effect()),
                missing == "record" ? null : new Action<object>(unused => effect())));
        }

        [DataTestMethod][DataRow(true)][DataRow(false)]
        public void OwnedTranscriptRowMustStillBelongToThePanelToSupplyOracleText(bool stillChild)
            => Assert.AreEqual(stillChild, OllamaOfficeUi.TranscriptChildCanBeObserved(42, 7, 42, 7, stillChild));

        [DataTestMethod][DataRow(true)][DataRow(false)]
        public void DisappearingReadOnlyTranscriptRowsAreNotMistakenForForeignOwnership(bool previousChildState)
            => Assert.IsFalse(OllamaOfficeUi.TranscriptChildCanBeObserved(42, 7, 0, 0, previousChildState));

        [DataTestMethod][DataRow(43, 7)][DataRow(42, 8)][DataRow(0, 7)][DataRow(42, 0)]
        public void ForeignOrInconsistentTranscriptOwnershipStillFailsImmediately(int process, int thread)
            => Assert.ThrowsException<InvalidOperationException>(() =>
                OllamaOfficeUi.TranscriptChildCanBeObserved(42, 7, process, (uint)thread, true));

        [DataTestMethod][DataRow(0, 7)][DataRow(42, 0)]
        public void MissingPanelOwnershipCannotAdmitEvenAnOtherwiseMatchingRow(int process, int thread)
            => Assert.ThrowsException<InvalidOperationException>(() =>
                OllamaOfficeUi.TranscriptChildCanBeObserved(process, (uint)thread, process, (uint)thread, true));

        [DataTestMethod][DataRow("length")][DataRow("text")]
        public void TemporarilyUnavailableOwnedTranscriptSuppliesNoOracleUntilAnotherCompleteObservation(string phase)
        {
            long elapsed = 0; int observations = 0;
            OllamaOfficeUi.ObserveWithinDeadline(() =>
            {
                if (++observations == 1)
                {
                    OllamaOfficeUi.RequireDeferredTranscriptRead(1460, 42, 7, 42, 7, true, phase);
                    return false;
                }
                return true;
            }, 1000, () => elapsed, cost => elapsed += cost, "complete native transcript");
            Assert.AreEqual(2, observations);
        }

        [TestMethod]
        public void PermanentlyUnavailableTranscriptFailsTheOriginalDeadline()
        {
            long elapsed = 0; int observations = 0;
            Assert.ThrowsException<TimeoutException>(() => OllamaOfficeUi.ObserveWithinDeadline(() =>
            {
                observations++;
                OllamaOfficeUi.RequireDeferredTranscriptRead(1460, 42, 7, 42, 7, true, "text");
                elapsed += 500;
                return false;
            }, 1000, () => elapsed, cost => elapsed += cost, "native transcript"));
            Assert.IsTrue(observations > 0);
            Assert.IsTrue(elapsed >= 1000);
        }

        [DataTestMethod][DataRow(0)][DataRow(5)][DataRow(6)][DataRow(13)][DataRow(1400)]
        public void NonTimeoutErrorsOnLiveTranscriptRemainFailures(int nativeError)
            => Assert.ThrowsException<InvalidOperationException>(() =>
                OllamaOfficeUi.RequireDeferredTranscriptRead(nativeError, 42, 7, 42, 7, true, "length"));

        [DataTestMethod][DataRow(43, 7)][DataRow(42, 8)][DataRow(0, 7)][DataRow(42, 0)]
        public void UnavailableTranscriptNeverDefersForeignOrInconsistentOwnership(int process, int thread)
            => Assert.ThrowsException<InvalidOperationException>(() =>
                OllamaOfficeUi.RequireDeferredTranscriptRead(1460, 42, 7, process, (uint)thread, true, "text"));

        [DataTestMethod][DataRow(0, 0, true)][DataRow(42, 7, false)]
        public void RecycledTranscriptReadCannotSupplyPartialOracleText(int process, int thread, bool stillChild)
            => OllamaOfficeUi.RequireDeferredTranscriptRead(5, 42, 7, process, (uint)thread, stillChild, "length");

        [TestMethod]
        public void NativeLookupRequiresOneNonzeroOriginalHandleBeforeReturningAPeer()
        {
            Assert.AreEqual(42L, OllamaOfficeUi.RequireUniqueFixedWindow("send", new[] { 42L }));
            foreach (var handles in new[] { null, new long[0], new[] { 0L }, new[] { 42L, 43L }, new[] { 42L, 42L } })
                Assert.ThrowsException<InvalidOperationException>(() => OllamaOfficeUi.RequireUniqueFixedWindow("send", handles));
            Assert.ThrowsException<InvalidOperationException>(() => OllamaOfficeUi.RequireUniqueFixedWindow(null, new[] { 42L }));
            Assert.ThrowsException<InvalidOperationException>(() => OllamaOfficeUi.RequireUniqueFixedWindow("", new[] { 42L }));
        }

        /// <summary>A peer with a familiar ID must still retain the exact original native ownership and visibility.</summary>
        [DataTestMethod]
        [DataRow("original")]
        [DataRow("handle")]
        [DataRow("id")]
        [DataRow("expectedPid")]
        [DataRow("expectedThread")]
        [DataRow("uiPid")]
        [DataRow("nativePid")]
        [DataRow("thread")]
        [DataRow("parent")]
        [DataRow("hidden")]
        [DataRow("none")]
        public void FixedAssistantControlsPreserveEveryOriginalIdentityGuard(string change)
        {
            Assert.AreEqual(change == "none", OllamaOfficeUi.FixedControlIdentityMatches(
                change == "original" ? 0 : 101, change == "handle" ? 102 : 101, "send",
                change == "id" ? "other" : "send", change == "expectedPid" ? 0 : 42,
                change == "expectedThread" ? 0u : 7u, change == "uiPid" ? 43 : 42,
                change == "nativePid" ? 43 : 42, change == "thread" ? 8u : 7u,
                change != "parent", change != "hidden"));
        }

        [DataTestMethod]
        [DataRow(0, 11, false)]
        [DataRow(11, 12, false)]
        [DataRow(11, 11, true)]
        public void PendingOrForeignObserverRetainsItsReferences(int original, int current, bool pending)
            => Assert.ThrowsException<InvalidOperationException>(() =>
                OllamaOfficeUi.RequireReferenceRelease(original, current, pending));

        [TestMethod]
        public void OriginalSettledObserverCanDropItsManagedReferences()
            => OllamaOfficeUi.RequireReferenceRelease(11, 11, false);

        [TestMethod]
        public void ObserverOwnerRefusalPrecedesAnyNativeWindowOrDesktopRead()
        {
            var refusal = new InvalidOperationException("Original parent ended");
            var ui = new OllamaOfficeUi(42, unused => Assert.Fail("No native observation expected."), () => { throw refusal; });
            var guard = typeof(OllamaOfficeUi).GetMethod("Guard", System.Reflection.BindingFlags.Instance | System.Reflection.BindingFlags.NonPublic);
            var error = Assert.ThrowsException<System.Reflection.TargetInvocationException>(() => guard.Invoke(ui, null));
            Assert.AreSame(refusal, error.InnerException);
        }

        [DataTestMethod]
        [DataRow(true, 30001L)][DataRow(false, 30001L)][DataRow(false, 30000L)]
        public void LateOrUnsuccessfulDeadlineReadCannotPassOrBeRepeated(bool result, long readCost)
        {
            long clock = 0; int reads = 0, pauses = 0;
            Assert.ThrowsException<TimeoutException>(() => OllamaOfficeUi.ObserveWithinDeadline(
                () => { reads++; clock = readCost; return result; }, 30000, () => clock, unused => pauses++, "catalogue"));
            Assert.AreEqual(1, reads); Assert.AreEqual(0, pauses);
        }

        [TestMethod]
        public void ExpiredWindowDoesNotStartAnotherNativeRead()
        {
            int reads = 0;
            Assert.ThrowsException<TimeoutException>(() => OllamaOfficeUi.ObserveWithinDeadline(
                () => { reads++; return true; }, 30000, () => 30001, unused => Assert.Fail("No pause expected."), "scope"));
            Assert.AreEqual(0, reads);
        }

        [DataTestMethod][DataRow(0L)][DataRow(30000L)]
        public void SuccessObservedInsideTheWindowIsAccepted(long readCost)
        {
            long clock = 0; int reads = 0;
            OllamaOfficeUi.ObserveWithinDeadline(() => { reads++; clock = readCost; return true; },
                30000, () => clock, unused => Assert.Fail("No pause expected."), "scope");
            Assert.AreEqual(1, reads);
        }

        [TestMethod]
        public void PollingPauseUsesOnlyTheRemainingWindow()
        {
            long clock = 0; int reads = 0, lastPause = -1;
            OllamaOfficeUi.ObserveWithinDeadline(() => ++reads == 2, 17, () => clock,
                delay => { lastPause = delay; clock += delay; }, "ready");
            Assert.AreEqual(17, lastPause); Assert.AreEqual(2, reads);
        }

        [DataTestMethod][DataRow(0L)][DataRow(-1L)]
        public void InvalidObservationWindowStartsNoNativeRead(long limit)
            => Assert.ThrowsException<ArgumentOutOfRangeException>(() => OllamaOfficeUi.ObserveWithinDeadline(
                () => { Assert.Fail("No native read expected."); return false; }, limit, () => 0, unused => { }, "ready"));

        [DataTestMethod]
        [DataRow("pid")]
        [DataRow("thread")]
        [DataRow("sitePid")]
        [DataRow("siteThread")]
        [DataRow("containerClass")]
        [DataRow("siteClass")]
        [DataRow("caption")]
        [DataRow("containerHidden")]
        [DataRow("siteHidden")]
        [DataRow("parent")]
        [DataRow("owner")]
        public void OriginalNativeHostAncestryRejectsEveryChangedIdentity(string changed)
        {
            Assert.IsFalse(OllamaOfficeUi.HostedNodesMatch(42, 7,
                changed == "pid" ? 43 : 42, changed == "thread" ? 8u : 7u,
                changed == "sitePid" ? 43 : 42, changed == "siteThread" ? 8u : 7u,
                changed == "containerClass" ? "Other" : "WindowsForms10.Window",
                changed == "siteClass" ? "GenericPaneOther" : "GenericPane",
                changed == "caption" ? "Other" : "VBAi",
                changed != "containerHidden", changed != "siteHidden", changed != "parent", changed != "owner"));
        }

        [TestMethod]
        public void ExactOriginalNativeAssistantAncestryIsAdmitted()
            => Assert.IsTrue(OllamaOfficeUi.HostedNodesMatch(42, 7, 42, 7, 42, 7,
                "WindowsForms10.Window", "GenericPane", "VBAi", true, true, true, true));
    }
}

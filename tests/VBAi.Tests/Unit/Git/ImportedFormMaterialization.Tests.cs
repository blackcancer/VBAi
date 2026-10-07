using Microsoft.VisualStudio.TestTools.UnitTesting;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.ExceptionServices;
using System.Threading;

namespace VBAi.Tests.Unit
{
    [TestClass, TestCategory("Unit")]
    public sealed class ImportedFormMaterializationTests
    {
        [DataTestMethod, DataRow(true), DataRow(false)]
        public void NoDeclaredFontPlanPerformsNoExportInitializationOrFontDelivery(bool absent)
        {
            var bindings = absent ? null : new FormStreamPadding.FormFontBinding[0];
            var result = ImportedFormMaterialization.Prepare(null, null, bindings,
                () => throw new AssertFailedException("Unexpected export"),
                _ => throw new AssertFailedException("Unexpected initialization"),
                () => throw new AssertFailedException("Unexpected native access"));
            Assert.AreSame(bindings, result);
        }

        [DataTestMethod, DataRow(false), DataRow(true)]
        public void ExactImportIncludingKnownPaddingDriftKeepsOriginalFontsWithoutActivation(bool padding)
        {
            var target = Snapshot(FormStreamPaddingTests.ContainerResourceBefore());
            var actual = Snapshot(padding ? FormStreamPaddingTests.ContainerResourceAfter() : target.Files["Form1.frx"]);
            byte[] saved = (byte[])actual.Files["Form1.frx"].Clone();
            int captures = 0;
            var result = ImportedFormMaterialization.Prepare(target, target.Manifest.Components[0], Bindings(target),
                () => { captures++; return actual; }, _ => throw new AssertFailedException("Exact imported resources must remain untouched"), () => { });
            Assert.AreEqual(0, result.Length); Assert.AreEqual(1, captures);
            CollectionAssert.AreEqual(saved, actual.Files["Form1.frx"]);
        }

        [TestMethod]
        public void MaterializedExactFrameRetainsIts827DescriptorWithoutAnyFontDelivery()
        {
            var target = Snapshot(FormStreamPaddingTests.ContainerResourceBefore());
            var before = ChangedFrame(target);
            int captures = 0, initialized = 0;
            var result = ImportedFormMaterialization.Prepare(target, target.Manifest.Components[0], Bindings(target),
                () => ++captures == 1 ? before : target, _ => { initialized++; return ImportedFormMaterialization.MaterializationOutcome.Rendered; }, () => { });
            Assert.AreEqual(1, initialized); Assert.AreEqual(2, captures); Assert.AreEqual(0, result.Length);
            Assert.AreEqual(82700u, BitConverter.ToUInt32(Bindings(target).Single(item => item.Type == 14).Descriptor, 6));
            Assert.IsFalse(target.SameFile(before, "Form1.frx"), "The strict font discrepancy must remain significant");
        }

        [TestMethod]
        public void PersistentFrameMismatchOffersOnlyThatOwnerAndPreservesCorrectRootFont()
        {
            var target = Snapshot(FormStreamPaddingTests.ContainerResourceBefore());
            var actual = ChangedFrame(target);
            var bindings = Bindings(target);
            var result = ImportedFormMaterialization.Prepare(target, target.Manifest.Components[0], bindings,
                () => actual, _ => ImportedFormMaterialization.MaterializationOutcome.Rendered, () => { });
            Assert.AreEqual(1, result.Length); Assert.AreSame(bindings.Single(item => item.Type == 14), result[0]);
            Assert.AreEqual("Controls/QualificationExtra", result[0].OwnerPath);
            Assert.AreEqual(82700u, BitConverter.ToUInt32(result[0].Descriptor, 6));
        }

        [TestMethod]
        public void NonFontResourceDiscrepancyDoesNotOverwriteAlreadyCorrectFontsOrBecomeAccepted()
        {
            var target = Snapshot(FormStreamPaddingTests.ContainerResourceBefore());
            var actual = Snapshot(target.Files["Form1.frx"].Concat(new byte[] { 71 }).ToArray());
            var result = ImportedFormMaterialization.Prepare(target, target.Manifest.Components[0], Bindings(target),
                () => actual, _ => ImportedFormMaterialization.MaterializationOutcome.Rendered, () => { });
            Assert.AreEqual(0, result.Length);
            Assert.IsFalse(target.SameFile(actual, "Form1.frx")); Assert.IsFalse(target.SameAs(actual));
        }

        [TestMethod]
        public void UnknownObservedResourceGraphRefusesBeforePlanningAnyOwnerDelivery()
        {
            var target = Snapshot(FormStreamPaddingTests.ContainerResourceBefore());
            byte[] bytes = (byte[])target.Files["Form1.frx"].Clone(); bytes[3048] ^= 1;
            var actual = Snapshot(bytes); int initialized = 0;
            Assert.ThrowsException<InvalidOperationException>(() => ImportedFormMaterialization.Prepare(target,
                target.Manifest.Components[0], Bindings(target), () => actual, _ => { initialized++; return ImportedFormMaterialization.MaterializationOutcome.Rendered; }, () => { }));
            Assert.AreEqual(1, initialized);
        }

        [DataTestMethod, DataRow(0), DataRow(1), DataRow(2)]
        public void ProjectIdentityFailureStopsAtItsBoundaryWithoutInitializationOrExportReplay(int boundary)
        {
            var target = Snapshot(FormStreamPaddingTests.ContainerResourceBefore()); var actual = ChangedFrame(target);
            var failure = new InvalidOperationException("Owned project identity changed");
            int validations = 0, captures = 0, initialized = 0;
            var thrown = Assert.ThrowsException<InvalidOperationException>(() => ImportedFormMaterialization.Prepare(target,
                target.Manifest.Components[0], Bindings(target), () => { captures++; return actual; }, _ => { initialized++; return ImportedFormMaterialization.MaterializationOutcome.Rendered; },
                () => { if (validations++ == boundary) throw failure; }));
            Assert.AreSame(failure, thrown);
            Assert.AreEqual(boundary == 0 ? 0 : 1, captures);
            Assert.AreEqual(boundary == 2 ? 1 : 0, initialized);
        }

        [TestMethod]
        public void FailedInitializationPreservesOriginalErrorAndNeverRecapturesOrRetries()
        {
            var target = Snapshot(FormStreamPaddingTests.ContainerResourceBefore()); var actual = ChangedFrame(target);
            var failure = new InvalidOperationException("Native initialization result uncertain"); int captures = 0, initialized = 0;
            var thrown = Assert.ThrowsException<InvalidOperationException>(() => ImportedFormMaterialization.Prepare(target,
                target.Manifest.Components[0], Bindings(target), () => { captures++; return actual; },
                _ => { initialized++; throw failure; }, () => { }));
            Assert.AreSame(failure, thrown); Assert.AreEqual(1, captures); Assert.AreEqual(1, initialized);
        }

        [DataTestMethod, DataRow(1), DataRow(2)]
        public void FailedExportIsNotRetriedAndCannotProduceAFontPlan(int failingCapture)
        {
            var target = Snapshot(FormStreamPaddingTests.ContainerResourceBefore()); var actual = ChangedFrame(target);
            var failure = new InvalidOperationException("Native export failed"); int captures = 0, initialized = 0;
            var thrown = Assert.ThrowsException<InvalidOperationException>(() => ImportedFormMaterialization.Prepare(target,
                target.Manifest.Components[0], Bindings(target), () => { if (++captures == failingCapture) throw failure; return actual; },
                _ => { initialized++; return ImportedFormMaterialization.MaterializationOutcome.Rendered; }, () => { }));
            Assert.AreSame(failure, thrown); Assert.AreEqual(failingCapture, captures); Assert.AreEqual(failingCapture - 1, initialized);
        }

        [DataTestMethod, DataRow(0), DataRow(302)]
        public void ExactOwnedDesignerAndVerifiedRootFallbackSelectOnlyTheOwningProcess(int designer)
        {
            Sta(() => Assert.AreEqual(new IntPtr(designer == 0 ? 301 : designer),
                ImportedFormMaterialization.SelectTarget(new IntPtr(301), new IntPtr(designer), true, true, true, 1, 7, _ => 7)));
        }

        [DataTestMethod]
        [DataRow(false, true, true, 1, 301, 7, 7)]
        [DataRow(true, false, true, 1, 301, 7, 7)]
        [DataRow(true, true, false, 1, 301, 7, 7)]
        [DataRow(true, true, true, 0, 301, 7, 7)]
        [DataRow(true, true, true, 1, 0, 7, 7)]
        [DataRow(true, true, true, 1, 301, 0, 7)]
        [DataRow(true, true, true, 1, 301, 7, 8)]
        public void ForeignHiddenInactiveOrUnidentifiedWindowNeverReachesNativeRendering(bool project, bool window,
            bool visible, int type, int root, int expected, int owner)
        {
            Sta(() => Assert.ThrowsException<InvalidOperationException>(() => ImportedFormMaterialization.SelectTarget(
                new IntPtr(root), new IntPtr(302), project, window, visible, type, (uint)expected, _ => (uint)owner)));
        }

        [TestMethod]
        public void ForeignChildCannotUseTheVerifiedRootOwnershipAsAnAlias()
        {
            Sta(() => Assert.ThrowsException<InvalidOperationException>(() => ImportedFormMaterialization.SelectTarget(
                new IntPtr(301), new IntPtr(302), true, true, true, 1, 7, pointer => pointer.ToInt64() == 301 ? 7u : 8u)));
        }

        [DataTestMethod, DataRow(false, true), DataRow(true, true), DataRow(false, false), DataRow(true, false)]
        public void NavigationRestoresThePriorFocusAndVisibilityOnlyAfterExpectedTransitions(bool visible, bool previous)
        {
            int focused = 0, hidden = 0;
            ImportedFormMaterialization.RestoreView(visible, previous, () => true, () => true,
                () => focused++, () => hidden++, () => { });
            Assert.AreEqual(previous ? 1 : 0, focused); Assert.AreEqual(visible ? 0 : 1, hidden);
        }

        [TestMethod]
        public void ConcurrentNavigationIsNeverOverriddenBeforeOrAfterPreviousFocus()
        {
            int focused = 0, hidden = 0;
            ImportedFormMaterialization.RestoreView(false, true, () => false, () => true,
                () => focused++, () => hidden++, () => { });
            Assert.AreEqual(0, focused); Assert.AreEqual(0, hidden);
            ImportedFormMaterialization.RestoreView(false, true, () => true, () => false,
                () => focused++, () => hidden++, () => { });
            Assert.AreEqual(1, focused); Assert.AreEqual(0, hidden);
        }

        [TestMethod]
        public void FailedFocusRestorationIsNotRetriedAndDoesNotHideAWindowWithUncertainNavigation()
        {
            var failure = new InvalidOperationException("Native focus result uncertain"); int focused = 0;
            var thrown = Assert.ThrowsException<InvalidOperationException>(() => ImportedFormMaterialization.RestoreView(
                false, true, () => true, () => true, () => { focused++; throw failure; },
                () => Assert.Fail("No hide after an uncertain focus operation"), () => { }));
            Assert.AreSame(failure, thrown); Assert.AreEqual(1, focused);
        }

        [TestMethod]
        public void RevokedProjectIdentityPreventsEvenNavigationReadbackDuringRestoration()
        {
            var failure = new InvalidOperationException("Project identity changed");
            var thrown = Assert.ThrowsException<InvalidOperationException>(() => ImportedFormMaterialization.RestoreView(
                false, true, () => throw new AssertFailedException("Unexpected view readback"), () => false,
                () => Assert.Fail("Unexpected focus"), () => Assert.Fail("Unexpected hide"), () => throw failure));
            Assert.AreSame(failure, thrown);
        }

        /// <summary>Separates both visibility failures without changing the guard or adding an ownership read.</summary>
        [TestMethod, DataRow(false, null), DataRow(true, false), DataRow(true, null)]
        public void DiagnosticRefusalSeparatesObservedVisibilityWithoutNativeRequery(bool main, bool? designer)
        {
            Sta(() =>
            {
                int reads = 0, descriptions = 0;
                var failure = Assert.ThrowsException<InvalidOperationException>(() => ImportedFormMaterialization.SelectObservedTarget(
                    "before-first-render", new IntPtr(11), IntPtr.Zero, true, true, main, designer, 1, 42,
                    _ => { reads++; return 42; }, () => { descriptions++; return "not-evaluated"; }));
                Assert.AreEqual(0, reads); Assert.AreEqual(1, descriptions);
                StringAssert.Contains(failure.Message, "MainVisible=" + main + ";DesignerVisible=" + (designer.HasValue ? designer.Value.ToString() : "not-evaluated"));
                StringAssert.Contains(failure.Message, "NativeOwners=not-evaluated");
                Assert.IsInstanceOfType(failure.InnerException, typeof(InvalidOperationException));
            });
        }

        /// <summary>Identifies every prepared render validation stage while preserving an inactive-window refusal.</summary>
        [TestMethod, DataRow("before-first-render"), DataRow("immediately-before-PrintWindow"), DataRow("after-PrintWindow")]
        public void DiagnosticInactiveDesignerKeepsTheStageAndRefusesBeforeNativeOwnership(string stage)
        {
            Sta(() =>
            {
                var failure = Assert.ThrowsException<InvalidOperationException>(() => ImportedFormMaterialization.SelectObservedTarget(
                    stage, new IntPtr(11), IntPtr.Zero, true, false, true, true, 1, 42,
                    _ => throw new AssertFailedException("Unexpected ownership read"), () => "not-evaluated"));
                StringAssert.Contains(failure.Message, "Stage=" + stage + ";Apartment=STA");
                StringAssert.Contains(failure.Message, "ProjectMatches=True;DesignerMatches=False");
            });
        }

        /// <summary>A foreign root reports only the ownership observation used in its original rejection.</summary>
        [TestMethod]
        public void DiagnosticForeignRootPreservesExactlyOneOwnerReadAndOriginalFailure()
        {
            Sta(() =>
            {
                int reads = 0;
                var failure = Assert.ThrowsException<InvalidOperationException>(() => ImportedFormMaterialization.SelectObservedTarget(
                    "before-first-render", new IntPtr(11), new IntPtr(12), true, true, true, true, 1, 42,
                    value => { reads++; Assert.AreEqual(new IntPtr(11), value); return 99; },
                    () => { Assert.AreEqual(1, reads); return "HWND=11,Pid=99,Thread=123,Error=0"; }));
                Assert.AreEqual(1, reads);
                StringAssert.Contains(failure.Message, "NativeOwners=HWND=11,Pid=99,Thread=123,Error=0");
                StringAssert.StartsWith(failure.InnerException.Message, "Only the exact active owned form designer");
            });
        }

        /// <summary>Successful selection keeps its established target and never evaluates failure-only diagnostics.</summary>
        [TestMethod]
        public void DiagnosticSuccessfulSelectionDoesNotCollectOrPublishARefusalDescription()
        {
            Sta(() => Assert.AreEqual(new IntPtr(12), ImportedFormMaterialization.SelectObservedTarget(
                "immediately-before-PrintWindow", new IntPtr(11), new IntPtr(12), true, true, true, true, 1, 42,
                _ => 42, () => throw new AssertFailedException("Unexpected refusal description"))));
        }

        /// <summary>Verifies production capture/show/probe/navigation callbacks and preserves the exact Frame descriptor.</summary>
        [DataTestMethod, DataRow(true, false), DataRow(false, false), DataRow(false, true)]
        public void ResourceProbeSettlesBeforeFocusAndRequiresDistinctCaptureAfterRestoration(bool earlyExact, bool residual)
        {
            var target = Snapshot(FormStreamPaddingTests.ContainerResourceBefore()); var different = ChangedFrame(target);
            var events = new List<string>(); int captures = 0;
            var result = ImportedFormMaterialization.Prepare(target, target.Manifest.Components[0], Bindings(target),
                () => { events.Add("capture " + ++captures); return captures == 1 || captures == 2 && !earlyExact || residual ? different : target; },
                probe =>
                {
                    bool designerActive = false;
                    var outcome = ImportedFormMaterialization.ShowAndInitialize(() => events.Add("show"), probe,
                        () => { designerActive = true; events.Add("focus"); }, () => events.Add("render"), () => { });
                    ImportedFormMaterialization.RestoreView(false, true, () => designerActive, () => !designerActive,
                        () => { designerActive = false; events.Add("restore focus"); }, () => events.Add("restore"), () => { },
                        outcome == ImportedFormMaterialization.MaterializationOutcome.ResourcesExactBeforeFocus);
                    return outcome;
                }, () => { });
            CollectionAssert.AreEqual(earlyExact ? new[] { "capture 1", "show", "capture 2", "restore", "capture 3" } :
                new[] { "capture 1", "show", "capture 2", "focus", "render", "restore focus", "restore", "capture 3" }, events);
            Assert.AreEqual(residual ? 1 : 0, result.Length);
            if (residual) Assert.AreEqual("Controls/QualificationExtra", result[0].OwnerPath);
            Assert.AreEqual(82700u, BitConverter.ToUInt32(Bindings(target).Single(value => value.Type == 14).Descriptor, 6));
        }

        /// <summary>Observed equality may never authorize later font repair when restoration changes resources.</summary>
        [DataTestMethod, DataRow(7), DataRow(14), DataRow(0)]
        public void ExactPrefocusResourcesThatDivergeAfterRestorationRefuseWithoutFontPlanOrAnotherInitialization(int changedOwner)
        {
            var target = Snapshot(FormStreamPaddingTests.ContainerResourceBefore());
            var final = changedOwner == 0 ? Snapshot(target.Files["Form1.frx"].Concat(new byte[] { 71 }).ToArray()) : ChangedOwner(target, changedOwner);
            int captures = 0, shows = 0;
            var failure = Assert.ThrowsException<InvalidOperationException>(() => ImportedFormMaterialization.Prepare(target,
                target.Manifest.Components[0], Bindings(target), () => ++captures == 2 ? target : captures == 1 ? ChangedFrame(target) : final,
                probe => ImportedFormMaterialization.ShowAndInitialize(() => shows++, probe,
                    () => Assert.Fail("No focus after exact resources"), () => Assert.Fail("No render after exact resources"), () => { }), () => { }));
            StringAssert.Contains(failure.Message, "font delivery is refused"); Assert.AreEqual(3, captures); Assert.AreEqual(1, shows);
        }

        /// <summary>An unrecognized outcome cannot reach capture or residual font planning.</summary>
        [TestMethod]
        public void UnknownInitializationOutcomeRefusesBeforePostRestorationCapture()
        {
            var target = Snapshot(FormStreamPaddingTests.ContainerResourceBefore()); int captures = 0;
            Assert.ThrowsException<InvalidOperationException>(() => ImportedFormMaterialization.Prepare(target, target.Manifest.Components[0],
                Bindings(target), () => { captures++; return ChangedFrame(target); }, _ => (ImportedFormMaterialization.MaterializationOutcome)0, () => { }));
            Assert.AreEqual(1, captures);
        }

        /// <summary>Native uncertainty is preserved regardless of whether a failed visibility setter already applied.</summary>
        [DataTestMethod, DataRow(false), DataRow(true)]
        public void VisibilityFailureCannotProbeFocusRenderOrBeReinterpretedAsSuccess(bool applied)
        {
            var failure = new InvalidOperationException("visibility uncertain"); int shows = 0; bool visible = false;
            var thrown = Assert.ThrowsException<InvalidOperationException>(() => ImportedFormMaterialization.ShowAndInitialize(
                () => { shows++; visible = applied; throw failure; }, () => throw new AssertFailedException("No probe after visibility failure"),
                () => Assert.Fail("No focus"), () => Assert.Fail("No render"), () => { }));
            Assert.AreSame(failure, thrown); Assert.AreEqual(1, shows); Assert.AreEqual(applied, visible);
        }

        /// <summary>A failed probe, focus or render stops the production sequence exactly once with its original exception.</summary>
        [DataTestMethod, DataRow(0), DataRow(1), DataRow(2)]
        public void ProbeFocusOrRenderFailureNeverReplaysOrAdvancesTheSequence(int boundary)
        {
            var failure = new InvalidOperationException("native " + boundary); var events = new List<string>();
            var thrown = Assert.ThrowsException<InvalidOperationException>(() => ImportedFormMaterialization.ShowAndInitialize(
                () => events.Add("show"), () => { events.Add("probe"); if (boundary == 0) throw failure; return false; },
                () => { events.Add("focus"); if (boundary == 1) throw failure; },
                () => { events.Add("render"); throw failure; }, () => { }));
            Assert.AreSame(failure, thrown);
            CollectionAssert.AreEqual(new[] { "show", "probe", "focus", "render" }.Take(boundary + 2).ToArray(), events);
        }

        /// <summary>Project revocation stops initialization before show, probe or focus on both exact and render paths.</summary>
        [DataTestMethod]
        [DataRow(0, true), DataRow(1, true), DataRow(2, true)]
        [DataRow(0, false), DataRow(1, false), DataRow(2, false)]
        public void InitializationGuardFailureStopsAtEveryNewBoundary(int boundary, bool exact)
        {
            var failure = new InvalidOperationException("identity changed"); int guards = 0; var events = new List<string>();
            var thrown = Assert.ThrowsException<InvalidOperationException>(() => ImportedFormMaterialization.ShowAndInitialize(
                () => events.Add("show"), () => { events.Add("probe"); return exact; }, () => events.Add("focus"), () => events.Add("render"),
                () => { if (guards++ == boundary) throw failure; }));
            Assert.AreSame(failure, thrown); Assert.AreEqual(boundary + 1, guards);
            CollectionAssert.AreEqual(new[] { "show", "probe" }.Take(boundary).ToArray(), events);
        }

        /// <summary>Each export failure is terminal and a failed pre-focus export performs no focus or rendering.</summary>
        [DataTestMethod, DataRow(1), DataRow(2), DataRow(3)]
        public void CaptureFailureAtInitialProbeOrRestoredViewIsNeverRetried(int failedCapture)
        {
            var target = Snapshot(FormStreamPaddingTests.ContainerResourceBefore()); int captures = 0, shows = 0;
            var failure = new InvalidOperationException("export failed");
            var thrown = Assert.ThrowsException<InvalidOperationException>(() => ImportedFormMaterialization.Prepare(target,
                target.Manifest.Components[0], Bindings(target), () =>
                {
                    if (++captures == failedCapture) throw failure; return captures == 1 ? ChangedFrame(target) : target;
                }, probe => ImportedFormMaterialization.ShowAndInitialize(() => shows++, probe,
                    () => Assert.Fail("The probe never returns nonexact"), () => Assert.Fail("No render"), () => { }), () => { }));
            Assert.AreSame(failure, thrown); Assert.AreEqual(failedCapture, captures); Assert.AreEqual(failedCapture == 1 ? 0 : 1, shows);
        }

        /// <summary>Only a settled exact probe and the still-current previous window allow an inactive designer to be hidden.</summary>
        [DataTestMethod]
        [DataRow(false, true, true, true, 1), DataRow(true, true, true, true, 0)]
        [DataRow(false, false, true, true, 0), DataRow(false, true, false, true, 0)]
        [DataRow(false, true, true, false, 0), DataRow(true, false, false, false, 0)]
        public void InactiveDesignerRestorationNeverStealsFocusAndRequiresTheExactSettledPreviousView(bool visible, bool previousDistinct,
            bool previousActive, bool exact, int expectedHides)
        {
            int hidden = 0;
            ImportedFormMaterialization.RestoreView(visible, previousDistinct, () => false, () => previousActive,
                () => Assert.Fail("No focus for an inactive designer"), () => hidden++, () => { }, exact);
            Assert.AreEqual(expectedHides, hidden);
        }

        /// <summary>Navigation changing during revalidation cancels the new visibility restoration.</summary>
        [DataTestMethod, DataRow(false), DataRow(true)]
        public void ConcurrentNavigationDuringExactInactiveRestorationPreventsHide(bool designerBecomesActive)
        {
            int checks = 0; bool changed = false;
            ImportedFormMaterialization.RestoreView(false, true, () => changed && designerBecomesActive,
                () => !changed, () => Assert.Fail("No focus"), () => Assert.Fail("No hide after concurrent navigation"),
                () => { if (++checks == 2) changed = true; }, true);
            Assert.AreEqual(2, checks);
        }

        /// <summary>Revocation at the final inactive-view guard forbids visibility mutation.</summary>
        [TestMethod]
        public void RevokedIdentityDuringExactInactiveRestorationPreservesItsFailureWithoutHide()
        {
            var failure = new InvalidOperationException("identity revoked"); int guards = 0;
            var thrown = Assert.ThrowsException<InvalidOperationException>(() => ImportedFormMaterialization.RestoreView(false, true,
                () => false, () => true, () => Assert.Fail("No focus"), () => Assert.Fail("No hide"),
                () => { if (++guards == 2) throw failure; }, true));
            Assert.AreSame(failure, thrown);
        }

        /// <summary>A failed hide is never attempted again even when it changed the designer visibility first.</summary>
        [DataTestMethod, DataRow(false), DataRow(true)]
        public void ExactInactiveHideFailureRemainsFatalWithoutRetry(bool applied)
        {
            var failure = new InvalidOperationException("hide uncertain"); int hides = 0; bool hidden = false;
            var thrown = Assert.ThrowsException<InvalidOperationException>(() => ImportedFormMaterialization.RestoreView(false, true,
                () => false, () => true, () => Assert.Fail("No focus"), () => { hides++; hidden = applied; throw failure; }, () => { }, true));
            Assert.AreSame(failure, thrown); Assert.AreEqual(1, hides); Assert.AreEqual(applied, hidden);
        }

        /// <summary>Project revocation inside the actual Prepare probe forbids export or any later initialization.</summary>
        [DataTestMethod, DataRow(4), DataRow(5)]
        public void PrefocusCaptureIdentityGuardsRefuseBeforeAndAfterTheProbeExport(int failedGuard)
        {
            var target = Snapshot(FormStreamPaddingTests.ContainerResourceBefore());
            var failure = new InvalidOperationException("probe identity changed"); int guards = 0, captures = 0, shows = 0;
            Action revalidate = () => { if (guards++ == failedGuard) throw failure; };
            var thrown = Assert.ThrowsException<InvalidOperationException>(() => ImportedFormMaterialization.Prepare(target,
                target.Manifest.Components[0], Bindings(target), () => { captures++; return captures == 1 ? ChangedFrame(target) : target; },
                probe => ImportedFormMaterialization.ShowAndInitialize(() => shows++, probe,
                    () => Assert.Fail("No focus after revoked probe identity"), () => Assert.Fail("No render"), revalidate), revalidate));
            Assert.AreSame(failure, thrown); Assert.AreEqual(failedGuard + 1, guards);
            Assert.AreEqual(failedGuard == 4 ? 1 : 2, captures); Assert.AreEqual(1, shows);
        }

        /// <summary>An unavailable navigation readback fails closed at each inactive-restoration query without mutation.</summary>
        [DataTestMethod, DataRow(0), DataRow(1), DataRow(2), DataRow(3)]
        public void ExactInactiveNavigationReadbackFailureCannotInventAHideOrFocusTarget(int failedRead)
        {
            var failure = new InvalidOperationException("active window unavailable"); int reads = 0;
            Func<bool, bool> read = result => { if (reads++ == failedRead) throw failure; return result; };
            var thrown = Assert.ThrowsException<InvalidOperationException>(() => ImportedFormMaterialization.RestoreView(false, true,
                () => read(false), () => read(true), () => Assert.Fail("No focus"), () => Assert.Fail("No hide"), () => { }, true));
            Assert.AreSame(failure, thrown); Assert.AreEqual(failedRead + 1, reads);
        }

        /// <summary>Changes one exact persisted font descriptor without rewriting the rest of the resource graph.</summary>
        private static VbaGitSnapshot ChangedOwner(VbaGitSnapshot target, int type)
        {
            CollectionAssert.AreEqual(FormStreamPaddingTests.ContainerResourceBefore(), target.Files["Form1.frx"]);
            return Snapshot(UserFormQualificationFontsTests.ResourceWithChangedDeclaredFont(type == 7 ? "Root" : "Frame", 9m));
        }

        /// <summary>Constructs a validated single-form snapshot from retained synthetic resources.</summary>
        private static VbaGitSnapshot Snapshot(byte[] resource)
        {
            return new VbaGitSnapshot(new VbaGitManifest
            {
                References = "",
                Components = new[] {
                new VbaGitComponent { Name = "Form1", Type = 3, HasResources = true } }
            }, new Dictionary<string, byte[]>
            {
                ["Form1.frm"] = VbaGitSnapshot.Utf8.GetBytes("VERSION 5.00\nBegin SyntheticForm\n OleObjectBlob = \"Form1.frx\":0000\nEnd\nAttribute VB_Name = \"Form1\"\n"),
                ["Form1.frx"] = (byte[])resource.Clone()
            });
        }

        /// <summary>Returns the exact parsed root and Frame descriptor plan.</summary>
        private static FormStreamPadding.FormFontBinding[] Bindings(VbaGitSnapshot snapshot)
        {
            return snapshot.FormFonts(snapshot.Manifest.Components[0]);
        }

        /// <summary>Creates the concrete 8.27 to 8.25 precision regression without weakening its comparison.</summary>
        private static VbaGitSnapshot ChangedFrame(VbaGitSnapshot target)
        {
            byte[] resource = (byte[])target.Files["Form1.frx"].Clone();
            byte[] descriptor = Bindings(target).Single(item => item.Type == 14).Descriptor;
            int position = Enumerable.Range(0, resource.Length - descriptor.Length + 1)
                .Single(index => resource.Skip(index).Take(descriptor.Length).SequenceEqual(descriptor));
            Array.Copy(BitConverter.GetBytes(82500u), 0, resource, position + 6, 4);
            return Snapshot(resource);
        }

        /// <summary>Runs only the pure ownership selector in an STA without creating a host or a window.</summary>
        private static void Sta(Action action)
        {
            Exception error = null;
            var thread = new Thread(() => { try { action(); } catch (Exception failure) { error = failure; } });
            thread.SetApartmentState(ApartmentState.STA); thread.Start();
            Assert.IsTrue(thread.Join(TimeSpan.FromSeconds(10)), "Managed ownership selector did not complete");
            if (error != null) ExceptionDispatchInfo.Capture(error).Throw();
        }
    }
}

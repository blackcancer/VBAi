using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.ExceptionServices;
using System.Threading;
using Microsoft.VisualStudio.TestTools.UnitTesting;

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
                () => throw new AssertFailedException("Unexpected initialization"),
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
                () => { captures++; return actual; }, () => Assert.Fail("Exact imported resources must remain untouched"), () => { });
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
                () => ++captures == 1 ? before : target, () => initialized++, () => { });
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
                () => actual, () => { }, () => { });
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
                () => actual, () => { }, () => { });
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
                target.Manifest.Components[0], Bindings(target), () => actual, () => initialized++, () => { }));
            Assert.AreEqual(1, initialized);
        }

        [DataTestMethod, DataRow(0), DataRow(1), DataRow(2)]
        public void ProjectIdentityFailureStopsAtItsBoundaryWithoutInitializationOrExportReplay(int boundary)
        {
            var target = Snapshot(FormStreamPaddingTests.ContainerResourceBefore()); var actual = ChangedFrame(target);
            var failure = new InvalidOperationException("Owned project identity changed");
            int validations = 0, captures = 0, initialized = 0;
            var thrown = Assert.ThrowsException<InvalidOperationException>(() => ImportedFormMaterialization.Prepare(target,
                target.Manifest.Components[0], Bindings(target), () => { captures++; return actual; }, () => initialized++,
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
                () => { initialized++; throw failure; }, () => { }));
            Assert.AreSame(failure, thrown); Assert.AreEqual(1, captures); Assert.AreEqual(1, initialized);
        }

        [DataTestMethod, DataRow(1), DataRow(2)]
        public void FailedExportIsNotRetriedAndCannotProduceAFontPlan(int failingCapture)
        {
            var target = Snapshot(FormStreamPaddingTests.ContainerResourceBefore()); var actual = ChangedFrame(target);
            var failure = new InvalidOperationException("Native export failed"); int captures = 0, initialized = 0;
            var thrown = Assert.ThrowsException<InvalidOperationException>(() => ImportedFormMaterialization.Prepare(target,
                target.Manifest.Components[0], Bindings(target), () => { if (++captures == failingCapture) throw failure; return actual; },
                () => initialized++, () => { }));
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

        /// <summary>Constructs a validated single-form snapshot from retained synthetic resources.</summary>
        private static VbaGitSnapshot Snapshot(byte[] resource)
        {
            return new VbaGitSnapshot(new VbaGitManifest { References = "", Components = new[] {
                new VbaGitComponent { Name = "Form1", Type = 3, HasResources = true } } }, new Dictionary<string, byte[]> {
                ["Form1.frm"] = VbaGitSnapshot.Utf8.GetBytes("VERSION 5.00\nBegin SyntheticForm\n OleObjectBlob = \"Form1.frx\":0000\nEnd\nAttribute VB_Name = \"Form1\"\n"),
                ["Form1.frx"] = (byte[])resource.Clone() });
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

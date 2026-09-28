namespace CodexVBE.Tests.Unit
{
    using System;
    using System.ComponentModel;
    using Microsoft.VisualStudio.TestTools.UnitTesting;
    using CodexVBE;

    // Matrix before execution: preview root/frame, inside-versus-outer decoration, scroll extents,
    // stale tree/non-design, path absent/ambiguous request, invalid padding, unknown container/child,
    // missing/read-only/scalar property, empty children, nonfinite/negative/oversize child geometry,
    // native setter failure, ignored writes, child movement or changed decoration => uncertain, no retry.
    [TestClass]
    [TestCategory("Unit")]
    public sealed partial class VbeFormsFitContentTests
    {
        [TestMethod]
        public void PreviewMeasuresOuterDecorationAndApplyVerifiesRootAndFrame()
        {
            foreach (bool nested in new[] { false, true })
            {
                var f = Create(nested); var request = f.Request();
                dynamic preview = f.Service.PreviewFitFormContent(request);
                Assert.IsTrue((bool)preview.ReadOnly);
                Assert.AreEqual(85d, (double)preview.Plan.WidthAfter);
                Assert.AreEqual(72d, (double)preview.Plan.HeightAfter);
                Assert.AreEqual(0, f.Writes);
                dynamic result = f.Service.ApplyFitFormContent(request);
                Assert.IsTrue((bool)result.Verified);
                Assert.AreEqual(85d, f.Width);
                Assert.AreEqual(72d, f.Height);
                Assert.AreEqual(2, f.Writes);
            }
        }

        [TestMethod]
        public void ScrollFitPreservesVisibleDimensionsAndContainsContent()
        {
            var f = Create(true); var request = f.Request("fit_scroll_extent");
            dynamic preview = f.Service.PreviewFitFormContent(request);
            Assert.AreEqual(290d, (double)preview.Plan.WidthAfter);
            Assert.AreEqual(170d, (double)preview.Plan.HeightAfter);
            dynamic result = f.Service.ApplyFitFormContent(request);
            Assert.IsTrue((bool)result.Verified);
            Assert.AreEqual(300d, f.Width); Assert.AreEqual(200d, f.Height);
            Assert.AreEqual(290d, f.ScrollWidth); Assert.AreEqual(170d, f.ScrollHeight);
        }

        [TestMethod]
        public void FitRejectsInvalidRequestsStaleTreeAndRuntimeModeBeforeWriting()
        {
            var f = Create();
            Assert.ThrowsException<ArgumentException>(() => f.Service.PreviewFitFormContent(null));
            Assert.ThrowsException<ArgumentException>(() => f.Service.ApplyFitFormContent(new Request()));
            for (int fault = 0; fault < 7; fault++)
            {
                f = Create(); var request = f.Request();
                if (fault == 0) request.Action = "runtime";
                if (fault == 1) request.Left = double.NaN;
                if (fault == 2) request.Top = -1;
                if (fault == 3) request.Left = 1001;
                if (fault == 4) { request.ControlPath = "Controls/Unknown"; request.ParentPath = "UserForm"; }
                if (fault == 5) request.ControlPath = "Controls/Unknown";
                if (fault == 6) request.ExpectedTreeVersion = "stale";
                if (fault < 6) Assert.ThrowsException<ArgumentException>(() => f.Service.ApplyFitFormContent(request));
                else Assert.ThrowsException<InvalidOperationException>(() => f.Service.ApplyFitFormContent(request));
                Assert.AreEqual(0, f.Writes);
            }
            f = Create(); var runtime = f.Request(); f.Project.Mode = 1;
            Assert.ThrowsException<InvalidOperationException>(() => f.Service.ApplyFitFormContent(runtime));
        }

        [TestMethod]
        public void FitRejectsUnknownEmptyUnavailableAndNonfiniteMeasurements()
        {
            for (int fault = 0; fault < 9; fault++)
            {
                var f = Create(true);
                if (fault == 0) f.Target.ClassName = "ThirdParty";
                if (fault == 1) f.Children[0].ClassName = "ThirdParty";
                if (fault == 2) f.Target.Controls = new object[0];
                if (fault == 3) f.ChildWidth = double.NaN;
                if (fault == 4) f.ChildTop = -1;
                if (fault == 5) f.ChildLeft = 32767;
                if (fault == 6) f.BorderX = -1;
                if (fault == 7) f.Target.Metadata = new PropertyDescriptorCollection(new PropertyDescriptor[0]);
                if (fault == 8) f.Width = f.BorderX;
                var request = f.Request();
                Assert.ThrowsException<InvalidOperationException>(() => f.Service.ApplyFitFormContent(request));
                Assert.AreEqual(0, f.Writes);
            }
        }

        [TestMethod]
        public void NativeSetterReadbackAndChildDriftReturnUncertaintyWithoutRetry()
        {
            for (int fault = 0; fault < 4; fault++)
            {
                var f = Create(); var request = f.Request();
                f.Fail = fault == 0; f.Ignore = fault == 1;
                f.ChangeChild = fault == 2; f.ChangeDecoration = fault == 3;
                dynamic result = f.Service.ApplyFitFormContent(request);
                Assert.IsTrue((bool)result.Uncertain);
                Assert.IsFalse((bool)result.Verified);
                Assert.IsFalse((bool)result.RetryAllowed);
                Assert.AreEqual(fault == 0 ? 1 : 2, f.Writes);
            }
        }
    }
}

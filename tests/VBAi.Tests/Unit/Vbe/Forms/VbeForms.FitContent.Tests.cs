namespace VBAi.Tests.Unit
{
    using System;
    using System.ComponentModel;
    using System.Linq;
    using Microsoft.VisualStudio.TestTools.UnitTesting;
    using VBAi;

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
        public void NativeScrollFitBypassesUnsafeDescriptorSetterAndStillVerifiesReadback()
        {
            var previous = VbeForms.NativeDesignerObject;
            var previousDpi = VbeForms.MeasureFitWindowDpi;
            try
            {
                foreach (bool nested in new[] { false, true })
                {
                    var f = Create(nested);
                    VbeForms.NativeDesignerObject = target => ReferenceEquals(target, f.Target);
                    VbeForms.MeasureFitWindowDpi = _ => 96;
                    int descriptorWrites = 0;
                    f.Target.Metadata = new PropertyDescriptorCollection(f.Target.Metadata.Cast<PropertyDescriptor>().Select(property =>
                        property.Name == "ScrollWidth" || property.Name == "ScrollHeight"
                        ? new VbeFormsCoverageTests.LiveProperty(property.Name, typeof(double), () => property.GetValue(f.Target), value => {
                            descriptorWrites++;
                            throw new InvalidOperationException("Unsafe COM descriptor setter must never run.");
                        }) : property).ToArray());
                    dynamic result = f.Service.ApplyFitFormContent(f.Request("fit_scroll_extent"));
                    Assert.IsTrue((bool)result.Verified, "Native scroll dimensions must use dispatch, then preserve the existing readback checks.");
                    Assert.AreEqual(0, descriptorWrites);
                    Assert.AreEqual(2, f.Writes);
                    Assert.AreEqual(290d, f.ScrollWidth); Assert.AreEqual(170d, f.ScrollHeight);
                    Assert.AreEqual(300d, f.Width); Assert.AreEqual(200d, f.Height);
                }
            }
            finally { VbeForms.NativeDesignerObject = previous; VbeForms.MeasureFitWindowDpi = previousDpi; }
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
            var wrongRoot = Create(); wrongRoot.Target.ClassName = "Frame";
            Assert.ThrowsException<InvalidOperationException>(() => wrongRoot.Service.ApplyFitFormContent(wrongRoot.Request()));
            Assert.AreEqual(0, wrongRoot.Writes);
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

        [STATestMethod]
        public void NativeDpiQuantizesContainerDimensionsAndNeverRoundsScrollExtents()
        {
            var nativeObject = VbeForms.NativeDesignerObject; var readDpi = VbeForms.MeasureFitWindowDpi;
            try
            {
                using (var window = new System.Windows.Forms.Form())
                {
                    uint actualDpi = readDpi(window.Handle); Assert.IsTrue(actualDpi >= 48 && actualDpi <= 768);
                    foreach (uint dpi in new uint[] { 0, 47, 48, 96, 144, 768, 769 }) foreach (string action in new[] { "fit_container", "fit_scroll_extent" })
                    {
                        var f = Create(); f.Host.MainWindow.HWnd = window.Handle.ToInt64();
                        VbeForms.NativeDesignerObject = target => ReferenceEquals(target, f.Target);
                        VbeForms.MeasureFitWindowDpi = handle => { Assert.AreEqual(window.Handle, handle); return dpi; };
                        var request = f.Request(action); request.ControlPath = "UserForm";
                        if (dpi < 48 || dpi > 768) { Assert.ThrowsException<InvalidOperationException>(() => f.Service.PreviewFitFormContent(request)); Assert.AreEqual(0, f.Writes); }
                        else
                        {
                            dynamic preview = f.Service.PreviewFitFormContent(request); double pixel = 72d / dpi;
                            Assert.AreEqual(pixel, (double)preview.Plan.TolerancePoints);
                            Assert.AreEqual(action == "fit_container" ? Math.Ceiling(85d / pixel) * pixel : 290d, (double)preview.Plan.WidthAfter, 0.0001);
                            dynamic result = f.Service.ApplyFitFormContent(request); Assert.IsTrue((bool)result.Verified);
                        }
                    }
                }
            }
            finally { VbeForms.NativeDesignerObject = nativeObject; VbeForms.MeasureFitWindowDpi = readDpi; }
        }

        [TestMethod]
        public void NativeComponentDimensionsExposeCompletePropertyDescriptorContracts()
        {
            var previous = VbeForms.NativeDesignerObject;
            try
            {
                VbeForms.NativeDesignerObject = target => true;
                var f = Create(); var properties = f.Target.Metadata.Cast<PropertyDescriptor>().Where(p => p.Name != "Width" && p.Name != "Height").ToArray();
                f.Target.Metadata = new PropertyDescriptorCollection(properties);
                var descriptor = (PropertyDescriptor)FitCall("RequireFitProperty", f.Target, "Width", true, f.Form);
                Assert.AreEqual(typeof(object), descriptor.ComponentType); Assert.AreEqual(typeof(double), descriptor.PropertyType); Assert.IsFalse(descriptor.IsReadOnly);
                Assert.AreEqual(300d, descriptor.GetValue(f.Target)); descriptor.SetValue(f.Target, 123d); Assert.AreEqual(123d, f.Form.Properties.Item("Width").Stored);
                FitCall("SetFitProperty", f.Target, descriptor, 124d);
                Assert.AreEqual(124d, f.Form.Properties.Item("Width").Stored, "Root dimensions must remain on VBIDE properties, not Designer.Width.");
                Assert.AreEqual(0, f.Writes);
                Assert.IsFalse(descriptor.CanResetValue(f.Target)); Assert.IsFalse(descriptor.ShouldSerializeValue(f.Target)); Assert.ThrowsException<NotSupportedException>(() => descriptor.ResetValue(f.Target));
                foreach (string fault in new[] { "absent", "indexed", "null", "type", "left" })
                {
                    f = Create(); f.Target.Metadata = new PropertyDescriptorCollection(new PropertyDescriptor[0]);
                    if (fault == "absent") f.Form.Properties.Remove(f.Form.Properties.Item("Width"));
                    if (fault == "indexed") f.Form.Properties.Item("Width").Indices = 1;
                    if (fault == "null") f.Form.Properties.Item("Width").Stored = null;
                    if (fault == "type") f.Form.Properties.Item("Width").Stored = "invalid";
                    Assert.ThrowsException<InvalidOperationException>(() => FitCall("RequireFitProperty", f.Target, fault == "left" ? "Left" : "Width", true, f.Form), fault);
                }
                var root = Create();
                foreach (string property in new[] { "Width", "Height" })
                {
                    var fallback = (PropertyDescriptor)FitCall("RequireFitProperty", new VbeFormsCoverageTests.Node(), property, true, root.Form);
                    Assert.IsNotNull(fallback);
                }
            }
            finally { VbeForms.NativeDesignerObject = previous; }
        }

        [TestMethod]
        public void NumericAndControlsContractsRejectNullNonNumericReadonlyUnboundedAndAmbiguousData()
        {
            foreach (object value in new object[] { null, -1d, 32768d, double.NaN, double.PositiveInfinity })
            {
                var property = new VbeFormsCoverageTests.LiveProperty("Width", typeof(double), () => value);
                Assert.ThrowsException<InvalidOperationException>(() => FitCall("ReadFitNumber", new object(), property));
            }
            foreach (Type type in new[] { typeof(float), typeof(double), typeof(decimal), typeof(int), typeof(short) })
            {
                var node = new VbeFormsCoverageTests.Node(); node.Metadata = new PropertyDescriptorCollection(new[] { new VbeFormsCoverageTests.LiveProperty("Width", type, () => Convert.ChangeType(100, type), value => { }) });
                Assert.IsNotNull(FitCall("RequireFitProperty", node, "Width", true, null));
            }
            foreach (string fault in new[] { "readonly", "type", "missing-controls", "non-enumerable", "duplicate", "count" })
            {
                var f = Create();
                if (fault == "readonly" || fault == "type")
                {
                    var property = fault == "readonly" ? new VbeFormsCoverageTests.LiveProperty("Width", typeof(double), () => 10d) : new VbeFormsCoverageTests.LiveProperty("Width", typeof(string), () => "10", value => { });
                    f.Target.Metadata = new PropertyDescriptorCollection(new[] { property });
                    Assert.ThrowsException<InvalidOperationException>(() => FitCall("RequireFitProperty", f.Target, "Width", true, null));
                }
                else
                {
                    if (fault == "missing-controls") f.Target.Metadata = new PropertyDescriptorCollection(new PropertyDescriptor[0]);
                    if (fault == "non-enumerable") f.Target.Controls = new object();
                    if (fault == "duplicate") f.Target.Controls = new[] { f.Children[0], f.Children[0] };
                    if (fault == "count") f.Target.Controls = Enumerable.Range(0, 513).Select(i => new VbeFormsCoverageTests.Node { Name = "Label" + i, ClassName = "Label", Parent = f.Target, Metadata = f.Children[0].Metadata }).ToArray();
                    Assert.ThrowsException<InvalidOperationException>(() => FitCall("ReadFitChildren", f.Target, f.Form.Name), fault);
                }
            }
            var foreign = Create(); foreign.Children[0].Parent = new VbeFormsCoverageTests.Node { Name = "Other" };
            Assert.AreEqual(0, ((System.Collections.IList)FitCall("ReadFitChildren", foreign.Target, foreign.Form.Name)).Count);
        }

        [TestMethod]
        public void FitCoversEveryPaddingViewportExtentAndReadbackBoundary()
        {
            foreach (string fault in new[] { "top-nan", "left-negative", "top-high", "inside-height", "outer-height", "zero-width", "zero-height", "height-high" })
            {
                var f = Create(); var request = f.Request();
                if (fault == "top-nan") request.Top = double.NaN;
                if (fault == "left-negative") request.Left = -1;
                if (fault == "top-high") request.Top = 1001;
                if (fault == "inside-height") f.Height = f.BorderY;
                if (fault == "outer-height") f.BorderY = -1;
                if (fault == "zero-width") { f.ChildLeft = f.ChildWidth = f.BorderX = 0; request.Left = 0; }
                if (fault == "zero-height") { f.ChildTop = f.ChildHeight = f.BorderY = 0; request.Top = 0; }
                if (fault == "height-high") f.ChildTop = 32767;
                if (fault.StartsWith("top-") || fault == "left-negative") Assert.ThrowsException<ArgumentException>(() => f.Service.ApplyFitFormContent(request), fault);
                else Assert.ThrowsException<InvalidOperationException>(() => f.Service.ApplyFitFormContent(request), fault);
                Assert.AreEqual(0, f.Writes);
            }
            foreach (string fault in new[] { "ignored-height", "inside-small-y", "inside-large-x", "inside-large-y", "scroll-large-x", "scroll-large-y" })
            {
                var f = Create(); var request = f.Request(fault.StartsWith("scroll") ? "fit_scroll_extent" : "fit_container");
                f.AfterWrite = () => { if (f.Writes != 2) return;
                    if (fault == "ignored-height") f.Height = 200;
                    if (fault == "inside-small-y") f.BorderY += 1;
                    if (fault == "inside-large-x") f.BorderX -= 1;
                    if (fault == "inside-large-y") f.BorderY -= 1;
                    if (fault == "scroll-large-x") f.Width += 1;
                    if (fault == "scroll-large-y") f.Height += 1;
                };
                dynamic result = f.Service.ApplyFitFormContent(request); Assert.IsTrue((bool)result.Uncertain, fault); Assert.AreEqual(2, f.Writes); Assert.IsFalse((bool)result.RetryAllowed);
            }
            var parent = Create(true); var valid = parent.Request(); valid.ParentPath = valid.ControlPath; valid.ControlPath = null;
            Assert.IsTrue((bool)((dynamic)parent.Service.PreviewFitFormContent(valid)).ReadOnly);
        }
    }
}

namespace VBAi.Tests.Unit
{
    using Microsoft.VisualStudio.TestTools.UnitTesting;
    using System;
    using VBAi;
    [TestClass, TestCategory("Unit"), DoNotParallelize]
    public sealed class VbeFormLayoutWorkflowTests
    {
        private static Request Layout(FormsWorkflowFixture f) { var r = f.Request("align_left"); r.Items = new[] { "Controls/A", "Controls/B" }; return r; }
        private static Request Tabs(FormsWorkflowFixture f) { var r = f.Request(); r.Items = new[] { "B", "A" }; return r; }
        [TestMethod]
        public void LayoutRequiresDistinctCanonicalControlPathsInOneContainerAndCurrentVersion()
        {
            using (var f = new FormsWorkflowFixture())
            {
                foreach (var items in new[] { (string[])null, new string[1], new string[65], new[] { "Controls/A", "controls/a" } })
                { var r = Layout(f); r.Items = items; Assert.ThrowsException<ArgumentException>(() => f.Service.LayoutControls(r, true)); }
                var request = Layout(f); request.ExpectedTreeVersion = null; Assert.ThrowsException<ArgumentException>(() => f.Service.LayoutControls(request, true));
                request = Layout(f); request.ExpectedTreeVersion = "stale"; Assert.ThrowsException<InvalidOperationException>(() => f.Service.LayoutControls(request, true));
                foreach (string missing in new[] { (string)null, "Controls/Missing" })
                { request = Layout(f); request.Items[1] = missing; Assert.ThrowsException<ArgumentException>(() => f.Service.LayoutControls(request, true)); }
                var frame = f.Form.Designer.Controls.Add("Frame", "Frame"); frame.Controls.Add("Label", "Child");
                var page = frame.Pages.Add("Page", "Page");
                request = Layout(f); request.Items[1] = "Controls/Frame/Pages/Page";
                Assert.ThrowsException<ArgumentException>(() => f.Service.LayoutControls(request, true));
                request = Layout(f); request.Items[1] = "Controls/Frame/Controls/Child";
                Assert.ThrowsException<ArgumentException>(() => f.Service.LayoutControls(request, true));
            }
        }
        [TestMethod]
        public void LayoutPreviewPreservesStateAndApplyWorksInRootAndNestedContainers()
        {
            using (var f = new FormsWorkflowFixture())
            {
                var r = Layout(f); dynamic preview = f.Service.LayoutControls(r, true);
                Assert.AreEqual(40, f.Form.Designer.Controls.Item("B").Left);
                Assert.AreEqual(0, ((FormLayoutBox[])preview.After)[1].Left);
                Assert.AreEqual(r.ExpectedTreeVersion, f.Request().ExpectedTreeVersion);
                dynamic applied = f.Service.LayoutControls(r, false); Assert.IsTrue((bool)applied.Verified); Assert.AreEqual(0, f.Form.Designer.Controls.Item("B").Left);
                var frame = f.Form.Designer.Controls.Add("Frame", "Frame"); frame.Controls.Add("Label", "X"); frame.Controls.Add("Label", "Y").Top = 40;
                r = f.Request("align_top"); r.Items = new[] { "Controls/Frame/Controls/X", "Controls/Frame/Controls/Y" };
                Assert.IsTrue((bool)((dynamic)f.Service.LayoutControls(r, false)).Applied); Assert.AreEqual(0, frame.Controls.Item("Y").Top);
            }
        }
        [TestMethod]
        public void LayoutReportsCompleteAndIncompleteRollbackAfterSetterFailure()
        {
            foreach (bool rollbackFails in new[] { false, true })
                using (var f = new FormsWorkflowFixture())
                {
                    var b = f.Form.Designer.Controls.Item("B"); int calls = 0;
                    b.OnSetGeometry = (name, value) => { if (name == "Width" && (++calls == 1 || rollbackFails)) throw new InvalidOperationException("geometry rejected"); };
                    var error = Assert.ThrowsException<InvalidOperationException>(() => f.Service.LayoutControls(Layout(f), false));
                    StringAssert.Contains(error.Message, rollbackFails ? "rollback incomplete" : "Original geometry restored");
                    Assert.AreEqual(40, b.Left); Assert.AreEqual(0, f.Form.Designer.Controls.Item("A").Left);
                }
        }
        [TestMethod]
        public void LayoutReadbackChecksEveryCoordinateAndNonfiniteNativeValues()
        {
            foreach (string field in new[] { "NaN", "Infinity", "Left", "Top", "Width", "Height" })
                using (var f = new FormsWorkflowFixture())
                {
                    var b = f.Form.Designer.Controls.Item("B");
                    b.OnSetGeometry = (name, value) => { if (name == "Top") b.OnReadGeometry = (property, current) => property == (field == "NaN" || field == "Infinity" ? "Left" : field) ? field == "NaN" ? double.NaN : field == "Infinity" ? double.PositiveInfinity : current + 1 : current; };
                    var error = Assert.ThrowsException<InvalidOperationException>(() => f.Service.LayoutControls(Layout(f), false));
                    StringAssert.Contains(error.Message, "did not retain"); StringAssert.Contains(error.Message, "rollback incomplete");
                }
        }
        [TestMethod]
        public void TabOrderRequiresEveryUniqueDirectControlAndCanonicalCurrentContainer()
        {
            using (var f = new FormsWorkflowFixture())
            {
                foreach (var items in new[] { (string[])null, new string[0], new string[65], new[] { "A", "a" } })
                { var r = Tabs(f); r.Items = items; Assert.ThrowsException<ArgumentException>(() => f.Service.SetTabOrder(r)); }
                foreach (string version in new[] { (string)null, "stale" })
                { var r = Tabs(f); r.ExpectedTreeVersion = version; Assert.ThrowsException<InvalidOperationException>(() => f.Service.SetTabOrder(r)); }
                var request = Tabs(f); request.ParentPath = "Controls/Missing"; Assert.ThrowsException<ArgumentException>(() => f.Service.SetTabOrder(request));
                foreach (var items in new[] { new[] { "A" }, new[] { "A", (string)null }, new[] { "A", "Missing" } })
                { request = Tabs(f); request.Items = items; Assert.ThrowsException<ArgumentException>(() => f.Service.SetTabOrder(request)); }
                Assert.IsTrue((bool)((dynamic)f.Service.SetTabOrder(Tabs(f))).Verified);
                Assert.AreEqual(0, f.Form.Designer.Controls.Item("B").TabIndex); Assert.AreEqual(1, f.Form.Designer.Controls.Item("A").TabIndex);
                var frame = f.Form.Designer.Controls.Add("Frame", "Frame"); frame.Controls.Add("Label", "X"); frame.Controls.Add("Label", "Y");
                request = f.Request(); request.ParentPath = "Controls/Frame"; request.Items = new[] { "Y", "X" };
                Assert.IsTrue((bool)((dynamic)f.Service.SetTabOrder(request)).Applied); Assert.AreEqual(0, frame.Controls.Item("Y").TabIndex);
            }
        }
        [TestMethod]
        public void TabOrderRestoresOnSetterOrReadbackFailureAndReportsRollbackFailure()
        {
            foreach (int scenario in new[] { 0, 1, 2, 3 })
                using (var f = new FormsWorkflowFixture())
                {
                    var b = f.Form.Designer.Controls.Item("B"); int calls = 0;
                    b.OnSetTab = value =>
                    {
                        calls++;
                        if (scenario == 1) b.IgnoreTab = true;
                        if (scenario == 3 && calls == 1) b.OnReadTab = current => current + 1;
                        if ((scenario == 0 || scenario == 3) && calls == 1 || scenario == 2) throw new InvalidOperationException("tab rejected");
                    };
                    var error = Assert.ThrowsException<InvalidOperationException>(() => f.Service.SetTabOrder(Tabs(f)));
                    StringAssert.Contains(error.Message, scenario >= 2 ? "Rollback failed" : "Original tab order restored");
                    if (scenario == 3) StringAssert.Contains(error.Message, "rollback readback mismatch");
                    if (scenario == 1) StringAssert.Contains(error.Message, "Tab order readback mismatch");
                }
        }
    }
}

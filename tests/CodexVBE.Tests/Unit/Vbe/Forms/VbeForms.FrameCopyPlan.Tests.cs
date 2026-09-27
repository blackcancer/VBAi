namespace CodexVBE.Tests.Unit
{
    using System;
    using System.Collections;
    using System.Linq;
    using CodexVBE;
    using Microsoft.VisualStudio.TestTools.UnitTesting;

    public sealed partial class VbeFormsFrameDuplicationTests
    {
        [TestMethod]
        public void FrameCopyPlanDescribesDirectLabelsWithoutMutatingTheForm()
        {
            var f = Create();
            var label = f.Frame.Controls.AddExisting("Label", "Title");
            label.Caption = "Hello";
            var request = f.Request();
            dynamic plan = f.Service.FrameCopyPlan(request);
            Assert.IsTrue((bool)plan.EligibleForLimitedProbe);
            Assert.IsTrue((bool)plan.ReadOnly);
            Assert.IsFalse((bool)plan.MutationVerified);
            Assert.AreEqual(1, (int)plan.CollectionCount);
            Assert.AreEqual(1, (int)plan.DirectChildCount);
            var child = ((IEnumerable)plan.Children).Cast<object>().Single();
            Assert.AreEqual("Controls/Frame1/Controls/Title", (string)((dynamic)child).SourcePath);
            Assert.AreEqual("Controls/FrameCopy/Controls/FrameCopy_Title", (string)((dynamic)child).ProposedPath);
            Assert.AreEqual("Label positive profile", (string)((dynamic)child).Profile);
            Assert.AreEqual(1, f.Form.Designer.Controls.Count);
            Assert.AreEqual((string)request.ExpectedTreeVersion, (string)((dynamic)f.Service.Tree(f.Project.Name, f.Form.Name)).TreeVersion);
        }

        [TestMethod]
        public void SimplePlanAcceptsTextBoxButLabelOnlyPlanRefusesIt()
        {
            var f = Create();
            f.Frame.Controls.AddExisting("TextBox", "Entry");
            var request = f.Request();
            dynamic labels = f.Service.FrameCopyPlan(request);
            dynamic simple = f.Service.FrameSimpleCopyPlan(request);
            Assert.IsFalse((bool)labels.EligibleForLimitedProbe);
            Assert.IsFalse((bool)((dynamic)((IEnumerable)labels.Children).Cast<object>().Single()).Eligible);
            StringAssert.Contains((string)((dynamic)((IEnumerable)labels.Children).Cast<object>().Single()).Issue, "Only direct Label");
            Assert.IsTrue((bool)simple.EligibleForLimitedProbe);
            Assert.AreEqual("TextBox text-value positive profile", (string)((dynamic)((IEnumerable)simple.Children).Cast<object>().Single()).Profile);
        }

        [TestMethod]
        public void PlanReportsNameCollisionAndOversizedChildNameWithoutWriting()
        {
            var f = Create();
            f.Frame.Controls.AddExisting("Label", "Title");
            f.Form.Designer.Controls.AddExisting("Label", "FrameCopy_Title");
            var request = f.Request();
            dynamic collision = f.Service.FrameCopyPlan(request);
            Assert.IsFalse((bool)collision.EligibleForLimitedProbe);
            Assert.IsTrue(((IEnumerable)collision.Issues).Cast<string>().Any(issue => issue.Contains("already exists")));
            request.NewName = new string ('F', 38);
            dynamic longChild = f.Service.FrameSimpleCopyPlan(request);
            Assert.IsFalse((bool)longChild.EligibleForLimitedProbe);
            Assert.IsTrue(((IEnumerable)longChild.Issues).Cast<string>().Any(issue => issue.Contains("40 characters")));
            Assert.AreEqual(2, f.Form.Designer.Controls.Count);
        }

        [TestMethod]
        public void PlanRejectsStalePathWrongTypeAndInvalidName()
        {
            var f = Create();
            var request = f.Request();
            request.ExpectedTreeVersion = "stale";
            Assert.ThrowsException<InvalidOperationException>(() => f.Service.FrameCopyPlan(request));
            request.ExpectedTreeVersion = ((dynamic)f.Service.Tree(f.Project.Name, f.Form.Name)).TreeVersion;
            request.ControlPath = "Controls/frame1";
            Assert.ThrowsException<InvalidOperationException>(() => f.Service.FrameCopyPlan(request));
            request.ControlPath = "Controls/Frame1";
            request.NewName = "2Bad";
            Assert.ThrowsException<ArgumentException>(() => f.Service.FrameCopyPlan(request));
            request.NewName = "FrameCopy";
            var label = f.Form.Designer.Controls.AddExisting("Label", "Label1");
            request.ControlPath = "Controls/Label1";
            request.ExpectedTreeVersion = ((dynamic)f.Service.Tree(f.Project.Name, f.Form.Name)).TreeVersion;
            Assert.ThrowsException<InvalidOperationException>(() => f.Service.FrameCopyPlan(request));
            Assert.AreEqual("Label1", label.Name);
        }
    }
}

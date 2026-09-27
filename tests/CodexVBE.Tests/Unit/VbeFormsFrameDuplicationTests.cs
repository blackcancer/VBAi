using System;
using System.Collections;
using System.Linq;
using CodexVBE;
using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace CodexVBE.Tests.Unit
{
    [TestClass]
    [TestCategory("Unit")]
    public sealed class VbeFormsFrameDuplicationTests
    {
        private sealed class Fixture
        {
            public VbeFormsPartialTests.FakeProject Project;
            public VbeFormsPartialTests.FakeForm Form;
            public VbeFormsPartialTests.FakeControl Frame;
            public VbeForms Service;

            public Request Request(string newName = "FrameCopy")
            {
                return new Request { Project = Project.Name, Form = Form.Name,
                    ControlPath = "Controls/Frame1", NewName = newName,
                    ExpectedTreeVersion = ((dynamic)Service.Tree(Project.Name, Form.Name)).TreeVersion };
            }
        }

        private static Fixture Create()
        {
            var project = new VbeFormsPartialTests.FakeProject();
            var form = new VbeFormsPartialTests.FakeForm();
            project.VBComponents.Add(form);
            var vbe = new VbeFormsPartialTests.FakeVbe();
            vbe.VBProjects.Add(project);
            var frame = form.Designer.Controls.AddExisting("Frame", "Frame1");
            frame.Caption = "Group";
            frame.Left = 4; frame.Top = 6; frame.Width = 90; frame.Height = 60;
            return new Fixture { Project = project, Form = form, Frame = frame,
                Service = new VbeForms(vbe) };
        }

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
            Assert.AreEqual((string)request.ExpectedTreeVersion,
                (string)((dynamic)f.Service.Tree(f.Project.Name, f.Form.Name)).TreeVersion);
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
            StringAssert.Contains((string)((dynamic)((IEnumerable)labels.Children).Cast<object>().Single()).Issue,
                "Only direct Label");
            Assert.IsTrue((bool)simple.EligibleForLimitedProbe);
            Assert.AreEqual("TextBox text-value positive profile",
                (string)((dynamic)((IEnumerable)simple.Children).Cast<object>().Single()).Profile);
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
            Assert.IsTrue(((IEnumerable)collision.Issues).Cast<string>()
                .Any(issue => issue.Contains("already exists")));
            request.NewName = new string('F', 38);
            dynamic longChild = f.Service.FrameSimpleCopyPlan(request);
            Assert.IsFalse((bool)longChild.EligibleForLimitedProbe);
            Assert.IsTrue(((IEnumerable)longChild.Issues).Cast<string>()
                .Any(issue => issue.Contains("40 characters")));
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

        [TestMethod]
        public void EmptyFrameDuplicationCopiesSupportedPropertiesAndChangesVersion()
        {
            var f = Create();
            var request = f.Request();
            dynamic result = f.Service.DuplicateEmptyFrame(request);
            var copy = f.Form.Designer.Controls.Item("FrameCopy");
            Assert.AreEqual("Controls/FrameCopy", (string)result.NewPath);
            Assert.AreEqual("Group", copy.Caption);
            Assert.AreEqual(4d, copy.Left);
            Assert.AreEqual(6d, copy.Top);
            Assert.AreEqual(90d, copy.Width);
            Assert.AreEqual(60d, copy.Height);
            Assert.AreEqual(0, copy.Controls.Count);
            Assert.AreEqual(0, (int)result.SourceChildCount);
            Assert.AreNotEqual((string)request.ExpectedTreeVersion, (string)result.Tree.TreeVersion);
        }

        [TestMethod]
        public void EmptyFrameDuplicationRefusesChildrenStaleVersionAndGeometry()
        {
            var f = Create();
            var request = f.Request();
            request.ExpectedTreeVersion = "stale";
            Assert.ThrowsException<InvalidOperationException>(() => f.Service.DuplicateEmptyFrame(request));
            request.ExpectedTreeVersion = ((dynamic)f.Service.Tree(f.Project.Name, f.Form.Name)).TreeVersion;
            f.Frame.Controls.AddExisting("Label", "Title");
            request.ExpectedTreeVersion = ((dynamic)f.Service.Tree(f.Project.Name, f.Form.Name)).TreeVersion;
            Assert.ThrowsException<InvalidOperationException>(() => f.Service.DuplicateEmptyFrame(request));
            f.Frame.Controls.Remove("Title");
            f.Frame.Width = 0;
            request.ExpectedTreeVersion = ((dynamic)f.Service.Tree(f.Project.Name, f.Form.Name)).TreeVersion;
            Assert.ThrowsException<InvalidOperationException>(() => f.Service.DuplicateEmptyFrame(request));
            Assert.AreEqual(1, f.Form.Designer.Controls.Count);
        }

        [TestMethod]
        public void EmptyFrameRollsBackWhenCopiedCaptionSetterFails()
        {
            var f = Create();
            var request = f.Request();
            f.Form.Designer.Controls.FailNextCaption = true;
            Assert.ThrowsException<InvalidOperationException>(() => f.Service.DuplicateEmptyFrame(request));
            Assert.AreEqual(1, f.Form.Designer.Controls.Count);
            Assert.AreEqual(1, f.Form.Designer.Controls.RemoveCount);
            Assert.AreEqual((string)request.ExpectedTreeVersion,
                (string)((dynamic)f.Service.Tree(f.Project.Name, f.Form.Name)).TreeVersion);
        }

        [TestMethod]
        public void LabelFrameDuplicationCopiesChildrenAndSourceRemainsUntouched()
        {
            var f = Create();
            var label = f.Frame.Controls.AddExisting("Label", "Title");
            label.Caption = "Before";
            label.Left = 2; label.Top = 3; label.Width = 50; label.Height = 12;
            label.BackColor = 255;
            label.Font.Name = "Calibri"; label.Font.Size = 11; label.Font.Bold = true;
            var request = f.Request();
            dynamic result = f.Service.DuplicateFrameWithLabels(request);
            var copy = f.Form.Designer.Controls.Item("FrameCopy");
            var child = copy.Controls.Item("FrameCopy_Title");
            Assert.AreEqual("Controls/FrameCopy", (string)result.NewPath);
            Assert.AreEqual(1, (int)result.DirectLabelsCopied);
            Assert.AreEqual("Before", child.Caption);
            Assert.AreEqual(2d, child.Left);
            Assert.AreEqual(255, child.BackColor);
            Assert.AreEqual("Calibri", child.Font.Name);
            Assert.IsTrue(child.Font.Bold);
            Assert.AreEqual("Title", label.Name);
            Assert.AreNotEqual((string)request.ExpectedTreeVersion, (string)result.Tree.TreeVersion);
        }

        [TestMethod]
        public void LabelFrameDuplicationRejectsEmptyOrUnsupportedChildrenBeforeAdd()
        {
            var f = Create();
            var request = f.Request();
            Assert.ThrowsException<InvalidOperationException>(() => f.Service.DuplicateFrameWithLabels(request));
            f.Frame.Controls.AddExisting("TextBox", "Entry");
            request.ExpectedTreeVersion = ((dynamic)f.Service.Tree(f.Project.Name, f.Form.Name)).TreeVersion;
            Assert.ThrowsException<InvalidOperationException>(() => f.Service.DuplicateFrameWithLabels(request));
            Assert.AreEqual(1, f.Form.Designer.Controls.Count);
        }

        [TestMethod]
        public void LabelFrameRollsBackWhenFrameSetterFails()
        {
            var f = Create();
            f.Frame.Controls.AddExisting("Label", "Title");
            var request = f.Request();
            f.Form.Designer.Controls.FailNextCaption = true;
            Assert.ThrowsException<InvalidOperationException>(() => f.Service.DuplicateFrameWithLabels(request));
            Assert.AreEqual(1, f.Form.Designer.Controls.Count);
            Assert.AreEqual(1, f.Form.Designer.Controls.RemoveCount);
            Assert.AreEqual((string)request.ExpectedTreeVersion,
                (string)((dynamic)f.Service.Tree(f.Project.Name, f.Form.Name)).TreeVersion);
        }

        [TestMethod]
        public void CommandButtonDuplicationCopiesNarrowProfileAndLeavesEventsOut()
        {
            var f = Create();
            var button = f.Form.Designer.Controls.AddExisting("CommandButton", "RunButton");
            button.Caption = "Run";
            button.Left = 7; button.Top = 9; button.Width = 45; button.Height = 18;
            var request = f.Request("RunCopy");
            request.ControlPath = "Controls/RunButton";
            request.ExpectedTreeVersion = ((dynamic)f.Service.Tree(f.Project.Name, f.Form.Name)).TreeVersion;
            dynamic result = f.Service.DuplicateCommandButton(request);
            var copy = f.Form.Designer.Controls.Item("RunCopy");
            Assert.AreEqual("Run", copy.Caption);
            Assert.AreEqual(7d, copy.Left);
            Assert.AreEqual(18d, copy.Height);
            Assert.IsFalse((bool)result.EventsCopied);
            Assert.AreEqual("Partial", (string)result.Completeness);
        }

        [TestMethod]
        public void CommandButtonDuplicationRefusesWrongTypeAndRollsBackSetterFailure()
        {
            var f = Create();
            var request = f.Request("RunCopy");
            Assert.ThrowsException<InvalidOperationException>(() => f.Service.DuplicateCommandButton(request));
            f.Form.Designer.Controls.AddExisting("CommandButton", "RunButton");
            request.ControlPath = "Controls/RunButton";
            request.ExpectedTreeVersion = ((dynamic)f.Service.Tree(f.Project.Name, f.Form.Name)).TreeVersion;
            f.Form.Designer.Controls.FailNextCaption = true;
            Assert.ThrowsException<InvalidOperationException>(() => f.Service.DuplicateCommandButton(request));
            Assert.AreEqual(2, f.Form.Designer.Controls.Count);
            Assert.AreEqual(1, f.Form.Designer.Controls.RemoveCount);
        }
    }
}

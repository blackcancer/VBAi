namespace CodexVBE.Tests.Unit
{
    using System;
    using System.Collections;
    using System.Linq;
    using CodexVBE;
    using Microsoft.VisualStudio.TestTools.UnitTesting;

    public sealed partial class VbeFormsValueDuplicationTests
    {
        [TestMethod]
        public void ProfiledFramePlanAndCopyCoverEveryPositiveChildProfile()
        {
            var f = Create("Frame");
            f.Source.Controls.AddExisting("Label", "Label1").Caption = "Heading";
            f.Source.Controls.AddExisting("TextBox", "Text1").Value = "payload";
            f.Source.Controls.AddExisting("CheckBox", "Check1").Value = true;
            f.Source.Controls.AddExisting("CommandButton", "Button1").Caption = "Run";
            f.Source.Controls.AddExisting("ComboBox", "Combo1").ListWidth = "90 pt";
            f.Source.Controls.AddExisting("OptionButton", "Option1").Value = true;
            var request = f.Request(name: "FrameCopy");
            dynamic plan = f.Service.FrameProfileCopyPlan(request);
            Assert.IsTrue((bool)plan.EligibleForLimitedProbe);
            Assert.AreEqual(6, (int)plan.DirectChildCount);
            Assert.IsTrue((bool)plan.ReadOnly);
            dynamic result = f.Service.DuplicateFrameProfiled(request);
            var copy = f.Form.Designer.Controls.Item("FrameCopy");
            Assert.AreEqual(6, (int)result.DirectChildrenCopied);
            Assert.AreEqual(6, copy.Controls.Count);
            Assert.AreEqual("payload", copy.Controls.Item("FrameCopy_Text1").Value);
            Assert.AreEqual(true, copy.Controls.Item("FrameCopy_Check1").Value);
            Assert.AreEqual("90 pt", copy.Controls.Item("FrameCopy_Combo1").ListWidth);
            Assert.AreEqual(false, copy.Controls.Item("FrameCopy_Option1").Value);
            Assert.AreNotEqual((string)request.ExpectedTreeVersion, (string)result.Tree.TreeVersion);
        }

        [TestMethod]
        public void ProfiledFramePlanRefusesUnprofiledChildAndBadValue()
        {
            var f = Create("Frame");
            f.Source.Controls.AddExisting("ToggleButton", "Toggle1");
            dynamic unsupported = f.Service.FrameProfileCopyPlan(f.Request(name: "FrameCopy"));
            Assert.IsFalse((bool)unsupported.EligibleForLimitedProbe);
            Assert.IsTrue(((IEnumerable)unsupported.Issues).Cast<string>().Any(issue => issue.Contains("no positive copy profile")));
            Assert.ThrowsException<InvalidOperationException>(() => f.Service.DuplicateFrameProfiled(f.Request(name: "FrameCopy")));
            f.Source.Controls.Remove("Toggle1");
            f.Source.Controls.AddExisting("CheckBox", "Check1").Value = null;
            dynamic invalid = f.Service.FrameProfileCopyPlan(f.Request(name: "FrameCopy"));
            Assert.IsFalse((bool)invalid.EligibleForLimitedProbe);
            Assert.IsTrue(((IEnumerable)invalid.Issues).Cast<string>().Any(issue => issue.Contains("Boolean")));
            Assert.AreEqual(1, f.Form.Designer.Controls.Count);
        }

        [TestMethod]
        public void ProfiledFrameRollsBackChildFailureAndReportsIncompleteRollback()
        {
            var f = Create("Frame");
            f.Source.Controls.AddExisting("Label", "Title");
            f.Form.Designer.Controls.FailNextChildCaption = true;
            Assert.ThrowsException<InvalidOperationException>(() => f.Service.DuplicateFrameProfiled(f.Request(name: "FrameCopy")));
            Assert.AreEqual(1, f.Form.Designer.Controls.Count);
            Assert.AreEqual(1, f.Form.Designer.Controls.RemoveCount);
            f.Form.Designer.Controls.FailNextCaption = true;
            f.Form.Designer.Controls.FailNextRemove = true;
            var error = Assert.ThrowsException<InvalidOperationException>(() => f.Service.DuplicateFrameProfiled(f.Request(name: "FrameCopy")));
            StringAssert.Contains(error.Message, "rollback was incomplete");
            Assert.AreEqual(2, f.Form.Designer.Controls.Count);
        }

        [TestMethod]
        public void ProfiledPlanReportsEmptyFrameNameCollisionAndInvalidGeometry()
        {
            var f = Create("Frame");
            dynamic empty = f.Service.FrameProfileCopyPlan(f.Request(name: "FrameCopy"));
            Assert.IsFalse((bool)empty.EligibleForLimitedProbe);
            Assert.IsTrue(((IEnumerable)empty.Issues).Cast<string>().Any(issue => issue.Contains("duplicate_empty_form_frame")));
            f.Source.Controls.AddExisting("Label", "Title");
            f.Form.Designer.Controls.AddExisting("Label", "FrameCopy_Title");
            dynamic collision = f.Service.FrameProfileCopyPlan(f.Request(name: "FrameCopy"));
            Assert.IsFalse((bool)collision.EligibleForLimitedProbe);
            Assert.IsTrue(((IEnumerable)collision.Issues).Cast<string>().Any(issue => issue.Contains("proposed child name already exists")));
            f.Source.Width = 0;
            dynamic badGeometry = f.Service.FrameProfileCopyPlan(f.Request(name: "FrameCopy"));
            Assert.IsTrue(((IEnumerable)badGeometry.Issues).Cast<string>().Any(issue => issue.Contains("geometry")));
        }

        [TestMethod]
        public void ProfiledPlanRefusesUnreadableChildProfilesWithoutMutation()
        {
            var f = Create("Frame");
            var label = f.Source.Controls.AddExisting("Label", "Title");
            label.Font.Size = 250;
            dynamic fontIssue = f.Service.FrameProfileCopyPlan(f.Request(name: "FrameCopy"));
            Assert.IsFalse((bool)fontIssue.EligibleForLimitedProbe);
            Assert.IsTrue(((IEnumerable)fontIssue.Issues).Cast<string>().Any(issue => issue.Contains("font is outside")));
            label.Font.Size = 10;
            f.Source.Controls.AddExisting("TextBox", "Entry").Value = 42;
            dynamic textIssue = f.Service.FrameProfileCopyPlan(f.Request(name: "FrameCopy"));
            Assert.IsTrue(((IEnumerable)textIssue.Issues).Cast<string>().Any(issue => issue.Contains("TextBox.Value must be text")));
            f.Source.Controls.Item("Entry").Value = "good";
            f.Source.Controls.AddExisting("ComboBox", "Choices").ListWidth = "";
            dynamic comboIssue = f.Service.FrameProfileCopyPlan(f.Request(name: "FrameCopy"));
            Assert.IsTrue(((IEnumerable)comboIssue.Issues).Cast<string>().Any(issue => issue.Contains("ListWidth must be nonempty")));
            Assert.AreEqual(1, f.Form.Designer.Controls.Count);
        }

        [TestMethod]
        public void ProfiledPlanRejectsStaleTreeAndDirectChildLimit()
        {
            var f = Create("Frame");
            var request = f.Request(name: "FrameCopy");
            request.ExpectedTreeVersion = "stale";
            Assert.ThrowsException<InvalidOperationException>(() => f.Service.FrameProfileCopyPlan(request));
            for (int index = 0; index < 129; index++)
                f.Source.Controls.AddExisting("Label", "Label" + index);
            request.ExpectedTreeVersion = ((dynamic)f.Service.Tree(f.Project.Name, f.Form.Name)).TreeVersion;
            Assert.ThrowsException<InvalidOperationException>(() => f.Service.FrameProfileCopyPlan(request));
            Assert.AreEqual(1, f.Form.Designer.Controls.Count);
        }
    }
}

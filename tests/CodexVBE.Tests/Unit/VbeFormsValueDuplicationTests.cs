using System;
using System.Collections;
using System.Linq;
using CodexVBE;
using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace CodexVBE.Tests.Unit
{
    [TestClass]
    [TestCategory("Unit")]
    public sealed class VbeFormsValueDuplicationTests
    {
        private sealed class Fixture
        {
            public VbeFormsPartialTests.FakeProject Project;
            public VbeFormsPartialTests.FakeForm Form;
            public VbeFormsPartialTests.FakeControl Source;
            public VbeForms Service;

            public Request Request(string path = null, string name = "Copy")
            {
                return new Request { Project = Project.Name, Form = Form.Name,
                    ControlPath = path ?? "Controls/" + Source.Name, NewName = name,
                    ExpectedTreeVersion = ((dynamic)Service.Tree(Project.Name, Form.Name)).TreeVersion };
            }
        }

        private static Fixture Create(string type)
        {
            var project = new VbeFormsPartialTests.FakeProject();
            var form = new VbeFormsPartialTests.FakeForm();
            project.VBComponents.Add(form);
            var vbe = new VbeFormsPartialTests.FakeVbe();
            vbe.VBProjects.Add(project);
            var source = form.Designer.Controls.AddExisting(type, type + "1");
            source.Caption = "Source";
            source.Left = 3; source.Top = 4; source.Width = 50; source.Height = 16;
            return new Fixture { Project = project, Form = form, Source = source,
                Service = new VbeForms(vbe) };
        }

        [TestMethod]
        public void TextBoxCopiesTextAndGeometryWithoutChangingTheSource()
        {
            var f = Create("TextBox");
            f.Source.Value = "alpha";
            var request = f.Request();
            dynamic result = f.Service.DuplicateTextBox(request);
            var copy = f.Form.Designer.Controls.Item("Copy");
            Assert.AreEqual("alpha", copy.Value);
            Assert.AreEqual(3d, copy.Left);
            Assert.AreEqual(50d, copy.Width);
            Assert.AreEqual("alpha", f.Source.Value);
            Assert.AreNotEqual((string)request.ExpectedTreeVersion, (string)result.Tree.TreeVersion);
        }

        [TestMethod]
        public void TextBoxAcceptsEmptyValueButRejectsNonTextBeforeAdd()
        {
            var f = Create("TextBox");
            dynamic empty = f.Service.DuplicateTextBox(f.Request());
            Assert.IsNull(f.Form.Designer.Controls.Item("Copy").Value);
            f.Source.Value = 12;
            var request = f.Request(name: "BadCopy");
            Assert.ThrowsException<InvalidOperationException>(() => f.Service.DuplicateTextBox(request));
            Assert.AreEqual(2, f.Form.Designer.Controls.Count);
            Assert.AreEqual("Controls/Copy", (string)empty.NewPath);
        }

        [TestMethod]
        public void TextBoxRejectsStaleTreeAndRollsBackFailedValueSetter()
        {
            var f = Create("TextBox");
            f.Source.Value = "alpha";
            var request = f.Request();
            request.ExpectedTreeVersion = "stale";
            Assert.ThrowsException<InvalidOperationException>(() => f.Service.DuplicateTextBox(request));
            request.ExpectedTreeVersion = ((dynamic)f.Service.Tree(f.Project.Name, f.Form.Name)).TreeVersion;
            f.Form.Designer.Controls.FailNextValue = true;
            Assert.ThrowsException<InvalidOperationException>(() => f.Service.DuplicateTextBox(request));
            Assert.AreEqual(1, f.Form.Designer.Controls.Count);
            Assert.AreEqual(1, f.Form.Designer.Controls.RemoveCount);
        }

        [TestMethod]
        public void CheckBoxCopiesBooleanTrueAndRefusesTriState()
        {
            var f = Create("CheckBox");
            f.Source.Value = true;
            var request = f.Request();
            dynamic result = f.Service.DuplicateCheckBox(request);
            var copy = f.Form.Designer.Controls.Item("Copy");
            Assert.AreEqual(true, copy.Value);
            Assert.AreEqual("Source", copy.Caption);
            Assert.AreEqual("Partial", (string)result.Completeness);
            f.Source.Value = null;
            Assert.ThrowsException<InvalidOperationException>(() => f.Service.DuplicateCheckBox(f.Request(name: "BadCopy")));
            Assert.AreEqual(2, f.Form.Designer.Controls.Count);
        }

        [TestMethod]
        public void CheckBoxRollsBackWhenBooleanSetterFails()
        {
            var f = Create("CheckBox");
            f.Source.Value = true;
            f.Form.Designer.Controls.FailNextValue = true;
            Assert.ThrowsException<InvalidOperationException>(() => f.Service.DuplicateCheckBox(f.Request()));
            Assert.AreEqual(1, f.Form.Designer.Controls.Count);
            Assert.AreEqual(1, f.Form.Designer.Controls.RemoveCount);
        }

        [TestMethod]
        public void ToggleButtonCopiesFalseDefaultWithoutWritingSelection()
        {
            var f = Create("ToggleButton");
            var request = f.Request();
            dynamic result = f.Service.DuplicateToggleButton(request);
            Assert.AreEqual(false, f.Form.Designer.Controls.Item("Copy").Value);
            Assert.AreEqual("Source", f.Form.Designer.Controls.Item("Copy").Caption);
            Assert.IsTrue(((string[])result.CopiedProperties).Last().Contains("no setter call"));
            f.Source.Value = true;
            Assert.ThrowsException<InvalidOperationException>(() => f.Service.DuplicateToggleButton(f.Request(name: "BadCopy")));
            Assert.AreEqual(2, f.Form.Designer.Controls.Count);
        }

        [TestMethod]
        public void ToggleButtonRollsBackFailedCaption()
        {
            var f = Create("ToggleButton");
            f.Form.Designer.Controls.FailNextCaption = true;
            Assert.ThrowsException<InvalidOperationException>(() => f.Service.DuplicateToggleButton(f.Request()));
            Assert.AreEqual(1, f.Form.Designer.Controls.Count);
            Assert.AreEqual(1, f.Form.Designer.Controls.RemoveCount);
        }

        [TestMethod]
        public void OptionButtonCopiesShellButLeavesSelectionAndGroupOut()
        {
            var f = Create("OptionButton");
            f.Source.Value = true;
            dynamic result = f.Service.DuplicateOptionButton(f.Request());
            var copy = f.Form.Designer.Controls.Item("Copy");
            Assert.AreEqual("Source", copy.Caption);
            Assert.AreEqual(false, copy.Value);
            Assert.IsFalse((bool)result.SelectionCopied);
            Assert.IsFalse((bool)result.GroupCopied);
            Assert.AreEqual(true, f.Source.Value);
        }

        [TestMethod]
        public void OptionButtonRejectsWrongTypeAndRollsBackSetterFailure()
        {
            var f = Create("OptionButton");
            var request = f.Request();
            request.ControlPath = "Controls/Missing";
            Assert.ThrowsException<InvalidOperationException>(() => f.Service.DuplicateOptionButton(request));
            f.Form.Designer.Controls.AddExisting("Label", "Label1");
            request.ControlPath = "Controls/Label1";
            request.ExpectedTreeVersion = ((dynamic)f.Service.Tree(f.Project.Name, f.Form.Name)).TreeVersion;
            Assert.ThrowsException<InvalidOperationException>(() => f.Service.DuplicateOptionButton(request));
            request.ControlPath = "Controls/OptionButton1";
            f.Form.Designer.Controls.FailNextCaption = true;
            Assert.ThrowsException<InvalidOperationException>(() => f.Service.DuplicateOptionButton(request));
            Assert.AreEqual(2, f.Form.Designer.Controls.Count);
        }

        [TestMethod]
        public void ComboBoxCopiesListWidthButNotItemsOrBindings()
        {
            var f = Create("ComboBox");
            f.Source.ListWidth = "83 pt";
            f.Source.Items.Add("selection");
            f.Source.RowSource = "Sheet1!A1:A2";
            dynamic result = f.Service.DuplicateComboBox(f.Request());
            var copy = f.Form.Designer.Controls.Item("Copy");
            Assert.AreEqual("83 pt", copy.ListWidth);
            Assert.AreEqual(0, copy.Items.Count);
            Assert.AreEqual("", copy.RowSource);
            Assert.IsFalse((bool)result.ItemsCopied);
            Assert.IsFalse((bool)result.BindingsCopied);
        }

        [TestMethod]
        public void ComboBoxRejectsNonTextWidthAndRollsBackSetterFailure()
        {
            var f = Create("ComboBox");
            f.Source.ListWidth = 83;
            Assert.ThrowsException<InvalidOperationException>(() => f.Service.DuplicateComboBox(f.Request()));
            f.Source.ListWidth = "83 pt";
            f.Form.Designer.Controls.FailNextListWidth = true;
            Assert.ThrowsException<InvalidOperationException>(() => f.Service.DuplicateComboBox(f.Request()));
            Assert.AreEqual(1, f.Form.Designer.Controls.Count);
            Assert.AreEqual(1, f.Form.Designer.Controls.RemoveCount);
        }

        [TestMethod]
        public void SimpleFrameCopiesLabelAndTextBoxValueWithDistinctPaths()
        {
            var f = Create("Frame");
            var label = f.Source.Controls.AddExisting("Label", "Title");
            label.Caption = "Heading";
            var text = f.Source.Controls.AddExisting("TextBox", "Entry");
            text.Value = "payload";
            var request = f.Request(name: "FrameCopy");
            dynamic result = f.Service.DuplicateFrameWithSimpleChildren(request);
            var copy = f.Form.Designer.Controls.Item("FrameCopy");
            Assert.AreEqual("Heading", copy.Controls.Item("FrameCopy_Title").Caption);
            Assert.AreEqual("payload", copy.Controls.Item("FrameCopy_Entry").Value);
            Assert.AreEqual(1, (int)result.DirectLabelsCopied);
            Assert.AreEqual(1, (int)result.DirectTextBoxesCopied);
            Assert.AreEqual(2, ((IEnumerable)result.CopiedChildPaths).Cast<string>().Count());
        }

        [TestMethod]
        public void SimpleFrameRefusesNoTextBoxOrNonTextValueBeforeMutation()
        {
            var f = Create("Frame");
            f.Source.Controls.AddExisting("Label", "Title");
            Assert.ThrowsException<InvalidOperationException>(() =>
                f.Service.DuplicateFrameWithSimpleChildren(f.Request(name: "FrameCopy")));
            var text = f.Source.Controls.AddExisting("TextBox", "Entry");
            text.Value = 42;
            Assert.ThrowsException<InvalidOperationException>(() =>
                f.Service.DuplicateFrameWithSimpleChildren(f.Request(name: "FrameCopy")));
            Assert.AreEqual(1, f.Form.Designer.Controls.Count);
        }

        [TestMethod]
        public void SimpleFrameRemovesCreatedChildAndFrameAfterChildSetterFailure()
        {
            var f = Create("Frame");
            f.Source.Controls.AddExisting("Label", "Title");
            f.Source.Controls.AddExisting("TextBox", "Entry").Value = "payload";
            f.Form.Designer.Controls.FailNextChildCaption = true;
            Assert.ThrowsException<InvalidOperationException>(() =>
                f.Service.DuplicateFrameWithSimpleChildren(f.Request(name: "FrameCopy")));
            Assert.AreEqual(1, f.Form.Designer.Controls.Count);
            Assert.AreEqual(1, f.Form.Designer.Controls.RemoveCount);
        }

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
            Assert.IsTrue(((IEnumerable)unsupported.Issues).Cast<string>()
                .Any(issue => issue.Contains("no positive copy profile")));
            Assert.ThrowsException<InvalidOperationException>(() =>
                f.Service.DuplicateFrameProfiled(f.Request(name: "FrameCopy")));
            f.Source.Controls.Remove("Toggle1");
            f.Source.Controls.AddExisting("CheckBox", "Check1").Value = null;
            dynamic invalid = f.Service.FrameProfileCopyPlan(f.Request(name: "FrameCopy"));
            Assert.IsFalse((bool)invalid.EligibleForLimitedProbe);
            Assert.IsTrue(((IEnumerable)invalid.Issues).Cast<string>()
                .Any(issue => issue.Contains("Boolean")));
            Assert.AreEqual(1, f.Form.Designer.Controls.Count);
        }

        [TestMethod]
        public void ProfiledFrameRollsBackChildFailureAndReportsIncompleteRollback()
        {
            var f = Create("Frame");
            f.Source.Controls.AddExisting("Label", "Title");
            f.Form.Designer.Controls.FailNextChildCaption = true;
            Assert.ThrowsException<InvalidOperationException>(() =>
                f.Service.DuplicateFrameProfiled(f.Request(name: "FrameCopy")));
            Assert.AreEqual(1, f.Form.Designer.Controls.Count);
            Assert.AreEqual(1, f.Form.Designer.Controls.RemoveCount);

            f.Form.Designer.Controls.FailNextCaption = true;
            f.Form.Designer.Controls.FailNextRemove = true;
            var error = Assert.ThrowsException<InvalidOperationException>(() =>
                f.Service.DuplicateFrameProfiled(f.Request(name: "FrameCopy")));
            StringAssert.Contains(error.Message, "rollback was incomplete");
            Assert.AreEqual(2, f.Form.Designer.Controls.Count);
        }
    }
}

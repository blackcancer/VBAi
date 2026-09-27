namespace CodexVBE.Tests.Unit
{
    using System;
    using System.Collections.Generic;
    using System.Linq;
    using CodexVBE;
    using Microsoft.VisualStudio.TestTools.UnitTesting;

    public sealed partial class VbeFormsContractTests
    {
        [TestMethod]
        public void FormListingIncludesOnlyUserFormsAndTheirDesignerState()
        {
            var project = new FakeProject
            {
                Name = "Projet"
            };
            project.VBComponents.Add(new FakeComponent { Name = "Module1", Type = 1 });
            project.VBComponents.Add(new FakeComponent { Name = "Accueil", Type = 3, HasOpenDesigner = true });
            project.VBComponents.Add(new FakeComponent { Name = "Détails", Type = 3, HasOpenDesigner = false });
            var host = new FakeVbe();
            host.VBProjects.Add(project);
            var forms = new VbeForms(host);
            dynamic list = forms.List("projet");
            Assert.AreEqual(2, ((IEnumerable<object>)list).Count());
            var rows = ((IEnumerable<object>)list).Cast<dynamic>().ToArray();
            Assert.AreEqual("Accueil", (string)rows[0].Name);
            Assert.IsTrue((bool)rows[0].DesignerOpen);
            Assert.AreEqual("Détails", (string)rows[1].Name);
            Assert.IsFalse((bool)rows[1].DesignerOpen);
            Assert.ThrowsException<ArgumentException>(() => forms.SetProperty(new Request { Property = "", Value = "value" }));
        }
    }
}

namespace CodexVBE.Tests.Unit
{
    using System;
    using System.Linq;
    using CodexVBE;
    using Microsoft.VisualStudio.TestTools.UnitTesting;

    [TestClass]
    [TestCategory("Unit")]
    public sealed partial class VbeFormsCoreBranchTests
    {
        [TestMethod]
        public void MultiPageAddAndRemoveUseCanonicalPagePathAndIndex()
        {
            var f = Create("MultiPage");
            var request = f.Request("Controls/MultiPage1");
            request.ParentPath = request.ControlPath;
            request.NewName = "PageA";
            request.Caption = "First";
            dynamic first = f.Service.AddPageOrTab(request, "Pages");
            Assert.AreEqual("Controls/MultiPage1/Pages/PageA", (string)first.AddedPath);
            Assert.AreEqual("First", f.Control.Pages.Single().Caption);
            request.ExpectedTreeVersion = first.Tree.TreeVersion;
            request.NewName = "PageB";
            request.InsertIndex = 0;
            dynamic second = f.Service.AddPageOrTab(request, "Pages");
            Assert.AreEqual("PageB", f.Control.Pages.First().Name);
            var remove = f.Request("Controls/MultiPage1/Pages/PageA");
            dynamic removed = f.Service.RemovePageOrTab(remove);
            Assert.IsTrue((bool)removed.Applied);
            Assert.AreEqual("PageB", f.Control.Pages.Single().Name);
            Assert.AreNotEqual((string)second.Tree.TreeVersion, (string)removed.Tree.TreeVersion);
        }

        [TestMethod]
        public void TabStripAddAndRemoveTrackNativeTabCollection()
        {
            var f = Create("TabStrip");
            var request = f.Request("Controls/TabStrip1");
            request.ParentPath = request.ControlPath;
            request.NewName = "TabA";
            dynamic added = f.Service.AddPageOrTab(request, "Tabs");
            Assert.AreEqual("TabA", f.Control.Tabs.Single().Name);
            Assert.AreEqual("TabA", f.Control.Tabs.Single().Caption);
            dynamic removed = f.Service.RemovePageOrTab(f.Request("Controls/TabStrip1/Tabs/TabA"));
            Assert.IsTrue((bool)removed.Applied);
            Assert.AreEqual(0, f.Control.Tabs.Count);
            Assert.AreNotEqual((string)added.Tree.TreeVersion, (string)removed.Tree.TreeVersion);
        }

        [TestMethod]
        public void PageAddRejectsStaleVersionDuplicateIndexAndWrongParent()
        {
            var f = Create("MultiPage");
            var request = f.Request("Controls/MultiPage1");
            request.ParentPath = request.ControlPath;
            request.NewName = "PageA";
            request.ExpectedTreeVersion = "stale";
            Assert.ThrowsException<InvalidOperationException>(() => f.Service.AddPageOrTab(request, "Pages"));
            request.ExpectedTreeVersion = f.Version;
            request.InsertIndex = -1;
            Assert.ThrowsException<ArgumentOutOfRangeException>(() => f.Service.AddPageOrTab(request, "Pages"));
            request.InsertIndex = null;
            f.Control.Pages.Add("PageA", "Existing");
            request.ExpectedTreeVersion = f.Version;
            Assert.ThrowsException<InvalidOperationException>(() => f.Service.AddPageOrTab(request, "Pages"));
            Assert.AreEqual(1, f.Control.Pages.Count);
            Assert.ThrowsException<InvalidOperationException>(() => f.Service.AddPageOrTab(request, "Tabs"));
        }

        [TestMethod]
        public void PageAddRollsBackAfterNativeAddFailureAndReportsFailedRollback()
        {
            var f = Create("MultiPage");
            var request = f.Request("Controls/MultiPage1");
            request.ParentPath = request.ControlPath;
            request.NewName = "PageA";
            f.Control.Pages.FailAfterAdd = true;
            var clean = Assert.ThrowsException<InvalidOperationException>(() => f.Service.AddPageOrTab(request, "Pages"));
            StringAssert.Contains(clean.Message, "absent after rollback");
            Assert.AreEqual(0, f.Control.Pages.Count);
            f.Control.Pages.FailRemove = true;
            var incomplete = Assert.ThrowsException<InvalidOperationException>(() => f.Service.AddPageOrTab(request, "Pages"));
            StringAssert.Contains(incomplete.Message, "rollback could not be verified");
            Assert.AreEqual(1, f.Control.Pages.Count);
        }

        [TestMethod]
        public void NodeScalarMutationUsesVersionAndConvertsToDeclaredType()
        {
            var f = Create("Label");
            var request = f.Request("Controls/Label1");
            request.Property = "Left";
            request.Value = "12.5";
            dynamic changed = f.Service.SetNodeProperty(request);
            Assert.AreEqual(12.5d, f.Control.Left);
            request.ExpectedTreeVersion = changed.Tree.TreeVersion;
            request.Property = "Caption";
            request.Value = "Updated";
            dynamic captioned = f.Service.SetNodeProperty(request);
            Assert.AreEqual("Updated", f.Control.Caption);
            Assert.AreNotEqual((string)changed.Tree.TreeVersion, (string)captioned.Tree.TreeVersion);
            request.ExpectedTreeVersion = "stale";
            Assert.ThrowsException<InvalidOperationException>(() => f.Service.SetNodeProperty(request));
        }

        [TestMethod]
        public void NodeMutationRefusesCrashProneNativeSettersBeforeWriting()
        {
            foreach (var caseInfo in new[]
            {
                Tuple.Create("ComboBox", "ColumnCount", (object)2),
                Tuple.Create("TextBox", "ScrollBars", (object)2),
                Tuple.Create("SpinButton", "Min", (object)1),
                Tuple.Create("ToggleButton", "Value", (object)true),
                Tuple.Create("Label", "Cancel", (object)true)
            }

            )
            {
                var f = Create(caseInfo.Item1);
                var request = f.Request("Controls/" + caseInfo.Item1 + "1");
                request.Property = caseInfo.Item2;
                request.Value = caseInfo.Item3;
                Assert.ThrowsException<InvalidOperationException>(() => f.Service.SetNodeProperty(request), caseInfo.Item1 + "." + caseInfo.Item2);
                Assert.AreEqual(1, f.Form.Designer.Controls.Count);
            }
        }

        [TestMethod]
        public void NodeMutationRejectsManagedObjectMemberAndUnknownProperty()
        {
            var f = Create("Label");
            var request = f.Request("Controls/Label1");
            request.Property = "Missing";
            request.Value = "x";
            Assert.ThrowsException<InvalidOperationException>(() => f.Service.SetNodeProperty(request));
            request.Property = "Font.Name";
            Assert.ThrowsException<InvalidOperationException>(() => f.Service.SetNodeProperty(request));
            Assert.AreEqual("Arial", f.Control.Font.Name);
        }
    }
}

namespace CodexVBE.Tests.Unit
{
    using System;
    using System.Collections;
    using System.Collections.Generic;
    using System.ComponentModel;
    using System.Linq;
    using System.Runtime.CompilerServices;
    using Microsoft.VisualStudio.TestTools.UnitTesting;
    using CodexVBE;

    public sealed partial class VbeFormsTests
    {
        [TestMethod]
        public void ListFiltersNonFormsAndTreeTracksCanonicalNestedPaths()
        {
            var fixture = NewFixture();
            fixture.Project.VBComponents.Add(new FakeModule("Module1"));
            var frame = fixture.Form.Designer.Controls.AddExisting("Frame1");
            frame.Controls.AddExisting("Child1");
            var forms = ((IEnumerable)fixture.Service.List(fixture.Project.Name)).Cast<object>().ToArray();
            Assert.AreEqual(1, forms.Length);
            Assert.AreEqual("Form1", (string)((dynamic)forms[0]).Name);
            dynamic tree = fixture.Service.Tree(fixture.Project.Name, fixture.Form.Name);
            Assert.AreEqual(2, (int)tree.NodeCount);
            var root = ((IEnumerable)tree.Controls).Cast<object>().Single();
            Assert.AreEqual("Controls/Frame1", (string)((dynamic)root).Path);
            var child = ((IEnumerable)((dynamic)root).Children).Cast<object>().Single();
            Assert.AreEqual("Controls/Frame1/Controls/Child1", (string)((dynamic)child).Path);
            Assert.AreEqual((string)tree.FormVersion, (string)tree.TreeVersion);
        }

        [TestMethod]
        public void SnapshotReportsOnlyDirectControlsAndParentProbeConfirmsOwnership()
        {
            var fixture = NewFixture();
            var frame = fixture.Form.Designer.Controls.AddExisting("Frame1");
            frame.Controls.AddExisting("Child1");
            dynamic state = fixture.Service.State(fixture.Project.Name, fixture.Form.Name);
            Assert.AreEqual("Original caption", (string)state.Caption);
            Assert.AreEqual(1, ((IEnumerable)state.Controls).Cast<object>().Count());
            Assert.AreEqual("Frame1", (string)((dynamic)((IEnumerable)state.Controls).Cast<object>().Single()).Name);
            dynamic parents = fixture.Service.ParentProbe(fixture.Project.Name, fixture.Form.Name);
            var row = ((IEnumerable)parents.Rows).Cast<object>().Single();
            Assert.IsTrue((bool)((dynamic)row).SameDesigner);
            Assert.IsFalse((bool)((dynamic)row).SameComponent);
        }

        [TestMethod]
        public void FormPropertiesDescribeScalarValueAndRejectStaleMutation()
        {
            var fixture = NewFixture();
            var properties = ((IEnumerable)fixture.Service.Properties(fixture.Project.Name, fixture.Form.Name)).Cast<VbePropertyInfo>().ToArray();
            var caption = properties.Single(p => p.Name == "Caption");
            Assert.AreEqual("scalar", caption.Kind);
            Assert.AreEqual("Original caption", caption.Value);
            var request = new Request
            {
                Project = fixture.Project.Name,
                Form = fixture.Form.Name,
                Property = "Caption",
                Value = "Edited",
                ExpectedFormVersion = "stale"
            };
            Assert.ThrowsException<InvalidOperationException>(() => fixture.Service.SetProperty(request));
            Assert.AreEqual("Original caption", fixture.Form.Designer.Caption);
        }

        [TestMethod]
        public void FormCaptionMutationCanBeReversedWithFreshVersion()
        {
            var fixture = NewFixture();
            dynamic before = fixture.Service.State(fixture.Project.Name, fixture.Form.Name);
            var request = new Request
            {
                Project = fixture.Project.Name,
                Form = fixture.Form.Name,
                Property = "Caption",
                Value = "Edited",
                ExpectedFormVersion = before.Version
            };
            dynamic changed = fixture.Service.SetProperty(request);
            Assert.AreEqual("Edited", fixture.Form.Designer.Caption);
            Assert.AreEqual("Edited", (string)changed.State.Caption);
            Assert.AreNotEqual((string)before.Version, (string)changed.State.Version);
            request.Value = "Original caption";
            request.ExpectedFormVersion = changed.State.Version;
            dynamic restored = fixture.Service.SetProperty(request);
            Assert.AreEqual("Original caption", (string)restored.State.Caption);
            Assert.AreEqual((string)before.Version, (string)restored.State.Version);
        }

        [TestMethod]
        public void AddAndRemoveControlRoundTripChangesTreeWithoutTouchingOtherControls()
        {
            var fixture = NewFixture();
            fixture.Form.Designer.Controls.AddExisting("Keep");
            dynamic before = fixture.Service.Tree(fixture.Project.Name, fixture.Form.Name);
            var request = new Request
            {
                Project = fixture.Project.Name,
                Form = fixture.Form.Name,
                Control = "Added",
                ControlType = "Forms.Label.1",
                Caption = "New",
                Left = 2,
                Top = 3,
                Width = 60,
                Height = 20,
                ExpectedFormVersion = before.TreeVersion
            };
            dynamic added = fixture.Service.AddControl(request);
            Assert.AreEqual(2, fixture.Form.Designer.Controls.Count);
            Assert.AreEqual("New", fixture.Form.Designer.Controls.Item("Added").Caption);
            Assert.AreNotEqual((string)before.TreeVersion, (string)added.Version);
            dynamic tree = fixture.Service.Tree(fixture.Project.Name, fixture.Form.Name);
            dynamic removed = fixture.Service.RemoveControl(new Request { Project = fixture.Project.Name, Form = fixture.Form.Name, ControlPath = "Controls/Added", ExpectedTreeVersion = tree.TreeVersion });
            Assert.IsTrue((bool)removed.Applied);
            Assert.AreEqual(1, fixture.Form.Designer.Controls.Count);
            Assert.AreEqual("Keep", fixture.Form.Designer.Controls.Single().Name);
            Assert.AreEqual((string)before.TreeVersion, (string)removed.Tree.TreeVersion);
        }

        [TestMethod]
        public void FailedPostAddSetterRollsBackTheControl()
        {
            var fixture = NewFixture();
            fixture.Form.Designer.Controls.FailNextCaption = true;
            dynamic before = fixture.Service.State(fixture.Project.Name, fixture.Form.Name);
            var error = Assert.ThrowsException<InvalidOperationException>(() => fixture.Service.AddControl(new Request { Project = fixture.Project.Name, Form = fixture.Form.Name, Control = "Broken", ControlType = "Forms.Label.1", Caption = "Fails", Left = 1, Top = 1, Width = 40, Height = 20, ExpectedFormVersion = before.Version }));
            StringAssert.Contains(error.Message, "no control with the requested name remains");
            Assert.AreEqual(0, fixture.Form.Designer.Controls.Count);
            dynamic after = fixture.Service.State(fixture.Project.Name, fixture.Form.Name);
            Assert.AreEqual((string)before.Version, (string)after.Version);
        }

        [TestMethod]
        public void NestedControlCanBeAddedAndRemovedByCanonicalPath()
        {
            var fixture = NewFixture();
            fixture.Form.Designer.Controls.AddExisting("Frame1");
            dynamic before = fixture.Service.Tree(fixture.Project.Name, fixture.Form.Name);
            dynamic added = fixture.Service.AddNestedControl(new Request { Project = fixture.Project.Name, Form = fixture.Form.Name, ParentPath = "Controls/Frame1", Control = "Nested", ControlType = "Forms.Label.1", Caption = "Inside", Left = 1, Top = 1, Width = 30, Height = 10, ExpectedTreeVersion = before.TreeVersion });
            Assert.AreEqual(2, (int)added.NodeCount);
            Assert.AreEqual(1, fixture.Form.Designer.Controls.Item("Frame1").Controls.Count);
            dynamic removed = fixture.Service.RemoveControl(new Request { Project = fixture.Project.Name, Form = fixture.Form.Name, ControlPath = "Controls/Frame1/Controls/Nested", ExpectedTreeVersion = added.TreeVersion });
            Assert.IsTrue((bool)removed.Applied);
            Assert.AreEqual(0, fixture.Form.Designer.Controls.Item("Frame1").Controls.Count);
            Assert.AreEqual((string)before.TreeVersion, (string)removed.Tree.TreeVersion);
        }

        [TestMethod]
        public void NodeCaptionMutationRequiresFreshTreeAndCanBeReversed()
        {
            var fixture = NewFixture();
            var label = fixture.Form.Designer.Controls.AddExisting("Label1");
            dynamic before = fixture.Service.Tree(fixture.Project.Name, fixture.Form.Name);
            var request = new Request
            {
                Project = fixture.Project.Name,
                Form = fixture.Form.Name,
                ControlPath = "Controls/Label1",
                Property = "Caption",
                Value = "After",
                ExpectedTreeVersion = before.TreeVersion
            };
            dynamic changed = fixture.Service.SetNodeProperty(request);
            Assert.AreEqual("After", label.Caption);
            Assert.AreNotEqual((string)before.TreeVersion, (string)changed.Tree.TreeVersion);
            request.Value = "Original control";
            request.ExpectedTreeVersion = changed.Tree.TreeVersion;
            dynamic restored = fixture.Service.SetNodeProperty(request);
            Assert.AreEqual("Original control", label.Caption);
            Assert.AreEqual((string)before.TreeVersion, (string)restored.Tree.TreeVersion);
        }

        [TestMethod]
        public void GeometryRejectsNonFiniteInputBeforeVbideMutation()
        {
            var fixture = NewFixture();
            fixture.Form.Designer.Controls.AddExisting("Label1");
            var request = new Request
            {
                Project = fixture.Project.Name,
                Form = fixture.Form.Name,
                Control = "Label1",
                Left = double.NaN,
                Top = 1,
                Width = 20,
                Height = 10
            };
            Assert.ThrowsException<ArgumentException>(() => fixture.Service.SetControlGeometry(request));
            Assert.AreEqual(0, fixture.Form.Designer.Controls.Item("Label1").Left);
        }

        [TestMethod]
        public void CreateAndOpenFormRequireValidIdentityAndDesignMode()
        {
            var fixture = NewFixture();
            var request = new Request
            {
                Project = fixture.Project.Name,
                Form = "NewForm"
            };
            fixture.Project.Mode = 1;
            Assert.ThrowsException<InvalidOperationException>(() => fixture.Service.Create(request));
            fixture.Project.Mode = 2;
            request.Form = "2Invalid";
            Assert.ThrowsException<ArgumentException>(() => fixture.Service.Create(request));
            request.Form = "form1";
            Assert.ThrowsException<InvalidOperationException>(() => fixture.Service.Create(request));
            request.Form = "NewForm";
            dynamic created = fixture.Service.Create(request);
            Assert.AreEqual("NewForm", (string)created.Form);
            var form = fixture.Project.VBComponents.OfType<FakeForm>().Single(item => item.Name == "NewForm");
            Assert.IsTrue(form.DesignerWindow().Visible);
            form.DesignerWindow().Visible = false;
            dynamic opened = fixture.Service.Open(fixture.Project.Name, "newform");
            Assert.IsTrue(form.DesignerWindow().Visible);
            Assert.AreEqual("NewForm", (string)opened.Form);
        }

        [TestMethod]
        public void FailedFormRenameRollsBackOnlyTheCreatedComponent()
        {
            var fixture = NewFixture();
            fixture.Project.VBComponents.RejectedCreatedName = "DeniedForm";
            var error = Assert.ThrowsException<InvalidOperationException>(() => fixture.Service.Create(new Request { Project = fixture.Project.Name, Form = "DeniedForm" }));
            StringAssert.Contains(error.Message, "absent after rollback");
            Assert.AreEqual(1, fixture.Project.VBComponents.Count);
            Assert.AreEqual("Form1", fixture.Project.VBComponents.OfType<FakeForm>().Single().Name);
        }

        [TestMethod]
        public void ControlGeometryNameCaptionAndFontRoundTripWithFreshVersion()
        {
            var fixture = NewFixture();
            var control = fixture.Form.Designer.Controls.AddExisting("Label1");
            dynamic before = fixture.Service.State(fixture.Project.Name, fixture.Form.Name);
            var request = new Request
            {
                Project = fixture.Project.Name,
                Form = fixture.Form.Name,
                Control = "Label1",
                ExpectedFormVersion = before.Version,
                Left = 12,
                Top = 7,
                Width = 80,
                Height = 21
            };
            dynamic moved = fixture.Service.SetControlGeometry(request);
            Assert.AreEqual(12d, control.Left);
            Assert.AreEqual(80d, control.Width);
            request.ExpectedFormVersion = moved.Version;
            request.NewName = "Heading";
            dynamic renamed = fixture.Service.RenameControl(request);
            Assert.AreEqual("Heading", control.Name);
            request.Control = "Heading";
            request.ExpectedFormVersion = renamed.Version;
            request.Caption = "New heading";
            dynamic captioned = fixture.Service.SetControlCaption(request);
            Assert.AreEqual("New heading", control.Caption);
            request.ExpectedFormVersion = captioned.Version;
            request.FontName = "Consolas";
            request.FontSize = 13;
            request.FontBold = true;
            dynamic font = fixture.Service.SetControlFont(request);
            Assert.AreEqual("Consolas", control.Font.Name);
            Assert.AreEqual(13d, control.Font.Size);
            Assert.IsTrue(control.Font.Bold);
            Assert.AreNotEqual((string)before.Version, (string)font.Version);
        }

        [TestMethod]
        public void ControlEditsRejectDuplicateNameMissingCaptionInvalidFontAndStaleState()
        {
            var fixture = NewFixture();
            fixture.Form.Designer.Controls.AddExisting("Label1");
            fixture.Form.Designer.Controls.AddExisting("Keep");
            dynamic state = fixture.Service.State(fixture.Project.Name, fixture.Form.Name);
            var request = new Request
            {
                Project = fixture.Project.Name,
                Form = fixture.Form.Name,
                Control = "Label1",
                ExpectedFormVersion = state.Version,
                NewName = "Keep"
            };
            Assert.ThrowsException<InvalidOperationException>(() => fixture.Service.RenameControl(request));
            request.NewName = "2Bad";
            Assert.ThrowsException<ArgumentException>(() => fixture.Service.RenameControl(request));
            request.Caption = null;
            Assert.ThrowsException<ArgumentException>(() => fixture.Service.SetControlCaption(request));
            request.FontName = "Arial";
            request.FontSize = double.NaN;
            Assert.ThrowsException<ArgumentException>(() => fixture.Service.SetControlFont(request));
            request.FontSize = 11;
            request.ExpectedFormVersion = "stale";
            Assert.ThrowsException<InvalidOperationException>(() => fixture.Service.SetControlFont(request));
            Assert.AreEqual("Label1", fixture.Form.Designer.Controls.Item("Label1").Name);
        }

        [TestMethod]
        public void ControlPropertiesIncludeCurrentNativeValuesAndRejectUnknownControl()
        {
            var fixture = NewFixture();
            fixture.Form.Designer.Controls.AddExisting("Label1").Caption = "Visible text";
            var properties = ((IEnumerable)fixture.Service.ControlProperties(fixture.Project.Name, fixture.Form.Name, "Label1")).Cast<object>().ToArray();
            var caption = properties.Single(property => (string)((dynamic)property).Name == "Caption");
            Assert.AreEqual("Visible text", (string)((dynamic)caption).Value);
            Assert.ThrowsException<InvalidOperationException>(() => fixture.Service.ControlProperties(fixture.Project.Name, fixture.Form.Name, "Missing"));
        }

        [TestMethod]
        public void FormPropertiesWriteScalarIdentityAndAllSupportedFontMembers()
        {
            var fixture = NewFixture();
            var request = new Request
            {
                Project = fixture.Project.Name,
                Form = fixture.Form.Name,
                ExpectedFormVersion = ((dynamic)fixture.Service.State(fixture.Project.Name, fixture.Form.Name)).Version
            };
            request.Property = "Width";
            request.Value = "345.5";
            dynamic resized = fixture.Service.SetProperty(request);
            Assert.AreEqual(345.5d, fixture.Form.Designer.Width);
            request.ExpectedFormVersion = resized.State.Version;
            request.Property = "Name";
            request.Value = "RenamedForm";
            dynamic renamed = fixture.Service.SetProperty(request);
            Assert.AreEqual("RenamedForm", fixture.Form.Name);
            request.Form = fixture.Form.Name;
            request.ExpectedFormVersion = renamed.State.Version;
            foreach (var edit in new[]
            {
                Tuple.Create("Font.Name", (object)"Consolas"),
                Tuple.Create("Font.Size", (object)12.5d),
                Tuple.Create("Font.Bold", (object)true),
                Tuple.Create("Font.Italic", (object)true),
                Tuple.Create("Font.Underline", (object)true),
                Tuple.Create("Font.Strikethrough", (object)true)
            }

            )
            {
                request.Property = edit.Item1;
                request.Value = edit.Item2;
                dynamic result = fixture.Service.SetProperty(request);
                request.ExpectedFormVersion = result.State.Version;
            }

            Assert.AreEqual("Consolas", fixture.Form.Designer.Font.Name);
            Assert.AreEqual(12.5d, fixture.Form.Designer.Font.Size);
            Assert.IsTrue(fixture.Form.Designer.Font.Bold);
            Assert.IsTrue(fixture.Form.Designer.Font.Italic);
            Assert.IsTrue(fixture.Form.Designer.Font.Underline);
            Assert.IsTrue(fixture.Form.Designer.Font.Strikethrough);
        }

        [TestMethod]
        public void FormPropertyWritesRejectUnknownUnsafeAndMalformedMembers()
        {
            var fixture = NewFixture();
            var request = new Request
            {
                Project = fixture.Project.Name,
                Form = fixture.Form.Name,
                ExpectedFormVersion = ((dynamic)fixture.Service.State(fixture.Project.Name, fixture.Form.Name)).Version,
                Value = "value"
            };
            foreach (string property in new[]
            {
                "Missing",
                "Caption.Other"
            }

            )
            {
                request.Property = property;
                Assert.ThrowsException<InvalidOperationException>(() => fixture.Service.SetProperty(request));
            }

            request.Property = "Font..Size";
            Assert.ThrowsException<ArgumentException>(() => fixture.Service.SetProperty(request));
            request.Property = "Font.Unknown";
            request.Value = true;
            Assert.ThrowsException<InvalidOperationException>(() => fixture.Service.SetProperty(request));
            request.Property = "Font.Size";
            request.Value = double.PositiveInfinity;
            Assert.ThrowsException<ArgumentException>(() => fixture.Service.SetProperty(request));
            request.Property = "Name";
            request.Value = "2Bad";
            Assert.ThrowsException<ArgumentException>(() => fixture.Service.SetProperty(request));
            Assert.AreEqual("Form1", fixture.Form.Name);
        }

        [TestMethod]
        public void EventCatalogUsesExactFormOrCanonicalControlPath()
        {
            var fixture = NewFixture();
            fixture.Form.Designer.Controls.AddExisting("Label1");
            dynamic formEvents = fixture.Service.EventCatalog(fixture.Project.Name, fixture.Form.Name, null);
            Assert.AreEqual("UserForm", (string)formEvents.ObjectName);
            dynamic controlEvents = fixture.Service.EventCatalog(fixture.Project.Name, fixture.Form.Name, "Controls/Label1");
            Assert.AreEqual("Label1", (string)controlEvents.ObjectName);
            Assert.ThrowsException<InvalidOperationException>(() => fixture.Service.EventCatalog(fixture.Project.Name, fixture.Form.Name, "Label1"));
        }

        [TestMethod]
        public void ZOrderInvokesNativeControlButReportsUnverifiedEffect()
        {
            var fixture = NewFixture();
            var control = fixture.Form.Designer.Controls.AddExisting("Label1");
            dynamic tree = fixture.Service.Tree(fixture.Project.Name, fixture.Form.Name);
            var request = new Request
            {
                Project = fixture.Project.Name,
                Form = fixture.Form.Name,
                ControlPath = "Controls/Label1",
                ExpectedTreeVersion = tree.TreeVersion,
                ZPosition = 0
            };
            dynamic moved = fixture.Service.ZOrderControl(request);
            Assert.AreEqual(1, control.ZOrderCount);
            Assert.AreEqual(0, control.LastZPosition);
            Assert.AreEqual("Unverified", (string)moved.Verification);
            Assert.IsTrue((bool)moved.VerificationPending);
            request.ZPosition = 2;
            Assert.ThrowsException<ArgumentOutOfRangeException>(() => fixture.Service.ZOrderControl(request));
            request.ZPosition = 1;
            request.ExpectedTreeVersion = "stale";
            Assert.ThrowsException<InvalidOperationException>(() => fixture.Service.ZOrderControl(request));
            Assert.AreEqual(1, control.ZOrderCount);
        }

        [TestMethod]
        public void PageTabAndPictureCommandsRejectUnverifiedTargetsBeforeNativeWrites()
        {
            var fixture = NewFixture();
            fixture.Form.Designer.Controls.AddExisting("Frame1");
            dynamic tree = fixture.Service.Tree(fixture.Project.Name, fixture.Form.Name);
            var page = new Request
            {
                Project = fixture.Project.Name,
                Form = fixture.Form.Name,
                ParentPath = "Controls/Frame1",
                NewName = "Page1",
                ExpectedTreeVersion = tree.TreeVersion
            };
            Assert.ThrowsException<InvalidOperationException>(() => fixture.Service.AddPageOrTab(page, "Pages"));
            Assert.ThrowsException<InvalidOperationException>(() => fixture.Service.AddPageOrTab(page, "Tabs"));
            var remove = new Request
            {
                Project = fixture.Project.Name,
                Form = fixture.Form.Name,
                ControlPath = "Controls/Frame1",
                ExpectedTreeVersion = tree.TreeVersion
            };
            Assert.ThrowsException<ArgumentException>(() => fixture.Service.RemovePageOrTab(remove));
            var picture = new Request
            {
                Project = fixture.Project.Name,
                Form = fixture.Form.Name,
                ControlPath = "Controls/Frame1",
                Property = "Caption",
                ExpectedTreeVersion = tree.TreeVersion
            };
            Assert.ThrowsException<InvalidOperationException>(() => fixture.Service.SetNodePicture(picture));
            Assert.AreEqual(1, fixture.Form.Designer.Controls.Count);
        }
    }
}

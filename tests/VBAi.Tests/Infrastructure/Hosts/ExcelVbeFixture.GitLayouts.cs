using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Drawing;
using System.Drawing.Imaging;
using System.Linq;
using System.Runtime.InteropServices;
using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace VBAi.Tests.Integration
{
    internal sealed partial class ExcelVbeFixture
    {
        /// <summary>Creates one owned native form layout with synthetic controls and inert code.</summary>
        internal void PrepareGitLayout(string form, string layout, string path, bool persistedBaseline = false,
            string rootFontSeedProfile = null)
        {
            // Suppress host events and all document macros before saving/reopening
            // this owned fixture. No user application or trust setting is changed.
            ((dynamic)application).EnableEvents = false;
            ((dynamic)application).AutomationSecurity = 3;
            PrepareGitForm(form, "Local Git " + layout, "LOCAL_GIT_LAYOUT_" + layout, path, rootFontSeedProfile);
            if (layout != "LabelButton")
                WithGitLayoutDesigner(form, (component, designer) => {
                    SetGitFormProperty(component, "Width", 350d);
                    SetGitFormProperty(component, "Height", 350d);
                    object controls = null, control = null;
                    try
                    {
                        controls = ((dynamic)designer).Controls;
                        string kind = layout == "FrameMultiPage" ? "Frame" : layout;
                        control = ((dynamic)controls).Add("Forms." + kind + ".1", "QualificationExtra", true);
                        SetGitLayoutGeometry(control, 12d, 108d, 240d, layout == "FrameMultiPage" ? 170d : 44d);
                        ((dynamic)control).Tag = "Original " + layout;
                        switch (layout)
                        {
                            case "TextBox": ((dynamic)control).Text = "Original text"; break;
                            case "ComboBox":
                            case "ListBox":
                                if (persistedBaseline) ((dynamic)control).ColumnCount = 1;
                                else
                                {
                                    ((dynamic)control).AddItem("First synthetic row");
                                    ((dynamic)control).AddItem("Second synthetic row");
                                    ((dynamic)control).ListIndex = 0;
                                }
                                break;
                            case "CheckBox":
                            case "OptionButton":
                            case "ToggleButton":
                                ((dynamic)control).Caption = "Synthetic choice";
                                ((dynamic)control).Value = false; break;
                            case "ScrollBar":
                            case "SpinButton":
                                ((dynamic)control).Min = 0; ((dynamic)control).Max = 100;
                                ((dynamic)control).Value = 10; break;
                            case "TabStrip": SetGitLayoutTabCaption(control, "First synthetic tab"); break;
                            case "Image": break; // The owning host loads its own OLE picture below.
                            case "FrameMultiPage": PrepareNestedGitLayout(control); break;
                            default: throw new ArgumentException("Unknown local Git form layout: " + layout);
                        }
                    }
                    finally { Release(control); Release(controls); }
                });
            if (layout == "Image") InstallGitLayoutPicture(form);
            ((dynamic)workbook).Save();
            // Qualification on the exact saved designer must begin after native
            // persistence. Independent baseline trials proved that initial geometry,
            // PNG representation and AddItem runtime rows can change without any
            // Git import. Never ignore those differences in snapshot comparison.
            if (persistedBaseline) Assert.AreEqual(0, ReopenAndReadProjectProtection(path));
        }

        /// <summary>Loads a synthetic bitmap through the production command in Excel, without crossing a process-local GDI handle.</summary>
        private void InstallGitLayoutPicture(string form)
        {
            string imagePath = File("local-git-layout-picture.bmp");
            using (var bitmap = new Bitmap(4, 4))
            {
                using (var graphics = Graphics.FromImage(bitmap)) graphics.Clear(Color.Crimson);
                bitmap.Save(imagePath, ImageFormat.Bmp);
            }
            object project = null;
            string projectName;
            try { project = ((dynamic)workbook).VBProject; projectName = ((dynamic)project).Name; }
            finally { Release(project); }
            var tree = GitLayoutCommandData(Command(new { Command = "form_tree", Project = projectName, Form = form }));
            var image = ((object[])tree["Controls"]).Select(VbeBridgeClient.Object)
                .Single(node => Convert.ToString(node["Name"]) == "QualificationExtra");
            Assert.AreEqual("Control", image["Kind"]);
            var installed = GitLayoutCommandData(Command(new { Command = "set_form_node_picture", Project = projectName,
                Form = form, ControlPath = image["Path"], ExpectedTreeVersion = tree["TreeVersion"], Property = "Picture", Path = imagePath }));
            Assert.AreEqual(image["Path"], installed["ControlPath"]);
            Assert.AreEqual("Picture", installed["Property"]);
            Assert.IsNotNull(installed["Tree"]);
            // ReadGitLayout independently reads the installed native Picture's
            // type and dimensions before any snapshot/recovery acceptance.
        }

        private static IDictionary<string, object> GitLayoutCommandData(IDictionary<string, object> response)
        {
            Assert.IsNotNull(response, "The production picture command must answer on the owned host bridge.");
            object error; response.TryGetValue("Error", out error);
            Assert.AreEqual(true, response["Ok"], "Native layout command failed: " + Convert.ToString(error));
            return VbeBridgeClient.Object(response["Data"]);
        }

        /// <summary>Changes a native persisted value appropriate to the selected control layout.</summary>
        internal void MutateGitLayout(string form, string layout, bool persistedBaseline = false)
        {
            WithGitLayoutControl(form, layout, control => {
                switch (layout)
                {
                    case "LabelButton": ((dynamic)control).Caption = "Changed synthetic label"; break;
                    case "TextBox":
                    case "FrameMultiPage": ((dynamic)control).Text = "Changed synthetic text"; break;
                    case "ComboBox":
                    case "ListBox":
                        if (persistedBaseline) ((dynamic)control).ColumnCount = 2;
                        else ((dynamic)control).ListIndex = 1;
                        break;
                    case "CheckBox":
                    case "OptionButton":
                    case "ToggleButton": ((dynamic)control).Value = true; break;
                    case "ScrollBar":
                    case "SpinButton": ((dynamic)control).Value = 20; break;
                    case "TabStrip": SetGitLayoutTabCaption(control, "Changed synthetic tab"); break;
                    case "Image": ((dynamic)control).Tag = "Changed synthetic image tag"; break;
                    default: throw new ArgumentException("Unknown local Git form layout: " + layout);
                }
            });
        }

        /// <summary>Reads stable native form/control values independently of Git snapshot comparison.</summary>
        internal IDictionary<string, object> ReadGitLayout(string form, string layout)
        {
            var result = new SortedDictionary<string, object>(ReadGitForm(form), StringComparer.Ordinal);
            result["Layout"] = layout;
            WithGitLayoutDesigner(form, (component, designer) => {
                result["ComponentCaption"] = ReadGitFormProperty(component, "Caption");
                object controls = null;
                try
                {
                    controls = ((dynamic)designer).Controls;
                    ReadGitLayoutCollection(controls, designer, form, "Root", result);
                    foreach (string name in new[] { "QualificationLabel", "QualificationButton" })
                    {
                        object control = null;
                        try { control = ((dynamic)controls).Item(name); ReadGitLayoutParent(control, designer, form, name, result); }
                        finally { Release(control); }
                    }
                    if (layout != "LabelButton")
                    {
                        object control = null;
                        try { control = ((dynamic)controls).Item("QualificationExtra"); ReadGitLayoutParent(control, designer, form, "QualificationExtra", result); }
                        finally { Release(control); }
                    }
                }
                finally { Release(controls); }
                if (layout == "FrameMultiPage") ReadNestedGitLayout(designer, form, result);
            });
            WithGitLayoutControl(form, layout, control => {
                ReadGitLayoutGeometry(control, "Control", result);
                result["Control.Tag"] = ((dynamic)control).Tag;
                switch (layout)
                {
                    case "LabelButton": result["Control.Caption"] = ((dynamic)control).Caption; break;
                    case "TextBox":
                    case "FrameMultiPage": result["Control.Text"] = ((dynamic)control).Text; break;
                    case "ComboBox":
                    case "ListBox":
                        int count = Convert.ToInt32(((dynamic)control).ListCount);
                        result["Control.ListCount"] = count;
                        result["Control.ListIndex"] = ((dynamic)control).ListIndex;
                        result["Control.ColumnCount"] = ((dynamic)control).ColumnCount;
                        for (int i = 0; i < count; i++) result["Control.List." + i] = ((dynamic)control).List[i, 0];
                        break;
                    case "CheckBox":
                    case "OptionButton":
                    case "ToggleButton":
                        result["Control.Caption"] = ((dynamic)control).Caption;
                        result["Control.Value"] = ((dynamic)control).Value; break;
                    case "ScrollBar":
                    case "SpinButton":
                        result["Control.Min"] = ((dynamic)control).Min;
                        result["Control.Max"] = ((dynamic)control).Max;
                        result["Control.Value"] = ((dynamic)control).Value; break;
                    case "TabStrip":
                        object tabs = null;
                        try
                        {
                            tabs = ((dynamic)control).Tabs;
                            int tabsCount = Convert.ToInt32(((dynamic)tabs).Count);
                            result["Control.TabCount"] = tabsCount;
                            result["Control.Value"] = ((dynamic)control).Value;
                            for (int i = 0; i < tabsCount; i++)
                            {
                                object tab = null;
                                try { tab = ((dynamic)tabs).Item(i); result["Control.Tab." + i + ".Caption"] = ((dynamic)tab).Caption; }
                                finally { Release(tab); }
                            }
                        }
                        finally { Release(tabs); }
                        break;
                    case "Image":
                        // StdPicture is process-local; native v2 proved its GET
                        // cannot marshal to the external test process. The
                        // production tree reads the image inside Excel below.
                        break;
                    default: throw new ArgumentException("Unknown local Git form layout: " + layout);
                }
            });
            if (layout == "Image") ReadGitLayoutPicture(form, result);
            return result;
        }

        /// <summary>Reads the actual installed image content through the production in-host descriptor.</summary>
        private void ReadGitLayoutPicture(string form, IDictionary<string, object> result)
        {
            object project = null;
            string projectName;
            try { project = ((dynamic)workbook).VBProject; projectName = ((dynamic)project).Name; }
            finally { Release(project); }
            var tree = GitLayoutCommandData(Command(new { Command = "form_tree", Project = projectName, Form = form }));
            var image = ((object[])tree["Controls"]).Select(VbeBridgeClient.Object)
                .Single(node => Convert.ToString(node["Name"]) == "QualificationExtra");
            var picture = ((object[])image["Properties"]).Select(VbeBridgeClient.Object)
                .Single(property => Convert.ToString(property["Name"]) == "Picture");
            Assert.IsTrue(string.IsNullOrEmpty(Convert.ToString(picture["Error"])), "The in-host Picture descriptor failed: " + picture["Error"]);
            Assert.AreEqual("object", picture["Kind"]);
            Assert.AreEqual(typeof(Bitmap).FullName, picture["Type"]);
            string digest = Convert.ToString(picture["Digest"]);
            Assert.IsTrue(digest.Length == 64 && digest.All(character => "0123456789abcdef".Contains(character)),
                "The in-host descriptor must hash the actual installed image, not report an empty or unreadable Picture.");
            result["Control.Picture.Type"] = picture["Type"];
            result["Control.Picture.Digest"] = digest;
            result["Control.Picture.ReadbackScope"] = "Production form_tree in-host ImageDigest; PNG content SHA-256. No external OLE handle read.";
        }

        /// <summary>Saves the owned fixture and reopens it with macros/events disabled.</summary>
        internal void SaveAndReopenGitLayout(string path)
        {
            ((dynamic)application).EnableEvents = false;
            ((dynamic)application).AutomationSecurity = 3;
            ((dynamic)workbook).Save();
            Assert.AreEqual(0, ReopenAndReadProjectProtection(path));
        }

        private void WithGitLayoutDesigner(string form, Action<object, object> action)
        {
            object project = null, components = null, component = null, designer = null;
            try
            {
                project = ((dynamic)workbook).VBProject;
                components = ((dynamic)project).VBComponents;
                component = ((dynamic)components).Item(form);
                designer = ((dynamic)component).Designer;
                action(component, designer);
            }
            finally { Release(designer); Release(component); Release(components); Release(project); }
        }

        private void WithGitLayoutControl(string form, string layout, Action<object> action)
        {
            WithGitLayoutDesigner(form, (component, designer) => {
                object controls = null, control = null, nestedControls = null, multi = null, pages = null, page = null, pageControls = null, leaf = null;
                try
                {
                    controls = ((dynamic)designer).Controls;
                    control = ((dynamic)controls).Item(layout == "LabelButton" ? "QualificationLabel" : "QualificationExtra");
                    if (layout != "FrameMultiPage") { action(control); return; }
                    nestedControls = ((dynamic)control).Controls;
                    multi = ((dynamic)nestedControls).Item("QualificationMultiPage");
                    pages = ((dynamic)multi).Pages; page = ((dynamic)pages).Item(0);
                    pageControls = ((dynamic)page).Controls;
                    leaf = ((dynamic)pageControls).Item("QualificationNestedText");
                    action(leaf);
                }
                finally { Release(leaf); Release(pageControls); Release(page); Release(pages); Release(multi); Release(nestedControls); Release(control); Release(controls); }
            });
        }

        private static void SetGitLayoutGeometry(object control, double left, double top, double width, double height)
        {
            ((dynamic)control).Left = left; ((dynamic)control).Top = top;
            ((dynamic)control).Width = width; ((dynamic)control).Height = height;
        }

        private static void ReadGitLayoutGeometry(object control, string prefix, IDictionary<string, object> result)
        {
            result[prefix + ".Name"] = ((dynamic)control).Name;
            result[prefix + ".Left"] = ((dynamic)control).Left;
            result[prefix + ".Top"] = ((dynamic)control).Top;
            result[prefix + ".Width"] = ((dynamic)control).Width;
            result[prefix + ".Height"] = ((dynamic)control).Height;
        }

        private static void SetGitLayoutTabCaption(object control, string caption)
        {
            object tabs = null, tab = null;
            try { tabs = ((dynamic)control).Tabs; tab = ((dynamic)tabs).Item(0); ((dynamic)tab).Caption = caption; }
            finally { Release(tab); Release(tabs); }
        }

        private static void PrepareNestedGitLayout(object frame)
        {
            object controls = null, multi = null, pages = null, page = null, pageControls = null, text = null;
            try
            {
                ((dynamic)frame).Caption = "Synthetic frame";
                controls = ((dynamic)frame).Controls;
                multi = ((dynamic)controls).Add("Forms.MultiPage.1", "QualificationMultiPage", true);
                SetGitLayoutGeometry(multi, 6d, 18d, 220d, 140d);
                pages = ((dynamic)multi).Pages; page = ((dynamic)pages).Item(0);
                ((dynamic)page).Caption = "Synthetic page";
                pageControls = ((dynamic)page).Controls;
                text = ((dynamic)pageControls).Add("Forms.TextBox.1", "QualificationNestedText", true);
                SetGitLayoutGeometry(text, 8d, 10d, 170d, 24d);
                ((dynamic)text).Text = "Original nested text";
            }
            finally { Release(text); Release(pageControls); Release(page); Release(pages); Release(multi); Release(controls); }
        }

        private static void ReadNestedGitLayout(object designer, string form, IDictionary<string, object> result)
        {
            object controls = null, frame = null, nestedControls = null, multi = null, pages = null;
            try
            {
                controls = ((dynamic)designer).Controls;
                frame = ((dynamic)controls).Item("QualificationExtra");
                ReadGitLayoutGeometry(frame, "Frame", result);
                result["Frame.Caption"] = ((dynamic)frame).Caption;
                nestedControls = ((dynamic)frame).Controls;
                result["Frame.ControlCount"] = ((dynamic)nestedControls).Count;
                ReadGitLayoutCollection(nestedControls, frame, form, "Frame", result);
                multi = ((dynamic)nestedControls).Item("QualificationMultiPage");
                ReadGitLayoutGeometry(multi, "MultiPage", result);
                ReadGitLayoutParent(multi, frame, form, "MultiPage", result);
                result["MultiPage.Value"] = ((dynamic)multi).Value;
                pages = ((dynamic)multi).Pages;
                int count = Convert.ToInt32(((dynamic)pages).Count);
                result["MultiPage.PageCount"] = count;
                for (int i = 0; i < count; i++)
                {
                    object page = null, pageControls = null;
                    try
                    {
                        page = ((dynamic)pages).Item(i);
                        result["MultiPage.Page." + i + ".Name"] = ((dynamic)page).Name;
                        result["MultiPage.Page." + i + ".Caption"] = ((dynamic)page).Caption;
                        pageControls = ((dynamic)page).Controls;
                        result["MultiPage.Page." + i + ".ControlCount"] = ((dynamic)pageControls).Count;
                        ReadGitLayoutParent(page, multi, form, "MultiPage.Page." + i, result);
                        ReadGitLayoutCollection(pageControls, page, form, "MultiPage.Page." + i, result);
                        if (i == 0)
                        {
                            object leaf = null;
                            try
                            {
                                leaf = ((dynamic)pageControls).Item("QualificationNestedText");
                                ReadGitLayoutParent(leaf, page, form, "NestedText", result);
                                ReadGitLayoutGeometry(leaf, "NestedText", result);
                                result["NestedText.Text"] = ((dynamic)leaf).Text;
                            }
                            finally { Release(leaf); }
                        }
                    }
                    finally { Release(pageControls); Release(page); }
                }
            }
            finally { Release(pages); Release(multi); Release(nestedControls); Release(frame); Release(controls); }
        }

        /// <summary>Reads the exact collection membership and independently identifies direct children by their parents.</summary>
        private static void ReadGitLayoutCollection(object controls, object owner, string form, string prefix,
            IDictionary<string, object> result)
        {
            var names = new SortedSet<string>(StringComparer.Ordinal);
            var direct = new SortedSet<string>(StringComparer.Ordinal);
            string ownerPath = GitLayoutContainerIdentity(owner, form);
            int count = Convert.ToInt32(((dynamic)controls).Count);
            Assert.IsTrue(count >= 0 && count <= 32, "The disposable layout collection must remain bounded.");
            for (int i = 0; i < count; i++)
            {
                object control = null, parent = null;
                try
                {
                    control = ((dynamic)controls).Item(i);
                    string name = ((dynamic)control).Name;
                    Assert.IsTrue(names.Add(name), "Duplicate native layout control: " + name);
                    parent = ((dynamic)control).Parent;
                    if (VbeProjectHostPath.SameProject(parent, owner) || GitLayoutContainerIdentity(parent, form) == ownerPath)
                        direct.Add(name);
                }
                finally
                {
                    // Parent may be an alias of an outer container RCW. Balance
                    // this acquisition rather than final-releasing that container.
                    ReleaseGitLayoutAlias(parent); Release(control);
                }
            }
            result[prefix + ".ControlNames"] = string.Join("|", names);
            result[prefix + ".DirectControlNames"] = string.Join("|", direct);
        }

        /// <summary>Verifies a native parent by COM identity or the complete typed container chain used by form_tree.</summary>
        private static void ReadGitLayoutParent(object control, object expectedParent, string form, string prefix,
            IDictionary<string, object> result)
        {
            object parent = null;
            try
            {
                parent = ((dynamic)control).Parent;
                string observed = GitLayoutContainerIdentity(parent, form);
                string expected = GitLayoutContainerIdentity(expectedParent, form);
                bool sameCom = VbeProjectHostPath.SameProject(parent, expectedParent);
                Assert.IsTrue(sameCom || observed == expected, "Native parent changed for " + prefix + ": " + observed + ", expected " + expected);
                result[prefix + ".Type"] = TypeDescriptor.GetClassName(control);
                result[prefix + ".ParentPath"] = observed;
                result[prefix + ".ExpectedParentPath"] = expected;
                result[prefix + ".SameComParent"] = sameCom;
                result[prefix + ".ParentVerified"] = true;
            }
            finally { ReleaseGitLayoutAlias(parent); }
        }

        private static string GitLayoutContainerIdentity(object container, string form)
        {
            var path = new List<string>();
            object current = container;
            bool acquired = false;
            try
            {
                for (int depth = 0; depth < 16; depth++)
                {
                    Assert.IsNotNull(current, "A native container ancestry ended before the UserForm.");
                    string type = TypeDescriptor.GetClassName(current);
                    bool isForm = string.Equals(type, "UserForm", StringComparison.OrdinalIgnoreCase);
                    string name;
                    // The design-time UserForm wrapper can omit Name; the exact
                    // VBComponent form identity is already held by the fixture.
                    if (isForm) name = form;
                    else name = ((dynamic)current).Name;
                    Assert.IsFalse(string.IsNullOrWhiteSpace(type) || string.IsNullOrWhiteSpace(name), "Native container identity is incomplete.");
                    path.Add(type + ":" + name);
                    if (isForm) return string.Join("/", path);
                    object next = ((dynamic)current).Parent;
                    if (acquired) ReleaseGitLayoutAlias(current);
                    current = next; acquired = true;
                }
                Assert.Fail("Native layout ancestry exceeds the bounded depth.");
                return null;
            }
            finally { if (acquired) ReleaseGitLayoutAlias(current); }
        }

        private static void ReleaseGitLayoutAlias(object alias)
        {
            if (alias != null && Marshal.IsComObject(alias)) Marshal.ReleaseComObject(alias);
        }
    }
}

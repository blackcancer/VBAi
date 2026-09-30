using System;
using System.Collections.Generic;
using System.Drawing;
using System.Windows.Forms;
using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace VBAi.Tests.Integration
{
    internal sealed partial class ExcelVbeFixture
    {
        /// <summary>Creates one owned native form layout with synthetic controls and inert code.</summary>
        internal void PrepareGitLayout(string form, string layout, string path)
        {
            // Suppress host events and all document macros before saving/reopening
            // this owned fixture. No user application or trust setting is changed.
            ((dynamic)application).EnableEvents = false;
            ((dynamic)application).AutomationSecurity = 3;
            PrepareGitForm(form, "Local Git " + layout, "LOCAL_GIT_LAYOUT_" + layout, path);
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
                                ((dynamic)control).AddItem("First synthetic row");
                                ((dynamic)control).AddItem("Second synthetic row");
                                ((dynamic)control).ListIndex = 0; break;
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
                            case "Image":
                                object picture = null;
                                using (var bitmap = new Bitmap(4, 4))
                                {
                                    using (var graphics = Graphics.FromImage(bitmap)) graphics.Clear(Color.Crimson);
                                    try
                                    {
                                        picture = GitLayoutPicture.ToPicture(bitmap);
                                        ((dynamic)control).Picture = picture;
                                    }
                                    finally { Release(picture); }
                                }
                                break;
                            case "FrameMultiPage": PrepareNestedGitLayout(control); break;
                            default: throw new ArgumentException("Unknown local Git form layout: " + layout);
                        }
                    }
                    finally { Release(control); Release(controls); }
                });
            ((dynamic)workbook).Save();
        }

        /// <summary>Changes a native persisted value appropriate to the selected control layout.</summary>
        internal void MutateGitLayout(string form, string layout)
        {
            WithGitLayoutControl(form, layout, control => {
                switch (layout)
                {
                    case "LabelButton": ((dynamic)control).Caption = "Changed synthetic label"; break;
                    case "TextBox":
                    case "FrameMultiPage": ((dynamic)control).Text = "Changed synthetic text"; break;
                    case "ComboBox":
                    case "ListBox": ((dynamic)control).ListIndex = 1; break;
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
                if (layout == "FrameMultiPage") ReadNestedGitLayout(designer, result);
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
                        object picture = null;
                        try
                        {
                            picture = ((dynamic)control).Picture;
                            Assert.IsNotNull(picture, "The Image layout must retain an actual picture.");
                            result["Control.Picture.Type"] = ((dynamic)picture).Type;
                            result["Control.Picture.Width"] = ((dynamic)picture).Width;
                            result["Control.Picture.Height"] = ((dynamic)picture).Height;
                        }
                        finally { Release(picture); }
                        break;
                    default: throw new ArgumentException("Unknown local Git form layout: " + layout);
                }
            });
            return result;
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

        private static void ReadNestedGitLayout(object designer, IDictionary<string, object> result)
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
                multi = ((dynamic)nestedControls).Item("QualificationMultiPage");
                ReadGitLayoutGeometry(multi, "MultiPage", result);
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
                    }
                    finally { Release(pageControls); Release(page); }
                }
            }
            finally { Release(pages); Release(multi); Release(nestedControls); Release(frame); Release(controls); }
        }

        /// <summary>Exposes WinForms' native StdPicture conversion without activating an ActiveX control.</summary>
        private sealed class GitLayoutPicture : AxHost
        {
            private GitLayoutPicture() : base("00000000-0000-0000-0000-000000000000") { }
            internal static object ToPicture(Image image) { return GetIPictureDispFromPicture(image); }
        }
    }
}

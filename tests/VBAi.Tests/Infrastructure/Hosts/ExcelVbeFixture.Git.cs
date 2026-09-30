using System;
using System.Collections.Generic;
using System.Drawing;
using System.Drawing.Imaging;
using System.Runtime.InteropServices;
using System.Threading;
using System.Windows.Forms;
using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace VBAi.Tests.Integration
{
    internal sealed partial class ExcelVbeFixture
    {
        [StructLayout(LayoutKind.Sequential)]
        private struct GitCaptureRect { public int Left, Top, Right, Bottom; }
        [DllImport("user32.dll")] private static extern bool GetWindowRect(IntPtr window, out GitCaptureRect rectangle);
        [DllImport("user32.dll")] private static extern bool PrintWindow(IntPtr window, IntPtr dc, uint flags);

        internal void CaptureGitFormDesigner(string form, string path)
        {
            object project = null, components = null, component = null, window = null;
            try
            {
                project = ((dynamic)workbook).VBProject; components = ((dynamic)project).VBComponents;
                component = ((dynamic)components).Item(form); window = ((dynamic)component).DesignerWindow();
                ((dynamic)window).Visible = true; ((dynamic)window).SetFocus();
                IntPtr handle = new IntPtr(Convert.ToInt64(((dynamic)window).HWnd));
                uint pid; GetWindowThreadProcessId(handle, out pid); Assert.AreEqual((uint)ProcessId, pid);
                for (int i = 0; i < 10; i++) { Application.DoEvents(); Thread.Sleep(30); }
                GitCaptureRect rectangle; Assert.IsTrue(GetWindowRect(handle, out rectangle));
                int width = rectangle.Right - rectangle.Left, height = rectangle.Bottom - rectangle.Top;
                Assert.IsTrue(width > 0 && height > 0 && width <= 4096 && height <= 4096);
                using (var bitmap = new Bitmap(width, height))
                using (var graphics = Graphics.FromImage(bitmap))
                {
                    IntPtr dc = graphics.GetHdc();
                    try { Assert.IsTrue(PrintWindow(handle, dc, 2), "Owned native designer capture failed."); }
                    finally { graphics.ReleaseHdc(dc); }
                    bitmap.Save(path, ImageFormat.Png);
                }
            }
            finally { Release(window); Release(component); Release(components); Release(project); }
        }

        internal void PrepareGitForm(string form, string caption, string marker, string path)
        {
            object project = null, components = null, component = null, designer = null, controls = null, label = null, button = null, code = null;
            try
            {
                project = ((dynamic)workbook).VBProject;
                components = ((dynamic)project).VBComponents;
                component = ((dynamic)components).Add(3);
                ((dynamic)component).Name = form;
                designer = ((dynamic)component).Designer;
                ((dynamic)designer).Caption = caption;
                SetGitFormProperty(component, "Width", 260d); SetGitFormProperty(component, "Height", 180d);
                controls = ((dynamic)designer).Controls;
                label = ((dynamic)controls).Add("Forms.Label.1", "QualificationLabel", true);
                ((dynamic)label).Caption = caption;
                ((dynamic)label).Left = 12d; ((dynamic)label).Top = 18d;
                ((dynamic)label).Width = 170d; ((dynamic)label).Height = 24d;
                button = ((dynamic)controls).Add("Forms.CommandButton.1", "QualificationButton", true);
                ((dynamic)button).Caption = "Synthetic button";
                ((dynamic)button).Left = 12d; ((dynamic)button).Top = 60d;
                ((dynamic)button).Width = 100d; ((dynamic)button).Height = 28d;
                code = ((dynamic)component).CodeModule;
                ((dynamic)code).InsertLines(1, "Option Explicit\r\nPrivate Const ObservedMarker As String = \"" + marker + "\"\r\nPrivate Sub QualificationButton_Click()\r\n    ' Synthetic event body; never executed.\r\nEnd Sub");
                ((dynamic)workbook).SaveAs(path, 52);
            }
            finally { Release(code); Release(button); Release(label); Release(controls); Release(designer); Release(component); Release(components); Release(project); }
        }

        private static void SetGitFormProperty(object component, string name, object value)
        {
            object properties = null, property = null;
            try { properties = ((dynamic)component).Properties; property = ((dynamic)properties).Item(name); ((dynamic)property).Value = value; Assert.AreEqual(Convert.ToDouble(value), Convert.ToDouble(((dynamic)property).Value), 0.1, "Native form dimension readback: " + name); }
            finally { Release(property); Release(properties); }
        }
        private static object ReadGitFormProperty(object component, string name)
        {
            object properties = null, property = null;
            try { properties = ((dynamic)component).Properties; property = ((dynamic)properties).Item(name); return ((dynamic)property).Value; }
            finally { Release(property); Release(properties); }
        }

        internal void WithGitProject(string path, Action<VbaGitProject> action)
        {
            object project = ((dynamic)workbook).VBProject;
            try { action(new VbaGitProject(() => project, path)); }
            finally { Release(project); }
        }

        internal IDictionary<string, object> ReadGitForm(string form)
        {
            object project = null, components = null, component = null, designer = null, controls = null, label = null, button = null, code = null;
            try
            {
                project = ((dynamic)workbook).VBProject; components = ((dynamic)project).VBComponents;
                component = ((dynamic)components).Item(form); designer = ((dynamic)component).Designer;
                controls = ((dynamic)designer).Controls; label = ((dynamic)controls).Item("QualificationLabel"); button = ((dynamic)controls).Item("QualificationButton");
                code = ((dynamic)component).CodeModule;
                int lines = ((dynamic)code).CountOfLines;
                return new Dictionary<string, object> {
                    ["Name"] = ((dynamic)component).Name, ["Type"] = ((dynamic)component).Type,
                    ["Caption"] = ((dynamic)designer).Caption, ["Width"] = ReadGitFormProperty(component, "Width"), ["Height"] = ReadGitFormProperty(component, "Height"),
                    ["ControlCount"] = ((dynamic)controls).Count,
                    ["LabelCaption"] = ((dynamic)label).Caption, ["LabelLeft"] = ((dynamic)label).Left, ["LabelTop"] = ((dynamic)label).Top,
                    ["LabelWidth"] = ((dynamic)label).Width, ["LabelHeight"] = ((dynamic)label).Height,
                    ["ButtonCaption"] = ((dynamic)button).Caption, ["ButtonLeft"] = ((dynamic)button).Left, ["ButtonTop"] = ((dynamic)button).Top,
                    ["Code"] = lines == 0 ? "" : ((dynamic)code).Lines[1, lines]
                };
            }
            finally { Release(code); Release(button); Release(label); Release(controls); Release(designer); Release(component); Release(components); Release(project); }
        }
    }
}

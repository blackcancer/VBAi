using System;
using System.Collections.Generic;
using System.Drawing;
using System.Drawing.Imaging;
using System.Runtime.InteropServices;
using System.Threading;
using System.Web.Script.Serialization;
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
            object project = null, components = null, component = null, window = null, editor = null, mainWindow = null;
            var evidence = new Dictionary<string, object> { ["Form"] = form, ["ExpectedProcessId"] = ProcessId,
                ["State"] = "PENDING", ["CapturePath"] = path, ["Scope"] = "UNQUALIFIED" };
            var observations = new List<object>(); evidence["Observations"] = observations;
            try
            {
                project = ((dynamic)workbook).VBProject; components = ((dynamic)project).VBComponents;
                component = ((dynamic)components).Item(form); window = ((dynamic)component).DesignerWindow();
                ((dynamic)window).Visible = true; ((dynamic)window).SetFocus();
                editor = ((dynamic)application).VBE; mainWindow = ((dynamic)editor).MainWindow;
                GitDesignerCaptureState state = null;
                // HWnd=0 is an accepted root-capture fallback. Do not keep pumping
                // after its exact active designer is ready: that gives unrelated
                // UI activation a larger opportunity to replace the capture target.
                for (int i = 0; i < 20; i++)
                {
                    Application.DoEvents(); Thread.Sleep(30);
                    state = ReadGitDesignerCaptureState(project, window, editor, mainWindow);
                    observations.Add(state);
                    if (state.MainHandle != 0 && state.ProjectIdentityMatches && state.DesignerIdentityMatches &&
                        state.DesignerVisible && state.MainVisible && state.DesignerType == 1 && state.ActiveType == 1 &&
                        !string.IsNullOrWhiteSpace(state.DesignerCaption) && state.DesignerCaption == state.ActiveCaption) break;
                }
                IntPtr handle = SelectGitDesignerCaptureTarget(state, (uint)ProcessId, target => {
                    uint owner; GetWindowThreadProcessId(target, out owner); return owner;
                });
                evidence["Scope"] = state.DesignerHandle == 0 ? "OwnedVbeRootWithExactActiveDesigner" : "OwnedNativeDesignerWindow";
                evidence["CaptureHandle"] = handle.ToInt64();
                uint pid; GetWindowThreadProcessId(handle, out pid); Assert.AreEqual((uint)ProcessId, pid);
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
                var after = ReadGitDesignerCaptureState(project, window, editor, mainWindow);
                observations.Add(after);
                Assert.AreEqual(handle, SelectGitDesignerCaptureTarget(after, (uint)ProcessId, target => {
                    uint owner; GetWindowThreadProcessId(target, out owner); return owner;
                }), "The exact capture target changed while taking the screenshot.");
                evidence["State"] = "CAPTURED_PENDING_VISUAL_REVIEW";
            }
            catch (Exception error) { evidence["State"] = "FAILED"; evidence["Failure"] = error.ToString(); throw; }
            finally
            {
                try { System.IO.File.WriteAllText(path + ".json", new JavaScriptSerializer().Serialize(evidence)); }
                finally { Release(mainWindow); Release(editor); Release(window); Release(component); Release(components); Release(project); }
            }
        }

        /// <summary>Read-only observations used to qualify a designer capture or its owning VBE-root fallback.</summary>
        internal sealed class GitDesignerCaptureState
        {
            public long DesignerHandle { get; set; }
            public long MainHandle { get; set; }
            public string DesignerCaption { get; set; }
            public string ActiveCaption { get; set; }
            public int DesignerType { get; set; }
            public int ActiveType { get; set; }
            public bool DesignerVisible { get; set; }
            public bool MainVisible { get; set; }
            public bool ProjectIdentityMatches { get; set; }
            public bool DesignerIdentityMatches { get; set; }
        }

        private static GitDesignerCaptureState ReadGitDesignerCaptureState(object project, object window, object editor, object mainWindow)
        {
            object activeProject = null, activeWindow = null;
            try
            {
                activeProject = ((dynamic)editor).ActiveVBProject;
                activeWindow = ((dynamic)editor).ActiveWindow;
                return new GitDesignerCaptureState {
                    DesignerHandle = Convert.ToInt64(((dynamic)window).HWnd), MainHandle = Convert.ToInt64(((dynamic)mainWindow).HWnd),
                    DesignerCaption = Convert.ToString(((dynamic)window).Caption),
                    ActiveCaption = activeWindow == null ? null : Convert.ToString(((dynamic)activeWindow).Caption),
                    DesignerType = Convert.ToInt32(((dynamic)window).Type), ActiveType = activeWindow == null ? -1 : Convert.ToInt32(((dynamic)activeWindow).Type),
                    DesignerVisible = Convert.ToBoolean(((dynamic)window).Visible), MainVisible = Convert.ToBoolean(((dynamic)mainWindow).Visible),
                    ProjectIdentityMatches = VbeProjectHostPath.SameProject(project, activeProject),
                    DesignerIdentityMatches = VbeProjectHostPath.SameProject(window, activeWindow)
                };
            }
            finally
            {
                // These getters can return the same RCWs held by the caller. Balance
                // only the acquired references; FinalRelease would invalidate aliases.
                if (activeWindow != null && Marshal.IsComObject(activeWindow)) Marshal.ReleaseComObject(activeWindow);
                if (activeProject != null && Marshal.IsComObject(activeProject)) Marshal.ReleaseComObject(activeProject);
            }
        }

        /// <summary>Pure capture boundary: HWnd=0 permits only the verified root and exact active designer identity.</summary>
        internal static IntPtr SelectGitDesignerCaptureTarget(GitDesignerCaptureState state, uint expectedPid, Func<IntPtr, uint> owner)
        {
            Assert.IsNotNull(state); Assert.IsNotNull(owner); Assert.IsTrue(expectedPid != 0);
            Assert.IsTrue(state.ProjectIdentityMatches && state.DesignerIdentityMatches, "The exact owned form/project is not active.");
            Assert.IsTrue(state.DesignerVisible && state.MainVisible, "The owned designer/VBE root is not visible.");
            Assert.AreEqual(1, state.DesignerType, "The requested VBIDE window is not a form designer.");
            Assert.AreEqual(state.DesignerType, state.ActiveType);
            Assert.IsFalse(string.IsNullOrWhiteSpace(state.DesignerCaption));
            Assert.AreEqual(state.DesignerCaption, state.ActiveCaption);
            var root = new IntPtr(state.MainHandle);
            Assert.IsTrue(root != IntPtr.Zero, "The owning VBE root has no native handle.");
            Assert.AreEqual(expectedPid, owner(root), "The VBE root belongs to another process.");
            var target = state.DesignerHandle == 0 ? root : new IntPtr(state.DesignerHandle);
            Assert.AreEqual(expectedPid, owner(target), "The native capture target belongs to another process.");
            return target;
        }

        internal void PrepareGitForm(string form, string caption, string marker, string path, string rootFontSeedProfile = null)
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
                if (rootFontSeedProfile != null)
                    LoadGitFontOnce(designer, "Form.Font", RootFontObservationManifest.SyntheticArial9Values(rootFontSeedProfile), true);
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

        /// <summary>Owns a project wrapper independently of temporary native readback aliases.</summary>
        private object OwnGitProjectRcw()
        {
            // Workbook.VBProject returns a shared RCW. Nested fixture reads can release
            // that wrapper; only the unique wrapper acquired here belongs to this scope.
            return AcquireIndependentGitProject(() => ((dynamic)workbook).VBProject,
                Marshal.GetIUnknownForObject, Marshal.GetUniqueObjectForIUnknown,
                identity => Marshal.Release(identity),
                alias => { if (Marshal.IsComObject(alias)) Marshal.ReleaseComObject(alias); });
        }

        /// <summary>Balances the shared acquisition and preserves every failure before handing off a unique wrapper.</summary>
        internal static object AcquireIndependentGitProject(Func<object> borrowedProject, Func<object, IntPtr> identify,
            Func<IntPtr, object> uniqueProject, Action<IntPtr> releaseIdentity, Action<object> releaseWrapper)
        {
            object borrowed = null, unique = null;
            IntPtr identity = IntPtr.Zero;
            var failures = new List<Exception>();
            try
            {
                borrowed = borrowedProject();
                identity = identify(borrowed);
                unique = uniqueProject(identity);
            }
            catch (Exception error) { failures.Add(error); }
            if (identity != IntPtr.Zero)
                try { releaseIdentity(identity); } catch (Exception error) { failures.Add(error); }
            if (borrowed != null)
                try { releaseWrapper(borrowed); } catch (Exception error) { failures.Add(error); }
            if (failures.Count != 0)
            {
                // A wrapper whose handoff failed remains ours, even when another
                // release failed. Shared and unique acquisitions are balanced once.
                if (unique != null)
                    try { releaseWrapper(unique); } catch (Exception error) { failures.Add(error); }
                if (failures.Count == 1) System.Runtime.ExceptionServices.ExceptionDispatchInfo.Capture(failures[0]).Throw();
                throw new AggregateException("Independent Git project acquisition and release failed.", failures);
            }
            return unique;
        }

        internal void WithGitProject(string path, Action<VbaGitProject> action)
        {
            object project = OwnGitProjectRcw();
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

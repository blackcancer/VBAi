using Microsoft.VisualStudio.TestTools.UnitTesting;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.Drawing.Imaging;
using System.Runtime.InteropServices;
using System.Threading;
using System.Web.Script.Serialization;
using System.Windows.Forms;

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
            var evidence = new Dictionary<string, object>
            {
                ["Form"] = form,
                ["ExpectedProcessId"] = ProcessId,
                ["State"] = "PENDING",
                ["CapturePath"] = path,
                ["Scope"] = "UNQUALIFIED"
            };
            var observations = new List<object>(); evidence["Observations"] = observations;
            try
            {
                project = NativeFixtureProgressTrace.Read<object>(NativeFixtureProgressTrace.Phase.CaptureProject, () => ((dynamic)workbook).VBProject);
                components = NativeFixtureProgressTrace.Read<object>(NativeFixtureProgressTrace.Phase.CaptureComponents, () => ((dynamic)project).VBComponents);
                component = NativeFixtureProgressTrace.Read<object>(NativeFixtureProgressTrace.Phase.CaptureComponent, () => ((dynamic)components).Item(form));
                window = NativeFixtureProgressTrace.Read<object>(NativeFixtureProgressTrace.Phase.DesignerWindow, () => ((dynamic)component).DesignerWindow());
                NativeFixtureProgressTrace.Run(NativeFixtureProgressTrace.Phase.DesignerVisible, () => ((dynamic)window).Visible = true);
                NativeFixtureProgressTrace.Run(NativeFixtureProgressTrace.Phase.DesignerFocus, () => ((dynamic)window).SetFocus());
                editor = NativeFixtureProgressTrace.Read<object>(NativeFixtureProgressTrace.Phase.CaptureVbe, () => ((dynamic)application).VBE);
                mainWindow = NativeFixtureProgressTrace.Read<object>(NativeFixtureProgressTrace.Phase.CaptureMain, () => ((dynamic)editor).MainWindow);
                GitDesignerCaptureState state = null;
                // HWnd=0 is an accepted root-capture fallback. Do not keep pumping
                // after its exact active designer is ready: that gives unrelated
                // UI activation a larger opportunity to replace the capture target.
                for (int i = 0; i < 20; i++)
                {
                    NativeFixtureProgressTrace.Run(NativeFixtureProgressTrace.Phase.PumpEvents, Application.DoEvents);
                    NativeFixtureProgressTrace.Run(NativeFixtureProgressTrace.Phase.PumpDelay, () => Thread.Sleep(30));
                    state = NativeFixtureProgressTrace.Read(NativeFixtureProgressTrace.Phase.CaptureState, () => ReadGitDesignerCaptureState(project, window, editor, mainWindow));
                    observations.Add(state);
                    if (state.MainHandle != 0 && state.ProjectIdentityMatches && state.DesignerIdentityMatches &&
                        state.DesignerVisible && state.MainVisible && state.DesignerType == 1 && state.ActiveType == 1 &&
                        !string.IsNullOrWhiteSpace(state.DesignerCaption) && state.DesignerCaption == state.ActiveCaption) break;
                }
                IntPtr handle = NativeFixtureProgressTrace.Read(NativeFixtureProgressTrace.Phase.SelectCaptureTarget, () => SelectGitDesignerCaptureTarget(state, (uint)ProcessId, target =>
                {
                    uint owner; GetWindowThreadProcessId(target, out owner); return owner;
                }));
                evidence["Scope"] = state.DesignerHandle == 0 ? "OwnedVbeRootWithExactActiveDesigner" : "OwnedNativeDesignerWindow";
                evidence["CaptureHandle"] = handle.ToInt64();
                uint pid; GetWindowThreadProcessId(handle, out pid); Assert.AreEqual((uint)ProcessId, pid);
                GitCaptureRect rectangle = default(GitCaptureRect); Assert.IsTrue(NativeFixtureProgressTrace.Read(NativeFixtureProgressTrace.Phase.CaptureRectangle, () => GetWindowRect(handle, out rectangle)));
                int width = rectangle.Right - rectangle.Left, height = rectangle.Bottom - rectangle.Top;
                Assert.IsTrue(width > 0 && height > 0 && width <= 4096 && height <= 4096);
                using (var bitmap = NativeFixtureProgressTrace.Read(NativeFixtureProgressTrace.Phase.CreateBitmap, () => new Bitmap(width, height)))
                using (var graphics = NativeFixtureProgressTrace.Read(NativeFixtureProgressTrace.Phase.CreateGraphics, () => Graphics.FromImage(bitmap)))
                {
                    IntPtr dc = NativeFixtureProgressTrace.Read(NativeFixtureProgressTrace.Phase.AcquireHdc, graphics.GetHdc);
                    try { Assert.IsTrue(NativeFixtureProgressTrace.Read(NativeFixtureProgressTrace.Phase.PrintWindow, () => PrintWindow(handle, dc, 2)), "Owned native designer capture failed."); }
                    finally { NativeFixtureProgressTrace.Run(NativeFixtureProgressTrace.Phase.ReleaseHdc, () => graphics.ReleaseHdc(dc)); }
                    NativeFixtureProgressTrace.Run(NativeFixtureProgressTrace.Phase.SavePng, () => bitmap.Save(path, ImageFormat.Png));
                }
                var after = NativeFixtureProgressTrace.Read(NativeFixtureProgressTrace.Phase.CaptureStateAfter, () => ReadGitDesignerCaptureState(project, window, editor, mainWindow));
                observations.Add(after);
                Assert.AreEqual(handle, NativeFixtureProgressTrace.Read(NativeFixtureProgressTrace.Phase.SelectCaptureTargetAfter, () => SelectGitDesignerCaptureTarget(after, (uint)ProcessId, target =>
                {
                    uint owner; GetWindowThreadProcessId(target, out owner); return owner;
                })), "The exact capture target changed while taking the screenshot.");
                evidence["State"] = "CAPTURED_PENDING_VISUAL_REVIEW";
            }
            catch (Exception error) { evidence["State"] = "FAILED"; evidence["Failure"] = error.ToString(); throw; }
            finally
            {
                try { NativeFixtureProgressTrace.Run(NativeFixtureProgressTrace.Phase.CaptureEvidence, () => System.IO.File.WriteAllText(path + ".json", new JavaScriptSerializer().Serialize(evidence))); }
                finally { NativeFixtureProgressTrace.Run(NativeFixtureProgressTrace.Phase.ReleaseMain, () => Release(mainWindow)); NativeFixtureProgressTrace.Run(NativeFixtureProgressTrace.Phase.ReleaseVbe, () => Release(editor)); NativeFixtureProgressTrace.Run(NativeFixtureProgressTrace.Phase.ReleaseWindow, () => Release(window)); NativeFixtureProgressTrace.Run(NativeFixtureProgressTrace.Phase.ReleaseComponent, () => Release(component)); NativeFixtureProgressTrace.Run(NativeFixtureProgressTrace.Phase.ReleaseComponents, () => Release(components)); NativeFixtureProgressTrace.Run(NativeFixtureProgressTrace.Phase.ReleaseProject, () => Release(project)); }
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
                activeProject = NativeFixtureProgressTrace.Read<object>(NativeFixtureProgressTrace.Phase.ActiveProject, () => ((dynamic)editor).ActiveVBProject);
                activeWindow = NativeFixtureProgressTrace.Read<object>(NativeFixtureProgressTrace.Phase.ActiveWindow, () => ((dynamic)editor).ActiveWindow);
                return new GitDesignerCaptureState
                {
                    DesignerHandle = NativeFixtureProgressTrace.Read<long>(NativeFixtureProgressTrace.Phase.DesignerHandle, () => Convert.ToInt64(((dynamic)window).HWnd)),
                    MainHandle = NativeFixtureProgressTrace.Read<long>(NativeFixtureProgressTrace.Phase.MainHandle, () => Convert.ToInt64(((dynamic)mainWindow).HWnd)),
                    DesignerCaption = NativeFixtureProgressTrace.Read<string>(NativeFixtureProgressTrace.Phase.DesignerCaption, () => Convert.ToString(((dynamic)window).Caption)),
                    ActiveCaption = NativeFixtureProgressTrace.Read<string>(NativeFixtureProgressTrace.Phase.ActiveCaption, () => activeWindow == null ? null : Convert.ToString(((dynamic)activeWindow).Caption)),
                    DesignerType = NativeFixtureProgressTrace.Read<int>(NativeFixtureProgressTrace.Phase.DesignerType, () => Convert.ToInt32(((dynamic)window).Type)),
                    ActiveType = NativeFixtureProgressTrace.Read<int>(NativeFixtureProgressTrace.Phase.ActiveType, () => activeWindow == null ? -1 : Convert.ToInt32(((dynamic)activeWindow).Type)),
                    DesignerVisible = NativeFixtureProgressTrace.Read<bool>(NativeFixtureProgressTrace.Phase.DesignerVisible, () => Convert.ToBoolean(((dynamic)window).Visible)),
                    MainVisible = NativeFixtureProgressTrace.Read<bool>(NativeFixtureProgressTrace.Phase.MainVisible, () => Convert.ToBoolean(((dynamic)mainWindow).Visible)),
                    ProjectIdentityMatches = NativeFixtureProgressTrace.Read<bool>(NativeFixtureProgressTrace.Phase.ProjectIdentity, () => VbeProjectHostPath.SameProject(project, activeProject)),
                    DesignerIdentityMatches = NativeFixtureProgressTrace.Read<bool>(NativeFixtureProgressTrace.Phase.DesignerIdentity, () => VbeProjectHostPath.SameProject(window, activeWindow))
                };
            }
            finally
            {
                // These getters can return the same RCWs held by the caller. Balance
                // only the acquired references; FinalRelease would invalidate aliases.
                if (activeWindow != null && Marshal.IsComObject(activeWindow)) NativeFixtureProgressTrace.Run(NativeFixtureProgressTrace.Phase.ReleaseActiveWindow, () => Marshal.ReleaseComObject(activeWindow));
                if (activeProject != null && Marshal.IsComObject(activeProject)) NativeFixtureProgressTrace.Run(NativeFixtureProgressTrace.Phase.ReleaseActiveProject, () => Marshal.ReleaseComObject(activeProject));
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
                project = NativeFixtureProgressTrace.Read<object>(NativeFixtureProgressTrace.Phase.FormProject, () => ((dynamic)workbook).VBProject);
                components = NativeFixtureProgressTrace.Read<object>(NativeFixtureProgressTrace.Phase.FormComponents, () => ((dynamic)project).VBComponents);
                component = NativeFixtureProgressTrace.Read<object>(NativeFixtureProgressTrace.Phase.FormAdd, () => ((dynamic)components).Add(3));
                NativeFixtureProgressTrace.Run(NativeFixtureProgressTrace.Phase.FormName, () => ((dynamic)component).Name = form);
                designer = NativeFixtureProgressTrace.Read<object>(NativeFixtureProgressTrace.Phase.FormDesigner, () => ((dynamic)component).Designer);
                NativeFixtureProgressTrace.Run(NativeFixtureProgressTrace.Phase.FormCaption, () => ((dynamic)designer).Caption = caption);
                NativeFixtureProgressTrace.Run(NativeFixtureProgressTrace.Phase.FormWidth, () => SetGitFormProperty(component, "Width", 260d)); NativeFixtureProgressTrace.Run(NativeFixtureProgressTrace.Phase.FormHeight, () => SetGitFormProperty(component, "Height", 180d));
                controls = NativeFixtureProgressTrace.Read<object>(NativeFixtureProgressTrace.Phase.FormControls, () => ((dynamic)designer).Controls);
                label = NativeFixtureProgressTrace.Read<object>(NativeFixtureProgressTrace.Phase.LabelAdd, () => ((dynamic)controls).Add("Forms.Label.1", "QualificationLabel", true));
                NativeFixtureProgressTrace.Run(NativeFixtureProgressTrace.Phase.LabelCaption, () => ((dynamic)label).Caption = caption);
                NativeFixtureProgressTrace.Run(NativeFixtureProgressTrace.Phase.LabelLeft, () => ((dynamic)label).Left = 12d); NativeFixtureProgressTrace.Run(NativeFixtureProgressTrace.Phase.LabelTop, () => ((dynamic)label).Top = 18d);
                NativeFixtureProgressTrace.Run(NativeFixtureProgressTrace.Phase.LabelWidth, () => ((dynamic)label).Width = 170d); NativeFixtureProgressTrace.Run(NativeFixtureProgressTrace.Phase.LabelHeight, () => ((dynamic)label).Height = 24d);
                button = NativeFixtureProgressTrace.Read<object>(NativeFixtureProgressTrace.Phase.ButtonAdd, () => ((dynamic)controls).Add("Forms.CommandButton.1", "QualificationButton", true));
                NativeFixtureProgressTrace.Run(NativeFixtureProgressTrace.Phase.ButtonCaption, () => ((dynamic)button).Caption = "Synthetic button");
                NativeFixtureProgressTrace.Run(NativeFixtureProgressTrace.Phase.ButtonLeft, () => ((dynamic)button).Left = 12d); NativeFixtureProgressTrace.Run(NativeFixtureProgressTrace.Phase.ButtonTop, () => ((dynamic)button).Top = 60d);
                NativeFixtureProgressTrace.Run(NativeFixtureProgressTrace.Phase.ButtonWidth, () => ((dynamic)button).Width = 100d); NativeFixtureProgressTrace.Run(NativeFixtureProgressTrace.Phase.ButtonHeight, () => ((dynamic)button).Height = 28d);
                code = NativeFixtureProgressTrace.Read<object>(NativeFixtureProgressTrace.Phase.FormCodeModule, () => ((dynamic)component).CodeModule);
                NativeFixtureProgressTrace.Run(NativeFixtureProgressTrace.Phase.FormCodeInsert, () => ((dynamic)code).InsertLines(1, "Option Explicit\r\nPrivate Const ObservedMarker As String = \"" + marker + "\"\r\nPrivate Sub QualificationButton_Click()\r\n    ' Synthetic event body; never executed.\r\nEnd Sub"));
                if (rootFontSeedProfile != null)
                    NativeFixtureProgressTrace.Run(NativeFixtureProgressTrace.Phase.FormSeedFont, () => LoadGitFontOnce(designer, "Form.Font", RootFontObservationManifest.SyntheticArial9Values(rootFontSeedProfile), true));
                NativeFixtureProgressTrace.Run(NativeFixtureProgressTrace.Phase.InitialSaveAs, () => ((dynamic)workbook).SaveAs(path, 52));
            }
            finally { NativeFixtureProgressTrace.Run(NativeFixtureProgressTrace.Phase.ReleaseCode, () => Release(code)); NativeFixtureProgressTrace.Run(NativeFixtureProgressTrace.Phase.ReleaseButton, () => Release(button)); NativeFixtureProgressTrace.Run(NativeFixtureProgressTrace.Phase.ReleaseLabel, () => Release(label)); NativeFixtureProgressTrace.Run(NativeFixtureProgressTrace.Phase.ReleaseControls, () => Release(controls)); NativeFixtureProgressTrace.Run(NativeFixtureProgressTrace.Phase.ReleaseDesigner, () => Release(designer)); NativeFixtureProgressTrace.Run(NativeFixtureProgressTrace.Phase.ReleaseComponent, () => Release(component)); NativeFixtureProgressTrace.Run(NativeFixtureProgressTrace.Phase.ReleaseComponents, () => Release(components)); NativeFixtureProgressTrace.Run(NativeFixtureProgressTrace.Phase.ReleaseProject, () => Release(project)); }
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
                return new Dictionary<string, object>
                {
                    ["Name"] = ((dynamic)component).Name,
                    ["Type"] = ((dynamic)component).Type,
                    ["Caption"] = ((dynamic)designer).Caption,
                    ["Width"] = ReadGitFormProperty(component, "Width"),
                    ["Height"] = ReadGitFormProperty(component, "Height"),
                    ["ControlCount"] = ((dynamic)controls).Count,
                    ["LabelCaption"] = ((dynamic)label).Caption,
                    ["LabelLeft"] = ((dynamic)label).Left,
                    ["LabelTop"] = ((dynamic)label).Top,
                    ["LabelWidth"] = ((dynamic)label).Width,
                    ["LabelHeight"] = ((dynamic)label).Height,
                    ["ButtonCaption"] = ((dynamic)button).Caption,
                    ["ButtonLeft"] = ((dynamic)button).Left,
                    ["ButtonTop"] = ((dynamic)button).Top,
                    ["Code"] = lines == 0 ? "" : ((dynamic)code).Lines[1, lines]
                };
            }
            finally { Release(code); Release(button); Release(label); Release(controls); Release(designer); Release(component); Release(components); Release(project); }
        }
    }
}

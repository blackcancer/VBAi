using System;
using System.Diagnostics;
using System.Drawing;
using System.Linq;
using System.Runtime.InteropServices;
using System.Runtime.ExceptionServices;
using System.Threading;

namespace VBAi
{
    /// <summary>Initializes an imported designer only when its persisted resources differ from the declared target.</summary>
    internal static class ImportedFormMaterialization
    {
        /// <summary>Preserves exact resources and returns only font owners still different after one initialization.</summary>
        internal static FormStreamPadding.FormFontBinding[] Prepare(VbaGitSnapshot target, VbaGitComponent component,
            FormStreamPadding.FormFontBinding[] bindings, Func<VbaGitSnapshot> capture, Action materialize, Action revalidate)
        {
            if (bindings == null || bindings.Length == 0) return bindings;
            revalidate();
            VbaGitSnapshot actual = capture();
            string resource = component.Name + ".frx";
            if (target.SameFile(actual, resource)) return new FormStreamPadding.FormFontBinding[0];
            revalidate();
            materialize();
            revalidate();
            actual = capture();
            if (target.SameFile(actual, resource)) return new FormStreamPadding.FormFontBinding[0];
            VbaGitComponent observedComponent = actual.Manifest.Components.Single(item => item.Name == component.Name && item.Type == 3);
            var observed = actual.FormFonts(observedComponent);
            if (observed == null)
                throw new InvalidOperationException("The imported form resource graph cannot be verified after designer initialization.");
            return bindings.Where(expected => !observed.Any(value => value.Type == expected.Type &&
                value.OwnerPath == expected.OwnerPath && value.Descriptor.SequenceEqual(expected.Descriptor))).ToArray();
        }

        /// <summary>Synchronously initializes only the verified designer on its actual owning VBE STA.</summary>
        internal static void Materialize(object project, object component, Action revalidate)
        {
            FormFontRestoration.RequireOwner(project);
            object window = null, editor = null, main = null, previous = null;
            bool originalVisible = false, originalViewKnown = false;
            Exception primary = null;
            try
            {
                revalidate();
                editor = ((dynamic)project).VBE;
                main = ((dynamic)editor).MainWindow;
                previous = ((dynamic)editor).ActiveWindow;
                window = ((dynamic)component).DesignerWindow();
                originalVisible = Convert.ToBoolean(((dynamic)window).Visible);
                originalViewKnown = true;
                revalidate();
                ((dynamic)window).Visible = true;
                revalidate();
                ((dynamic)window).SetFocus();
                // The caller is inside a native import transaction, possibly under
                // a modal Git dialog. Do not pump unrelated WinForms messages or
                // force-enable the modal owner. PrintWindow renders synchronously.
                Render(project, window, editor, main, revalidate);
            }
            catch (Exception error) { primary = error; throw; }
            finally
            {
                Exception restoreFailure = null;
                try
                {
                    if (originalViewKnown)
                        RestoreView(originalVisible, previous != null && !VbeProjectHostPath.SameProject(previous, window),
                            () => IsActive(editor, window), () => IsActive(editor, previous),
                            () => ((dynamic)previous).SetFocus(), () => { ((dynamic)window).Visible = false; }, revalidate);
                }
                catch (Exception error) { restoreFailure = error; }
                Exception combined = restoreFailure == null ? primary : primary == null ? restoreFailure :
                    new AggregateException("Imported designer initialization and navigation restoration failed.", primary, restoreFailure);
                FormFontRestoration.ReleaseOwnedReferences(new[] { window, editor, main, previous }, value => {
                    if (value != null && Marshal.IsComObject(value)) Marshal.ReleaseComObject(value);
                }, combined);
                if (restoreFailure != null) ExceptionDispatchInfo.Capture(combined).Throw();
            }
        }

        /// <summary>Restores the prior view only while the expected native navigation remains current.</summary>
        internal static void RestoreView(bool originallyVisible, bool previousDistinct, Func<bool> designerActive,
            Func<bool> previousActive, Action focusPrevious, Action hideDesigner, Action revalidate)
        {
            revalidate();
            if (!designerActive()) return;
            if (previousDistinct)
            {
                focusPrevious();
                revalidate();
                if (!previousActive()) return;
            }
            else
            {
                revalidate();
                if (!designerActive()) return;
            }
            if (!originallyVisible) hideDesigner();
        }

        /// <summary>Balances the acquired active-window alias without changing user navigation.</summary>
        private static bool IsActive(object editor, object expected)
        {
            object current = null;
            try
            {
                current = ((dynamic)editor).ActiveWindow;
                return expected != null && VbeProjectHostPath.SameProject(current, expected);
            }
            finally { if (current != null && Marshal.IsComObject(current)) Marshal.ReleaseComObject(current); }
        }

        /// <summary>Reads exact window/project ownership around one bounded native render.</summary>
        private static void Render(object project, object window, object editor, object main, Action revalidate)
        {
            revalidate();
            IntPtr target = ReadTarget(project, window, editor, main);
            Rectangle native;
            if (!GetWindowRect(target, out native)) throw new InvalidOperationException("The imported designer bounds are unavailable.");
            int width = checked(native.Right - native.Left), height = checked(native.Bottom - native.Top);
            if (width <= 0 || height <= 0 || width > 4096 || height > 4096)
                throw new InvalidOperationException("The imported designer render exceeds the bounded native surface.");
            using (var bitmap = new Bitmap(width, height))
            using (var graphics = Graphics.FromImage(bitmap))
            {
                IntPtr dc = graphics.GetHdc();
                try
                {
                    revalidate();
                    if (ReadTarget(project, window, editor, main) != target)
                        throw new InvalidOperationException("The imported designer render target changed.");
                    if (!PrintWindow(target, dc, 2))
                        throw new InvalidOperationException("The imported designer could not be initialized by native rendering.");
                }
                finally { graphics.ReleaseHdc(dc); }
            }
            revalidate();
            if (ReadTarget(project, window, editor, main) != target)
                throw new InvalidOperationException("The imported designer changed during native initialization.");
        }

        /// <summary>Refuses a foreign, hidden or inactive designer before acquiring its drawing surface.</summary>
        private static IntPtr ReadTarget(object project, object window, object editor, object main)
        {
            object activeProject = null, activeWindow = null;
            try
            {
                activeProject = ((dynamic)editor).ActiveVBProject;
                activeWindow = ((dynamic)editor).ActiveWindow;
                return SelectTarget(new IntPtr(Convert.ToInt64(((dynamic)main).HWnd)),
                    new IntPtr(Convert.ToInt64(((dynamic)window).HWnd)),
                    VbeProjectHostPath.SameProject(project, activeProject), VbeProjectHostPath.SameProject(window, activeWindow),
                    (bool)Convert.ToBoolean(((dynamic)main).Visible) && (bool)Convert.ToBoolean(((dynamic)window).Visible),
                    (int)Convert.ToInt32(((dynamic)window).Type),
                    (uint)Process.GetCurrentProcess().Id, handle => {
                        uint pid; GetWindowThreadProcessId(handle, out pid); return pid;
                    });
            }
            finally
            {
                if (activeWindow != null && Marshal.IsComObject(activeWindow)) Marshal.ReleaseComObject(activeWindow);
                if (activeProject != null && Marshal.IsComObject(activeProject)) Marshal.ReleaseComObject(activeProject);
            }
        }

        /// <summary>Allows the owning VBE root only when the exact active designer has no HWND.</summary>
        internal static IntPtr SelectTarget(IntPtr root, IntPtr designer, bool projectMatches, bool designerMatches,
            bool visible, int designerType, uint expectedPid, Func<IntPtr, uint> owner)
        {
            if (Thread.CurrentThread.GetApartmentState() != ApartmentState.STA || expectedPid == 0 ||
                root == IntPtr.Zero || !projectMatches || !designerMatches || !visible || designerType != 1 || owner(root) != expectedPid)
                throw new InvalidOperationException("Only the exact active owned form designer can be initialized.");
            IntPtr target = designer == IntPtr.Zero ? root : designer;
            if (owner(target) != expectedPid)
                throw new InvalidOperationException("The imported designer belongs to another native process.");
            return target;
        }

        /// <summary>Native screen rectangle used solely to bound an owned rendering surface.</summary>
        [StructLayout(LayoutKind.Sequential)]
        private struct Rectangle { internal int Left, Top, Right, Bottom; }
        /// <summary>Reads bounds of the already verified owned window.</summary>
        [DllImport("user32.dll")] private static extern bool GetWindowRect(IntPtr window, out Rectangle bounds);
        /// <summary>Synchronously renders the owned window into an in-memory surface without storing an image.</summary>
        [DllImport("user32.dll")] private static extern bool PrintWindow(IntPtr window, IntPtr dc, uint flags);
        /// <summary>Reads native process ownership without activating another window.</summary>
        [DllImport("user32.dll")] private static extern uint GetWindowThreadProcessId(IntPtr window, out uint pid);
    }
}

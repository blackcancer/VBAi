using System;
using System.Diagnostics;
using System.Collections.Generic;
using System.Globalization;
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
        /// <summary>Records how initialization settled, independently of the final import acceptance.</summary>
        internal enum MaterializationOutcome { ResourcesExactBeforeFocus = 1, Rendered = 2 }

        /// <summary>Preserves exact resources and returns only font owners still different after one initialization.</summary>
        internal static FormStreamPadding.FormFontBinding[] Prepare(VbaGitSnapshot target, VbaGitComponent component,
            FormStreamPadding.FormFontBinding[] bindings, Func<VbaGitSnapshot> capture, Func<Func<bool>, MaterializationOutcome> materialize, Action revalidate)
        {
            if (bindings == null || bindings.Length == 0) return bindings;
            revalidate();
            VbaGitSnapshot actual = capture();
            string resource = component.Name + ".frx";
            if (target.SameFile(actual, resource)) return new FormStreamPadding.FormFontBinding[0];
            revalidate();
            MaterializationOutcome outcome = materialize(() => {
                revalidate();
                VbaGitSnapshot probe = capture();
                revalidate();
                return target.SameFile(probe, resource);
            });
            if (outcome != MaterializationOutcome.ResourcesExactBeforeFocus && outcome != MaterializationOutcome.Rendered)
                throw new InvalidOperationException("The imported designer initialization outcome is unrecognized.");
            revalidate();
            actual = capture();
            if (target.SameFile(actual, resource)) return new FormStreamPadding.FormFontBinding[0];
            if (outcome == MaterializationOutcome.ResourcesExactBeforeFocus)
                throw new InvalidOperationException("The imported form resources changed after restoring its prior view; font delivery is refused.");
            VbaGitComponent observedComponent = actual.Manifest.Components.Single(item => item.Name == component.Name && item.Type == 3);
            var observed = actual.FormFonts(observedComponent);
            if (observed == null)
                throw new InvalidOperationException("The imported form resource graph cannot be verified after designer initialization.");
            return bindings.Where(expected => !observed.Any(value => value.Type == expected.Type &&
                value.OwnerPath == expected.OwnerPath && value.Descriptor.SequenceEqual(expected.Descriptor))).ToArray();
        }

        /// <summary>Synchronously initializes only the verified designer on its actual owning VBE STA.</summary>
        internal static MaterializationOutcome Materialize(object project, object component, Func<bool> resourcesExactBeforeFocus, Action revalidate)
        {
            FormFontRestoration.RequireOwner(project);
            object window = null, editor = null, main = null, previous = null;
            bool originalVisible = false, originalViewKnown = false;
            Exception primary = null;
            MaterializationOutcome? outcome = null;
            try
            {
                revalidate();
                editor = ((dynamic)project).VBE;
                main = ((dynamic)editor).MainWindow;
                previous = ((dynamic)editor).ActiveWindow;
                window = ((dynamic)component).DesignerWindow();
                originalVisible = Convert.ToBoolean(((dynamic)window).Visible);
                originalViewKnown = true;
                outcome = ShowAndInitialize(() => { ((dynamic)window).Visible = true; }, resourcesExactBeforeFocus,
                    () => ((dynamic)window).SetFocus(), () => Render(project, window, editor, main, revalidate), revalidate);
                return outcome.Value;
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
                            () => ((dynamic)previous).SetFocus(), () => { ((dynamic)window).Visible = false; }, revalidate,
                            outcome == MaterializationOutcome.ResourcesExactBeforeFocus);
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

        /// <summary>Opens the owned designer once and avoids focus and rendering when an exact resource probe suffices.</summary>
        internal static MaterializationOutcome ShowAndInitialize(Action show, Func<bool> resourcesExactBeforeFocus,
            Action focus, Action render, Action revalidate)
        {
            revalidate();
            show();
            revalidate();
            bool exact = resourcesExactBeforeFocus();
            revalidate();
            if (exact) return MaterializationOutcome.ResourcesExactBeforeFocus;
            focus();
            // Native import may run under a modal Git dialog. Rendering keeps all
            // exact owning-STA, project, active-designer, visibility and PID guards.
            // Do not pump messages, force-enable an owner or retry a native mutation.
            render();
            return MaterializationOutcome.Rendered;
        }

        /// <summary>Restores the prior view only while the expected native navigation remains current.</summary>
        internal static void RestoreView(bool originallyVisible, bool previousDistinct, Func<bool> designerActive,
            Func<bool> previousActive, Action focusPrevious, Action hideDesigner, Action revalidate, bool exactBeforeFocus = false)
        {
            revalidate();
            if (!designerActive())
            {
                // Only a settled exact pre-focus probe permits restoring visibility
                // without stealing focus from the still-current previous window.
                if (!exactBeforeFocus || originallyVisible || !previousDistinct || !previousActive()) return;
                revalidate();
                if (designerActive() || !previousActive()) return;
                hideDesigner();
                return;
            }
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
            IntPtr target = ReadTarget(project, window, editor, main, "before-first-render");
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
                    if (ReadTarget(project, window, editor, main, "immediately-before-PrintWindow") != target)
                        throw new InvalidOperationException("The imported designer render target changed.");
                    if (!PrintWindow(target, dc, 2))
                        throw new InvalidOperationException("The imported designer could not be initialized by native rendering.");
                }
                finally { graphics.ReleaseHdc(dc); }
            }
            revalidate();
            if (ReadTarget(project, window, editor, main, "after-PrintWindow") != target)
                throw new InvalidOperationException("The imported designer changed during native initialization.");
        }

        /// <summary>Refuses a foreign, hidden or inactive designer before acquiring its drawing surface.</summary>
        private static IntPtr ReadTarget(object project, object window, object editor, object main, string stage)
        {
            object activeProject = null, activeWindow = null;
            try
            {
                activeProject = ((dynamic)editor).ActiveVBProject;
                activeWindow = ((dynamic)editor).ActiveWindow;
                IntPtr root = new IntPtr(Convert.ToInt64(((dynamic)main).HWnd));
                IntPtr designer = new IntPtr(Convert.ToInt64(((dynamic)window).HWnd));
                bool projectMatches = VbeProjectHostPath.SameProject(project, activeProject);
                bool designerMatches = VbeProjectHostPath.SameProject(window, activeWindow);
                bool mainVisible = Convert.ToBoolean(((dynamic)main).Visible);
                bool? designerVisible = mainVisible ? (bool?)Convert.ToBoolean(((dynamic)window).Visible) : null;
                int type = Convert.ToInt32(((dynamic)window).Type);
                // Read native ownership only if the unchanged selector evaluates
                // it. Reuse that observation in the refusal, never query again to
                // manufacture evidence after the failed guard.
                var owners = new Dictionary<IntPtr, uint>();
                var observations = new Dictionary<IntPtr, string>();
                Func<IntPtr, uint> owner = handle => {
                    uint pid;
                    if (owners.TryGetValue(handle, out pid)) return pid;
                    uint thread = GetWindowThreadProcessId(handle, out pid);
                    int error = thread == 0 ? Marshal.GetLastWin32Error() : 0;
                    owners.Add(handle, pid);
                    observations.Add(handle, string.Format(CultureInfo.InvariantCulture,
                        "HWND={0},Pid={1},Thread={2},Error={3}", handle.ToInt64(), pid, thread, error));
                    return pid;
                };
                return SelectObservedTarget(stage, root, designer, projectMatches, designerMatches,
                    mainVisible, designerVisible, type, (uint)Process.GetCurrentProcess().Id, owner,
                    () => observations.Count == 0 ? "not-evaluated" : string.Join("|",
                        observations.OrderBy(pair => pair.Key.ToInt64()).Select(pair => pair.Value)));
            }
            finally
            {
                if (activeWindow != null && Marshal.IsComObject(activeWindow)) Marshal.ReleaseComObject(activeWindow);
                if (activeProject != null && Marshal.IsComObject(activeProject)) Marshal.ReleaseComObject(activeProject);
            }
        }

        /// <summary>Preserves the native refusal and reports only the operands observed at that validation stage.</summary>
        internal static IntPtr SelectObservedTarget(string stage, IntPtr root, IntPtr designer, bool projectMatches,
            bool designerMatches, bool mainVisible, bool? designerVisible, int designerType, uint expectedPid,
            Func<IntPtr, uint> owner, Func<string> ownershipObservation)
        {
            try
            {
                return SelectTarget(root, designer, projectMatches, designerMatches, mainVisible && designerVisible == true,
                    designerType, expectedPid, owner);
            }
            catch (InvalidOperationException failure)
            {
                string context = string.Format(CultureInfo.InvariantCulture,
                    " Stage={0};Apartment={1};ExpectedPid={2};Root={3};Designer={4};ProjectMatches={5};" +
                    "DesignerMatches={6};MainVisible={7};DesignerVisible={8};DesignerType={9};NativeOwners={10}.",
                    stage, Thread.CurrentThread.GetApartmentState(), expectedPid, root.ToInt64(), designer.ToInt64(),
                    projectMatches, designerMatches, mainVisible, designerVisible.HasValue ? designerVisible.Value.ToString() : "not-evaluated", designerType, ownershipObservation());
                throw new InvalidOperationException(failure.Message + context, failure);
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
        [DllImport("user32.dll", SetLastError = true)] private static extern uint GetWindowThreadProcessId(IntPtr window, out uint pid);
    }
}

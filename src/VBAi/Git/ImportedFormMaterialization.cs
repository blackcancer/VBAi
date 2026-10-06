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
        internal enum MaterializationOutcome {

/// <summary>Identifies the resources exact before focus case of materialization outcome.</summary>
ResourcesExactBeforeFocus = 1,

/// <summary>Identifies the rendered case of materialization outcome.</summary>
Rendered = 2 }

        /// <summary>Preserves exact resources and returns only font owners still different after one initialization.</summary>
        /// <param name="target">vba git snapshot that supplies the target for this operation.</param>
        /// <param name="component">vba git component that supplies the component for this operation.</param>
        /// <param name="bindings">form font binding[] that supplies the bindings for this operation.</param>
        /// <param name="capture">func&lt;vba git snapshot&gt; that supplies the capture for this operation.</param>
        /// <param name="materialize">func&lt;func&lt;bool&gt;, materialization outcome&gt; that supplies the materialize for this operation.</param>
        /// <param name="revalidate">action that supplies the revalidate for this operation.</param>
        /// <returns>form font binding[] produced by the operation for prepare on imported form materialization.</returns>
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
        /// <param name="project">object that supplies the project for this operation.</param>
        /// <param name="component">object that supplies the component for this operation.</param>
        /// <param name="resourcesExactBeforeFocus">func&lt;bool&gt; that supplies the resources exact before focus for this operation.</param>
        /// <param name="revalidate">action that supplies the revalidate for this operation.</param>
        /// <returns>materialization outcome produced by the operation for materialize on imported form materialization.</returns>
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
                    () => ((dynamic)window).SetFocus(), () => Render(project, window, editor, main, previous, revalidate), revalidate);
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
        /// <param name="show">action that supplies the show for this operation.</param>
        /// <param name="resourcesExactBeforeFocus">func&lt;bool&gt; that supplies the resources exact before focus for this operation.</param>
        /// <param name="focus">action that supplies the focus for this operation.</param>
        /// <param name="render">action that supplies the render for this operation.</param>
        /// <param name="revalidate">action that supplies the revalidate for this operation.</param>
        /// <returns>materialization outcome produced by the operation for show and initialize on imported form materialization.</returns>
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
        /// <param name="originallyVisible">Indicates whether originally visible is enabled.</param>
        /// <param name="previousDistinct">Indicates whether previous distinct is enabled.</param>
        /// <param name="designerActive">func&lt;bool&gt; that supplies the designer active for this operation.</param>
        /// <param name="previousActive">func&lt;bool&gt; that supplies the previous active for this operation.</param>
        /// <param name="focusPrevious">action that supplies the focus previous for this operation.</param>
        /// <param name="hideDesigner">action that supplies the hide designer for this operation.</param>
        /// <param name="revalidate">action that supplies the revalidate for this operation.</param>
        /// <param name="exactBeforeFocus">Indicates whether exact before focus is enabled.</param>
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
        /// <param name="editor">object that supplies the editor for this operation.</param>
        /// <param name="expected">object that supplies the expected for this operation.</param>
        /// <returns>Boolean indicating the result of the check for is active on imported form materialization.</returns>
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
        /// <param name="project">object that supplies the project for this operation.</param>
        /// <param name="window">object that supplies the window for this operation.</param>
        /// <param name="editor">object that supplies the editor for this operation.</param>
        /// <param name="main">object that supplies the main for this operation.</param>
        /// <param name="previous">object that supplies the previous for this operation.</param>
        /// <param name="revalidate">action that supplies the revalidate for this operation.</param>
        private static void Render(object project, object window, object editor, object main, object previous, Action revalidate)
        {
            revalidate();
            IntPtr target = ReadTarget(project, window, editor, main, previous, "before-first-render");
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
                    if (ReadTarget(project, window, editor, main, previous, "immediately-before-PrintWindow") != target)
                        throw new InvalidOperationException("The imported designer render target changed.");
                    if (!PrintWindow(target, dc, 2))
                        throw new InvalidOperationException("The imported designer could not be initialized by native rendering.");
                }
                finally { graphics.ReleaseHdc(dc); }
            }
            revalidate();
            if (ReadTarget(project, window, editor, main, previous, "after-PrintWindow") != target)
                throw new InvalidOperationException("The imported designer changed during native initialization.");
        }

        /// <summary>Refuses a foreign, hidden or inactive designer before acquiring its drawing surface.</summary>
        /// <param name="project">object that supplies the project for this operation.</param>
        /// <param name="window">object that supplies the window for this operation.</param>
        /// <param name="editor">object that supplies the editor for this operation.</param>
        /// <param name="main">object that supplies the main for this operation.</param>
        /// <param name="previous">object that supplies the previous for this operation.</param>
        /// <param name="stage">Text that supplies the stage value. Use the format required by the calling operation.</param>
        /// <returns>int ptr produced by the operation for read target on imported form materialization.</returns>
        private static IntPtr ReadTarget(object project, object window, object editor, object main, object previous, string stage)
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
                        observations.OrderBy(pair => pair.Key.ToInt64()).Select(pair => pair.Value)),
                    () => ImportedFormFocusDiagnostic.Format(ImportedFormFocusDiagnostic.Observe(
                        activeWindow != null, designerMatches, previous != null, root != IntPtr.Zero, type, designer.ToInt64(),
                        () => Convert.ToInt32(((dynamic)activeWindow).Type),
                        () => Convert.ToInt64(((dynamic)activeWindow).HWnd),
                        () => Convert.ToString(((dynamic)activeWindow).Caption),
                        () => Convert.ToString(((dynamic)window).Caption),
                        () => IsWindowEnabled(root), () => VbeProjectHostPath.SameProject(activeWindow, previous))));
            }
            finally
            {
                if (activeWindow != null && Marshal.IsComObject(activeWindow)) Marshal.ReleaseComObject(activeWindow);
                if (activeProject != null && Marshal.IsComObject(activeProject)) Marshal.ReleaseComObject(activeProject);
            }
        }

        /// <summary>Preserves the native refusal and reports only the operands observed at that validation stage.</summary>
        /// <param name="stage">Text that supplies the stage value. Use the format required by the calling operation.</param>
        /// <param name="root">Native handle that supplies the root for this operation.</param>
        /// <param name="designer">Native handle that supplies the designer for this operation.</param>
        /// <param name="projectMatches">Indicates whether project matches is enabled.</param>
        /// <param name="designerMatches">Indicates whether designer matches is enabled.</param>
        /// <param name="mainVisible">Indicates whether main visible is enabled.</param>
        /// <param name="designerVisible">bool that supplies the designer visible for this operation.</param>
        /// <param name="designerType">int that supplies the designer type for this operation.</param>
        /// <param name="expectedPid">uint that supplies the expected pid for this operation.</param>
        /// <param name="owner">func&lt;int ptr, uint&gt; that supplies the owner for this operation.</param>
        /// <param name="ownershipObservation">func&lt;string&gt; that supplies the ownership observation for this operation.</param>
        /// <param name="failureObservation">func&lt;string&gt; that supplies the failure observation for this operation.</param>
        /// <returns>int ptr produced by the operation for select observed target on imported form materialization.</returns>
        internal static IntPtr SelectObservedTarget(string stage, IntPtr root, IntPtr designer, bool projectMatches,
            bool designerMatches, bool mainVisible, bool? designerVisible, int designerType, uint expectedPid,
            Func<IntPtr, uint> owner, Func<string> ownershipObservation, Func<string> failureObservation = null)
        {
            try
            {
                return SelectTarget(root, designer, projectMatches, designerMatches, mainVisible && designerVisible == true,
                    designerType, expectedPid, owner);
            }
            catch (InvalidOperationException failure)
            {
                string context;
                try
                {
                    context = string.Format(CultureInfo.InvariantCulture,
                        " Stage={0};Apartment={1};ExpectedPid={2};Root={3};Designer={4};ProjectMatches={5};" +
                        "DesignerMatches={6};MainVisible={7};DesignerVisible={8};DesignerType={9};NativeOwners={10}.",
                        stage, Thread.CurrentThread.GetApartmentState(), expectedPid, root.ToInt64(), designer.ToInt64(),
                        projectMatches, designerMatches, mainVisible, designerVisible.HasValue ? designerVisible.Value.ToString() : "not-evaluated", designerType, ownershipObservation());
                    if (failureObservation != null) context += failureObservation();
                }
                catch (Exception)
                {
                    // A diagnostic failure must not replace the exact original selection refusal.
                    ExceptionDispatchInfo.Capture(failure).Throw();
                    throw;
                }
                throw new InvalidOperationException(failure.Message + context, failure);
            }
        }

        /// <summary>Allows the owning VBE root only when the exact active designer has no HWND.</summary>
        /// <param name="root">Native handle that supplies the root for this operation.</param>
        /// <param name="designer">Native handle that supplies the designer for this operation.</param>
        /// <param name="projectMatches">Indicates whether project matches is enabled.</param>
        /// <param name="designerMatches">Indicates whether designer matches is enabled.</param>
        /// <param name="visible">Indicates whether visible is enabled.</param>
        /// <param name="designerType">int that supplies the designer type for this operation.</param>
        /// <param name="expectedPid">uint that supplies the expected pid for this operation.</param>
        /// <param name="owner">func&lt;int ptr, uint&gt; that supplies the owner for this operation.</param>
        /// <returns>int ptr produced by the operation for select target on imported form materialization.</returns>
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
        private struct Rectangle {

/// <summary>Maintains the left and top and right and bottom state for rectangle.</summary>
internal int Left, Top, Right, Bottom; }

        /// <summary>Reads bounds of the already verified owned window.</summary>
        /// <param name="window">Native handle that supplies the window for this operation.</param>
        /// <param name="bounds">rectangle that supplies the bounds for this operation.</param>
        /// <returns>Boolean indicating the result of the check for get window rect on imported form materialization.</returns>
        [DllImport("user32.dll")] private static extern bool GetWindowRect(IntPtr window, out Rectangle bounds);

        /// <summary>Synchronously renders the owned window into an in-memory surface without storing an image.</summary>
        /// <param name="window">Native handle that supplies the window for this operation.</param>
        /// <param name="dc">Native handle that supplies the dc for this operation.</param>
        /// <param name="flags">uint that supplies the flags for this operation.</param>
        /// <returns>Boolean indicating the result of the check for print window on imported form materialization.</returns>
        [DllImport("user32.dll")] private static extern bool PrintWindow(IntPtr window, IntPtr dc, uint flags);

        /// <summary>Reads native process ownership without activating another window.</summary>
        /// <param name="window">Native handle that supplies the window for this operation.</param>
        /// <param name="pid">uint that supplies the pid for this operation.</param>
        /// <returns>uint produced by the operation for get window thread process id on imported form materialization.</returns>
        [DllImport("user32.dll", SetLastError = true)] private static extern uint GetWindowThreadProcessId(IntPtr window, out uint pid);

        /// <summary>Observes the already acquired owner root only after an identity refusal; never enables it.</summary>
        /// <param name="window">Native handle that supplies the window for this operation.</param>
        /// <returns>Boolean indicating the result of the check for is window enabled on imported form materialization.</returns>
        [DllImport("user32.dll")] private static extern bool IsWindowEnabled(IntPtr window);
    }
}

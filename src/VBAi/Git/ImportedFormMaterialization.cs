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
        /// <param name="target">Repository snapshot whose serialized form resources must remain byte-exact.</param>
        /// <param name="component">Imported UserForm component whose designer may need one initialization pass.</param>
        /// <param name="bindings">Font descriptors captured before import, used only if native initialization rewrites persisted resources.</param>
        /// <param name="capture">Reads the current project snapshot at each guarded observation point.</param>
        /// <param name="materialize">Performs one guarded designer initialization and reports whether bytes were exact before focus or rendered.</param>
        /// <param name="revalidate">Rechecks project/revision/owner admission around every snapshot and native boundary.</param>
        /// <returns>Only font owners still missing or changed after a rendered initialization; empty when exact bytes remain, or the input when no plan exists.</returns>
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
        /// <param name="project">Imported project's VBE project object, which must remain the active project.</param>
        /// <param name="component">Imported form component whose designer window is opened and rendered once if required.</param>
        /// <param name="resourcesExactBeforeFocus">Checks whether target resource bytes already match before the designer receives focus.</param>
        /// <param name="revalidate">Rechecks the native import's owner, project, revision, and approval conditions.</param>
        /// <returns>Whether resources were exact before focus or the verified designer required one native render.</returns>
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
        /// <param name="show">Makes the imported designer visible without pumping additional messages.</param>
        /// <param name="resourcesExactBeforeFocus">Probes persisted bytes before any focus or rendering side effect.</param>
        /// <param name="focus">Focuses the exact verified designer only when its resource bytes are not already exact.</param>
        /// <param name="render">Performs the bounded native render used to initialize the owned designer.</param>
        /// <param name="revalidate">Rechecks import admission around each native boundary.</param>
        /// <returns><see cref="MaterializationOutcome.ResourcesExactBeforeFocus"/> or <see cref="MaterializationOutcome.Rendered"/>.</returns>
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
        /// <param name="originallyVisible">Whether the designer was visible before materialization.</param>
        /// <param name="previousDistinct">Whether a different prior window was active before this operation.</param>
        /// <param name="designerActive">Rechecks whether the imported designer currently owns focus.</param>
        /// <param name="previousActive">Rechecks whether the prior window is still active.</param>
        /// <param name="focusPrevious">Restores focus to the prior window after a rendered initialization.</param>
        /// <param name="hideDesigner">Hides the designer only after guarded navigation checks pass.</param>
        /// <param name="revalidate">Revalidates the native import before changing visibility or focus.</param>
        /// <param name="exactBeforeFocus"><see langword="true"/> only when a settled pre-focus probe proved resources exact.</param>
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
        /// <param name="editor">VBE automation object whose ActiveWindow is read.</param>
        /// <param name="expected">Window identity to compare with the currently active window.</param>
        /// <returns><see langword="true"/> when the active window represents the expected project/window identity.</returns>
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
        /// <param name="project">Expected imported project used for active-project identity checks.</param>
        /// <param name="window">Imported form designer whose HWND may be rendered.</param>
        /// <param name="editor">Owning VBE automation object.</param>
        /// <param name="main">VBE main window used as the allowed fallback render target.</param>
        /// <param name="previous">Previously active VBE window used for diagnostics and restoration.</param>
        /// <param name="revalidate">Import admission check called before and after native rendering.</param>
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
        /// <param name="project">Project that must still be the active VBE project.</param>
        /// <param name="window">Exact imported designer whose visibility and active identity are checked.</param>
        /// <param name="editor">VBE automation object providing current project/window identities.</param>
        /// <param name="main">Root HWND owner allowed only when the exact active form has no HWND.</param>
        /// <param name="previous">Previously active window included in the refusal diagnostic.</param>
        /// <param name="stage">Named validation stage included in any refusal detail.</param>
        /// <returns>Verified target HWND after project, designer, visibility, type, and process checks pass.</returns>
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
        /// <param name="stage">Validation stage used to annotate a native refusal.</param>
        /// <param name="root">VBE main-window HWND.</param><param name="designer">Imported designer HWND.</param>
        /// <param name="projectMatches">Whether the active project is the expected imported project.</param>
        /// <param name="designerMatches">Whether the active VBE window is the exact imported designer.</param>
        /// <param name="mainVisible">Whether the owning VBE main window is visible.</param>
        /// <param name="designerVisible">Designer visibility, or null when the root is hidden and this observation was not made.</param>
        /// <param name="designerType">VBE window type reported for the designer.</param><param name="expectedPid">Current process ID captured as the expected native owner.</param>
        /// <param name="owner">Reads the owning process ID for an HWND; each queried handle is observed once.</param>
        /// <param name="ownershipObservation">Formats only HWND/PID/thread/error values observed by the owner reader.</param>
        /// <param name="failureObservation">Optional bounded active-window snapshot appended only after validation refuses the target.</param>
        /// <returns>The selected HWND when all invariants hold; otherwise throws an annotated refusal with observed operands.</returns>
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
        /// <param name="root">Owning VBE main-window HWND.</param><param name="designer">Exact active designer HWND, or zero when the host exposes none.</param>
        /// <param name="projectMatches">Whether active-project identity matches the imported project.</param>
        /// <param name="designerMatches">Whether active-window identity matches the imported designer.</param>
        /// <param name="visible">Whether both the owner and designer are visible.</param>
        /// <param name="designerType">VBE window type; only UserForm designer type 1 is accepted.</param>
        /// <param name="expectedPid">Current process ID that must own both root and selected target.</param>
        /// <param name="owner">Native HWND-to-process lookup used for ownership validation.</param>
        /// <returns>The designer HWND, or the VBE root only when the verified active designer HWND is zero.</returns>
        /// <exception cref="InvalidOperationException">The call is off-STA or any exact project, designer, visibility, type, or process guard fails.</exception>
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

/// <summary>Screen-coordinate bounds used to cap the in-memory render surface at 4096 pixels per side.</summary>
internal int Left, Top, Right, Bottom; }

        /// <summary>Reads bounds of the already verified owned window.</summary>
        /// <param name="window">Already verified native render target.</param><param name="bounds">Receives its screen-coordinate bounds.</param>
        /// <returns><see langword="true"/> when Windows returns a rectangle.</returns>
        [DllImport("user32.dll")] private static extern bool GetWindowRect(IntPtr window, out Rectangle bounds);

        /// <summary>Synchronously renders the owned window into an in-memory surface without storing an image.</summary>
        /// <param name="window">Verified designer HWND rendered synchronously.</param><param name="dc">In-memory bitmap device context receiving the image.</param>
        /// <param name="flags">PrintWindow mode flags; value 2 requests rendering even when not foreground.</param>
        /// <returns><see langword="true"/> when the native window was rendered.</returns>
        [DllImport("user32.dll")] private static extern bool PrintWindow(IntPtr window, IntPtr dc, uint flags);

        /// <summary>Reads native process ownership without activating another window.</summary>
        /// <param name="window">Native window whose process and thread ownership are checked.</param><param name="pid">Receives the owning process ID.</param>
        /// <returns>Owning native thread ID, or zero when ownership cannot be read.</returns>
        [DllImport("user32.dll", SetLastError = true)] private static extern uint GetWindowThreadProcessId(IntPtr window, out uint pid);

        /// <summary>Observes the already acquired owner root only after an identity refusal; never enables it.</summary>
        /// <param name="window">VBE owner HWND observed only for refusal diagnostics.</param>
        /// <returns><see langword="true"/> when Windows reports that the owner accepts input.</returns>
        [DllImport("user32.dll")] private static extern bool IsWindowEnabled(IntPtr window);
    }
}

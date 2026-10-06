using System;
using System.Diagnostics;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;

namespace VBAi
{

    /// <summary>Owns the vbe debug windows state and operations.</summary>
    internal static partial class VbeDebugWindows
    {

        /// <summary>Handles native context window style for vbe debug windows.</summary>
        /// <param name="handle">Native handle that supplies the handle for this operation.</param>
        /// <param name="index">int that supplies the index for this operation.</param>
        /// <returns>int produced by the operation for native context window style on vbe debug windows.</returns>
        [DllImport("user32.dll", EntryPoint = "GetWindowLongW")]
        private static extern int NativeContextWindowStyle(IntPtr handle, int index);

        /// <summary>Maintains the context window style state for vbe debug windows.</summary>
        internal static Func<IntPtr, int> ContextWindowStyle = handle => NativeContextWindowStyle(handle, -16);

        /// <summary>Reads the native read-only Locals context, not the active code cursor.</summary>
        /// <returns>Text produced by the operation for read locals context on vbe debug windows.</returns>
        internal static string ReadLocalsContext()
        {
            IntPtr root = FindVbeRoot();
            if (root == IntPtr.Zero) throw new InvalidOperationException("The visible VBE root is unavailable.");
            var panes = ChildWindows(root).Where(h => WindowText(h) == "Variables locales" || WindowText(h) == "Locals").ToArray();
            if (panes.Length != 1) throw new InvalidOperationException("Exactly one visible Locals pane is required; open_debug_pane can open it.");
            IntPtr edit = GetDlgItem(panes[0], 4604), button = GetDlgItem(panes[0], 4601);
            uint owner;
            GetWindowThreadProcessId(edit, out owner);
            if (edit == IntPtr.Zero || owner != (uint)Process.GetCurrentProcess().Id || ClassName(edit) != "Edit" ||
                !IsWindowVisible(edit) || (ContextWindowStyle(edit) & 0x800) == 0 ||
                button == IntPtr.Zero || ClassName(button) != "Button" || !IsWindowVisible(button))
                throw new InvalidOperationException("The read-only native Locals context control is unavailable.");
            var text = new StringBuilder(1024);
            int length = GetWindowText(edit, text, text.Capacity);
            if (length <= 0 || length >= text.Capacity - 1)
                throw new InvalidOperationException("The native Locals context is empty or truncated.");
            return text.ToString();
        }

        /// <summary>Ensures no quick watch dialog for vbe debug windows.</summary>
        internal static void EnsureNoQuickWatchDialog()
        {
            if (FindDialog("Espion express", "Quick Watch") != IntPtr.Zero)
                throw new InvalidOperationException("Close the existing Quick Watch dialog before starting an inspection.");
            if (FindDialog("Microsoft Visual Basic pour Applications", "Microsoft Visual Basic for Applications", "Microsoft Visual Basic") != IntPtr.Zero)
                throw new InvalidOperationException("Close the existing VBA diagnostic before starting an inspection.");
        }

        /// <summary>Waits for the owned dialog to close before another selected expression can be inspected.</summary>
        /// <param name="request">request that supplies the request for this operation.</param>
        /// <returns>object produced by the operation for read scalar quick watch on vbe debug windows.</returns>
        internal static object ReadScalarQuickWatch(Request request)
        {
            var trace = VbeInspectionTrace.Current;
            trace?.Record(VbeInspectionTrace.Phase.ObserverEntered);
            observerTextDepth.Value++;
            Exception failure = null;
            try { return ReadScalarQuickWatchCore(request); }
            catch (Exception error) { failure = error; throw; }
            finally
            {
                observerTextDepth.Value--;
                trace?.Record(VbeInspectionTrace.Phase.ObserverTerminal, failure);
            }
        }

        /// <summary>Reads scalar quick watch core for vbe debug windows.</summary>
        /// <param name="request">request that supplies the request for this operation.</param>
        /// <returns>object produced by the operation for read scalar quick watch core on vbe debug windows.</returns>
        private static object ReadScalarQuickWatchCore(Request request)
        {
            // Watch the diagnostic too: a lost selection opens an error instead of Quick Watch.
            // Only acknowledge the exact diagnostic reproduced by this operation, after the absence preflight.
            bool opened = false;
            for (int attempt = 0; attempt < 60; attempt++)
            {
                if (FindDialog("Espion express", "Quick Watch") != IntPtr.Zero) { opened = true; VbeInspectionTrace.Current?.Record(VbeInspectionTrace.Phase.ObserverDialogFound); break; }
                IntPtr diagnostic = FindDialog("Microsoft Visual Basic pour Applications", "Microsoft Visual Basic for Applications", "Microsoft Visual Basic");
                if (diagnostic != IntPtr.Zero)
                {
                    var controls = new NativeProbe().DialogControls(diagnostic);
                    string message = controls.Where(c => c.Kind == "Static").Select(c => c.Text)
                        .FirstOrDefault(t => !string.IsNullOrWhiteSpace(t));
                    IntPtr ok = GetDlgItem(diagnostic, 2);
                    if (message == "Pas d'expression espionne sélectionnée" && ok != IntPtr.Zero && ClassName(ok) == "Button" && WindowText(ok) == "OK")
                    {
                        PostMessage(ok, BmClick, IntPtr.Zero, IntPtr.Zero);
                        for (int closeAttempt = 0; closeAttempt < 40 && IsWindowVisible(diagnostic); closeAttempt++) PauseNative(25);
                    }
                    throw new InvalidOperationException("Native Quick Watch diagnostic: " + (message ?? "unavailable"));
                }
                PauseNative(50);
            }
            if (!opened) throw new InvalidOperationException("The Quick Watch dialog did not open.");
            object result = null;
            Exception readError = null;
            try { result = CompleteQuickWatch(request); VbeInspectionTrace.Current?.Record(VbeInspectionTrace.Phase.ObserverReadComplete); }
            catch (Exception ex) { readError = ex; }
            for (int attempt = 0; attempt < 40; attempt++)
            {
                if (FindDialog("Espion express", "Quick Watch") == IntPtr.Zero)
                {
                    if (readError != null) throw new InvalidOperationException(readError.Message, readError);
                    return result;
                }
                PauseNative(25);
            }
            throw new InvalidOperationException("Quick Watch did not close; the remaining inspection was stopped.");
        }
    }
}

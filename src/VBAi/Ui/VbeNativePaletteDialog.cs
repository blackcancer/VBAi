using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading;
using System.Windows.Automation;
using Row = VBAi.VbeNativePaletteState.ColorRow;

namespace VBAi
{

    /// <summary>Reads and writes native editor colors through an owned Options dialog.</summary>
    internal static class VbeNativePaletteDialog
    {

        /// <summary>Callback signature used by the native window-enumeration functions.</summary>
        /// <param name="window">Window handle reported by Windows.</param><param name="parameter">Caller-supplied enumeration context.</param>
        /// <returns><see langword="true"/> to continue enumeration.</returns>
        private delegate bool EnumProc(IntPtr window, IntPtr parameter);

        /// <summary>Enumerates top-level desktop windows.</summary>
        /// <param name="callback">Callback invoked for each top-level window.</param><param name="parameter">Context passed to each callback.</param>
        /// <returns><see langword="true"/> when enumeration succeeds.</returns>
        [DllImport("user32.dll")] private static extern bool EnumWindows(EnumProc callback, IntPtr parameter);

        /// <summary>Enumerates child windows of a native parent.</summary>
        /// <param name="parent">Parent window handle.</param><param name="callback">Callback invoked for each child.</param>
        /// <param name="parameter">Context passed to each callback.</param><returns><see langword="true"/> when enumeration succeeds.</returns>
        [DllImport("user32.dll")] private static extern bool EnumChildWindows(IntPtr parent, EnumProc callback, IntPtr parameter);

        /// <summary>Gets the class name associated with a native window handle.</summary>
        /// <param name="window">Window handle to query.</param><param name="name">Buffer that receives the class name.</param>
        /// <param name="capacity">Character capacity of the buffer.</param><returns>Number of characters copied.</returns>
        [DllImport("user32.dll", CharSet = CharSet.Unicode)] private static extern int GetClassName(IntPtr window, StringBuilder name, int capacity);

        /// <summary>Gets a dialog control's identifier.</summary><param name="window">Control handle.</param><returns>Control identifier.</returns>
        [DllImport("user32.dll")] private static extern int GetDlgCtrlID(IntPtr window);

        /// <summary>Reads a native window style value.</summary><param name="window">Window handle.</param><param name="index">Style index.</param><returns>Style bits.</returns>
        [DllImport("user32.dll", EntryPoint = "GetWindowLongW")] private static extern int GetWindowStyle(IntPtr window, int index);

        /// <summary>Gets the process and thread that created a native window.</summary><param name="window">Window handle.</param>
        /// <param name="process">Receives the owning process identifier.</param><returns>Owning thread identifier.</returns>
        [DllImport("user32.dll")] private static extern uint GetWindowThreadProcessId(IntPtr window, out uint process);

        /// <summary>Gets a related native window according to a Windows relationship command.</summary><param name="window">Starting window.</param>
        /// <param name="command">Relationship selector.</param><returns>Related window handle, or zero when absent.</returns>
        [DllImport("user32.dll")] private static extern IntPtr GetWindow(IntPtr window, uint command);

        /// <summary>Gets the immediate parent of a child window.</summary><param name="window">Child window.</param><returns>Parent window handle.</returns>
        [DllImport("user32.dll")] private static extern IntPtr GetParent(IntPtr window);

        /// <summary>Gets a dialog item by its control identifier.</summary><param name="window">Dialog window handle.</param>
        /// <param name="id">Control identifier.</param><returns>Control handle, or zero when absent.</returns>
        [DllImport("user32.dll")] private static extern IntPtr GetDlgItem(IntPtr window, int id);

        /// <summary>Checks whether a native window is visible.</summary><param name="window">Window handle.</param><returns>Visibility state.</returns>
        [DllImport("user32.dll")] private static extern bool IsWindowVisible(IntPtr window);

        /// <summary>Checks whether a native window accepts input.</summary><param name="window">Window handle.</param><returns>Enabled state.</returns>
        [DllImport("user32.dll")] private static extern bool IsWindowEnabled(IntPtr window);

        /// <summary>Posts a message to a native window's queue.</summary><param name="window">Target window.</param><param name="message">Message identifier.</param>
        /// <param name="wParam">Message-specific first value.</param><param name="lParam">Message-specific second value.</param>
        /// <returns><see langword="true"/> when the message was posted.</returns>
        [DllImport("user32.dll")] private static extern bool PostMessage(IntPtr window, uint message, IntPtr wParam, IntPtr lParam);

        /// <summary>Sends a message to a native window and waits for its result.</summary><param name="window">Target window.</param>
        /// <param name="message">Message identifier.</param><param name="wParam">Message-specific first value.</param>
        /// <param name="lParam">Message-specific second value.</param><returns>Window procedure result.</returns>
        [DllImport("user32.dll")] private static extern IntPtr SendMessage(IntPtr window, uint message, IntPtr wParam, IntPtr lParam);

        /// <summary>Bounded UIA/open, worker-join, native-close, and color-page wait limits, in milliseconds.</summary>
        internal static int OpenTimeoutMilliseconds = 10000, WorkerTimeoutMilliseconds = 30000,
            CloseTimeoutMilliseconds = 2000, PageTimeoutMilliseconds = 1000;

        /// <summary>Message-posting route used to close or control only the dialog owned by this transaction.</summary>
        internal static Func<IntPtr, uint, IntPtr, IntPtr, bool> PostDialogMessage = PostMessage;

        /// <summary>Worker-join route used with the explicit bounded shutdown timeout.</summary>
        internal static Func<Thread, int, bool> WaitWorker = (worker, timeout) => worker.Join(timeout);

        /// <summary>Snapshot of native descendants owned by the VBE before opening Options, used to distinguish the new dialog.</summary>
        internal static Func<IntPtr, HashSet<IntPtr>> OwnedWindows = Windows;

        /// <summary>Opens the VBE Options dialog, reads its ten color rows, optionally updates them, then closes it.</summary>
        /// <param name="vbe">VBE automation object used to invoke the Options command.</param>
        /// <param name="update">Callback receiving the observed rows and returning replacement rows, or <see langword="null"/> to leave them unchanged.</param>
        /// <returns>The rows observed before the optional update.</returns>
        /// <exception cref="InvalidOperationException">The dialog, color page, or requested update cannot be validated.</exception>
        internal static Row[] Visit(object vbe, Func<Row[], Row[]> update)
        {
            dynamic editor = vbe;
            IntPtr owner = new IntPtr(Convert.ToInt64(editor.MainWindow.HWnd));
            if (!IsWindowEnabled(owner) || !IsWindowVisible(owner))
                throw new InvalidOperationException("The VBE must be visible and have no modal dialog open.");
            var existing = OwnedWindows(owner);
            dynamic command = editor.CommandBars.FindControl(1, 522);
            if (command == null || !(bool)command.Enabled) throw new InvalidOperationException("The native Options command is unavailable.");
            Exception failure = null;
            Row[] observed = null;
            var cancellation = new CancellationTokenSource();
            IntPtr dialog = IntPtr.Zero;
            var worker = new Thread(() =>
            {
                bool accept = false;
                try
                {
                    var wait = Stopwatch.StartNew();
                    while (dialog == IntPtr.Zero && wait.ElapsedMilliseconds < OpenTimeoutMilliseconds)
                    {
                        cancellation.Token.ThrowIfCancellationRequested();
                        var candidates = OwnedWindows(owner).Where(window => !existing.Contains(window) &&
                            ClassName(window) == "#32770" && GetDlgItem(window, 1) != IntPtr.Zero &&
                            GetDlgItem(window, 2) != IntPtr.Zero && Descendants(window, "SysTabControl32").Count == 1).ToArray();
                        if (candidates.Length > 1) throw new InvalidOperationException("The owned Options dialog is ambiguous.");
                        dialog = candidates.SingleOrDefault();
                        if (dialog == IntPtr.Zero) Thread.Sleep(50);
                    }
                    if (dialog == IntPtr.Zero) throw new InvalidOperationException("The owned Options dialog did not open.");
                    IntPtr listHandle = FindColorPage(dialog, cancellation.Token);
                    var list = AutomationElement.FromHandle(listHandle);
                    var names = list.FindAll(TreeScope.Children, new PropertyCondition(AutomationElement.ControlTypeProperty, ControlType.ListItem))
                        .Cast<AutomationElement>().Select(item => item.Current.Name).ToArray();
                    if (names.Length != 10 || SendMessage(listHandle, 0x18b, IntPtr.Zero, IntPtr.Zero).ToInt32() != 10)
                        throw new InvalidOperationException("Unexpected native color category count.");
                    IntPtr foreground = Control(dialog, 4913), background = Control(dialog, 4914), indicator = Control(dialog, 4935);
                    observed = Read(listHandle, names, foreground, background, indicator);
                    Row[] desired = update(observed);
                    if (desired != null)
                    {
                        VbeNativePaletteState.ValidateRows(desired);
                        if (!names.SequenceEqual(desired.Select(row => row.Name))) throw new InvalidOperationException("The native color categories changed.");
                        for (int index = 0; index < desired.Length; index++)
                        {
                            cancellation.Token.ThrowIfCancellationRequested();
                            Select(listHandle, index);
                            SetColor(foreground, 4913, desired[index].Foreground);
                            SetColor(background, 4914, desired[index].Background);
                            SetColor(indicator, 4935, desired[index].Indicator);
                        }
                        if (!VbeNativePaletteState.Equal(desired, Read(listHandle, names, foreground, background, indicator)))
                            throw new InvalidOperationException("The native color controls did not retain the requested values.");
                        accept = true;
                    }
                }
                catch (Exception error) { failure = error; }
                finally
                {
                    if (dialog != IntPtr.Zero)
                    {
                        IntPtr button = GetDlgItem(dialog, accept ? 1 : 2);
                        if (button == IntPtr.Zero || !PostDialogMessage(button, 0xf5, IntPtr.Zero, IntPtr.Zero))
                            failure = failure ?? new InvalidOperationException("The native Options dialog could not be closed.");
                    }
                }
            }) { IsBackground = true, Name = "VBE native palette" };
            worker.SetApartmentState(ApartmentState.MTA);
            worker.Start();
            try
            {
                command.Execute();
                // Some VBE hosts return before the modal dialog is visible. Keep
                // pumping the owning STA until the worker closes that dialog.
                var completion = Stopwatch.StartNew();
                while (worker.IsAlive && completion.ElapsedMilliseconds < WorkerTimeoutMilliseconds)
                {
                    System.Windows.Forms.Application.DoEvents();
                    Thread.Sleep(10);
                }
            }
            finally { cancellation.Cancel(); }
            if (!WaitWorker(worker, 1500)) throw new InvalidOperationException("The native palette worker did not finish.");
            cancellation.Dispose();
            if (failure != null) throw new InvalidOperationException("Native editor colors could not be updated.", failure);
            var closing = Stopwatch.StartNew();
            while (dialog != IntPtr.Zero && IsWindowVisible(dialog) && closing.ElapsedMilliseconds < CloseTimeoutMilliseconds)
            {
                System.Windows.Forms.Application.DoEvents();
                Thread.Sleep(10);
            }
            if (IsWindowVisible(dialog))
                throw new InvalidOperationException("The native Options dialog did not finish closing.");
            return observed;
        }

        /// <summary>Lists visible top-level windows owned by the process and nested under the supplied VBE window.</summary>
        /// <param name="owner">VBE main-window handle used as the ownership boundary.</param>
        /// <returns>Handles of visible owned top-level windows.</returns>
        private static HashSet<IntPtr> Windows(IntPtr owner)
        {
            var result = new HashSet<IntPtr>();
            uint process;
            GetWindowThreadProcessId(owner, out process);
            EnumWindows((window, unused) =>
            {
                uint candidate;
                GetWindowThreadProcessId(window, out candidate);
                if (candidate == process && IsWindowVisible(window))
                {
                    IntPtr parent = GetWindow(window, 4);
                    for (int depth = 0; parent != IntPtr.Zero && depth < 16; depth++, parent = GetWindow(parent, 4))
                        if (parent == owner) { result.Add(window); break; }
                }
                return true;
            }, IntPtr.Zero);
            return result;
        }

        /// <summary>Reads a native window's class name.</summary>
        /// <param name="window">Window handle to query.</param>
        /// <returns>The class name returned by Windows, or an empty string.</returns>
        private static string ClassName(IntPtr window)
        {
            var name = new StringBuilder(256);
            GetClassName(window, name, name.Capacity);
            return name.ToString();
        }

        /// <summary>Finds visible descendants matching a native class and optionally a dialog control identifier.</summary>
        /// <param name="parent">Parent window whose descendants are enumerated.</param>
        /// <param name="className">Required native class name.</param>
        /// <param name="id">Required control identifier, or a negative value to ignore the identifier.</param>
        /// <returns>Matching descendant window handles in enumeration order.</returns>
        private static List<IntPtr> Descendants(IntPtr parent, string className, int id = -1)
        {
            var matches = new List<IntPtr>();
            EnumChildWindows(parent, (window, unused) =>
            {
                if (IsWindowVisible(window) && ClassName(window) == className &&
                    (id < 0 || GetDlgCtrlID(window) == id)) matches.Add(window);
                return true;
            }, IntPtr.Zero);
            return matches;
        }

        /// <summary>Selects the Options tab that contains the native syntax-color controls.</summary>
        /// <param name="dialog">Handle of the owned VBE Options dialog.</param>
        /// <param name="cancellation">Token checked while switching tabs and waiting for controls.</param>
        /// <returns>Handle of the syntax-category list on the color page.</returns>
        /// <exception cref="InvalidOperationException">The tab structure or native color page is unexpected.</exception>
        internal static IntPtr FindColorPage(IntPtr dialog, CancellationToken cancellation)
        {
            IntPtr tabs = Descendants(dialog, "SysTabControl32").Single();
            if (!IsWindowEnabled(tabs)) throw new InvalidOperationException("The Options tabs are disabled.");
            // TCM_SETCURFOCUS notifies the parent for ordinary tabs. It does not
            // require activating the VBE or sending keyboard shortcuts.
            if ((GetWindowStyle(tabs, -16) & 0x100) != 0)
                throw new InvalidOperationException("Unexpected button-style Options tabs.");
            int count = SendMessage(tabs, 0x1304, IntPtr.Zero, IntPtr.Zero).ToInt32();
            if (count < 1 || count > 32) throw new InvalidOperationException("Unexpected Options tab count.");
            for (int index = 0; index < count; index++)
            {
                cancellation.ThrowIfCancellationRequested();
                SendMessage(tabs, 0x1330, new IntPtr(index), IntPtr.Zero);
                var wait = Stopwatch.StartNew();
                do
                {
                    cancellation.ThrowIfCancellationRequested();
                    var lists = Descendants(dialog, "ListBox", 4912);
                    if (lists.Count == 1 && SendMessage(tabs, 0x130b, IntPtr.Zero, IntPtr.Zero).ToInt32() == index &&
                        new[] { 4913, 4914, 4935 }.All(id => Descendants(dialog, "ComboBox", id).Count == 1))
                        return lists[0];
                    Thread.Sleep(50);
                } while (wait.ElapsedMilliseconds < PageTimeoutMilliseconds);
            }
            throw new InvalidOperationException("The native color formatting page was not found in " + count + " Options tabs.");
        }

        /// <summary>Finds one color combo box and verifies its native color-index item count.</summary>
        /// <param name="dialog">Options dialog handle.</param><param name="id">Dialog control identifier.</param>
        /// <returns>The matching combo-box handle.</returns>
        /// <exception cref="InvalidOperationException">The control is missing, duplicated, or has an unexpected item count.</exception>
        private static IntPtr Control(IntPtr dialog, int id)
        {
            var matches = Descendants(dialog, "ComboBox", id);
            IntPtr handle = matches.Count == 1 ? matches[0] : IntPtr.Zero;
            if (handle == IntPtr.Zero || SendMessage(handle, 0x146, IntPtr.Zero, IntPtr.Zero).ToInt32() != 17)
                throw new InvalidOperationException("Unexpected native color selector " + id);
            return handle;
        }

        /// <summary>Selects one syntax-category row in the native color list.</summary>
        /// <param name="list">Handle of the color-category list box.</param><param name="index">Zero-based row index.</param>
        /// <exception cref="InvalidOperationException">The list rejects the selection or does not report it.</exception>
        private static void Select(IntPtr list, int index)
        {
            if (SendMessage(list, 0x186, new IntPtr(index), IntPtr.Zero).ToInt32() != index)
                throw new InvalidOperationException("The native color category could not be selected.");
            SendMessage(GetParent(list), 0x111, new IntPtr((1 << 16) | 4912), list);
        }

        /// <summary>Reads the three selected color indices for every syntax-category row.</summary>
        /// <param name="list">Handle of the category list box.</param><param name="names">Category names in list order.</param>
        /// <param name="foreground">Foreground color combo-box handle.</param><param name="background">Background color combo-box handle.</param>
        /// <param name="indicator">Indicator color combo-box handle.</param>
        /// <returns>One color row per name, in the same order.</returns>
        private static Row[] Read(IntPtr list, string[] names, IntPtr foreground, IntPtr background, IntPtr indicator)
        {
            var rows = new Row[names.Length];
            for (int index = 0; index < names.Length; index++)
            {
                Select(list, index);
                rows[index] = new Row { Name = names[index], Foreground = Current(foreground), Background = Current(background), Indicator = Current(indicator) };
            }
            VbeNativePaletteState.ValidateRows(rows);
            return rows;
        }

        /// <summary>Gets the selected color index from a native combo box.</summary>
        /// <param name="combo">Color combo-box handle.</param>
        /// <returns>Zero-based selection index, or the native no-selection value.</returns>
        private static int Current(IntPtr combo) => SendMessage(combo, 0x147, IntPtr.Zero, IntPtr.Zero).ToInt32();

        /// <summary>Sets and verifies the selected color index in a native combo box.</summary>
        /// <param name="combo">Color combo-box handle.</param><param name="id">Dialog control identifier used in errors.</param>
        /// <param name="index">Zero-based color index to select.</param>
        /// <exception cref="InvalidOperationException">The native control does not retain the requested selection.</exception>
        private static void SetColor(IntPtr combo, int id, int index)
        {
            if (SendMessage(combo, 0x14e, new IntPtr(index), IntPtr.Zero).ToInt32() != index)
                throw new InvalidOperationException("The native color selection was rejected.");
            SendMessage(GetParent(combo), 0x111, new IntPtr((1 << 16) | id), combo);
            if (Current(combo) != index) throw new InvalidOperationException("The native color selection was not retained.");
        }
    }
}

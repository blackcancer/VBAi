using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading;
using System.Windows.Automation;
using Row = CodexVBE.VbeNativePaletteState.ColorRow;

namespace CodexVBE
{
    /// <summary>Reads and writes native editor colors through an owned Options dialog.</summary>
    internal static class VbeNativePaletteDialog
    {
        private delegate bool EnumProc(IntPtr window, IntPtr parameter);
        [DllImport("user32.dll")] private static extern bool EnumWindows(EnumProc callback, IntPtr parameter);
        [DllImport("user32.dll")] private static extern bool EnumChildWindows(IntPtr parent, EnumProc callback, IntPtr parameter);
        [DllImport("user32.dll", CharSet = CharSet.Unicode)] private static extern int GetClassName(IntPtr window, StringBuilder name, int capacity);
        [DllImport("user32.dll")] private static extern int GetDlgCtrlID(IntPtr window);
        [DllImport("user32.dll", EntryPoint = "GetWindowLongW")] private static extern int GetWindowStyle(IntPtr window, int index);
        [DllImport("user32.dll")] private static extern uint GetWindowThreadProcessId(IntPtr window, out uint process);
        [DllImport("user32.dll")] private static extern IntPtr GetWindow(IntPtr window, uint command);
        [DllImport("user32.dll")] private static extern IntPtr GetParent(IntPtr window);
        [DllImport("user32.dll")] private static extern IntPtr GetDlgItem(IntPtr window, int id);
        [DllImport("user32.dll")] private static extern bool IsWindowVisible(IntPtr window);
        [DllImport("user32.dll")] private static extern bool IsWindowEnabled(IntPtr window);
        [DllImport("user32.dll")] private static extern bool PostMessage(IntPtr window, uint message, IntPtr wParam, IntPtr lParam);
        [DllImport("user32.dll")] private static extern IntPtr SendMessage(IntPtr window, uint message, IntPtr wParam, IntPtr lParam);

        internal static int OpenTimeoutMilliseconds = 10000, WorkerTimeoutMilliseconds = 30000,
            CloseTimeoutMilliseconds = 2000, PageTimeoutMilliseconds = 1000;
        internal static Func<IntPtr, uint, IntPtr, IntPtr, bool> PostDialogMessage = PostMessage;
        internal static Func<Thread, int, bool> WaitWorker = (worker, timeout) => worker.Join(timeout);
        internal static Func<IntPtr, HashSet<IntPtr>> OwnedWindows = Windows;

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

        private static string ClassName(IntPtr window)
        {
            var name = new StringBuilder(256);
            GetClassName(window, name, name.Capacity);
            return name.ToString();
        }

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

        private static IntPtr Control(IntPtr dialog, int id)
        {
            var matches = Descendants(dialog, "ComboBox", id);
            IntPtr handle = matches.Count == 1 ? matches[0] : IntPtr.Zero;
            if (handle == IntPtr.Zero || SendMessage(handle, 0x146, IntPtr.Zero, IntPtr.Zero).ToInt32() != 17)
                throw new InvalidOperationException("Unexpected native color selector " + id);
            return handle;
        }

        private static void Select(IntPtr list, int index)
        {
            if (SendMessage(list, 0x186, new IntPtr(index), IntPtr.Zero).ToInt32() != index)
                throw new InvalidOperationException("The native color category could not be selected.");
            SendMessage(GetParent(list), 0x111, new IntPtr((1 << 16) | 4912), list);
        }

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

        private static int Current(IntPtr combo) => SendMessage(combo, 0x147, IntPtr.Zero, IntPtr.Zero).ToInt32();

        private static void SetColor(IntPtr combo, int id, int index)
        {
            if (SendMessage(combo, 0x14e, new IntPtr(index), IntPtr.Zero).ToInt32() != index)
                throw new InvalidOperationException("The native color selection was rejected.");
            SendMessage(GetParent(combo), 0x111, new IntPtr((1 << 16) | id), combo);
            if (Current(combo) != index) throw new InvalidOperationException("The native color selection was not retained.");
        }
    }
}

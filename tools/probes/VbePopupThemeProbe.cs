using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Drawing;
using System.IO;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading;

// Captures a newly opened native dialog. Never treats an arbitrary editor image
// as evidence that the requested dialog or menu was actually displayed.
public static class VbePopupThemeProbe
{
    delegate bool EnumProc(IntPtr window, IntPtr parameter);
    [StructLayout(LayoutKind.Sequential)] struct Rect { public int Left, Top, Right, Bottom; }
    [DllImport("user32.dll")] static extern bool EnumWindows(EnumProc callback, IntPtr parameter);
    [DllImport("user32.dll")] static extern uint GetWindowThreadProcessId(IntPtr window, out uint process);
    [DllImport("user32.dll")] static extern bool GetWindowRect(IntPtr window, out Rect bounds);
    [DllImport("user32.dll")] static extern bool IsWindowVisible(IntPtr window);
    [DllImport("user32.dll")] static extern IntPtr GetWindow(IntPtr window, uint command);
    [DllImport("user32.dll", CharSet=CharSet.Unicode)] static extern int GetClassName(IntPtr window, StringBuilder name, int capacity);
    [DllImport("user32.dll", CharSet=CharSet.Unicode)] static extern int GetWindowText(IntPtr window, StringBuilder name, int capacity);
    [DllImport("user32.dll")] static extern bool PostMessage(IntPtr window, uint message, IntPtr wParam, IntPtr lParam);
    [DllImport("user32.dll")] static extern IntPtr GetDlgItem(IntPtr window, int id);
    [DllImport("user32.dll")] static extern bool PrintWindow(IntPtr window, IntPtr dc, uint flags);
    static Thread worker;
    public static string Error;

    static bool OwnedBy(IntPtr window, IntPtr editor)
    {
        for (int depth = 0; window != IntPtr.Zero && depth < 20; depth++)
        {
            window = GetWindow(window, 4);
            if (window == editor) return true;
        }
        return false;
    }

    static List<IntPtr> Dialogs(uint process, IntPtr editor)
    {
        var result = new List<IntPtr>();
        EnumWindows((window, parameter) => {
            uint owner; GetWindowThreadProcessId(window, out owner);
            if (owner == process && IsWindowVisible(window) && OwnedBy(window, editor))
            {
                var name = new StringBuilder(256); GetClassName(window, name, name.Capacity);
                if (name.ToString() == "#32770") result.Add(window);
            }
            return true;
        }, IntPtr.Zero);
        return result;
    }

    public static void Begin(uint process, IntPtr editor, string output, bool dialog = false)
    {
        if (!dialog) throw new NotSupportedException("Menu capture must first identify the native popup; ShowPopup is unsupported by the tested VBE.");
        if (worker != null && worker.IsAlive) throw new InvalidOperationException("A capture is already running.");
        var previous = new HashSet<IntPtr>(Dialogs(process, editor));
        Error = null;
        worker = new Thread(() => {
            IntPtr target = IntPtr.Zero;
            try
            {
                var timeout = Stopwatch.StartNew();
                while (timeout.ElapsedMilliseconds < 6000)
                {
                    var candidates = Dialogs(process, editor).FindAll(window => !previous.Contains(window));
                    if (candidates.Count > 1) throw new InvalidOperationException("Several new dialogs appeared; refusing to choose one.");
                    if (candidates.Count == 1 && GetDlgItem(candidates[0], 1) != IntPtr.Zero && GetDlgItem(candidates[0], 2) != IntPtr.Zero)
                    {
                        target = candidates[0];
                        break;
                    }
                    Thread.Sleep(100);
                }
                if (target == IntPtr.Zero) throw new TimeoutException("The requested native dialog was not observed.");
                Thread.Sleep(300);
                if (!IsWindowVisible(target)) throw new InvalidOperationException("The dialog disappeared before capture.");
                var title = new StringBuilder(256); GetWindowText(target, title, title.Capacity);
                Rect bounds;
                if (!GetWindowRect(target, out bounds)) throw new InvalidOperationException("Cannot read dialog bounds.");
                using (var image = new Bitmap(bounds.Right - bounds.Left, bounds.Bottom - bounds.Top))
                {
                    using (var graphics = Graphics.FromImage(image))
                    {
                        IntPtr dc = graphics.GetHdc();
                        try { if (!PrintWindow(target, dc, 2)) throw new InvalidOperationException("Dialog capture failed."); }
                        finally { graphics.ReleaseHdc(dc); }
                    }
                    image.Save(output + ".png", System.Drawing.Imaging.ImageFormat.Png);
                }
                File.WriteAllText(output + ".txt", "Handle=" + target + "\r\nTitle=" + title + "\r\nClass=#32770\r\nOwnerVerified=True\r\n");
            }
            catch (Exception error) { Error = error.ToString(); }
            finally
            {
                if (target != IntPtr.Zero && IsWindowVisible(target))
                {
                    PostMessage(GetDlgItem(target, 2), 0x00F5, IntPtr.Zero, IntPtr.Zero);
                    var timeout = Stopwatch.StartNew();
                    while (IsWindowVisible(target) && timeout.ElapsedMilliseconds < 3000) Thread.Sleep(50);
                    if (IsWindowVisible(target)) Error = (Error ?? "") + " Dialog did not close after Cancel.";
                }
            }
        });
        worker.IsBackground = true;
        worker.Start();
    }

    static List<IntPtr> VisibleWindows(uint process)
    {
        var result = new List<IntPtr>();
        EnumWindows((window, parameter) => {
            uint owner; GetWindowThreadProcessId(window, out owner);
            if (owner == process && IsWindowVisible(window)) result.Add(window);
            return true;
        }, IntPtr.Zero);
        return result;
    }

    public static void BeginMenu(uint process, IntPtr editor, string output)
    {
        if (worker != null && worker.IsAlive) throw new InvalidOperationException("A capture is already running.");
        var previous = new HashSet<IntPtr>(VisibleWindows(process));
        Error = null;
        worker = new Thread(() => {
            var appeared = new List<IntPtr>();
            try
            {
                var timeout = Stopwatch.StartNew();
                while (timeout.ElapsedMilliseconds < 4000)
                {
                    appeared = VisibleWindows(process).FindAll(window => !previous.Contains(window));
                    if (appeared.Count > 0) break;
                    Thread.Sleep(100);
                }
                if (appeared.Count == 0) throw new TimeoutException("No popup window appeared.");
                Thread.Sleep(300);
                appeared = VisibleWindows(process).FindAll(window => !previous.Contains(window));
                bool menuFound = appeared.Exists(window => {
                    var name = new StringBuilder(256); GetClassName(window, name, name.Capacity);
                    return name.ToString() == "MsoCommandBarPopup" && OwnedBy(window, editor);
                });
                if (!menuFound) throw new InvalidOperationException("No VBE-owned Office popup was identified.");
                var report = new StringBuilder();
                int index = 0;
                foreach (IntPtr window in appeared)
                {
                    var name = new StringBuilder(256); GetClassName(window, name, name.Capacity);
                    var title = new StringBuilder(256); GetWindowText(window, title, title.Capacity);
                    uint owner; uint thread = GetWindowThreadProcessId(window, out owner);
                    report.AppendLine("Handle=" + window + " Class=" + name + " Title=" + title + " Owner=" + GetWindow(window, 4) + " Thread=" + thread + " VbeOwned=" + OwnedBy(window, editor));
                    Rect bounds;
                    if (!GetWindowRect(window, out bounds) || bounds.Right <= bounds.Left || bounds.Bottom <= bounds.Top) continue;
                    using (var image = new Bitmap(bounds.Right - bounds.Left, bounds.Bottom - bounds.Top))
                    {
                        using (var graphics = Graphics.FromImage(image))
                        {
                            IntPtr dc = graphics.GetHdc();
                            try { if (!PrintWindow(window, dc, 2)) continue; }
                            finally { graphics.ReleaseHdc(dc); }
                        }
                        image.Save(output + "-" + index++ + ".png", System.Drawing.Imaging.ImageFormat.Png);
                    }
                }
                File.WriteAllText(output + ".txt", report.ToString());
            }
            catch (Exception error) { Error = error.ToString(); }
            finally
            {
                PostMessage(editor, 0x001F, IntPtr.Zero, IntPtr.Zero);
                foreach (IntPtr window in appeared) PostMessage(window, 0x001F, IntPtr.Zero, IntPtr.Zero);
                var timeout = Stopwatch.StartNew();
                while (appeared.Exists(IsWindowVisible) && timeout.ElapsedMilliseconds < 2000) Thread.Sleep(50);
                if (appeared.Exists(IsWindowVisible)) Error = (Error ?? "") + " Popup remained visible after cancellation.";
            }
        });
        worker.IsBackground = true;
        worker.Start();
    }

    public static bool Wait() { return worker == null || worker.Join(10000); }
}

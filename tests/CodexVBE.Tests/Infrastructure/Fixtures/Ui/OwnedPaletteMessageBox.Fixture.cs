using System;
using System.Diagnostics;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading;
using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace CodexVBE.Tests.Unit
{
    internal static class OwnedPaletteMessageBox
    {
        private delegate bool EnumProc(IntPtr window, IntPtr parameter);
        [DllImport("kernel32.dll")] private static extern uint GetCurrentThreadId();
        [DllImport("user32.dll")] private static extern bool EnumThreadWindows(uint thread, EnumProc callback, IntPtr parameter);
        [DllImport("user32.dll")] private static extern bool EnumChildWindows(IntPtr parent, EnumProc callback, IntPtr parameter);
        [DllImport("user32.dll", CharSet = CharSet.Unicode)] private static extern int GetClassName(IntPtr window, StringBuilder text, int capacity);
        [DllImport("user32.dll", CharSet = CharSet.Unicode)] private static extern int GetWindowText(IntPtr window, StringBuilder text, int capacity);
        [DllImport("user32.dll")] private static extern bool PostMessage(IntPtr window, uint message, IntPtr wParam, IntPtr lParam);
        [DllImport("user32.dll")] private static extern bool PostThreadMessage(uint thread, uint message, IntPtr wParam, IntPtr lParam);
        internal static string Dismiss(Action show, string marker)
        {
            uint ownerThread = GetCurrentThreadId(); string body = null; int closed = 0;
            var worker = new Thread(() =>
            {
                var watch = Stopwatch.StartNew();
                while (watch.ElapsedMilliseconds < 5000 && closed == 0)
                {
                    EnumThreadWindows(ownerThread, (window, parameter) =>
                    {
                        var name = new StringBuilder(128); GetClassName(window, name, name.Capacity);
                        if (name.ToString() != "#32770") return true;
                        string text = null;
                        EnumChildWindows(window, (child, unused) =>
                        {
                            var value = new StringBuilder(4096); GetWindowText(child, value, value.Capacity);
                            if (value.ToString().Contains(marker)) text = value.ToString(); return true;
                        }, IntPtr.Zero);
                        if (text != null && PostMessage(window, 0x111, new IntPtr(1), IntPtr.Zero)) { body = text; closed++; }
                        return true;
                    }, IntPtr.Zero);
                    if (closed == 0) Thread.Sleep(10);
                }
                // This disposable STA is created solely for the fixture. Ensure a failed
                // locator cannot leave its modal loop alive after the bounded test.
                if (closed == 0) PostThreadMessage(ownerThread, 0x12, IntPtr.Zero, IntPtr.Zero);
            }) { IsBackground = true };
            worker.Start(); show(); Assert.IsTrue(worker.Join(6000)); Assert.AreEqual(1, closed);
            return body;
        }
    }
}

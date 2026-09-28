using System;
using System.Diagnostics;
using System.Runtime.InteropServices;
using System.Text;
using System.Windows.Forms;
using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace CodexVBE.Tests.Unit
{
    internal static class OwnedPaletteMessageBox
    {
        private delegate bool EnumProc(IntPtr window, IntPtr parameter);
        private delegate IntPtr HookProc(int code, IntPtr first, IntPtr second);
        [StructLayout(LayoutKind.Sequential)] private struct WindowMessage
        { internal IntPtr Parameter, First; internal uint Message; internal IntPtr Window; }
        [DllImport("kernel32.dll")] private static extern uint GetCurrentThreadId();
        [DllImport("user32.dll")] private static extern bool EnumThreadWindows(uint thread, EnumProc callback, IntPtr parameter);
        [DllImport("user32.dll")] private static extern bool EnumChildWindows(IntPtr parent, EnumProc callback, IntPtr parameter);
        [DllImport("user32.dll", CharSet = CharSet.Unicode)] private static extern int GetClassName(IntPtr window, StringBuilder text, int capacity);
        [DllImport("user32.dll", CharSet = CharSet.Unicode)] private static extern int GetWindowText(IntPtr window, StringBuilder text, int capacity);
        [DllImport("user32.dll")] private static extern IntPtr GetDlgItem(IntPtr window, int id);
        [DllImport("user32.dll")] private static extern int GetDlgCtrlID(IntPtr window);
        [DllImport("user32.dll")] private static extern bool PostMessage(IntPtr window, uint message, IntPtr wParam, IntPtr lParam);
        [DllImport("user32.dll")] private static extern IntPtr SetWindowsHookEx(int kind, HookProc callback, IntPtr module, uint thread);
        [DllImport("user32.dll")] private static extern bool UnhookWindowsHookEx(IntPtr hook);
        [DllImport("user32.dll")] private static extern IntPtr CallNextHookEx(IntPtr hook, int code, IntPtr first, IntPtr second);

        internal static string Dismiss(Action show, string marker)
        {
            uint thread = GetCurrentThreadId();
            IntPtr button = IntPtr.Zero, dialog = IntPtr.Zero, cbt = IntPtr.Zero, messages = IntPtr.Zero;
            string body = null; int posts = 0, delivered = 0, destroyed = 0;
            Exception observationError = null;
            var watch = Stopwatch.StartNew();
            Action<IntPtr> inspect = window =>
            {
                if (posts != 0) return;
                var name = new StringBuilder(128); GetClassName(window, name, name.Capacity);
                if (name.ToString() != "#32770") return;
                string text = null;
                EnumChildWindows(window, (child, unused) =>
                {
                    var value = new StringBuilder(4096); GetWindowText(child, value, value.Capacity);
                    if (value.ToString().Contains(marker)) text = value.ToString();
                    return true;
                }, IntPtr.Zero);
                IntPtr ok = GetDlgItem(window, 1);
                if (text == null || ok == IntPtr.Zero) return;
                dialog = window; button = ok; body = text;
                if (PostMessage(ok, 0x00f5, IntPtr.Zero, IntPtr.Zero)) posts++;
            };
            HookProc observe = (code, first, second) =>
            {
                try
                {
                    if (code == 5) inspect(first);
                    if (code == 4 && first == dialog) destroyed++;
                }
                catch (Exception error) { observationError = error; }
                return CallNextHookEx(cbt, code, first, second);
            };
            HookProc delivery = (code, first, second) =>
            {
                if (code >= 0)
                {
                    var message = Marshal.PtrToStructure<WindowMessage>(second);
                    if (message.Window == dialog && message.Message == 0x111 && message.Parameter == button && (message.First.ToInt64() & 0xffff) == GetDlgCtrlID(button)) delivered++;
                }
                return CallNextHookEx(messages, code, first, second);
            };
            using (var timer = new Timer { Interval = 50 })
            {
                timer.Tick += (sender, args) =>
                {
                    EnumThreadWindows(thread, (window, unused) => { inspect(window); return true; }, IntPtr.Zero);
                    if (watch.ElapsedMilliseconds >= 5000 && dialog != IntPtr.Zero && destroyed == 0)
                    { PostMessage(dialog, 0x10, IntPtr.Zero, IntPtr.Zero); timer.Stop(); }
                };
                cbt = SetWindowsHookEx(5, observe, IntPtr.Zero, thread);
                messages = SetWindowsHookEx(4, delivery, IntPtr.Zero, thread);
                Assert.AreNotEqual(IntPtr.Zero, cbt); Assert.AreNotEqual(IntPtr.Zero, messages);
                try { timer.Start(); show(); }
                finally
                {
                    timer.Stop(); UnhookWindowsHookEx(messages); UnhookWindowsHookEx(cbt);
                    GC.KeepAlive(observe); GC.KeepAlive(delivery);
                }
            }
            Assert.IsNull(observationError, "The owned STA observer failed.");
            Assert.AreEqual(1, posts, "Exactly one BM_CLICK must be posted to the owned OK button.");
            Assert.AreEqual(1, delivered, "The owned native BUTTON procedure must deliver its actual command to the modal procedure.");
            Assert.AreEqual(1, destroyed, "The owned MessageBox must be destroyed before the fixture returns.");
            return body;
        }
    }
}

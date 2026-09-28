using System;
using System.Collections.Generic;
using System.Reflection;
using System.Runtime.InteropServices;
using CodexVBE;
using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace CodexVBE.Tests.Unit
{
    internal sealed class NativePaletteSchedulerFixture : IDisposable
    {
        private static readonly FieldInfo Update = typeof(VbeNativePalette).GetField("updateInProgress", BindingFlags.Static | BindingFlags.NonPublic);
        private readonly int previousUpdate = (int)Update.GetValue(null);
        private readonly Action<string, string> previousLog = LoadLog.AppendText;
        internal readonly NativePaletteFixture Recovery = new NativePaletteFixture();
        internal readonly NativeThemeVbe Vbe = new NativeThemeVbe();
        internal readonly List<bool> Targets = new List<bool>();
        internal readonly List<Exception> Errors = new List<Exception>();
        internal readonly List<string> Messages = new List<string>();
        internal readonly IntPtr Window;
        internal readonly VbeNativePalette Service;
        internal Action DuringChange;
        internal Exception Failure;

        [DllImport("user32.dll", CharSet = CharSet.Unicode)] private static extern IntPtr CreateWindowEx(int exStyle, string className,
            string title, int style, int x, int y, int width, int height, IntPtr parent, IntPtr id, IntPtr instance, IntPtr parameter);
        [DllImport("user32.dll")] private static extern bool DestroyWindow(IntPtr window);
        [DllImport("user32.dll")] private static extern bool ShowWindow(IntPtr window, int command);
        [DllImport("user32.dll")] private static extern bool EnableWindow(IntPtr window, bool enabled);

        internal NativePaletteSchedulerFixture()
        {
            Window = CreateWindowEx(0, "STATIC", "Synthetic palette scheduler", unchecked((int)0x90000000),
                -10000, -10000, 8, 8, IntPtr.Zero, IntPtr.Zero, IntPtr.Zero, IntPtr.Zero);
            Assert.AreNotEqual(IntPtr.Zero, Window);
            Update.SetValue(null, 0); LoadLog.AppendText = (path, message) => Messages.Add(message);
            Service = new VbeNativePalette(Vbe, Window, Recovery.PathName, (vbe, enabled, path) =>
            {
                Assert.AreSame(Vbe, vbe); Assert.AreEqual(Recovery.PathName, path); Targets.Add(enabled);
                Assert.AreEqual(1, Busy); DuringChange?.Invoke(); if (Failure != null) throw Failure;
            }, error => Errors.Add(error));
        }
        internal bool TimerEnabled => ((System.Windows.Forms.Timer)Field("timer")).Enabled;
        internal object Field(string name) => typeof(VbeNativePalette).GetField(name, BindingFlags.Instance | BindingFlags.NonPublic).GetValue(Service);
        internal int Busy { get => (int)Update.GetValue(null); set => Update.SetValue(null, value); }
        internal void Visible(bool enabled) => ShowWindow(Window, enabled ? 5 : 0);
        internal void Enabled(bool enabled) => EnableWindow(Window, enabled);
        internal void Tick() => typeof(VbeNativePalette).GetMethod("ApplyPending", BindingFlags.Instance | BindingFlags.NonPublic)
            .Invoke(Service, new object[] { null, EventArgs.Empty });
        public void Dispose()
        {
            Service.Dispose(); DestroyWindow(Window); Recovery.Dispose();
            Update.SetValue(null, previousUpdate); LoadLog.AppendText = previousLog;
        }
    }
}

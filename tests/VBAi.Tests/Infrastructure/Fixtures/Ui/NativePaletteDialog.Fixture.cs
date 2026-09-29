using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Threading;
using System.Windows.Forms;
using VBAi;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using Row = VBAi.VbeNativePaletteState.ColorRow;

namespace VBAi.Tests.Unit
{
    public sealed class SyntheticPaletteMainWindow { public long HWnd { get; set; } }
    public sealed class SyntheticPaletteCommandBars
    {
        public SyntheticPaletteCommand Command { get; set; }
        public object FindControl(int kind, int id) { Assert.AreEqual(1, kind); Assert.AreEqual(522, id); return Command; }
    }
    public sealed class SyntheticPaletteCommand
    {
        public bool Enabled { get; set; } = true;
        public Action Open { get; set; }
        public void Execute() => Open();
    }
    public sealed class SyntheticPaletteVbe
    {
        public SyntheticPaletteMainWindow MainWindow { get; set; }
        public SyntheticPaletteCommandBars CommandBars { get; set; }
    }

    internal sealed class NativePaletteDialogFixture : IDisposable
    {
        [StructLayout(LayoutKind.Sequential, CharSet = CharSet.Unicode)]
        private struct TabItem
        {
            internal uint Mask, State, StateMask;
            [MarshalAs(UnmanagedType.LPWStr)] internal string Text;
            internal int Capacity, Image;
            internal IntPtr Data;
        }
        [StructLayout(LayoutKind.Sequential)] private struct Notification { internal IntPtr Window; internal UIntPtr Id; internal int Code; }
        [DllImport("user32.dll", CharSet = CharSet.Unicode)] private static extern IntPtr CreateWindowEx(int exStyle, string className,
            string title, int style, int x, int y, int width, int height, IntPtr parent, IntPtr id, IntPtr instance, IntPtr parameter);
        [DllImport("user32.dll")] private static extern bool DestroyWindow(IntPtr window);
        [DllImport("user32.dll")] internal static extern bool ShowWindow(IntPtr window, int command);
        [DllImport("user32.dll")] internal static extern bool EnableWindow(IntPtr window, bool enabled);
        [DllImport("user32.dll", EntryPoint = "SendMessageW")] internal static extern IntPtr Send(IntPtr window, uint message, IntPtr wParam, IntPtr lParam);
        [DllImport("user32.dll", EntryPoint = "SendMessageW", CharSet = CharSet.Unicode)] private static extern IntPtr Text(IntPtr window, uint message, IntPtr wParam, string text);
        [DllImport("user32.dll", EntryPoint = "SendMessageW", CharSet = CharSet.Unicode)] private static extern IntPtr InsertTab(IntPtr window, uint message, IntPtr index, ref TabItem item);
        [DllImport("user32.dll", EntryPoint = "SetWindowLongW")] internal static extern int SetStyle(IntPtr window, int index, int style);

        private readonly int open = VbeNativePaletteDialog.OpenTimeoutMilliseconds, worker = VbeNativePaletteDialog.WorkerTimeoutMilliseconds,
            close = VbeNativePaletteDialog.CloseTimeoutMilliseconds, page = VbeNativePaletteDialog.PageTimeoutMilliseconds;
        private readonly Func<IntPtr, uint, IntPtr, IntPtr, bool> previousPost = VbeNativePaletteDialog.PostDialogMessage;
        private readonly Func<Thread, int, bool> previousJoin = VbeNativePaletteDialog.WaitWorker;
        private readonly Func<IntPtr, HashSet<IntPtr>> previousWindows = VbeNativePaletteDialog.OwnedWindows;
        private readonly ManualResetEventSlim dialogsPublished = new ManualResetEventSlim();
        internal readonly Form Owner = new Form { ShowInTaskbar = false, StartPosition = FormStartPosition.Manual,
            Location = new System.Drawing.Point(-10000, -10000), Size = new System.Drawing.Size(480, 360) };
        internal readonly SyntheticPaletteVbe Vbe;
        internal Row[] Current = NativePaletteFixture.Rows();
        internal readonly List<Dialog> Dialogs = new List<Dialog>();
        internal int Categories = 10, PaletteColors = 17;
        internal bool IgnoreWrites, RevertSelection, IgnoreClose;
        internal int Accepts, Cancels, Notifications;
        internal Action AfterOpen { get; set; }
        internal Dialog Latest => Dialogs[Dialogs.Count - 1];

        internal NativePaletteDialogFixture()
        {
            VbeNativePaletteDialog.OpenTimeoutMilliseconds = 500;
            VbeNativePaletteDialog.WorkerTimeoutMilliseconds = 10000;
            VbeNativePaletteDialog.CloseTimeoutMilliseconds = 250;
            VbeNativePaletteDialog.PageTimeoutMilliseconds = 80;
            Owner.Show();
            using (var initialize = new TabControl()) { var unused = initialize.Handle; }
            Vbe = new SyntheticPaletteVbe { MainWindow = new SyntheticPaletteMainWindow { HWnd = Owner.Handle.ToInt64() },
                CommandBars = new SyntheticPaletteCommandBars { Command = new SyntheticPaletteCommand { Open = () => { Open(); AfterOpen?.Invoke(); } } } };
        }
        internal void ConfigureAmbiguousOpen()
        {
            int ownerThread = Thread.CurrentThread.ManagedThreadId;
            VbeNativePaletteDialog.OwnedWindows = owner =>
            {
                if (Thread.CurrentThread.ManagedThreadId != ownerThread)
                    Assert.IsTrue(dialogsPublished.Wait(5000), "Both owned Options dialogs must be published before worker discovery.");
                return previousWindows(owner);
            };
            Vbe.CommandBars.Command.Open = () => { Open(); Open(); dialogsPublished.Set(); };
        }
        internal Dialog Open()
        {
            var dialog = new Dialog(this); Dialogs.Add(dialog); return dialog;
        }
        internal static object Invoke(string name, params object[] args)
        {
            try { return typeof(VbeNativePaletteDialog).GetMethod(name, BindingFlags.Static | BindingFlags.NonPublic).Invoke(null, args); }
            catch (TargetInvocationException error) { System.Runtime.ExceptionServices.ExceptionDispatchInfo.Capture(error.InnerException).Throw(); throw; }
        }
        internal sealed class Dialog : NativeWindow, IDisposable
        {
            private readonly NativePaletteDialogFixture fixture;
            internal IntPtr Tabs, List, Foreground, Background, Indicator, Ok, Cancel;
            internal bool MissingPage;
            internal Dialog(NativePaletteDialogFixture fixture)
            {
                this.fixture = fixture;
                IntPtr handle = CreateWindowEx(0, "#32770", "Synthetic palette Options", unchecked((int)0x90000000),
                    -10000, -10000, 400, 300, fixture.Owner.Handle, IntPtr.Zero, IntPtr.Zero, IntPtr.Zero);
                Assert.AreNotEqual(IntPtr.Zero, handle); AssignHandle(handle);
                Tabs = Child("SysTabControl32", 10, true, 0);
                var item = new TabItem { Mask = 1, Text = "General" }; InsertTab(Tabs, 0x133e, IntPtr.Zero, ref item);
                item.Text = "Editor Format"; InsertTab(Tabs, 0x133e, new IntPtr(1), ref item);
                List = Child("ListBox", 4912, false, 0x41);
                for (int index = 0; index < fixture.Categories; index++) Text(List, 0x180, IntPtr.Zero, "Category " + index);
                Foreground = ColorCombo(4913); Background = ColorCombo(4914); Indicator = ColorCombo(4935);
                Ok = Child("BUTTON", 1, true, 0); Cancel = Child("BUTTON", 2, true, 0);
            }
            internal IntPtr Child(string kind, int id, bool visible, int extraStyle)
            {
                IntPtr value = CreateWindowEx(0, kind, "", 0x40000000 | (visible ? 0x10000000 : 0) | extraStyle,
                    8, 8, 180, 120, Handle, new IntPtr(id), IntPtr.Zero, IntPtr.Zero);
                Assert.AreNotEqual(IntPtr.Zero, value); return value;
            }
            private IntPtr ColorCombo(int id)
            {
                IntPtr combo = Child("ComboBox", id, false, 3);
                for (int index = 0; index < fixture.PaletteColors; index++) Text(combo, 0x143, IntPtr.Zero, "Color " + index);
                return combo;
            }
            internal void ColorPage()
            { foreach (var window in new[] { List, Foreground, Background, Indicator }) ShowWindow(window, MissingPage ? 0 : 5); }
            internal void DeleteButtons() { Send(Handle, 0x8001, IntPtr.Zero, IntPtr.Zero); }
            protected override void WndProc(ref Message message)
            {
                if (message.Msg == 0x8001) { DestroyWindow(Ok); DestroyWindow(Cancel); Ok = Cancel = IntPtr.Zero; message.Result = IntPtr.Zero; return; }
                if (message.Msg == 0x4e)
                {
                    var notification = Marshal.PtrToStructure<Notification>(message.LParam);
                    if (notification.Window == Tabs && notification.Code == -551)
                    {
                        if (Send(Tabs, 0x130b, IntPtr.Zero, IntPtr.Zero).ToInt32() == 1) ColorPage();
                        else foreach (var window in new[] { List, Foreground, Background, Indicator }) ShowWindow(window, 0);
                    }
                }
                if (message.Msg == 0x111)
                {
                    int id = (int)(message.WParam.ToInt64() & 0xffff), code = (int)(message.WParam.ToInt64() >> 16);
                    if (id == 1 || id == 2)
                    {
                        if (!fixture.IgnoreClose) { if (id == 1) fixture.Accepts++; else fixture.Cancels++; ShowWindow(Handle, 0); }
                        message.Result = IntPtr.Zero; return;
                    }
                    if (id == 4912 && code == 1)
                    {
                        int index = Send(List, 0x188, IntPtr.Zero, IntPtr.Zero).ToInt32(); fixture.Notifications++;
                        if (index >= 0 && index < fixture.Current.Length)
                        {
                            var row = fixture.Current[index];
                            Send(Foreground, 0x14e, new IntPtr(row.Foreground), IntPtr.Zero);
                            Send(Background, 0x14e, new IntPtr(row.Background), IntPtr.Zero);
                            Send(Indicator, 0x14e, new IntPtr(row.Indicator), IntPtr.Zero);
                        }
                    }
                    if ((id == 4913 || id == 4914 || id == 4935) && code == 1)
                    {
                        fixture.Notifications++; int index = Send(List, 0x188, IntPtr.Zero, IntPtr.Zero).ToInt32();
                        int value = Send(message.LParam, 0x147, IntPtr.Zero, IntPtr.Zero).ToInt32();
                        if (fixture.RevertSelection) Send(message.LParam, 0x14e, new IntPtr(0), IntPtr.Zero);
                        else if (!fixture.IgnoreWrites)
                        {
                            if (id == 4913) fixture.Current[index].Foreground = value;
                            if (id == 4914) fixture.Current[index].Background = value;
                            if (id == 4935) fixture.Current[index].Indicator = value;
                        }
                    }
                }
                base.WndProc(ref message);
            }
            public void Dispose() { IntPtr window = Handle; ReleaseHandle(); if (window != IntPtr.Zero) DestroyWindow(window); }
        }
        public void Dispose()
        {
            foreach (var dialog in Dialogs) dialog.Dispose(); Owner.Dispose();
            VbeNativePaletteDialog.OpenTimeoutMilliseconds = open; VbeNativePaletteDialog.WorkerTimeoutMilliseconds = worker;
            VbeNativePaletteDialog.CloseTimeoutMilliseconds = close; VbeNativePaletteDialog.PageTimeoutMilliseconds = page;
            VbeNativePaletteDialog.PostDialogMessage = previousPost; VbeNativePaletteDialog.WaitWorker = previousJoin;
            VbeNativePaletteDialog.OwnedWindows = previousWindows; dialogsPublished.Set(); dialogsPublished.Dispose();
        }
    }
}

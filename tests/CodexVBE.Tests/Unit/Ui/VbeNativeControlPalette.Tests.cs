using System;
using System.Runtime.InteropServices;
using System.Threading;
using System.Windows.Forms;
using CodexVBE;
using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace CodexVBE.Tests.Unit
{
    [TestClass, TestCategory("Unit")]
    public sealed class VbeNativeControlPaletteTests
    {
        [DllImport("user32.dll")] private static extern int SendMessage(IntPtr window, uint message, IntPtr wParam, IntPtr lParam);

        [TestMethod]
        public void NativeTreeRestoresCustomColorsAndSystemColorSentinel()
        {
            OnSta(() =>
            {
                using (var tree = new TreeView())
                {
                    IntPtr handle = tree.Handle;
                    foreach (int background in new[] { 0x00345678, -1 })
                    {
                        SendMessage(handle, 0x111d, IntPtr.Zero, new IntPtr(background));
                        SendMessage(handle, 0x111e, IntPtr.Zero, new IntPtr(0x00123456));
                        try
                        {
                            VbeNativeTheme.ApplyControlPalette(handle, "SysTreeView32");
                            Assert.AreNotEqual(background, SendMessage(handle, 0x111f, IntPtr.Zero, IntPtr.Zero));
                            // Repeated activation must not replace the original snapshot.
                            VbeNativeTheme.ApplyControlPalette(handle, "SysTreeView32");
                        }
                        finally { VbeNativeTheme.RestoreControlPalette(handle, "SysTreeView32"); }
                        Assert.AreEqual(background, SendMessage(handle, 0x111f, IntPtr.Zero, IntPtr.Zero));
                        Assert.AreEqual(0x00123456, SendMessage(handle, 0x1120, IntPtr.Zero, IntPtr.Zero));
                    }
                }
            });
        }

        [TestMethod]
        public void NativeListRestoresThreeIndependentColors()
        {
            OnSta(() =>
            {
                using (var list = new ListView())
                {
                    IntPtr handle = list.Handle;
                    SendMessage(handle, 0x1001, IntPtr.Zero, new IntPtr(0x00112233));
                    SendMessage(handle, 0x1024, IntPtr.Zero, new IntPtr(0x00445566));
                    SendMessage(handle, 0x1026, IntPtr.Zero, new IntPtr(0x00778899));
                    try { VbeNativeTheme.ApplyControlPalette(handle, "SysListView32"); }
                    finally { VbeNativeTheme.RestoreControlPalette(handle, "SysListView32"); }
                    Assert.AreEqual(0x00112233, SendMessage(handle, 0x1000, IntPtr.Zero, IntPtr.Zero));
                    Assert.AreEqual(0x00445566, SendMessage(handle, 0x1023, IntPtr.Zero, IntPtr.Zero));
                    Assert.AreEqual(0x00778899, SendMessage(handle, 0x1025, IntPtr.Zero, IntPtr.Zero));
                }
            });
        }

        private static void OnSta(Action action)
        {
            Exception failure = null;
            var thread = new Thread(() => { try { action(); } catch (Exception error) { failure = error; } });
            thread.SetApartmentState(ApartmentState.STA);
            thread.Start();
            Assert.IsTrue(thread.Join(5000), "The native control test did not complete.");
            if (failure != null) throw failure;
        }
    }
}

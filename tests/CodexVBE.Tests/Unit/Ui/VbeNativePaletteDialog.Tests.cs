using System;
using System.Drawing;
using System.Runtime.InteropServices;
using System.Threading;
using System.Windows.Forms;
using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace CodexVBE.Tests.Unit
{
    [TestClass, TestCategory("Unit")]
    public sealed class VbeNativePaletteDialogTests
    {
        [TestMethod]
        public void ColorPageSelectionUsesNativeVisibilityAndAllowsDisabledIndicator()
        {
            Exception failure = null;
            var thread = new Thread(() =>
            {
                try
                {
                    using (var fixture = new PageFixture())
                    {
                        fixture.Show();
                        Assert.AreEqual(fixture.ColorList,
                            VbeNativePaletteDialog.FindColorPage(fixture.Handle, CancellationToken.None));
                        Assert.AreEqual(1, PageFixture.SendMessage(fixture.Tabs, 0x130b, IntPtr.Zero, IntPtr.Zero).ToInt32());
                        Assert.IsFalse(PageFixture.IsWindowEnabled(fixture.Indicator));
                    }
                }
                catch (Exception error) { failure = error; }
            });
            thread.SetApartmentState(ApartmentState.STA);
            thread.Start();
            Assert.IsTrue(thread.Join(10000), "Native page selection did not finish.");
            if (failure != null) throw failure;
        }

        private sealed class PageFixture : Form
        {
            [StructLayout(LayoutKind.Sequential, CharSet = CharSet.Unicode)]
            private struct TabItem
            {
                public uint Mask, State, StateMask;
                [MarshalAs(UnmanagedType.LPWStr)] public string Text;
                public int TextLength, Image;
                public IntPtr Parameter;
            }
            [StructLayout(LayoutKind.Sequential)]
            private struct Notification { public IntPtr Window; public UIntPtr Id; public int Code; }
            [DllImport("user32.dll", CharSet = CharSet.Unicode)]
            private static extern IntPtr CreateWindowEx(int exStyle, string className, string title, int style,
                int x, int y, int width, int height, IntPtr parent, IntPtr id, IntPtr instance, IntPtr parameter);
            [DllImport("user32.dll")] public static extern IntPtr SendMessage(IntPtr window, uint message, IntPtr wParam, IntPtr lParam);
            [DllImport("user32.dll", EntryPoint = "SendMessageW", CharSet = CharSet.Unicode)]
            private static extern IntPtr InsertTab(IntPtr window, uint message, IntPtr index, ref TabItem item);
            [DllImport("user32.dll")] private static extern bool ShowWindow(IntPtr window, int command);
            [DllImport("user32.dll")] private static extern bool EnableWindow(IntPtr window, bool enabled);
            [DllImport("user32.dll")] public static extern bool IsWindowEnabled(IntPtr window);

            public IntPtr Tabs, ColorList, Indicator;
            private IntPtr[] page = Array.Empty<IntPtr>();

            public PageFixture()
            {
                ShowInTaskbar = false;
                StartPosition = FormStartPosition.Manual;
                Location = new Point(-10000, -10000);
                // Initialize common controls, then use the actual native classes.
                using (var initialize = new TabControl()) { var handle = initialize.Handle; }
                Tabs = Child("SysTabControl32", 10, true);
                var item = new TabItem { Mask = 1, Text = "General" };
                InsertTab(Tabs, 0x133e, IntPtr.Zero, ref item);
                item.Text = "Colors";
                InsertTab(Tabs, 0x133e, new IntPtr(1), ref item);
                Child("ListBox", 4912, false); // A hidden duplicate must not be selected.
                ColorList = Child("ListBox", 4912, false);
                Indicator = Child("ComboBox", 4935, false);
                page = new[] { ColorList, Child("ComboBox", 4913, false), Child("ComboBox", 4914, false), Indicator };
                EnableWindow(Indicator, false);
            }

            private IntPtr Child(string className, int id, bool visible)
            {
                IntPtr window = CreateWindowEx(0, className, "", 0x40000000 | (visible ? 0x10000000 : 0),
                    0, 0, 180, 80, Handle, new IntPtr(id), IntPtr.Zero, IntPtr.Zero);
                Assert.AreNotEqual(IntPtr.Zero, window);
                return window;
            }

            protected override void WndProc(ref Message message)
            {
                if (message.Msg == 0x4e)
                {
                    var notification = Marshal.PtrToStructure<Notification>(message.LParam);
                    if (notification.Window == Tabs && notification.Code == -551)
                    {
                        bool colors = SendMessage(Tabs, 0x130b, IntPtr.Zero, IntPtr.Zero).ToInt32() == 1;
                        foreach (IntPtr window in page) ShowWindow(window, colors ? 5 : 0);
                    }
                }
                base.WndProc(ref message);
            }
        }
    }
}

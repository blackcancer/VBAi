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

namespace CodexVBE.Tests.Unit
{
    [TestClass, TestCategory("Unit"), DoNotParallelize]
    public sealed class VbeNativePaletteDialogTransactionTests
    {
        private static void WithFixture(Action<NativePaletteDialogFixture> action) => NativeThemeFixture.OnSta(() =>
        { using (var fixture = new NativePaletteDialogFixture()) action(fixture); });
        private static InvalidOperationException Failure(NativePaletteDialogFixture fixture, Func<VbeNativePaletteState.ColorRow[], VbeNativePaletteState.ColorRow[]> update,
            string fragment)
        {
            var actual = Assert.ThrowsException<InvalidOperationException>(() => VbeNativePaletteDialog.Visit(fixture.Vbe, update));
            StringAssert.Contains(actual.GetBaseException().Message, fragment);
            System.Windows.Forms.Application.DoEvents(); return actual;
        }

        [TestMethod]
        public void OwnerAndCommandGuardsRejectBeforeCreatingAnyOptionsDialog()
        {
            WithFixture(fixture =>
            {
                long owner = fixture.Vbe.MainWindow.HWnd;
                fixture.Vbe.MainWindow.HWnd = 0; Failure(fixture, rows => null, "must be visible");
                fixture.Vbe.MainWindow.HWnd = owner; fixture.Owner.Hide(); Failure(fixture, rows => null, "must be visible");
                fixture.Owner.Show(); NativePaletteDialogFixture.EnableWindow(fixture.Owner.Handle, false); Failure(fixture, rows => null, "must be visible");
                NativePaletteDialogFixture.EnableWindow(fixture.Owner.Handle, true);
                var command = fixture.Vbe.CommandBars.Command;
                fixture.Vbe.CommandBars.Command = null; Failure(fixture, rows => null, "command is unavailable");
                fixture.Vbe.CommandBars.Command = command; command.Enabled = false; Failure(fixture, rows => null, "command is unavailable");
                Assert.AreEqual(0, fixture.Dialogs.Count);
            });
        }

        [TestMethod]
        public void RealOwnedDialogReadCancelsAndWriteAcceptsAfterVerifiedNotificationReadback()
        {
            WithFixture(fixture =>
            {
                var expected = NativePaletteFixture.Rows();
                var observed = VbeNativePaletteDialog.Visit(fixture.Vbe, rows => null);
                Assert.IsTrue(VbeNativePaletteState.Equal(expected, observed)); Assert.AreEqual(1, fixture.Cancels); Assert.AreEqual(0, fixture.Accepts);
                var desired = VbeNativePaletteState.Dark(expected);
                var before = VbeNativePaletteDialog.Visit(fixture.Vbe, rows => desired);
                Assert.IsTrue(VbeNativePaletteState.Equal(expected, before)); Assert.IsTrue(VbeNativePaletteState.Equal(desired, fixture.Current));
                Assert.AreEqual(1, fixture.Accepts); Assert.IsTrue(fixture.Notifications >= 60);
                var verified = VbeNativePaletteDialog.Visit(fixture.Vbe, rows => null);
                Assert.IsTrue(VbeNativePaletteState.Equal(desired, verified)); Assert.AreEqual(2, fixture.Cancels);
            });
        }

        [TestMethod]
        public void InvalidDesiredRowsRenamedCategoriesAndCallbackErrorsCancelTheOwnedDialog()
        {
            WithFixture(fixture =>
            {
                Failure(fixture, rows => new VbeNativePaletteState.ColorRow[0], "incomplete or invalid");
                Assert.AreEqual(1, fixture.Cancels);
                Failure(fixture, rows => { var desired = VbeNativePaletteState.Dark(rows); desired[0].Name = "Renamed"; return desired; }, "categories changed");
                Assert.AreEqual(2, fixture.Cancels);
                var expected = new InvalidOperationException("synthetic update failed");
                var actual = Failure(fixture, rows => { throw expected; }, "synthetic update failed");
                Assert.AreSame(expected, actual.InnerException); Assert.AreEqual(3, fixture.Cancels); Assert.AreEqual(0, fixture.Accepts);
            });
        }

        [TestMethod]
        public void NativeControlsRejectInvalidCategoryColorAndUnretainedSelection()
        {
            WithFixture(fixture =>
            {
                var dialog = fixture.Open(); dialog.ColorPage();
                Assert.AreEqual(dialog.Foreground, NativePaletteDialogFixture.Invoke("Control", dialog.Handle, 4913));
                Assert.ThrowsException<InvalidOperationException>(() => NativePaletteDialogFixture.Invoke("Control", dialog.Handle, 9999));
                Assert.ThrowsException<InvalidOperationException>(() => NativePaletteDialogFixture.Invoke("Select", dialog.List, 100));
                NativePaletteDialogFixture.Invoke("Select", dialog.List, 0);
                Assert.AreEqual(0, NativePaletteDialogFixture.Invoke("Current", dialog.Foreground));
                Assert.ThrowsException<InvalidOperationException>(() => NativePaletteDialogFixture.Invoke("SetColor", dialog.Foreground, 4913, 100));
                NativePaletteDialogFixture.Invoke("SetColor", dialog.Foreground, 4913, 8);
                Assert.AreEqual(8, fixture.Current[0].Foreground);
                fixture.RevertSelection = true;
                Assert.ThrowsException<InvalidOperationException>(() => NativePaletteDialogFixture.Invoke("SetColor", dialog.Foreground, 4913, 8));
                var read = (VbeNativePaletteState.ColorRow[])NativePaletteDialogFixture.Invoke("Read", dialog.List,
                    Array.ConvertAll(NativePaletteFixture.Rows(), row => row.Name), dialog.Foreground, dialog.Background, dialog.Indicator);
                Assert.IsTrue(VbeNativePaletteState.Equal(fixture.Current, read));
            });
            WithFixture(fixture =>
            {
                fixture.PaletteColors = 16; var dialog = fixture.Open(); dialog.ColorPage();
                Assert.ThrowsException<InvalidOperationException>(() => NativePaletteDialogFixture.Invoke("Control", dialog.Handle, 4913));
            });
        }

        [TestMethod]
        public void IgnoredWritesAndUnretainedSelectionsNeverAcceptAChangedPalette()
        {
            WithFixture(fixture =>
            {
                fixture.IgnoreWrites = true;
                Failure(fixture, rows => VbeNativePaletteState.Dark(rows), "did not retain the requested values");
                Assert.AreEqual(1, fixture.Cancels); Assert.AreEqual(0, fixture.Accepts);
                Assert.IsTrue(VbeNativePaletteState.Equal(NativePaletteFixture.Rows(), fixture.Current));
                fixture.IgnoreWrites = false; fixture.RevertSelection = true;
                Failure(fixture, rows => VbeNativePaletteState.Dark(rows), "selection was not retained");
                Assert.AreEqual(2, fixture.Cancels); Assert.AreEqual(0, fixture.Accepts);
            });
        }

        [TestMethod]
        public void WrongCategoryCountAndSelectorCountCancelBeforeAnyUpdateCallback()
        {
            WithFixture(fixture =>
            {
                fixture.Categories = 9;
                Failure(fixture, rows => { Assert.Fail("Invalid categories must not reach update."); return null; }, "category count");
                Assert.AreEqual(1, fixture.Cancels);
            });
            WithFixture(fixture =>
            {
                fixture.PaletteColors = 16;
                Failure(fixture, rows => { Assert.Fail("Invalid selectors must not reach update."); return null; }, "selector");
                Assert.AreEqual(1, fixture.Cancels);
            });
        }

        [TestMethod]
        public void TabsRejectDisabledButtonStyleInvalidCountsCancellationAndMissingPage()
        {
            WithFixture(fixture =>
            {
                var dialog = fixture.Open();
                NativePaletteDialogFixture.EnableWindow(dialog.Tabs, false);
                Assert.ThrowsException<InvalidOperationException>(() => VbeNativePaletteDialog.FindColorPage(dialog.Handle, CancellationToken.None));
                NativePaletteDialogFixture.EnableWindow(dialog.Tabs, true);
                NativePaletteDialogFixture.SetStyle(dialog.Tabs, -16, 0x40000100 | 0x10000000);
                Assert.ThrowsException<InvalidOperationException>(() => VbeNativePaletteDialog.FindColorPage(dialog.Handle, CancellationToken.None));
                NativePaletteDialogFixture.SetStyle(dialog.Tabs, -16, 0x50000000);
                using (var canceled = new CancellationTokenSource())
                { canceled.Cancel(); Assert.ThrowsException<OperationCanceledException>(() => VbeNativePaletteDialog.FindColorPage(dialog.Handle, canceled.Token)); }
                dialog.MissingPage = true;
                Assert.ThrowsException<InvalidOperationException>(() => VbeNativePaletteDialog.FindColorPage(dialog.Handle, CancellationToken.None));
                NativePaletteDialogFixture.Send(dialog.Tabs, 0x1309, IntPtr.Zero, IntPtr.Zero);
                Assert.ThrowsException<InvalidOperationException>(() => VbeNativePaletteDialog.FindColorPage(dialog.Handle, CancellationToken.None));
            });
        }

        [TestMethod]
        public void PreexistingDialogIsExcludedAndAmbiguousNewDialogsAreRejected()
        {
            WithFixture(fixture =>
            {
                var previous = fixture.Open();
                var observed = VbeNativePaletteDialog.Visit(fixture.Vbe, rows => null);
                Assert.IsTrue(VbeNativePaletteState.Equal(NativePaletteFixture.Rows(), observed));
                Assert.AreEqual(1, fixture.Cancels); Assert.AreEqual(2, fixture.Dialogs.Count);
            });
            WithFixture(fixture =>
            {
                fixture.ConfigureAmbiguousOpen();
                Failure(fixture, rows => null, "ambiguous"); Assert.AreEqual(0, fixture.Cancels);
            });
        }

        [TestMethod]
        public void MissingDialogAndCommandFailureAreBoundedAndReleaseTheWorker()
        {
            WithFixture(fixture =>
            {
                fixture.Vbe.CommandBars.Command.Open = () => { };
                Failure(fixture, rows => null, "did not open"); Assert.AreEqual(0, fixture.Dialogs.Count);
            });
            WithFixture(fixture =>
            {
                var expected = new InvalidOperationException("synthetic Options execution failed");
                fixture.Vbe.CommandBars.Command.Open = () => { throw expected; };
                Assert.AreSame(expected, Assert.ThrowsException<InvalidOperationException>(() => VbeNativePaletteDialog.Visit(fixture.Vbe, rows => null)));
            });
        }

        [TestMethod]
        public void MissingCloseButtonPostFailureJoinRefusalAndVisibleDialogAreReported()
        {
            WithFixture(fixture => Failure(fixture, rows => { fixture.Latest.DeleteButtons(); return null; }, "could not be closed"));
            WithFixture(fixture =>
            {
                int calls = 0;
                VbeNativePaletteDialog.PostDialogMessage = (window, message, first, second) => { calls++; Assert.AreEqual((uint)0xf5, message); return false; };
                Failure(fixture, rows => null, "could not be closed"); Assert.AreEqual(1, calls);
            });
            WithFixture(fixture =>
            {
                VbeNativePaletteDialog.WaitWorker = (thread, timeout) => { Assert.IsTrue(thread.Join(timeout)); return false; };
                Failure(fixture, rows => null, "worker did not finish");
            });
            WithFixture(fixture =>
            {
                fixture.IgnoreClose = true;
                Failure(fixture, rows => null, "did not finish closing");
            });
        }
        [TestMethod]
        public void NativeVisibilityRejectsNullAndInvalidOwnedDialogHandles()
        {
            Assert.IsFalse(IsWindowVisible(IntPtr.Zero)); Assert.IsFalse(IsWindowVisible(new IntPtr(-1)));
            WithFixture(fixture =>
            {
                var dialog = fixture.Open(); Assert.IsTrue(IsWindowVisible(dialog.Handle));
                NativePaletteDialogFixture.ShowWindow(dialog.Handle, 0); Assert.IsFalse(IsWindowVisible(dialog.Handle));
            });
        }

        [DllImport("user32.dll")] private static extern bool IsWindowVisible(IntPtr window);
    }
}

using System.Drawing;
using System;
using System.Diagnostics;
using System.Linq;
using System.Runtime.InteropServices;
using System.Threading;
using System.Windows.Forms;
using VBAi;
using VBAi.Tests.Infrastructure;
using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace VBAi.Tests.Unit
{
    [TestClass]
    public sealed class UiInputControlsTests
    {
        [DllImport("user32.dll", SetLastError = true)]
        private static extern IntPtr SendMessageTimeout(IntPtr hwnd, uint message, IntPtr wParam,
            IntPtr lParam, uint flags, uint milliseconds, out IntPtr result);

        [STATestMethod]
        public void ExternalListSelectionUpdatesCachedItemsAndEventsBeforeHandleDestruction()
        {
            using (var list = new UiListBox())
            {
                list.Items.Add("First"); list.Items.Add("Second");
                var handle = list.Handle;
                // Enumeration materializes the managed selection cache, as UIA
                // selection queries do, before an external native selection.
                foreach (object item in list.SelectedItems) Assert.Fail("Unexpected initial selection.");
                int changes = 0; list.SelectedIndexChanged += (sender, args) => changes++;
                SendExternalSelection(handle, 0);
                Assert.AreEqual("First", list.SelectedItem);
                Assert.AreEqual("First", list.Text);
                Assert.AreEqual(1, changes);
                SendExternalSelection(handle, 1);
                Assert.AreEqual("Second", list.SelectedItem);
                CollectionAssert.AreEqual(new[] { "Second" }, new System.Collections.Generic.List<string>(
                    System.Linq.Enumerable.Cast<string>(list.SelectedItems)).ToArray());
                Assert.AreEqual(2, changes);
                // Destroying the handle reads Text in .NET Framework. This must
                // preserve both selection and items when the handle is recreated.
                list.BorderStyle = BorderStyle.None;
                Assert.AreEqual("Second", list.SelectedItem);
                Assert.AreEqual(2, list.Items.Count);
            }
        }

        [STATestMethod]
        public void ExternalListDeselectionAndRejectedRequestsDoNotInventSelectionEvents()
        {
            using (var list = new UiListBox())
            {
                list.Items.Add("First"); var handle = list.Handle;
                list.SelectedIndex = 0;
                foreach (object item in list.SelectedItems) Assert.AreEqual("First", item);
                int changes = 0; list.SelectedIndexChanged += (sender, args) => changes++;
                SendExternalSelection(handle, 0);
                SendExternalSelection(handle, 99);
                Assert.AreEqual(0, changes);
                Assert.AreEqual("First", list.SelectedItem);
                SendExternalSelection(handle, -1);
                Assert.AreEqual(1, changes);
                Assert.IsNull(list.SelectedItem); Assert.AreEqual("", list.Text);
                Assert.AreEqual(0, list.SelectedItems.Count);
            }
        }

        [STATestMethod]
        public void ManagedListSelectionKeepsOneEventPerChangeAndNormalItemRefresh()
        {
            using (var list = new UiListBox())
            {
                list.Items.Add("First"); list.Items.Add("Second"); var handle = list.Handle;
                int changes = 0; list.SelectedIndexChanged += (sender, args) => changes++;
                list.SelectedIndex = 0; list.SelectedIndex = 0; list.SelectedIndex = 1;
                Assert.AreEqual(2, changes); Assert.AreEqual("Second", list.SelectedItem);
                list.Items.Clear(); list.Items.Add("Replacement"); list.SelectedIndex = 0;
                Assert.AreEqual("Replacement", list.SelectedItem);
                Assert.AreEqual("Replacement", list.Text);
            }
        }

        [STATestMethod]
        public void MultiModeNativeSelectionRepairsPrimedCacheWithoutDuplicateManagedEvents()
        {
            foreach (var mode in new[] { SelectionMode.MultiExtended, SelectionMode.MultiSimple })
            using (var list = new UiListBox { SelectionMode = mode })
            {
                list.Items.Add("First"); list.Items.Add("Second");
                IntPtr handle = list.Handle;
                foreach (object item in list.SelectedItems) Assert.Fail("Unexpected initial selection.");
                int changes = 0; list.SelectedIndexChanged += (sender, args) => changes++;
                SendExternalSelection(handle, 1, 0x185, 0);
                CollectionAssert.AreEqual(new[] { "First" }, System.Linq.Enumerable.Cast<string>(list.SelectedItems).ToArray());
                Assert.AreEqual(1, changes);
                SendExternalSelection(handle, 1, 0x185, 0);
                SendExternalSelection(handle, 1, 0x185, 99);
                Assert.AreEqual(1, changes);
                list.SetSelected(1, true);
                CollectionAssert.AreEqual(new[] { "First", "Second" }, System.Linq.Enumerable.Cast<string>(list.SelectedItems).ToArray());
                Assert.AreEqual(2, changes);
                SendExternalSelection(handle, 0, 0x185, 0);
                CollectionAssert.AreEqual(new[] { "Second" }, System.Linq.Enumerable.Cast<string>(list.SelectedItems).ToArray());
                Assert.AreEqual(3, changes);
                list.BorderStyle = BorderStyle.None;
                CollectionAssert.AreEqual(new[] { "Second" }, System.Linq.Enumerable.Cast<string>(list.SelectedItems).ToArray());
            }
        }

        private static void SendExternalSelection(IntPtr handle, int index)
        {
            SendExternalSelection(handle, index, 0x186, 0);
        }

        private static void SendExternalSelection(IntPtr handle, int index, uint message, int lParam)
        {
            bool done = false; IntPtr delivery = IntPtr.Zero;
            var sender = new Thread(() => {
                IntPtr result;
                delivery = SendMessageTimeout(handle, message, new IntPtr(index), new IntPtr(lParam), 2, 5000, out result);
                Volatile.Write(ref done, true);
            }) { IsBackground = true };
            sender.Start(); var clock = Stopwatch.StartNew();
            while (!Volatile.Read(ref done) && clock.ElapsedMilliseconds < 6000)
            { Application.DoEvents(); Thread.Sleep(1); }
            Assert.IsTrue(Volatile.Read(ref done), "External list message exceeded its bounded delivery deadline.");
            Assert.IsTrue(sender.Join(1000));
            Assert.AreNotEqual(IntPtr.Zero, delivery, "Native message delivery failed.");
        }

        /// <summary>Paints owner-drawn selections and native edit backgrounds without changing values or retaining stale brushes.</summary>
        [STATestMethod]
        public void ComboSelectionAndNativeBrushFollowPaletteDirectionAndEnabledState()
        {
            using (var combo = new UiComboBox { Width = 180, DropDownStyle = ComboBoxStyle.DropDown })
            using (var image = new Bitmap(180, 40))
            using (var graphics = Graphics.FromImage(image))
            {
                var handle = combo.Handle; combo.Items.Add("Selected model"); combo.SelectedIndex = 0;
                foreach (var direction in new[] { RightToLeft.No, RightToLeft.Yes })
                foreach (bool selected in new[] { false, true })
                {
                    combo.RightToLeft = direction; combo.BackColor = Color.FromArgb(30, 34, 42); combo.ForeColor = Color.White;
                    UiInvoke.Call(typeof(UiComboBox), "OnDrawItem", combo, new DrawItemEventArgs(graphics, combo.Font, new Rectangle(0, 0, 180, 40), 0, selected ? DrawItemState.Selected : DrawItemState.None));
                    Assert.AreEqual((selected ? SystemColors.Highlight : combo.BackColor).ToArgb(), image.GetPixel(1, 20).ToArgb());
                    Assert.AreEqual(0, combo.SelectedIndex);
                }
                UiInvoke.Call(typeof(UiComboBox), "OnDrawItem", combo, new DrawItemEventArgs(graphics, combo.Font, new Rectangle(0, 0, 180, 40), -1, DrawItemState.None));
                var dc = graphics.GetHdc();
                try
                {
                    foreach (bool enabled in new[] { true, false })
                    foreach (int messageId in new[] { 0x133, 0x138 })
                    {
                        combo.Enabled = enabled;
                        var args = new object[] { Message.Create(handle, messageId, dc, System.IntPtr.Zero) };
                        UiInvoke.Call(typeof(UiComboBox), "WndProc", combo, args);
                        Assert.AreNotEqual(System.IntPtr.Zero, ((Message)args[0]).Result);
                    }
                }
                finally { graphics.ReleaseHdc(dc); }
                Assert.AreNotEqual(System.IntPtr.Zero, UiInvoke.Field<System.IntPtr>(combo, "inputBrush"));
                combo.BackColor = Color.White;
                Assert.AreEqual(System.IntPtr.Zero, UiInvoke.Field<System.IntPtr>(combo, "inputBrush"));
            }
        }

        [STATestMethod]
        public void SharedFieldsKeepNativeEditingSelectionAndCheckStateWhenThemesChange()
        {
            using (var theme = new ThemeScope())
            using (var form = new Form())
            using (var text = new UiTextBox { Text = "Module1", Dock = DockStyle.Top })
            using (var list = new UiCheckedListBox { Dock = DockStyle.Fill })
            using (var combo = new UiComboBox { Dock = DockStyle.Bottom, DropDownStyle = ComboBoxStyle.DropDown })
            {
                form.Controls.Add(list); form.Controls.Add(combo); form.Controls.Add(text);
                list.Items.Add("Module1", true); list.Items.Add("Module2", false);
                combo.Items.Add("main"); combo.Text = "feature/custom";
                text.Select(6, 1); text.SelectedText = "2";
                foreach (var choice in new[] { ThemeChoice.Dark, ThemeChoice.Light })
                {
                    ThemeScope.SetChoice(choice); UiTheme.Apply(form);
                    Assert.AreEqual("Module2", text.Text);
                    Assert.AreEqual("feature/custom", combo.Text);
                    Assert.IsTrue(list.GetItemChecked(0)); Assert.IsFalse(list.GetItemChecked(1));
                    Assert.AreEqual(text.BackColor, combo.BackColor);
                    Assert.AreEqual(text.BackColor, list.BackColor);
                    Assert.AreEqual(BorderStyle.FixedSingle, text.BorderStyle);
                    Assert.AreEqual(BorderStyle.FixedSingle, list.BorderStyle);
                    using (var bitmap = new Bitmap(400, 300)) form.DrawToBitmap(bitmap, new Rectangle(0, 0, 400, 300));
                }
            }
        }

        [STATestMethod]
        public void ChatAndSettingsSelectorsShareNativeBehaviorAndVisualBase()
        {
            using (var chat = new ChatChoiceBox())
            using (var settings = new ThemedComboBox())
            {
                Assert.IsInstanceOfType(chat, typeof(UiComboBox));
                Assert.IsInstanceOfType(settings, typeof(UiComboBox));
                foreach (UiComboBox combo in new UiComboBox[] { chat, settings })
                {
                    combo.Items.Add("First"); combo.Items.Add("Second"); combo.SelectedIndex = 1;
                    Assert.AreEqual("Second", combo.SelectedItem);
                    Assert.IsTrue(combo.ItemHeight >= combo.Font.Height);
                    combo.RightToLeft = RightToLeft.Yes;
                    using (var bitmap = new Bitmap(180, 30)) combo.DrawToBitmap(bitmap, new Rectangle(0, 0, 180, 30));
                    Assert.AreEqual(1, combo.SelectedIndex);
                }
            }
        }
    }
}

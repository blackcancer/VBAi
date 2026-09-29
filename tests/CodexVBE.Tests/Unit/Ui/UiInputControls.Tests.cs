using System.Drawing;
using System.Windows.Forms;
using CodexVBE;
using CodexVBE.Tests.Infrastructure;
using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace CodexVBE.Tests.Unit
{
    [TestClass]
    public sealed class UiInputControlsTests
    {
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

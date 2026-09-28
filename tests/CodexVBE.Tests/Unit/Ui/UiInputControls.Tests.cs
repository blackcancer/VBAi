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

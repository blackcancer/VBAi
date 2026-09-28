using System;
using System.Drawing;
using System.Linq;
using System.Windows.Forms;
using CodexVBE;
using CodexVBE.Tests.Infrastructure;
using Microsoft.VisualStudio.TestTools.UnitTesting;
namespace CodexVBE.Tests.Unit
{
    [TestClass]
    public sealed class ChatDesignerControlsTests
    {
        [STATestMethod, TestCategory("Unit")]
        public void RoundedButtonsAndPanelsPaintSizingParentHoverPressAndFocusStates()
        {
            using (var bitmap = new Bitmap(240, 100))
            using (var graphics = Graphics.FromImage(bitmap))
            using (var form = new Form { Left = -10000, Top = -10000, ShowInTaskbar = false, Size = new Size(400, 200) })
            using (var button = new ChatActionButton { Text = "Action", Size = new Size(180, 40) })
            using (var panel = new ChatComposerPanel { Size = new Size(180, 60) })
            {
                using (var path = ChatControlPainting.Rounded(new Rectangle(0, 0, 20, 20), -1)) Assert.IsTrue(path.PointCount > 0);
                foreach (var size in new[] { new Size(1, 40), new Size(40, 1), new Size(180, 40) }) { button.Size = size; UiInvoke.Call(typeof(ChatActionButton), "OnPaint", button, new PaintEventArgs(graphics, button.ClientRectangle)); panel.Size = size; UiInvoke.Call(typeof(ChatComposerPanel), "OnPaintBackground", panel, new PaintEventArgs(graphics, panel.ClientRectangle)); }
                button.Size = new Size(180, 40); panel.Size = new Size(180, 60);
                form.Controls.Add(button); form.Controls.Add(panel); panel.Location = new Point(0, 50); form.Show();
                foreach (var enabled in new[] { false, true })
                {
                    button.Enabled = enabled;
                    foreach (var mouse in new[] { "OnMouseLeave", "OnMouseEnter", "OnMouseDown", "OnMouseUp" })
                    {
                        UiInvoke.Call(typeof(ChatActionButton), mouse, button, mouse == "OnMouseDown" || mouse == "OnMouseUp" ? (object)new MouseEventArgs(MouseButtons.Left, 1, 0, 0, 0) : EventArgs.Empty);
                        UiInvoke.Call(typeof(ChatActionButton), "OnPaint", button, new PaintEventArgs(graphics, button.ClientRectangle));
                    }
                }
                button.Focus(); UiInvoke.Call(typeof(Control), "WndProc", button, new object[] { Message.Create(button.Handle, 0x0128, new IntPtr(0x10002), IntPtr.Zero) }); UiInvoke.Call(typeof(ChatActionButton), "OnPaint", button, new PaintEventArgs(graphics, button.ClientRectangle)); Assert.IsTrue(button.Focused);
                UiInvoke.Call(typeof(ChatComposerPanel), "OnPaintBackground", panel, new PaintEventArgs(graphics, panel.ClientRectangle)); Assert.AreEqual(panel.BackColor.ToArgb(), bitmap.GetPixel(90, 30).ToArgb());
            }
        }
        [STATestMethod, TestCategory("Unit")]
        public void ChoiceBoxPaintsItemsAndNativeMessagesAcrossSelectionFocusAndSizing()
        {
            using (var combo = new ChatChoiceBox { Size = new Size(180, 30) })
            using (var form = new Form { Left = -10000, Top = -10000, ShowInTaskbar = false, Size = new Size(300, 200) })
            using (var bitmap = new Bitmap(200, 60))
            using (var graphics = Graphics.FromImage(bitmap))
            {
                combo.Items.Add("Choice");
                foreach (var index in new[] { -1, 0 }) foreach (var selected in new[] { false, true }) { UiInvoke.Call(typeof(ChatChoiceBox), "OnDrawItem", combo, new DrawItemEventArgs(graphics, combo.Font, new Rectangle(0, 0, 180, 30), index, selected ? DrawItemState.Selected : DrawItemState.None)); if (index == 0) Assert.AreEqual((selected ? Color.FromArgb(239, 246, 255) : combo.BackColor).ToArgb(), bitmap.GetPixel(175, 15).ToArgb()); }
                foreach (var parent in new[] { false, true })
                {
                    if (parent) { form.Controls.Add(combo); form.Show(); combo.Focus(); } else { var handle = combo.Handle; }
                    foreach (var enabled in new[] { false, true }) foreach (var selection in new[] { -1, 0 })
                    {
                        combo.Enabled = enabled; combo.SelectedIndex = selection; if (parent && enabled) combo.Focus();
                        foreach (var msg in new[] { 0, 0x000F, 0x0317, 0x0318 })
                        {
                            IntPtr hdc = msg == 0x0317 || msg == 0x0318 ? graphics.GetHdc() : IntPtr.Zero;
                            try { UiInvoke.Call(typeof(ChatChoiceBox), "WndProc", combo, new object[] { Message.Create(combo.Handle, msg, hdc, IntPtr.Zero) }); } finally { if (hdc != IntPtr.Zero) graphics.ReleaseHdc(hdc); }
                        }
                    }
                }
                UiInvoke.Call(typeof(ChatChoiceBox), "OnGotFocus", combo, EventArgs.Empty); UiInvoke.Call(typeof(ChatChoiceBox), "OnLostFocus", combo, EventArgs.Empty);
                foreach (var size in new[] { new Size(2, 30), new Size(30, 2) }) { combo.Size = size; UiInvoke.Call(typeof(ChatChoiceBox), "WndProc", combo, new object[] { Message.Create(combo.Handle, 0x000F, IntPtr.Zero, IntPtr.Zero) }); }
            }
        }
    }
}

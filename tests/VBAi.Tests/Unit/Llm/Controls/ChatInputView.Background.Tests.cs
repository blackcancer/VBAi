using Microsoft.VisualStudio.TestTools.UnitTesting;
using System.Drawing;
using System.Windows.Media;
using Color = System.Drawing.Color;
using VBAi.Tests.Infrastructure;

namespace VBAi.Tests.Unit
{
    /// <summary>Checks opaque composition without instantiating live add-in services.</summary>
    [TestClass, TestCategory("Unit")]
    public sealed class ChatInputViewBackgroundTests
    {
        /// <summary>Checks lazy creation, palette changes and preservation of the draft.</summary>
        [STATestMethod]
        public void LazyEditorUsesOpaqueThemeAndRetainsDraftAcrossEveryPaletteChange()
        {
            using (var view = new ChatInputView())
            {
                view.BackColor = Color.FromArgb(248, 250, 252);
                Assert.IsNull(UiInvoke.Field<System.Windows.Controls.TextBox>(view, "editor"));
                var editor = view.Editor;
                editor.Text = "Conserve ce brouillon #Projet.Module";
                editor.Select(4, 8);
                foreach (var color in new[] { Color.FromArgb(248, 250, 252), Color.FromArgb(22, 24, 28), Color.White, Color.FromArgb(31, 41, 55) })
                {
                    view.BackColor = color;
                    var brush = (SolidColorBrush)editor.Background;
                    Assert.AreEqual((byte)255, brush.Color.A);
                    Assert.AreEqual(color.R, brush.Color.R); Assert.AreEqual(color.G, brush.Color.G); Assert.AreEqual(color.B, brush.Color.B);
                    Assert.AreEqual(color, UiInvoke.Field<ChatContentHost>(view, "host").BackColor);
                    Assert.AreEqual(color, UiInvoke.Field<System.Windows.Forms.TextBox>(view, "previewEditor").BackColor);
                    Assert.AreEqual("Conserve ce brouillon #Projet.Module", editor.Text);
                    Assert.AreEqual(4, editor.SelectionStart); Assert.AreEqual(8, editor.SelectionLength);
                    Assert.AreSame(editor, view.Editor);
                }
            }
        }
    }
}

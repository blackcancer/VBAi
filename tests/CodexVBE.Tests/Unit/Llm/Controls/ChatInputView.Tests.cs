using System.Drawing;
using System.Reflection;
using System.Windows.Forms;
using CodexVBE;
using Microsoft.VisualStudio.TestTools.UnitTesting;
namespace CodexVBE.Tests.Unit
{
    [TestClass, TestCategory("Unit")]
    public sealed class ChatInputViewTests
    {
        [STATestMethod]
        public void DesignerInputHasNoEngineAndAppearanceFlowsToLazyEditor()
        {
            using (var view = new ChatInputView())
            using (var font = new Font("Segoe UI", 12F, FontStyle.Bold | FontStyle.Italic))
            {
                Assert.IsNull(typeof(ChatInputView).GetField("editor", BindingFlags.Instance | BindingFlags.NonPublic).GetValue(view));
                Assert.AreEqual(1, view.Controls.Find("previewEditor", true).Length);
                view.InputPadding = new Padding(4, 6, 8, 10);
                Assert.AreEqual(view.InputPadding, view.Controls.Find("previewPanel", true)[0].Padding);
                view.Font = font; view.ForeColor = Color.Blue; view.SpellCheckEnabled = false;
                var editor = view.Editor;
                Assert.AreSame(editor, view.Editor);
                Assert.AreEqual(16.0, editor.FontSize, 0.001);
                Assert.AreEqual(System.Windows.FontWeights.Bold, editor.FontWeight);
                Assert.AreEqual(System.Windows.FontStyles.Italic, editor.FontStyle);
                Assert.AreEqual(new System.Windows.Thickness(4, 6, 8, 10), editor.Padding);
                Assert.IsFalse(editor.SpellCheck.IsEnabled);
                view.SpellCheckEnabled = true;
                Assert.IsTrue(editor.SpellCheck.IsEnabled);
                Assert.ThrowsException<System.ArgumentOutOfRangeException>(() => view.InputPadding = new Padding(-1));
            }
        }
    }
}

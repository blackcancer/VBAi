using System.Drawing;
using System.Reflection;
using System.Windows.Forms;
using VBAi;
using VBAi.Tests.Infrastructure;
using Microsoft.VisualStudio.TestTools.UnitTesting;
namespace VBAi.Tests.Unit
{
    [TestClass, TestCategory("Unit")]
    public sealed class ChatInputViewTests
    {
        /// <summary>Qualifie le cycle de propriétés Designer et chacune des quatre bornes de marge.</summary>
        [STATestMethod]
        public void DesignerPaddingSerializationResetAndEveryInvalidEdgeAreChecked()
        {
            using (var font = new Font("Segoe UI", 11F, FontStyle.Regular))
            using (var view = new ChatInputView())
            {
                Assert.AreEqual(false, UiInvoke.Call(typeof(ChatInputView), "ShouldSerializeInputPadding", view));
                view.InputPadding = new Padding(1, 2, 3, 4);
                Assert.AreEqual(true, UiInvoke.Call(typeof(ChatInputView), "ShouldSerializeInputPadding", view));
                UiInvoke.Call(typeof(ChatInputView), "ResetInputPadding", view);
                Assert.AreEqual(new Padding(12), view.InputPadding);
                foreach (var padding in new[] { new Padding(-1, 0, 0, 0), new Padding(0, -1, 0, 0),
                    new Padding(0, 0, -1, 0), new Padding(0, 0, 0, -1) })
                    Assert.ThrowsException<System.ArgumentOutOfRangeException>(() => view.InputPadding = padding);
                var editor = view.Editor;
                view.Font = font;
                view.ForeColor = Color.Green;
                Assert.AreEqual(System.Windows.FontWeights.Normal, editor.FontWeight);
                Assert.AreEqual(System.Windows.FontStyles.Normal, editor.FontStyle);
                Assert.AreEqual(System.Windows.Media.Colors.Green, ((System.Windows.Media.SolidColorBrush)editor.Foreground).Color);
                view.InputPadding = new Padding(3);
                Assert.AreEqual(new System.Windows.Thickness(3), editor.Padding);
                UiInvoke.Call(typeof(ChatInputView), "Dispose", view, false);
                Assert.IsFalse(view.IsDisposed);
                var components = UiInvoke.Field<System.ComponentModel.IContainer>(view, "components");
                components.Dispose();
                typeof(ChatInputView).GetField("components", BindingFlags.Instance | BindingFlags.NonPublic).SetValue(view, null);
            }
        }

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
                Assert.AreEqual(editor.ToolTip, System.Windows.Automation.AutomationProperties.GetHelpText(editor));
                StringAssert.Contains((string)editor.ToolTip, UiText.Get("Write your request. Enter sends it; Shift+Enter adds a line. Use # or @ to reference VBA code and / to see chat commands."));
                Assert.AreEqual(16.0, editor.FontSize, 0.001);
                Assert.AreEqual(System.Windows.FontWeights.Bold, editor.FontWeight);
                Assert.AreEqual(System.Windows.FontStyles.Italic, editor.FontStyle);
                Assert.AreEqual(new System.Windows.Thickness(4, 6, 8, 10), editor.Padding);
                Assert.IsFalse(editor.SpellCheck.IsEnabled);
                view.SpellCheckEnabled = true;
                Assert.IsTrue(editor.SpellCheck.IsEnabled);
                Assert.IsTrue(view.SpellCheckEnabled);
                Assert.ThrowsException<System.ArgumentOutOfRangeException>(() => view.InputPadding = new Padding(-1));
            }
        }
    }
}

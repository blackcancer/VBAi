using System.Reflection;
using System.Windows.Forms;
using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace VBAi.Tests.Unit
{
    [TestClass, TestCategory("Unit")]
    public sealed class TestSupportReviewDialogDesignerTests
    {
        [STATestMethod]
        public void DesignerDisposalSupportsFinalizerStyleAndAlreadyAbsentComponents()
        {
            var dialog = new TestSupportReviewDialog();
            typeof(TestSupportReviewDialog).GetMethod("Dispose", BindingFlags.Instance | BindingFlags.NonPublic).Invoke(dialog, new object[] { false });
            dialog.Dispose();
            var noComponents = new TestSupportReviewDialog();
            var field = typeof(TestSupportReviewDialog).GetField("components", BindingFlags.Instance | BindingFlags.NonPublic);
            ((System.ComponentModel.IContainer)field.GetValue(noComponents)).Dispose();
            field.SetValue(noComponents, null);
            noComponents.Dispose();
            Assert.IsTrue(noComponents.IsDisposed);
        }
        [STATestMethod]
        public void ReviewStartsWithoutHostAndRequiresExplicitApply()
        {
            using (var dialog = new TestSupportReviewDialog())
            {
                Assert.AreEqual(DialogResult.None, dialog.DialogResult);
                Assert.AreEqual(DialogResult.Cancel, ((Button)dialog.CancelButton).DialogResult);
                Assert.AreEqual(DialogResult.OK, ((Button)dialog.AcceptButton).DialogResult);
                Assert.IsInstanceOfType(dialog.AcceptButton, typeof(UiActionButton));
                Assert.IsInstanceOfType(typeof(TestSupportReviewDialog).GetField("diff", BindingFlags.NonPublic | BindingFlags.Instance).GetValue(dialog), typeof(CodeDiffView));
                Assert.AreEqual(AutoScaleMode.Dpi, dialog.AutoScaleMode);
            }
        }
    }
}

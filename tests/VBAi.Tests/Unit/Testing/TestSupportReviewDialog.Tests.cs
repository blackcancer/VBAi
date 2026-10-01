using System.Reflection;
using System.Windows.Forms;
using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace VBAi.Tests.Unit
{
    [TestClass, TestCategory("Unit"), DoNotParallelize]
    public sealed class TestSupportReviewDialogTests
    {
        [STATestMethod]
        public void ReviewShowsExactProjectAndDiffAndReturnsOnlyExplicitApproval()
        {
            using (var owner = new Form())
            foreach (bool nullSource in new[] { false, true })
            foreach (var decision in new[] { DialogResult.OK, DialogResult.Cancel })
            {
                bool seen = false;
                bool accepted = TestSupportReviewDialog.Confirm(owner, "Disposable project", nullSource ? null : "Before", nullSource ? null : "After", (dialog, suppliedOwner) =>
                {
                    Assert.AreSame(owner, suppliedOwner);
                    StringAssert.Contains(((Label)typeof(TestSupportReviewDialog).GetField("projectLabel", BindingFlags.Instance | BindingFlags.NonPublic).GetValue(dialog)).Text, "Disposable project");
                    var diff = (CodeDiffView)typeof(TestSupportReviewDialog).GetField("diff", BindingFlags.Instance | BindingFlags.NonPublic).GetValue(dialog);
                    Assert.AreEqual(nullSource ? "" : "Before", typeof(CodeDiffView).GetField("before", BindingFlags.Instance | BindingFlags.NonPublic).GetValue(diff));
                    Assert.AreEqual(nullSource ? "" : "After", typeof(CodeDiffView).GetField("after", BindingFlags.Instance | BindingFlags.NonPublic).GetValue(diff));
                    seen = true;
                    return decision;
                });
                Assert.IsTrue(seen);
                Assert.AreEqual(decision == DialogResult.OK, accepted);
            }
        }

        [STATestMethod]
        public void NativeModalReviewRequiresAnExplicitDialogResult()
        {
            using (var timer = new Timer { Interval = 50 })
            {
                timer.Tick += (_, __) =>
                {
                    foreach (Form form in Application.OpenForms)
                    {
                        if (!(form is TestSupportReviewDialog)) continue;
                        timer.Stop(); form.DialogResult = DialogResult.Cancel; form.Close(); break;
                    }
                };
                timer.Start();
                Assert.IsFalse(TestSupportReviewDialog.Confirm(null, "Disposable", "Before", "After"));
            }
        }

    }
}

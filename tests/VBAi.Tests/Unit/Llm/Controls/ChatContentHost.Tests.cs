using Microsoft.VisualStudio.TestTools.UnitTesting;
using System;
using System.ComponentModel;
using System.Runtime.InteropServices;
using VBAi.Tests.Infrastructure;
namespace VBAi.Tests.Unit
{
    [TestClass, TestCategory("Unit")]
    public sealed class ChatContentHostTests
    {
        [STATestMethod]
        public void GotFocusToleratesDetachedSiteWithoutRepeatingNotification()
        {
            using (var host = new ChatContentHost())
            {
                int attempts = 0;
                host.GotFocus += (sender, args) => { attempts++; throw new InvalidComObjectException(); };
                LlmBoundaryScope.Call(host, "OnGotFocus", EventArgs.Empty);
                Assert.AreEqual(1, attempts);
                Assert.IsFalse(host.IsDisposed);
            }
        }

        [STATestMethod]
        public void GotFocusForwardsNormalNotificationOnce()
        {
            using (var host = new ChatContentHost())
            {
                int attempts = 0;
                host.GotFocus += (sender, args) => attempts++;
                LlmBoundaryScope.Call(host, "OnGotFocus", EventArgs.Empty);
                Assert.AreEqual(1, attempts);
            }
        }

        [STATestMethod]
        public void GotFocusPropagatesUnrelatedErrorsWithoutRepeatingNotification()
        {
            using (var host = new ChatContentHost())
            {
                int attempts = 0;
                var failure = new InvalidOperationException("Unrelated focus failure");
                host.GotFocus += (sender, args) => { attempts++; throw failure; };
                var wrapper = Assert.ThrowsException<System.Reflection.TargetInvocationException>(
                    () => LlmBoundaryScope.Call(host, "OnGotFocus", EventArgs.Empty));
                Assert.AreSame(failure, wrapper.InnerException);
                Assert.AreEqual(1, attempts);
            }
        }

        [STATestMethod]
        public void ContentHostUsesStandardWinFormsDesignerAndStartsEmpty()
        {
            using (var host = new ChatContentHost())
            {
                Assert.IsNull(host.Child);
                var attribute = (DesignerAttribute)TypeDescriptor.GetAttributes(host)[typeof(DesignerAttribute)];
                StringAssert.StartsWith(attribute.DesignerTypeName, "System.Windows.Forms.Design.ControlDesigner,");
            }
        }
    }
}

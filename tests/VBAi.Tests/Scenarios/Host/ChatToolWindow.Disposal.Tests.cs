using System.ComponentModel;
using System.Reflection;
using VBAi;
using VBAi.Tests.Infrastructure;
using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace VBAi.Tests.Unit
{
    [TestClass]
    [TestCategory("Unit")]
    public sealed class ChatToolWindowDisposalTests
    {
        [STATestMethod]
        public void ToolWindowCanDisposeWithNoComponentContainer()
        {
            using (var window = new ChatToolWindow())
            {
                var field = typeof(ChatToolWindow).GetField("components", BindingFlags.Instance | BindingFlags.NonPublic);
                var original = (IContainer)field.GetValue(window);
                try { field.SetValue(window, null); LlmBoundaryScope.Call(window, "Dispose", true); }
                finally { original.Dispose(); }
                Assert.IsTrue(window.IsDisposed);
            }
        }
    }
}

using System.ComponentModel;
using System.Reflection;
using CodexVBE;
using CodexVBE.Tests.Infrastructure;
using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace CodexVBE.Tests.Unit
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

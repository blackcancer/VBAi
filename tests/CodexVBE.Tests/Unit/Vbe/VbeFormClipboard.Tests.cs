using System;
using CodexVBE;
using Microsoft.VisualStudio.TestTools.UnitTesting;
namespace CodexVBE.Tests.Unit
{
    [TestClass, TestCategory("Unit")]
    public sealed class VbeFormClipboardTests
    {
        [TestMethod]
        public void ClipboardAndSelectionRevisionsAreBothRequired()
        {
            var request = new Request { ExpectedDesignerSelectionVersion = "tree-selection", ExpectedClipboardVersion = "123" };
            VbeForms.RequireClipboardRevision(request, "tree-selection", "123");
            Assert.ThrowsException<InvalidOperationException>(() => VbeForms.RequireClipboardRevision(request, "changed", "123"));
            Assert.ThrowsException<InvalidOperationException>(() => VbeForms.RequireClipboardRevision(request, "tree-selection", "124"));
            request.ExpectedClipboardVersion = "0";
            Assert.ThrowsException<InvalidOperationException>(() => VbeForms.RequireClipboardRevision(request, "tree-selection", "0"));
            request.ExpectedDesignerSelectionVersion = null;
            Assert.ThrowsException<InvalidOperationException>(() => VbeForms.RequireClipboardRevision(request, "tree-selection", "123"));
        }
    }
}

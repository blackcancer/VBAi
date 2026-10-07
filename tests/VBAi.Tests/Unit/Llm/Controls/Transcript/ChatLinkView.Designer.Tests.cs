using Microsoft.VisualStudio.TestTools.UnitTesting;
using VBAi.Tests.Infrastructure;
namespace VBAi.Tests.Unit
{
    [TestClass, TestCategory("Unit")]
    public sealed class ChatLinkViewDesignerContractTests
    {
        [STATestMethod]
        public void DesignerLayoutAndManagedDisposalAreComplete() { TranscriptFixture.LayoutAndLifecycle<ChatLinkView>(); }
    }
}
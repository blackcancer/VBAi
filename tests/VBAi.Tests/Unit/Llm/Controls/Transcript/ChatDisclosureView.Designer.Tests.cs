using System;
using System.ComponentModel;
using System.Drawing;
using System.Windows.Forms;
using VBAi;
using VBAi.Tests.Infrastructure;
using Microsoft.VisualStudio.TestTools.UnitTesting;
namespace VBAi.Tests.Unit
{
    [TestClass, TestCategory("Unit")]
    public sealed class ChatDisclosureViewDesignerContractTests
    {
        [STATestMethod]
        public void DesignerLayoutAndManagedDisposalAreComplete() { TranscriptFixture.LayoutAndLifecycle<ChatDisclosureView>(); }
    }
}
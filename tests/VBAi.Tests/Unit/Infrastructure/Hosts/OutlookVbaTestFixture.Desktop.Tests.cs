using Microsoft.VisualStudio.TestTools.UnitTesting;
using System;
using VBAi.Tests.Integration;

namespace VBAi.Tests.Unit
{
    [TestClass, TestCategory("Unit")]
    public sealed class OutlookPrivateDesktopTests
    {
        [TestMethod]
        public void MissingOrForeignPrivateDesktopIsRefusedBeforeExecutableOrActivation()
        {
            Assert.ThrowsException<InvalidOperationException>(() => OutlookVbaTestFixture.RequirePrivateOutlookExecutable(null, null, null));
            Assert.ThrowsException<InvalidOperationException>(() => OutlookVbaTestFixture.RequirePrivateOutlookExecutable(null, "owned", "other"));
        }
    }
}

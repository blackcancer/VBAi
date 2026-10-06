using System;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using VBAi.Tests.Integration;

namespace VBAi.Tests.Unit
{
    /// <summary>Real-desktop mode cannot inherit a private scope or run without its explicit scenario.</summary>
    [TestClass]
    public sealed class OllamaOfficeDesktopTests
    {
        [DataTestMethod]
        [DataRow(null, null, null)]
        [DataRow("0", null, null)]
        [DataRow("1", "Default", null)]
        [DataRow("1", null, "Default")]
        [DataRow("1", "VBAiTests_private", "VBAiTests_private")]
        public void ConflictingOrMissingOptInsRefuseBeforeDesktopAccess(string scenario, string required, string configured)
            => Assert.ThrowsException<InvalidOperationException>(() => OllamaOfficeDesktop.RequireMainOptIn(scenario, required, configured));

        [TestMethod]
        public void ExplicitMainScenarioHasNoPrivateDescriptor()
            => OllamaOfficeDesktop.RequireMainOptIn("1", null, null);
    }
}

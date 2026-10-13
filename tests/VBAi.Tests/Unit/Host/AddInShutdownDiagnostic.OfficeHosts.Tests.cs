using Microsoft.VisualStudio.TestTools.UnitTesting;
using System;

namespace VBAi.Tests.Unit
{
    /// <summary>Checks the expanded host-name boundary without enabling runtime observation.</summary>
    [TestClass, TestCategory("Unit")]
    public sealed class AddInShutdownDiagnosticOfficeHostsTests
    {
        [DataTestMethod, DataRow("POWERPNT"), DataRow("MSACCESS"), DataRow("MSPUB"), DataRow("OUTLOOK")]
        public void SupportedOfficeHostNamesAreAcceptedWithoutGuessingAliases(string name)
        {
            AddInShutdownDiagnostic.RequireObservedHostProcessName(name);
            AddInShutdownDiagnostic.RequireObservedHostProcessName(name.ToLowerInvariant());
            foreach (string invalid in new[] { name + ".EXE", name + " ", " " + name, name + "_TEST", @"C:\Office\" + name })
                Assert.ThrowsException<InvalidOperationException>(() => AddInShutdownDiagnostic.RequireObservedHostProcessName(invalid));
        }
    }
}

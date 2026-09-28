using System;
using CodexVBE;
using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace CodexVBE.Tests.Unit
{
    public sealed partial class VbeProcedureMutationTests
    {
        [TestMethod]
        public void SymbolSearchValidatesEveryBoundFiltersModulesAndIsolatesFailedSourceReads()
        {
            var fixture = new Fixture("Sub Run()\r\nEnd Sub");
            foreach (var request in new[] { new Request { Offset = -1 }, new Request { Limit = -1 },
                new Request { Limit = 201 }, new Request { Query = new string('x', 201) } })
                Assert.ThrowsException<ArgumentException>(() => fixture.Navigation.ProjectSymbols(request));
            dynamic result = fixture.Navigation.ProjectSymbols(new Request { Project = "Projet", Module = "missing", Query = "Sub" });
            Assert.AreEqual(0, (int)result.Total);
            result = fixture.Navigation.ProjectSymbols(new Request { Project = "Projet", Module = fixture.Component.Name, Query = "Run", MatchCase = true });
            Assert.AreEqual(1, (int)result.Total);
            fixture.Component.CodeModule = null;
            result = fixture.Navigation.ProjectSymbols(new Request { Project = "Projet", Query = "" });
            Assert.AreEqual(1, (int)result.Errors.Count);
        }
    }
}

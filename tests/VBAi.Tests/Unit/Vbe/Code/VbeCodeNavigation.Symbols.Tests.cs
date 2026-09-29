using System;
using System.Linq;
using VBAi;
using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace VBAi.Tests.Unit
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

        [TestMethod]
        public void DeclarationSymbolsMapAccessorRangesAndKeepModuleDeclarationsUnbound()
        {
            string code = "Dim moduleValue As Long\r\nProperty Get Item() As Long\r\nDim getterValue As Long\r\nEnd Property\r\nProperty Let Item(ByVal input As Long)\r\nDim setterValue As Long\r\nEnd Property\r\nSub Run()\r\nDim localValue As Long\r\nEnd Sub";
            var fixture = new Fixture(code);
            dynamic result = fixture.Navigation.ProjectSymbols(new Request { Project = "Projet", Query = "Value" });
            Assert.AreEqual(0, (int)result.Errors.Count);
            var symbols = ((System.Collections.IEnumerable)result.Symbols).Cast<object>().Select(x => (dynamic)x).ToArray();
            Assert.AreEqual(4, symbols.Length);
            Assert.IsNull((object)symbols.Single(x => (string)x.Name == "moduleValue").ProcKind);
            Assert.AreEqual(3, (int)symbols.Single(x => (string)x.Name == "getterValue").ProcKind);
            Assert.AreEqual(1, (int)symbols.Single(x => (string)x.Name == "setterValue").ProcKind);
            Assert.AreEqual(0, (int)symbols.Single(x => (string)x.Name == "localValue").ProcKind);
            result = fixture.Navigation.ProjectSymbols(new Request { Project = "Projet", Query = "absent", WholeWord = true });
            Assert.AreEqual(0, (int)result.Total);
            var empty = new Fixture("");
            result = empty.Navigation.ProjectSymbols(new Request { Project = "Projet" });
            Assert.AreEqual(1, (int)result.Total);
        }
    }
}

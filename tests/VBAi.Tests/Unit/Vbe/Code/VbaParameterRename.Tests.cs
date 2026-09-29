using System;
using VBAi;
using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace VBAi.Tests.Unit
{
    /// <summary>Vérifie les déclarations de paramètres privés et leurs appelants nommés.</summary>
    [TestClass, TestCategory("Unit")]
    public sealed class VbaParameterRenameTests
    {
        /// <summary>Construit l'identité de la déclaration de paramètre inspectée.</summary>
        private static Request Request(string header = "Private Sub Run(ByVal value As Long)") =>
            new Request { Module = "Module1", Procedure = "Run", Query = "value", NewName = "amount",
                StartLine = 1, StartColumn = header.IndexOf("value", StringComparison.Ordinal) + 1 };

        /// <summary>Met à jour déclaration/usages et les appels liés sans toucher les arguments des autres méthodes.</summary>
        [TestMethod]
        public void PrivateParameterAndNamedCallersAreRenamedWithoutChangingOtherBindings()
        {
            const string header = "Private Sub Run(ByVal value As Long)";
            string before = header + "\nDebug.Print value, obj.value, \"value\" ' value\nCall Run(value:=value)\nEnd Sub\n" +
                "Sub Caller()\nCall Run(value:=7)\nRun value:=8\nModule1.Run value:=9\nCall Module1.Run(value:=10)\n" +
                "If True Then Run value:=11 Else Run value:=12\nobj.Run value:=13\nCall obj.Run(value:=14)\n" +
                "Run value:=Other(value:=15)\nIf True Then Run value:=16 Else Other value:=17\nEnd Sub";
            string after = VbaParameterRename.Transform(before, Request(), 1, 4);
            StringAssert.Contains(after, "Private Sub Run(ByVal amount As Long)");
            StringAssert.Contains(after, "Debug.Print amount, obj.value, \"value\" ' value");
            StringAssert.Contains(after, "Call Run(amount:=amount)");
            StringAssert.Contains(after, "Call Run(amount:=7)");
            StringAssert.Contains(after, "Run amount:=8");
            StringAssert.Contains(after, "Module1.Run amount:=9");
            StringAssert.Contains(after, "Call Module1.Run(amount:=10)");
            StringAssert.Contains(after, "If True Then Run amount:=11 Else Run amount:=12");
            StringAssert.Contains(after, "obj.Run value:=13");
            StringAssert.Contains(after, "Call obj.Run(value:=14)");
            StringAssert.Contains(after, "Run amount:=Other(value:=15)");
            StringAssert.Contains(after, "If True Then Run amount:=16 Else Other value:=17");
        }

        /// <summary>Couvre Function, récursion et profondeur des arguments imbriqués.</summary>
        [TestMethod]
        public void FunctionCallsNestedArgumentsAndPhysicalContinuationsKeepTheirBinding()
        {
            const string header = "Private Function Run(Optional ByVal value As Long = 3) As Long";
            string before = header + "\r\nRun = value\r\nEnd Function\r\nSub Caller()\r\n" +
                "Debug.Print Run(value:=Run(value:=1))\r\nDebug.Print Run(value:=Other(value:=2))\r\n" +
                "Debug.Print Run( _\r\nvalue:=3)\r\nDebug.Print obj.Run(value:=4), value\r\nOther Run, 1\r\nDebug.Print Run\r\nWith obj\r\n.Run value:=5\r\nEnd With\r\nEnd Sub";
            string after = VbaParameterRename.Transform(before, Request(header), 1, 3);
            StringAssert.Contains(after, "ByVal amount As Long = 3");
            StringAssert.Contains(after, "Run = amount\r\n");
            StringAssert.Contains(after, "Run(amount:=Run(amount:=1))");
            StringAssert.Contains(after, "Run(amount:=Other(value:=2))");
            StringAssert.Contains(after, "Run( _\r\namount:=3)");
            StringAssert.Contains(after, "obj.Run(value:=4), value");
            StringAssert.Contains(after, "Other Run, 1"); StringAssert.Contains(after, ".Run value:=5");
        }

        /// <summary>Refuse les signatures publiques, ambiguës, conditionnelles et les noms masqués.</summary>
        [TestMethod]
        public void PublicPropertiesShadowedNamesAndConditionalBindingsAreRefused()
        {
            const string basic = "Private Sub Run(ByVal value As Long)\nDebug.Print value\nEnd Sub";
            foreach (string source in new[] { basic.Replace("Private", "Public"), basic.Replace("Private ", ""),
                basic.Replace("Private Sub", "Private Property Get"), basic.Replace("Run(", "Other("),
                basic + "\nPrivate Sub Run()\nEnd Sub", basic + "\n#Const Flag = True", basic + "\nDim Run As Object",
                basic + "\nDim Module1 As Object", basic + "\nSub Caller()\nCall Project.Module1.Run(value:=1)\nEnd Sub" })
                Assert.ThrowsException<InvalidOperationException>(() => VbaParameterRename.Transform(source, Request(), 1, 3));
            var request = Request(); request.ProcKind = 1;
            Assert.ThrowsException<InvalidOperationException>(() => VbaParameterRename.Transform(basic, Request(), 42, 43));
            var shifted = Request(); shifted.StartLine = 2;
            StringAssert.Contains(VbaParameterRename.Transform("\n" + basic, shifted, 2, 4), "ByVal amount");
            Assert.ThrowsException<InvalidOperationException>(() => VbaParameterRename.Transform(basic, request, 1, 3));
            request = Request(); request.StartColumn++;
            Assert.ThrowsException<InvalidOperationException>(() => VbaParameterRename.Transform(basic, request, 1, 3));
            request = Request(); request.NewName = "Debug";
            Assert.ThrowsException<ArgumentException>(() => VbaParameterRename.Transform(basic, request, 1, 3));
            request = Request(); request.NewName = "Run";
            Assert.ThrowsException<InvalidOperationException>(() => VbaParameterRename.Transform(basic, request, 1, 3));
            foreach (string source in new[] { "", "Private\n", basic + "\n\n" })
            {
                if (source.StartsWith("Private Sub")) StringAssert.Contains(VbaParameterRename.Transform(source, Request(), 1, 3), "ByVal amount");
                else Assert.ThrowsException<InvalidOperationException>(() => VbaParameterRename.Transform(source, Request(), 1, 3));
            }
        }
    }
}

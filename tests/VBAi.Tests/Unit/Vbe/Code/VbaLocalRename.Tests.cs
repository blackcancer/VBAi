namespace VBAi.Tests.Unit
{
    using Microsoft.VisualStudio.TestTools.UnitTesting;
    using System;
    using VBAi;

    [TestClass, TestCategory("Unit")]
    public sealed class VbaLocalRenameTests
    {
        private static Request Rename() => new Request { Procedure = "Run", Query = "value", NewName = "amount", StartLine = 2, StartColumn = 5 };
        [TestMethod]
        public void LocalRenameKeepsMembersLabelsStringsCommentsAndOtherProcedures()
        {
            string source = "Sub Run()\nDim value As Long\nvalue = obj.value + value\n" +
                "Debug.Print \"value\", value: Debug.Print value\nCall Other(value:=value)\n" +
                "GoTo value\nvalue: Debug.Print value\n' value\nEnd Sub\nSub Other()\nDim value As Long\nEnd Sub";
            string actual = VbaLocalRename.Transform(source, Rename(), 1, 12);
            Assert.AreEqual("Sub Run()\nDim amount As Long\namount = obj.value + amount\n" +
                "Debug.Print \"value\", amount: Debug.Print amount\nCall Other(value:=amount)\n" +
                "GoTo value\nvalue: Debug.Print amount\n' value\nEnd Sub\nSub Other()\nDim value As Long\nEnd Sub", actual);
        }
        [TestMethod]
        public void LocalRenameRefusesCollisionConditionalParametersAndStaleDeclarationCoordinates()
        {
            string basic = "Sub Run()\nDim value As Long\nDebug.Print value\nEnd Sub";
            var request = Rename(); request.NewName = "Debug";
            Assert.ThrowsException<ArgumentException>(() => VbaLocalRename.Transform(basic, request, 1, 4));
            request.NewName = "amount";
            Assert.ThrowsException<InvalidOperationException>(() => VbaLocalRename.Transform(basic.Replace("Debug.Print value", "Debug.Print amount, value"), request, 1, 4));
            Assert.ThrowsException<InvalidOperationException>(() => VbaLocalRename.Transform(basic.Replace("Dim value As Long", "#If VBA7 Then\nDim value As Long\n#End If"), request, 1, 6));
            request.StartColumn = 6;
            Assert.ThrowsException<InvalidOperationException>(() => VbaLocalRename.Transform(basic, request, 1, 4));
            request.StartLine = 1; request.StartColumn = 16;
            Assert.ThrowsException<InvalidOperationException>(() => VbaLocalRename.Transform("Sub Run(ByVal value As Long)\nEnd Sub", request, 1, 2));
        }
        [TestMethod]
        public void TypeContextsBangMembersAndTypeSuffixesKeepTheirOriginalBinding()
        {
            string source = "Sub Run()\nDim value As value\nSet value = New value\nDebug.Print record!value, value\nEnd Sub";
            string actual = VbaLocalRename.Transform(source, Rename(), 1, 5);
            StringAssert.Contains(actual, "Dim amount As value");
            StringAssert.Contains(actual, "Set amount = New value");
            StringAssert.Contains(actual, "record!value, amount");
            source = "Sub Run()\nDim value$\nvalue$ = \"text\"\nEnd Sub";
            actual = VbaLocalRename.Transform(source, Rename(), 1, 4);
            StringAssert.Contains(actual, "Dim amount$"); StringAssert.Contains(actual, "amount$ =");
        }

        [TestMethod]
        public void ProcedureBoundsHeadersScopeAndCompilationGuardsAreCheckedBeforeAnyReplacement()
        {
            const string source = "Sub Run()\nDim value As Long\nvalue = 1\nEnd Sub";
            foreach (var bounds in new[] { new[] { 0, 4, 2, 5 }, new[] { 3, 2, 2, 5 }, new[] { 3, 4, 2, 5 }, new[] { 1, 1, 2, 5 }, new[] { 1, 4, 2, 0 } })
            {
                var request = Rename(); request.StartLine = bounds[2]; request.StartColumn = bounds[3];
                Assert.ThrowsException<ArgumentException>(() => VbaLocalRename.Transform(source, request, bounds[0], bounds[1]));
            }
            foreach (string query in new[] { null, "different" })
            { var request = Rename(); request.Query = query; Assert.ThrowsException<InvalidOperationException>(() => VbaLocalRename.Transform(source, request, 1, 4)); }
            foreach (string scope in new[] { null, "Other" })
            { var request = Rename(); request.Procedure = scope; Assert.ThrowsException<InvalidOperationException>(() => VbaLocalRename.Transform(source, request, 1, 4)); }
            foreach (string changed in new[] {
                source.Replace("Sub Run()", "Sub Run()\n"), source.Replace("Sub Run()", "\nSub Run()"),
                source.Replace("End Sub", "End Function"), source.Replace("End Sub", ""),
                source.Replace("value = 1", "Dim value As String"),
                source.Replace("value = 1", "#Const Feature = True\nvalue = 1"),
                "DefInt A-Z\n" + source })
            {
                var request = Rename();
                if (changed.StartsWith("DefInt")) { request.StartLine = 3; Assert.ThrowsException<InvalidOperationException>(() => VbaLocalRename.Transform(changed, request, 2, 5)); }
                else if (changed.StartsWith("\n")) { request.StartLine = 3; Assert.ThrowsException<InvalidOperationException>(() => VbaLocalRename.Transform(changed, request, 1, 5)); }
                else Assert.ThrowsException<InvalidOperationException>(() => VbaLocalRename.Transform(changed, request, 1, changed.Split('\n').Length));
            }
            var conditional = Rename(); conditional.StartLine = 3;
            Assert.ThrowsException<InvalidOperationException>(() => VbaLocalRename.Transform("Sub Run()\n#If VBA7 Then\nDim value As Long\n#End If\nEnd Sub", conditional, 1, 5));
            var module = Rename(); module.Procedure = "Module"; module.StartLine = 1;
            Assert.ThrowsException<InvalidOperationException>(() => VbaLocalRename.Transform("Dim value As Long\nEnd Sub", module, 1, 2));
        }

        [TestMethod]
        public void FunctionsPropertiesConstantsCaseChangesAndEveryExcludedBindingArePreserved()
        {
            foreach (var signature in new[] { new[] { "Function Run() As Long", "End Function" }, new[] { "Property Get Run() As Long", "End Property" }, new[] { "Sub Run()", "End Sub" } })
            {
                string source = signature[0] + "\nConst value As Long = 2\n" +
                    "Debug.Print value \t, value\nCall Other(value:=value)\n" +
                    "GoSub value\nResume value\nCall Register(AddressOf value)\n" +
                    "If TypeOf obj Is value Then Debug.Print value\nIf obj Is value Then Debug.Print value\n" +
                    "Debug.Print obj!value, obj ! value, obj.value\nvalue :\n" + signature[1] + "\nDim outside As Long";
                var request = Rename(); request.StartColumn = 7;
                string actual = VbaLocalRename.Transform(source, request, 1, source.Split('\n').Length);
                StringAssert.Contains(actual, "Const amount"); StringAssert.Contains(actual, "Other(value:=amount)");
                StringAssert.Contains(actual, "GoSub value"); StringAssert.Contains(actual, "Resume value");
                StringAssert.Contains(actual, "AddressOf value"); StringAssert.Contains(actual, "TypeOf obj Is value Then Debug.Print amount");
                StringAssert.Contains(actual, "If obj Is amount Then Debug.Print amount");
                StringAssert.Contains(actual, "obj!value, obj ! value, obj.value"); StringAssert.Contains(actual, "value :");
                request.NewName = "VALUE";
                actual = VbaLocalRename.Transform(source, request, 1, source.Split('\n').Length);
                StringAssert.Contains(actual, "Const VALUE"); StringAssert.Contains(actual, "Other(value:=VALUE)");
            }
        }

        [TestMethod]
        public void PhysicalBoundariesAndMalformedNamedDeclarationsFailWithoutGuessingBinding()
        {
            string source = "Dim moduleValue As Long\nSub Run()\nDim value As Long\nDebug.Print value \t\nEnd Sub:value:";
            var request = Rename(); request.StartLine = 3;
            string result = VbaLocalRename.Transform(source, request, 2, 5);
            StringAssert.Contains(result, "Dim moduleValue"); StringAssert.Contains(result, "Dim amount");
            StringAssert.Contains(result, "End Sub:value:");
            source = source.TrimEnd(':');
            result = VbaLocalRename.Transform(source, request, 2, 5);
            StringAssert.Contains(result, "End Sub:amount");
            result = VbaLocalRename.Transform(source.Replace("End Sub:value", "End Sub:Debug.Print value:"), request, 2, 5);
            StringAssert.Contains(result, "End Sub:Debug.Print amount:");
            Assert.ThrowsException<InvalidOperationException>(() => VbaLocalRename.Transform("Sub Run()\nDim value:=0\nEnd Sub", Rename(), 1, 3));
        }
    }
}

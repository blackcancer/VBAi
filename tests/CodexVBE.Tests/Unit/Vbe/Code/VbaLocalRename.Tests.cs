namespace CodexVBE.Tests.Unit
{
    using System;
    using CodexVBE;
    using Microsoft.VisualStudio.TestTools.UnitTesting;

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
    }
}

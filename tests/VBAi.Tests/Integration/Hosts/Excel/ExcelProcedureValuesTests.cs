using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading;
using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace VBAi.Tests.Integration
{
    /// <summary>Scénarios Excel natifs préparés; exécution séparément opt-in, état initial NOT_RUN.</summary>
    [TestClass, TestCategory("Excel")]
    public sealed class ExcelProcedureValuesTests
    {
        /// <summary>Qualifie ParamArray vide, trente scalaires, Variant Null et un tableau transporté comme une valeur.</summary>
        [STATestMethod]
        public void NativeParamArrayCallsPreserveArityValuesAndSingleInvocation()
        {
            ExcelVbeFixture.Run(host =>
            {
                host.EnableProcedureValuesTeardownTrace(nameof(NativeParamArrayCallsPreserveArityValuesAndSingleInvocation));
                string project = Convert.ToString(VbeBridgeClient.Object(((object[])host.Command("list_projects")["Data"]).Single())["Name"]);
                const string module = "ParamArrayProbe";
                var created = Data(host.Command(new { Command = "create_module", Project = project, Module = module, ExpectedMode = 2 }));
                var initial = Data(host.Command(new { Command = "read_module", Project = project, Module = module }));
                const string source = "Option Explicit\r\nPublic Function Values(ParamArray items() As Variant) As Variant\r\n" +
                    "ThisWorkbook.Worksheets(1).Range(\"A1\").Value2 = ThisWorkbook.Worksheets(1).Range(\"A1\").Value2 + 1\r\nValues = items\r\nEnd Function\r\n" +
                    "Public Function Tagged(ByVal tag As String, ParamArray items()) As Variant\r\n" +
                    "ThisWorkbook.Worksheets(1).Range(\"A1\").Value2 = ThisWorkbook.Worksheets(1).Range(\"A1\").Value2 + 1\r\n" +
                    "Tagged = tag & \"|\" & CStr(UBound(items) + 1)\r\n" +
                    "If UBound(items) >= 0 Then\r\nIf IsArray(items(0)) Then Tagged = Tagged & \"|array\"\r\nEnd If\r\nEnd Function";
                Data(host.Command(new { Command = "replace_lines", Project = project, Module = module, ExpectedSha256 = initial["Sha256"], StartLine = 1, Count = created["Lines"], Text = source }));
                var inspected = Data(host.Command(new { Command = "read_module", Project = project, Module = module }));
                var properties = Data(host.Command(new { Command = "project_properties", Project = project }));
                string path = host.File("paramarray.xlsm");
                Data(host.Command(new { Command = "save_host_document_as", Project = project, Path = path, ExpectedProjectVersion = properties["Version"] }));
                int calls = 0;
                foreach (object[] arguments in new[] { new object[0], new object[] { "quoted\"", 7, true, null }, Enumerable.Range(1, 30).Cast<object>().ToArray() })
                {
                    var queued = Data(host.Command(new { Command = "run_procedure_values", Project = path, Module = module, Procedure = "Values",
                        ExpectedHostPath = path, ExpectedSha256 = inspected["Sha256"], ExpectedMode = 2, Arguments = arguments }));
                    var status = Wait(host, path, queued); var returned = VbeBridgeClient.Object(status["Output"]);
                    Assert.AreEqual("Array", returned["Kind"]); CollectionAssert.AreEqual(arguments, (object[])returned["Value"]);
                    Assert.AreEqual(++calls, Convert.ToInt32(host.ReadCell("A1")));
                    Data(host.Command(new { Command = "procedure_values_status", Project = path, Query = queued["Query"] }));
                    Assert.AreEqual(calls, Convert.ToInt32(host.ReadCell("A1")), "Polling must not repeat the native invocation.");
                }
                foreach (object[] arguments in new[] { new object[] { "native" }, new object[] { "native", new object[] { 3, 4 } } })
                {
                    var queued = Data(host.Command(new { Command = "run_procedure_values", Project = path, Module = module, Procedure = "Tagged",
                        ExpectedHostPath = path, ExpectedSha256 = inspected["Sha256"], ExpectedMode = 2, Arguments = arguments }));
                    var returned = VbeBridgeClient.Object(Wait(host, path, queued)["Output"]);
                    Assert.AreEqual(arguments.Length == 1 ? "native|0" : "native|1|array", returned["Value"]);
                    Assert.AreEqual(++calls, Convert.ToInt32(host.ReadCell("A1")));
                }
                var refused = host.Command(new { Command = "run_procedure_values", Project = path, Module = module, Procedure = "Tagged",
                    ExpectedHostPath = path, ExpectedSha256 = inspected["Sha256"], ExpectedMode = 2,
                    Arguments = new object[] { "native" }, ArgumentNames = new[] { "tag" } });
                Assert.AreEqual(false, refused["Ok"]); Assert.AreEqual(calls, Convert.ToInt32(host.ReadCell("A1")));
                Assert.AreEqual(inspected["Code"], Data(host.Command(new { Command = "read_module", Project = path, Module = module }))["Code"]);
            });
        }
        /// <summary>Qualifie SAFEARRAY vector/matrix, bornes retournées, Null et unique effet de bord par invocation.</summary>
        [STATestMethod]
        public void NativeVariantArraysRoundTripWithBoundsAndOneInvocation()
        {
            ExcelVbeFixture.Run(host =>
            {
                host.EnableProcedureValuesTeardownTrace(nameof(NativeVariantArraysRoundTripWithBoundsAndOneInvocation));
                string project = Convert.ToString(VbeBridgeClient.Object(((object[])host.Command("list_projects")["Data"]).Single())["Name"]);
                const string module = "ArrayProbe";
                var created = Data(host.Command(new { Command = "create_module", Project = project, Module = module, ExpectedMode = 2 }));
                var initial = Data(host.Command(new { Command = "read_module", Project = project, Module = module }));
                const string source = "Option Explicit\r\nPublic Function Vector(ByVal values As Variant, Optional ByVal text As String = \"default\") As Variant\r\n" +
                    "ThisWorkbook.Worksheets(1).Range(\"A1\").Value2 = ThisWorkbook.Worksheets(1).Range(\"A1\").Value2 + 1\r\n" +
                    "Dim result(1 To 2, -2 To -1) As Variant\r\nresult(1, -2) = values(0)\r\nresult(1, -1) = values(1)\r\nresult(2, -2) = values(2)\r\nresult(2, -1) = text\r\nVector = result\r\nEnd Function\r\n" +
                    "Public Function Matrix(ByVal values As Variant) As Variant\r\n" +
                    "ThisWorkbook.Worksheets(1).Range(\"A1\").Value2 = ThisWorkbook.Worksheets(1).Range(\"A1\").Value2 + 1\r\nMatrix = values\r\nEnd Function";
                Data(host.Command(new { Command = "replace_lines", Project = project, Module = module, ExpectedSha256 = initial["Sha256"], StartLine = 1, Count = created["Lines"], Text = source }));
                var inspected = Data(host.Command(new { Command = "read_module", Project = project, Module = module }));
                var properties = Data(host.Command(new { Command = "project_properties", Project = project }));
                string path = host.File("variant-arrays.xlsm");
                Data(host.Command(new { Command = "save_host_document_as", Project = project, Path = path, ExpectedProjectVersion = properties["Version"] }));
                var queued = Data(host.Command(new { Command = "run_procedure_values", Project = path, Module = module, Procedure = "Vector",
                    ExpectedHostPath = path, ExpectedSha256 = inspected["Sha256"], ExpectedMode = 2,
                    Arguments = new object[] { new object[] { 7, "quoted\" value", null } } }));
                var first = Wait(host, path, queued); var returned = VbeBridgeClient.Object(first["Output"]);
                Assert.AreEqual("Returned", first["State"], Convert.ToString(first["Error"])); Assert.AreEqual(true, first["ReturnValueVerified"]);
                CollectionAssert.AreEqual(new object[] { 1, -2 }, (object[])returned["LowerBounds"]);
                var rows = (object[])returned["Value"]; Assert.AreEqual(7, ((object[])rows[0])[0]); Assert.IsNull(((object[])rows[1])[0]);
                Assert.AreEqual("default", ((object[])rows[1])[1]); Assert.AreEqual(1d, Convert.ToDouble(host.ReadCell("A1")));
                Data(host.Command(new { Command = "procedure_values_status", Project = path, Query = queued["Query"] }));
                Assert.AreEqual(1d, Convert.ToDouble(host.ReadCell("A1")), "Polling must not invoke the function again.");
                queued = Data(host.Command(new { Command = "run_procedure_values", Project = path, Module = module, Procedure = "Matrix",
                    ExpectedHostPath = path, ExpectedSha256 = inspected["Sha256"], ExpectedMode = 2,
                    Arguments = new object[] { new object[] { new object[] { 1, 2 }, new object[] { 3, 4 } } } }));
                var matrix = VbeBridgeClient.Object(Wait(host, path, queued)["Output"]);
                CollectionAssert.AreEqual(new object[] { 0, 0 }, (object[])matrix["LowerBounds"]);
                rows = (object[])matrix["Value"]; Assert.AreEqual(4, ((object[])rows[1])[1]);
                Assert.AreEqual(2d, Convert.ToDouble(host.ReadCell("A1")));
                Assert.AreEqual(inspected["Code"], Data(host.Command(new { Command = "read_module", Project = path, Module = module }))["Code"]);
            });
        }
        /// <summary>Attend uniquement le statut jusqu'à la borne fixée, sans relancer la macro.</summary>
        private static IDictionary<string, object> Wait(ExcelVbeFixture host, string project, IDictionary<string, object> operation)
        {
            DateTime deadline = DateTime.UtcNow.AddSeconds(15);
            while (Convert.ToBoolean(operation["Pending"]) && DateTime.UtcNow < deadline)
            { Thread.Sleep(100); operation = Data(host.Command(new { Command = "procedure_values_status", Project = project, Query = operation["Query"] })); }
            Assert.AreEqual("Returned", operation["State"], Convert.ToString(operation["Error"])); return operation;
        }
        /// <summary>Exige une réponse de pont valide avant inspection des données natives.</summary>
        private static IDictionary<string, object> Data(IDictionary<string, object> response)
        { Assert.AreEqual(true, response["Ok"], Convert.ToString(response["Error"])); return VbeBridgeClient.Object(response["Data"]); }
    }
}

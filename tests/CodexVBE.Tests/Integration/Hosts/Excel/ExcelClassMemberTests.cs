using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading;
using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace CodexVBE.Tests.Integration
{
    /// <summary>Renommages privés qualifiés sur des classes VBA ordinaires d'un classeur possédé.</summary>
    [TestClass, TestCategory("Excel")]
    public sealed class ExcelClassMemberTests
    {
        /// <summary>Vérifie Function, familles Get/Let et Get/Set, compilation, résultats et annulation exacte.</summary>
        [STATestMethod]
        public void PrivateClassMembersAndAccessorFamiliesRetainNativeBehaviorAndUndo()
        {
            using (var host = ExcelVbeFixture.Start())
            {
                string project = Convert.ToString(VbeBridgeClient.Object(((object[])host.Command("list_projects")["Data"]).Single())["Name"]);
                const string module = "Calculator";
                const string source = "Option Explicit\r\nPrivate stored As Long\r\nPrivate storedObject As Object\r\n" +
                    "Private Function Compute(ByVal number As Long) As Long\r\nCompute = number * 2\r\nEnd Function\r\n" +
                    "Public Function Calculate(ByVal number As Long) As Long\r\nCalculate = Compute(number) + Compute(1)\r\nEnd Function\r\n" +
                    "Private Property Get Value() As Long\r\nValue = stored\r\nEnd Property\r\n" +
                    "Private Property Let Value(ByVal rhs As Long)\r\nstored = rhs\r\nEnd Property\r\n" +
                    "Public Function VerifyValue() As Long\r\nValue = 21\r\nVerifyValue = Value\r\nEnd Function\r\n" +
                    "Private Property Get Reference() As Object\r\nSet Reference = storedObject\r\nEnd Property\r\n" +
                    "Private Property Set Reference(ByVal rhs As Object)\r\nSet storedObject = rhs\r\nEnd Property\r\n" +
                    "Public Function VerifyObject() As Long\r\nSet Reference = ThisWorkbook.Worksheets(1).Range(\"B1\")\r\n" +
                    "Reference.Value2 = 6\r\nVerifyObject = Reference.Value2\r\nEnd Function";
                var created = Data(host.Command(new { Command = "create_class", Project = project, Module = module, ExpectedMode = 2 }));
                var initial = Data(host.Command(new { Command = "read_module", Project = project, Module = module }));
                Data(host.Command(new { Command = "replace_lines", Project = project, Module = module, ExpectedSha256 = initial["Sha256"], StartLine = 1, Count = created["Lines"], Text = source }));
                var original = Data(host.Command(new { Command = "read_module", Project = project, Module = module }));
                const string caller = "Option Explicit\r\nPublic Sub Main()\r\nDim instance As Calculator\r\nSet instance = New Calculator\r\n" +
                    "ThisWorkbook.Worksheets(1).Range(\"A1\").Value2 = instance.Calculate(20)\r\n" +
                    "ThisWorkbook.Worksheets(1).Range(\"A2\").Value2 = instance.VerifyValue()\r\n" +
                    "ThisWorkbook.Worksheets(1).Range(\"A3\").Value2 = instance.VerifyObject()\r\n" +
                    "ThisWorkbook.Worksheets(1).Range(\"C1\").Value2 = ThisWorkbook.Worksheets(1).Range(\"C1\").Value2 + 1\r\nEnd Sub";
                created = Data(host.Command(new { Command = "create_module", Project = project, Module = "ClassCaller", ExpectedMode = 2 }));
                initial = Data(host.Command(new { Command = "read_module", Project = project, Module = "ClassCaller" }));
                Data(host.Command(new { Command = "replace_lines", Project = project, Module = "ClassCaller", ExpectedSha256 = initial["Sha256"], StartLine = 1, Count = created["Lines"], Text = caller }));
                VerifyNativeResult(host, project);
                foreach (var rename in new[] { new { Old = "Compute", Name = "DoubleNumber", Kind = 0 },
                    new { Old = "Value", Name = "StoredValue", Kind = 3 }, new { Old = "Reference", Name = "StoredReference", Kind = 3 } })
                {
                    var inspected = Data(host.Command(new { Command = "read_module", Project = project, Module = module }));
                    string text = (string)inspected["Code"];
                    var lines = text.Replace("\r\n", "\n").Split('\n');
                    string prefix = rename.Kind == 0 ? "Private Function " : "Private Property Get ";
                    int line = Array.FindIndex(lines, value => value.StartsWith(prefix + rename.Old + "(", StringComparison.Ordinal));
                    Assert.IsTrue(line >= 0);
                    var preview = Data(host.Command(new { Command = "preview_class_member_rename", Project = project, Module = module,
                        Query = rename.Old, NewName = rename.Name, ProcKind = rename.Kind,
                        StartLine = line + 1, StartColumn = prefix.Length + 1, ExpectedSha256 = inspected["Sha256"] }));
                    var applied = Data(host.Command(new { Command = "apply_class_member_rename", Project = project, Module = module,
                        Query = rename.Old, NewName = rename.Name, ProcKind = rename.Kind,
                        StartLine = line + 1, StartColumn = prefix.Length + 1, ExpectedSha256 = inspected["Sha256"],
                        ExpectedProjectVersion = preview["ExpectedProjectVersion"], ExpectedMode = 2 }));
                    Assert.AreEqual(true, applied["ReadbackVerified"]);
                    VerifyNativeResult(host, project);
                }
                for (int undo = 0; undo < 3; undo++)
                {
                    var current = Data(host.Command(new { Command = "read_module", Project = project, Module = module }));
                    Data(host.Command(new { Command = "undo_code_edit", Project = project, Module = module, ExpectedSha256 = current["Sha256"] }));
                }
                Assert.AreEqual(original["Code"], Data(host.Command(new { Command = "read_module", Project = project, Module = module }))["Code"]);
                VerifyNativeResult(host, project);
            }
        }

        /// <summary>Compile puis exécute une macro publique et relit les cellules indépendamment du pont.</summary>
        private static void VerifyNativeResult(ExcelVbeFixture host, string project)
        {
            host.FocusCodeModule("ClassCaller");
            var compiled = Data(host.Command(new { Command = "compile_project", Project = project, ExpectedMode = 2 }));
            Assert.AreEqual(true, compiled["Compiled"], Convert.ToString(compiled["Diagnostic"]));
            var inspected = Data(host.Command(new { Command = "read_module", Project = project, Module = "ClassCaller" }));
            int expectedCalls = Convert.ToInt32(host.ReadCell("C1")) + 1;
            Data(host.Command(new { Command = "run_sub", Project = project, Module = "ClassCaller", Procedure = "Main",
                ExpectedMode = 2, ExpectedSha256 = inspected["Sha256"] }));
            DateTime deadline = DateTime.UtcNow.AddSeconds(10);
            while (Convert.ToInt32(host.ReadCell("C1")) < expectedCalls && DateTime.UtcNow < deadline)
            {
                Data(host.Command(new { Command = "debug_state", Project = project }));
                Thread.Sleep(50);
            }
            Assert.AreEqual(expectedCalls, Convert.ToInt32(host.ReadCell("C1")), "The renamed class must execute once; prior cell values are insufficient.");
            Assert.AreEqual(42d, Convert.ToDouble(host.ReadCell("A1")));
            Assert.AreEqual(21d, Convert.ToDouble(host.ReadCell("A2")));
            Assert.AreEqual(6d, Convert.ToDouble(host.ReadCell("A3")));
            Assert.AreEqual(6d, Convert.ToDouble(host.ReadCell("B1")));
        }

        /// <summary>Exige une réponse de pont réussie avant lecture des données.</summary>
        private static IDictionary<string, object> Data(IDictionary<string, object> response)
        { Assert.AreEqual(true, response["Ok"], Convert.ToString(response["Error"])); return VbeBridgeClient.Object(response["Data"]); }
    }
}

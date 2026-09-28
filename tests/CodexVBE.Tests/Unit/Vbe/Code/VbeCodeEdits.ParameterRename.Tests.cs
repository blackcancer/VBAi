using System;
using CodexVBE;
using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace CodexVBE.Tests.Unit
{
    public sealed partial class VbeCodeEditsTests
    {
        /// <summary>Exerce toutes les lectures, gardes et erreurs de l'édition puis l'historique géré.</summary>
        [TestMethod]
        public void PrivateParameterRenameChecksComponentCatalogModeAndReadbackBeforeHistoryReplay()
        {
            const string original = "Private Sub Run(ByVal value As Long)\nDebug.Print value\nEnd Sub\nSub Caller()\nRun value:=7\nEnd Sub";
            for (int scenario = 0; scenario < 15; scenario++)
            {
                int current = scenario, writes = 0;
                string code = original;
                VbeCodeEdits edits = null;
                var request = new Request { Project = "P", Module = "Module1", Procedure = "Run", Query = "value", NewName = "amount",
                    ProcKind = 0, StartLine = 1, StartColumn = original.IndexOf("value", StringComparison.Ordinal) + 1, ExpectedSha256 = Hash(code), ExpectedMode = 2 };
                edits = new VbeCodeEdits(command => {
                    switch (command.Command)
                    {
                        case "read_module": return Response.Success(new Snapshot { Code = code });
                        case "component_properties": return current == 0 ? Response.Failure("component failed") : Response.Success(new ComponentKind { Type = current == 1 ? 2 : 1 });
                        case "list_procedures":
                            if (current == 2) return Response.Failure("catalog failed");
                            var selected = new ProcedureRow { Name = "Run", Kind = 0, BodyLine = 1, EndLine = 3 };
                            return Response.Success(new ProcedureCatalog { Sha256 = current == 3 ? "stale" : Hash(code),
                                Procedures = current == 4 ? new ProcedureRow[0] : current == 5 ? new[] { selected, selected } :
                                new[] { new ProcedureRow { Name = "Other", Kind = 0 }, new ProcedureRow { Name = "Run", Kind = 1 }, selected } });
                        case "debug_state": return current == 7 ? Response.Failure("state failed") : Response.Success(new DesignState { Mode = current == 8 ? 1 : 2 });
                        case "replace_lines":
                            writes++;
                            if (current == 9) return Response.Failure("write failed");
                            string before = code;
                            code = current == 10 ? "unexpected readback" : command.Text;
                            edits.Record("P", "Module1", before, code);
                            return Response.Success("written");
                        default: throw new InvalidOperationException("unexpected command " + command.Command);
                    }
                });
                if (scenario == 6) request.ExpectedMode = 1;
                if (scenario == 11) request.ExpectedSha256 = "stale";
                if (scenario <= 11)
                {
                    if (scenario == 6) Assert.ThrowsException<ArgumentException>(() => edits.RenameParameter(request, false));
                    else Assert.ThrowsException<InvalidOperationException>(() => edits.RenameParameter(request, false));
                    Assert.AreEqual(scenario == 9 || scenario == 10 ? 1 : 0, writes);
                }
                else
                {
                    dynamic preview = edits.RenameParameter(request, true);
                    Assert.IsTrue((bool)preview.Changed); Assert.AreEqual(0, writes); Assert.AreEqual(original, code);
                    if (scenario == 12) continue;
                    edits.RenameParameter(request, false);
                    StringAssert.Contains(code, "ByVal amount As Long"); StringAssert.Contains(code, "Run amount:=7");
                    request.ExpectedSha256 = Hash(code); edits.Replay(request, false); Assert.AreEqual(original, code);
                    if (scenario == 14) { request.ExpectedSha256 = Hash(code); edits.Replay(request, true); StringAssert.Contains(code, "ByVal amount As Long"); }
                }
            }
        }
    }
}

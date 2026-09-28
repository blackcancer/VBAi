using System;
using System.Collections.Generic;
using System.Linq;
using System.Web.Script.Serialization;
using CodexVBE;
using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace CodexVBE.Tests.Unit
{
    public sealed partial class LlmVbeToolsBoundaryTests
    {
        /// <summary>Publie les différences de tous les modules, même si le renommage n'a été appliqué qu'en partie.</summary>
        [TestMethod]
        public void ProcedureRenameDiffsReflectSuccessPartialFailureAndUnreadableModules()
        {
            foreach (string failure in new[] { "none", "response", "exception", "readback" })
            foreach (bool subscribed in new[] { false, true })
            {
                var tools = new ToolFixture().Tools;
                var changes = new List<CodeChange>();
                if (subscribed) tools.CodeEdited += changes.Add;
                bool invoked = false;
                tools.Execute = request => {
                    if (request.Command == "preview_procedure_rename") return Response.Success(new {
                        ExpectedProjectVersion = "v", Edits = new[] {
                            new { Module = "A", Before = "old A", ExpectedSha256 = "a" },
                            new { Module = "B", Before = "old B", ExpectedSha256 = "b" } } });
                    if (request.Command == "apply_procedure_rename") {
                        invoked = true;
                        if (failure == "exception") throw new InvalidOperationException("partial write");
                        return failure == "response" ? Response.Failure("partial write") : Response.Success(new { Applied = true });
                    }
                    Assert.AreEqual("read_module", request.Command);
                    Assert.IsTrue(invoked);
                    if (failure == "readback" && request.Module == "A") return Response.Failure("unreadable A");
                    bool changed = request.Module == "A" || failure == "none" || failure == "readback";
                    return Response.Success(new LiveSourceResult { Code = (changed ? "new " : "old ") + request.Module,
                        Sha256 = changed ? "changed" : request.Module.ToLowerInvariant() });
                };
                var args = Arguments("apply_procedure_rename"); args["ExpectedProjectVersion"] = "v";
                var result = Json.Deserialize<Response>(tools.Invoke("apply_procedure_rename", Json.Serialize(args)));
                Assert.AreEqual(failure == "none" || failure == "readback", result.Ok);
                Assert.AreEqual(subscribed ? (failure == "none" ? 2 : 1) : 0, changes.Count);
                foreach (var change in changes) {
                    Assert.AreEqual("old " + change.Module, change.Before);
                    Assert.AreEqual("new " + change.Module, change.After);
                }
                if (subscribed && failure == "readback") Assert.AreEqual("B", changes.Single().Module);
            }
        }

        /// <summary>Une prévisualisation illisible ou périmée interdit toute écriture et notification.</summary>
        [TestMethod]
        public void ProcedureRenamePreflightRejectsMissingAndStalePlansBeforeWriting()
        {
            foreach (string fault in new[] { "failure", "missing", "stale" })
            {
                var tools = new ToolFixture().Tools; int writes = 0, changes = 0;
                tools.CodeEdited += _ => changes++;
                tools.Execute = request => {
                    if (request.Command != "preview_procedure_rename") { writes++; return Response.Success(new object()); }
                    if (fault == "failure") return Response.Failure("preview unavailable");
                    return Response.Success(fault == "missing" ? (object)new { Edits = new object[0] } : new { ExpectedProjectVersion = "stale", Edits = new object[0] });
                };
                var args = Arguments("apply_procedure_rename"); args["ExpectedProjectVersion"] = "v";
                Failed(tools.Invoke("apply_procedure_rename", Json.Serialize(args)), fault);
                Assert.AreEqual(0, writes); Assert.AreEqual(0, changes);
            }
        }
    }
}

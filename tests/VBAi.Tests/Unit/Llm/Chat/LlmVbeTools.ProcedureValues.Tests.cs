namespace VBAi.Tests.Unit
{
    using System;
    using System.Linq;
    using System.Windows.Forms;
    using VBAi;
    using Microsoft.VisualStudio.TestTools.UnitTesting;

    public sealed partial class LlmVbeToolsBoundaryTests
    {
        /// <summary>Le catalogue expose deux rangs bornés et conserve les préconditions de l'exécution.</summary>
        [TestMethod]
        public void ProcedureValuesCatalogueDescribesBoundedVectorsMatricesAndExecutionGuards()
        {
            var definition = LlmVbeTools.Definitions.Select(d => Dict(Dict(Json.DeserializeObject(Json.Serialize(d)))["function"]))
                .Single(d => (string)d["name"] == "run_procedure_values");
            var parameters = Dict(definition["parameters"]);
            var properties = Dict(parameters["properties"]);
            var arguments = Dict(properties["Arguments"]);
            Assert.AreEqual("array", arguments["type"]);
            Assert.AreEqual(30, arguments["maxItems"]);
            var choices = (object[])Dict(arguments["items"])["anyOf"];
            Assert.AreEqual(3, choices.Length);
            var vector = Dict(choices[1]); var matrix = Dict(choices[2]);
            Assert.AreEqual(1, vector["minItems"]); Assert.AreEqual(1024, vector["maxItems"]);
            Assert.AreEqual("array", vector["type"]); Assert.AreEqual("array", matrix["type"]);
            Assert.AreEqual("array", Dict(matrix["items"])["type"]);
            Assert.AreEqual(4, ((object[])Dict(vector["items"])["anyOf"]).Length);
            var required = ((object[])parameters["required"]).Cast<string>().ToArray();
            foreach (string guard in new[] { "Project", "Module", "Procedure", "ExpectedSha256", "ExpectedHostPath", "ExpectedMode", "Arguments" })
                CollectionAssert.Contains(required, guard);
        }

        /// <summary>Les vraies validations JSON refusent les tableaux non transportables avant Execute.</summary>
        [TestMethod]
        public void ProcedureValuesBoundaryRejectsMalformedAndOversizedValuesWithoutDispatch()
        {
            var fixture = new ToolFixture(); int calls = 0;
            fixture.Tools.Execute = request => { calls++; return Response.Success("unexpected"); };
            foreach (object rejected in RejectedProcedureValues())
            {
                var values = ProcedureValuesArguments(); values["Arguments"] = rejected;
                Failed(fixture.Tools.Invoke("run_procedure_values", Json.Serialize(values)), "invalid scalar/array values");
            }
            foreach (string missing in new[] { "Project", "Module", "Procedure", "ExpectedSha256", "ExpectedHostPath", "ExpectedMode", "Arguments" })
            {
                var values = ProcedureValuesArguments(); values.Remove(missing);
                Failed(fixture.Tools.Invoke("run_procedure_values", Json.Serialize(values)), "missing " + missing);
            }
            Assert.AreEqual(0, calls);
        }

        /// <summary>La requête garde les valeurs JSON, noms et préconditions pour la liaison VBA côté hôte.</summary>
        [TestMethod]
        public void ProcedureValuesBoundaryDeliversValuesAndNativeGuardsExactlyOnce()
        {
            var fixture = new ToolFixture(); int calls = 0; Request captured = null;
            fixture.Tools.Execute = request => { calls++; captured = request; return Response.Success(new { Query = "queued" }); };
            var values = ProcedureValuesArguments();
            Success(fixture.Tools.Invoke("run_procedure_values", Json.Serialize(values)), "vector and matrix");
            Assert.AreEqual(1, calls); Assert.AreEqual("run_procedure_values", captured.Command);
            Assert.AreEqual("P", captured.Project); Assert.AreEqual("Module1", captured.Module);
            Assert.AreEqual("Values", captured.Procedure); Assert.AreEqual(2, captured.ExpectedMode);
            Assert.AreEqual("source-revision", captured.ExpectedSha256);
            Assert.AreEqual(@"C:\Temp\fixture.xlsm", captured.ExpectedHostPath);
            CollectionAssert.AreEqual(new[] { "vector", "matrix" }, captured.ArgumentNames);
            var vector = (object[])captured.Arguments[0];
            Assert.AreEqual("quote\"", vector[0]); Assert.IsNull(vector[1]); Assert.AreEqual(true, vector[2]);
            Assert.AreEqual(4, ((object[])((object[])captured.Arguments[1])[1])[1]);
            values["Arguments"] = new object[] { 1 }; values["ArgumentNames"] = new[] { "first\nEnd" };
            Failed(fixture.Tools.Invoke("run_procedure_values", Json.Serialize(values)), "invalid parameter name");
            Assert.AreEqual(1, calls);
        }

        /// <summary>L'exécution respecte mode, portée et approbation; l'inspection reste sans exécution.</summary>
        [TestMethod]
        public void ProcedureValuesBoundaryEnforcesExecutionPermissionsAndAllowsReadOnlyStatus()
        {
            var fixture = new ToolFixture(); int runs = 0, reads = 0;
            fixture.Tools.Execute = request =>
            {
                if (request.Command == "procedure_values_status") reads++; else runs++;
                return Response.Success(new { InvocationInvoked = false });
            };
            var run = Json.Serialize(ProcedureValuesArguments());
            var status = Json.Serialize(new { Project = "P", Query = "queued" });
            fixture.Tools.Mode = ChatMode.Plan;
            Failed(fixture.Tools.Invoke("run_procedure_values", run), "plan execution");
            Success(fixture.Tools.Invoke("procedure_values_status", status), "plan inspection");
            fixture.Tools.Mode = ChatMode.Agent; fixture.Tools.BoundProject = "OtherProject";
            Failed(fixture.Tools.Invoke("run_procedure_values", run), "another project");
            fixture.Tools.BoundProject = "P"; fixture.Settings.VbeEditApproval = "ReadOnly";
            Failed(fixture.Tools.Invoke("procedure_values_status", status), "shared context not authorized");
            fixture.Tools.SetReadAccess(new string[0], true);
            Failed(fixture.Tools.Invoke("run_procedure_values", run), "read-only execution");
            Success(fixture.Tools.Invoke("procedure_values_status", status), "read-only inspection");
            fixture.Settings.VbeEditApproval = "Unknown";
            Failed(fixture.Tools.Invoke("run_procedure_values", run), "unknown policy");
            fixture.Settings.VbeEditApproval = "AskEachTime";
            fixture.Tools.ShowApproval = (dialog, owner) => DialogResult.No;
            Failed(fixture.Tools.Invoke("run_procedure_values", run), "approval refusal");
            Assert.AreEqual(0, runs); Assert.AreEqual(2, reads);
            fixture.Tools.ShowApproval = (dialog, owner) => DialogResult.Yes;
            Success(fixture.Tools.Invoke("run_procedure_values", run), "approved execution");
            Assert.AreEqual(1, runs);
        }
    }
}

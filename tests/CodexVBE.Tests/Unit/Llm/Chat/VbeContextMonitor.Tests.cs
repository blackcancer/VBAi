using System;
using CodexVBE;
using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace CodexVBE.Tests.Unit
{
    /// <summary>Vérifie les sondages incrémentaux, leurs erreurs et les formes JSON inattendues.</summary>
    [TestClass, TestCategory("Unit")]
    public sealed class VbeContextMonitorTests
    {
        /// <summary>Ignore les données non structurées et distingue les changements avec ou sans abonnement.</summary>
        [TestMethod]
        public void PollingHandlesFailuresMissingFieldsAndUnobservedChanges()
        {
            bool fail = false; object modules = null, code = null; int projects = 0;
            var monitor = new VbeContextMonitor(request => {
                if (fail) return Response.Failure("closed");
                if (request.Command == "list_projects") return Response.Success(new { Version = projects++ });
                if (request.Command == "list_modules") return Response.Success(modules);
                if (request.Command == "read_module") return Response.Success(code);
                return Response.Success(new object[0]);
            });
            monitor.Step(null); monitor.Step(""); monitor.Step("");
            modules = new object[] { "invalid", new { Other = "missing" }, new { Name = "M" } };
            monitor.Step("P"); monitor.Step("P"); monitor.Step("P");
            code = new { Other = "no SHA" }; monitor.Step("P"); monitor.Step("P"); monitor.Step("P");
            code = new { Sha256 = "one" }; monitor.Step("P"); monitor.Step("P"); monitor.Step("P");
            int changed = 0; monitor.Changed += () => changed++;
            code = new { Sha256 = "two" }; monitor.Step("P"); monitor.Step("P"); monitor.Step("P"); Assert.IsTrue(changed > 0);
            fail = true; monitor.Step("Other"); monitor.Step("Other");
            fail = false; modules = new { NotAnArray = true }; monitor.Step("Third"); monitor.Step("Third");
        }
    }
}

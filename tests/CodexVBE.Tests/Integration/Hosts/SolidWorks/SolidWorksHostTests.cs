using System;
using System.Diagnostics;
using System.Linq;
using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace CodexVBE.Tests.Integration
{
    /// <summary>Vérifie une instance SolidWorks VBE fournie par l’environnement, sans en lancer une nouvelle.</summary>
    [TestClass]
    [TestCategory("SolidWorks")]
    public sealed class SolidWorksHostTests
    {
        /// <summary>Confirme que le PID fourni est un processus SLDWORKS vivant avec CodexVBE connecté.</summary>
        [TestMethod]
        public void ExistingSolidWorksVbeHasConnectedCodexAddIn()
        {
            var suppliedPid = Environment.GetEnvironmentVariable("CODEXVBE_SOLIDWORKS_PID");
            if (string.IsNullOrWhiteSpace(suppliedPid))
                Assert.Inconclusive("SolidWorks is opt-in. Supply an existing preloaded SLDWORKS PID in CODEXVBE_SOLIDWORKS_PID.");

            int processId;
            Assert.IsTrue(int.TryParse(suppliedPid, out processId) && processId > 0,
                "CODEXVBE_SOLIDWORKS_PID must be a positive integer.");

            using (var process = Process.GetProcessById(processId))
            {
                Assert.AreEqual("SLDWORKS", process.ProcessName, true,
                    "The supplied PID must identify an existing SLDWORKS process.");
                Assert.IsFalse(process.HasExited, "The supplied SLDWORKS process already exited.");

                var environment = VbeBridgeClient.Read(processId, "vbe_environment");
                Assert.IsNotNull(environment, "The supplied SLDWORKS PID has no CodexVBE bridge; preload its VBE and add-in.");
                Assert.AreEqual(true, environment["Ok"]);
                var fields = VbeBridgeClient.Object(VbeBridgeClient.Object(environment["Data"])["Properties"]);
                Assert.IsTrue(Convert.ToInt32(fields["ProjectCount"]) >= 1,
                    "The preloaded SolidWorks VBE has no project.");

                var addIns = VbeBridgeClient.Read(processId, "list_addins");
                Assert.IsNotNull(addIns, "The CodexVBE bridge disconnected while reading VBE add-ins.");
                Assert.AreEqual(true, addIns["Ok"]);
                Assert.IsTrue(((object[])VbeBridgeClient.Object(addIns["Data"])["AddIns"]).Any(addIn => {
                    var properties = VbeBridgeClient.Object(VbeBridgeClient.Object(addIn)["Properties"]);
                    return Convert.ToString(properties["ProgId"]) == "CodexVBE.AddIn" &&
                        Convert.ToBoolean(properties["Connect"]);
                }), "CodexVBE.AddIn is not connected in the supplied SolidWorks VBE.");

                Assert.IsFalse(process.HasExited, "The supplied SLDWORKS process exited during inspection.");
            }
        }
    }
}

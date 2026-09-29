using System;
using System.IO;
using System.Linq;
using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace VBAi.Tests.Integration
{
    /// <summary>Vérifie les commandes de sauvegarde de projet dans une instance Excel isolée.</summary>
    [TestClass]
    [TestCategory("Excel")]
    public sealed class ExcelProjectPersistenceTests
    {
        /// <summary>Sauvegarde le classeur macro puis vérifie le chemin, l’état et le résultat relus par le pont.</summary>
        [TestMethod]
        [STATestMethod]
        public void IsolatedExcelCanSaveMacroProjectThroughItsOwnVbeBridge()
        {
            using (var other = ExcelVbeFixture.Start())
            using (var host = ExcelVbeFixture.Start())
            {
                Assert.AreNotEqual(other.ProcessId, host.ProcessId, "The ownership test requires two distinct disposable Excel processes.");
                var projectsResponse = host.Command("list_projects");
                Assert.AreEqual(true, projectsResponse["Ok"]);
                var projects = (object[])projectsResponse["Data"];
                Assert.AreEqual(1, projects.Length, "The isolated Excel instance must expose exactly one test project.");
                string project = Convert.ToString(VbeBridgeClient.Object(projects.Single())["Name"]);

                var beforeResponse = host.Command(new { Command = "project_persistence_status", Project = project });
                Assert.AreEqual(true, beforeResponse["Ok"]);
                var before = VbeBridgeClient.Object(beforeResponse["Data"]);
                Assert.AreEqual(true, before["HostAvailable"], "Host lookup: " + Convert.ToString(before["Reason"]));
                Assert.AreEqual(false, before["HostHasPath"]);

                var projectResponse = host.Command(new { Command = "project_properties", Project = project });
                Assert.AreEqual(true, projectResponse["Ok"]);
                string version = Convert.ToString(VbeBridgeClient.Object(projectResponse["Data"])["Version"]);
                string path = host.File("isolated.xlsm");
                var saveAsResponse = host.Command(new { Command = "save_host_document_as", Project = project,
                    Path = path, ExpectedProjectVersion = version });
                Assert.AreEqual(true, saveAsResponse["Ok"], Convert.ToString(saveAsResponse["Error"]));
                var saveAs = VbeBridgeClient.Object(saveAsResponse["Data"]);
                Assert.AreEqual(path, Convert.ToString(saveAs["HostPath"]));
                Assert.IsTrue(File.Exists(path));

                var afterResponse = host.Command(new { Command = "project_persistence_status", Project = project });
                Assert.AreEqual(true, afterResponse["Ok"]);
                var after = VbeBridgeClient.Object(afterResponse["Data"]);
                Assert.AreEqual(true, after["HostAvailable"]);
                Assert.AreEqual(path, Convert.ToString(after["HostPath"]));
                Assert.AreEqual(true, after["HostSaved"]);
                projectResponse = host.Command(new { Command = "project_properties", Project = project });
                Assert.AreEqual(true, projectResponse["Ok"]);
                version = Convert.ToString(VbeBridgeClient.Object(projectResponse["Data"])["Version"]);
                var saveResponse = host.Command(new { Command = "save_host_document", Project = project,
                    ExpectedProjectVersion = version, ExpectedHostPath = path });
                Assert.AreEqual(true, saveResponse["Ok"], Convert.ToString(saveResponse["Error"]));
                Assert.AreEqual(true, VbeBridgeClient.Object(saveResponse["Data"])["SaveInvoked"]);
                var otherProjects = (object[])other.Command("list_projects")["Data"];
                string otherProject = Convert.ToString(VbeBridgeClient.Object(otherProjects.Single())["Name"]);
                var otherState = VbeBridgeClient.Object(other.Command(new { Command = "project_persistence_status", Project = otherProject })["Data"]);
                Assert.AreEqual(true, otherState["HostAvailable"]);
                Assert.AreEqual(false, otherState["HostHasPath"], "Saving one VBE project must leave the other disposable workbook unsaved.");
            }
        }
    }
}

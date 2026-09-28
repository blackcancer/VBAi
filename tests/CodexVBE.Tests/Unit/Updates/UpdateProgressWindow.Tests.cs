using System;
using System.IO;
using System.Linq;
using System.Text;
using System.Collections.Generic;
using System.Diagnostics;
using System.Net;
using System.Net.Http;
using System.Threading;
using System.Threading.Tasks;
using System.Web.Script.Serialization;
using System.Windows.Forms;
using CodexVBE;
using CodexVBE.Tests.Infrastructure;
using Microsoft.VisualStudio.TestTools.UnitTesting;
namespace CodexVBE.Tests.Unit
{
    [TestClass, TestCategory("Unit")]
    public sealed class UpdateProgressWindowTests
    {
        [STATestMethod]
        public void ExternalProgressWindowCancellationPersistsWithoutLaunchingAnything()
        {
            using (var scope = new UpdateScope())
            using (var window = new UpdateProgressWindow())
            {
                var job = UpdateInstallerRunnerTests.Job(scope); job.Culture = "fr-FR";
                window.Configure(scope.Root, job, false);
                Assert.AreEqual("Mises à jour VBAi", window.Text);
                LlmBoundaryScope.Call(window, "Close_Click", null, EventArgs.Empty);
                Assert.AreEqual("Update cancelled.", UpdateInstallJob.Load(scope.Root).Status);
            }
        }
    }
}

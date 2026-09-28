using System;
using System.IO;
using CodexVBE;
using CodexVBE.Tests.Infrastructure;
using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace CodexVBE.Tests.Unit
{
    [TestClass, TestCategory("Unit")]
    public sealed class ProjectHelpTests
    {
        /// <summary>Le handle est rapporté sans certifier la rubrique ni exécuter une application arbitraire.</summary>
        [TestMethod]
        public void ProjectHelpRequiresTheInspectedConfiguredLocalChmAndUnsignedContext()
        {
            string root = Path.Combine(Path.GetTempPath(), "VBAi-help-" + Guid.NewGuid().ToString("N")); Directory.CreateDirectory(root);
            try
            {
                string file = Path.Combine(root, "help.chm"); File.WriteAllText(file, "test");
                var vbe = new IdeSurfaceFixture.Vbe(); var project = new IdeSurfaceFixture.Project { HelpFile = file }; vbe.VBProjects.Add(project);
                var service = new VbeProjectComponents(vbe, new VbeForms(vbe)); int calls = 0;
                service.HelpLauncher = (path, context) => { calls++; Assert.AreEqual(file, path); Assert.AreEqual(Convert.ToUInt32(project.HelpContextID), context); return context == 0 ? IntPtr.Zero : new IntPtr(1); };
                foreach (object context in new object[] { 0, "12", uint.MaxValue })
                {
                    project.HelpContextID = context;
                    dynamic before = service.ProjectProperties("P");
                    dynamic result = service.OpenProjectHelp(new Request { Project = "P", ExpectedProjectVersion = before.Version });
                    Assert.IsFalse((bool)result.TopicVerified); Assert.AreEqual(Convert.ToUInt32(context) != 0, (bool)result.HelpWindowCreated);
                }
                foreach (string invalid in new[] { null, "", "relative.chm", Path.Combine(root, "help.hlp") })
                {
                    project.HelpFile = invalid; dynamic before = service.ProjectProperties("P");
                    Assert.ThrowsException<InvalidOperationException>(() => service.OpenProjectHelp(new Request { Project = "P", ExpectedProjectVersion = before.Version }));
                }
                project.HelpFile = Path.Combine(root, "missing.chm"); dynamic missing = service.ProjectProperties("P");
                Assert.ThrowsException<FileNotFoundException>(() => service.OpenProjectHelp(new Request { Project = "P", ExpectedProjectVersion = missing.Version }));
                project.HelpFile = file;
                foreach (object context in new object[] { -1, null, "not a number", "4294967296" })
                {
                    project.HelpContextID = context; dynamic before = service.ProjectProperties("P");
                    Assert.ThrowsException<InvalidOperationException>(() => service.OpenProjectHelp(new Request { Project = "P", ExpectedProjectVersion = before.Version }));
                }
                Assert.AreEqual(3, calls);
                Assert.ThrowsException<ArgumentException>(() => service.OpenProjectHelp(null));
                Assert.ThrowsException<ArgumentException>(() => service.OpenProjectHelp(new Request { Project = "P" }));
                Assert.ThrowsException<InvalidOperationException>(() => service.OpenProjectHelp(new Request { Project = "P", ExpectedProjectVersion = "stale" }));
            }
            finally { Directory.Delete(root, true); }
        }
    }
}

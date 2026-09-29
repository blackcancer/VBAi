using System;
using System.IO;
using VBAi;
using VBAi.Tests.Infrastructure;
using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace VBAi.Tests.Unit
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

namespace VBAi.Tests.Unit
{
    [Microsoft.VisualStudio.TestTools.UnitTesting.TestClass, Microsoft.VisualStudio.TestTools.UnitTesting.TestCategory("Unit"), Microsoft.VisualStudio.TestTools.UnitTesting.DoNotParallelize]
    public sealed class ProjectHelpNativeBoundaryTests
    {
        [Microsoft.VisualStudio.TestTools.UnitTesting.TestMethod]
        public void DefaultHelpAdapterPassesNativeContextCommandAndUnverifiedHandle()
        {
            string file = System.IO.Path.Combine(System.IO.Path.GetTempPath(), "owned-help-" + System.Guid.NewGuid().ToString("N") + ".chm");
            var original = VbeProjectComponents.NativeHelp;
            try
            {
                System.IO.File.WriteAllText(file, "owned placeholder");
                var vbe = new VBAi.Tests.Infrastructure.IdeSurfaceFixture.Vbe();
                var project = new VBAi.Tests.Infrastructure.IdeSurfaceFixture.Project { HelpFile = file };
                vbe.VBProjects.Add(project);
                var service = new VbeProjectComponents(vbe, new VbeForms(vbe));
                foreach (uint context in new[] { 0U, 12U, uint.MaxValue })
                {
                    project.HelpContextID = context; int calls = 0;
                    VbeProjectComponents.NativeHelp = (owner, path, command, data) => {
                        calls++;
                        Microsoft.VisualStudio.TestTools.UnitTesting.Assert.AreEqual(System.IntPtr.Zero, owner);
                        Microsoft.VisualStudio.TestTools.UnitTesting.Assert.AreEqual(file, path);
                        Microsoft.VisualStudio.TestTools.UnitTesting.Assert.AreEqual(context == 0 ? 0U : 15U, command);
                        Microsoft.VisualStudio.TestTools.UnitTesting.Assert.AreEqual((ulong)context, data.ToUInt64());
                        return context == 0 ? System.IntPtr.Zero : new System.IntPtr(2);
                    };
                    dynamic metadata = service.ProjectProperties("P");
                    dynamic result = service.OpenProjectHelp(new Request { Project = "P", ExpectedProjectVersion = metadata.Version });
                    Microsoft.VisualStudio.TestTools.UnitTesting.Assert.AreEqual(1, calls);
                    Microsoft.VisualStudio.TestTools.UnitTesting.Assert.AreEqual(context != 0, (bool)result.HelpWindowCreated);
                    Microsoft.VisualStudio.TestTools.UnitTesting.Assert.IsFalse((bool)result.TopicVerified);
                }
                VbeProjectComponents.NativeHelp = (owner, path, command, data) => { throw new System.IO.IOException("native help unavailable"); };
                dynamic before = service.ProjectProperties("P");
                Microsoft.VisualStudio.TestTools.UnitTesting.Assert.ThrowsException<System.IO.IOException>(() => service.OpenProjectHelp(new Request { Project = "P", ExpectedProjectVersion = before.Version }));
            }
            finally { VbeProjectComponents.NativeHelp = original; System.IO.File.Delete(file); }
        }
    }
}
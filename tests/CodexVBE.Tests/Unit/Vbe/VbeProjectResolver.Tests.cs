namespace CodexVBE.Tests.Unit
{
    using System;
    using System.Collections.Generic;
    using System.IO;
    using CodexVBE;
    using Microsoft.VisualStudio.TestTools.UnitTesting;

    [TestClass]
    [TestCategory("Unit")]
    public sealed partial class ProjectResolverTests
    {
        [TestMethod]
        public void AmbiguousProjectNamesRequireTheExactSavedDocumentPath()
        {
            var first = new FakeProject
            {
                Name = "VBAProject",
                FileName = @"C:\Temp\First.xlsm"
            };
            var second = new FakeProject
            {
                Name = "VBAProject",
                FileName = @"C:\Temp\Second.xlsm"
            };
            var host = new FakeVbe
            {
                VBProjects = new List<FakeProject>
                {
                    first,
                    second
                }
            };
            Assert.ThrowsException<ArgumentException>(() => VbeProjectResolver.Resolve(host, " "));
            Assert.ThrowsException<InvalidOperationException>(() => VbeProjectResolver.Resolve(host, "VBAProject"));
            Assert.AreSame(second, (object)VbeProjectResolver.Resolve(host, second.FileName.ToLowerInvariant()));
            Assert.ThrowsException<InvalidOperationException>(() => VbeProjectResolver.Resolve(host, @"C:\Temp\Missing.xlsm"));
        }

        [TestMethod]
        public void NameLookupIsCaseInsensitiveForAUniqueProject()
        {
            var project = new FakeProject
            {
                Name = "Été",
                FileName = Path.Combine(Path.GetTempPath(), "été.xlsm")
            };
            var host = new FakeVbe
            {
                VBProjects = new List<FakeProject>
                {
                    project
                }
            };
            Assert.AreSame(project, (object)VbeProjectResolver.Resolve(host, "ÉTÉ"));
        }
    }
}

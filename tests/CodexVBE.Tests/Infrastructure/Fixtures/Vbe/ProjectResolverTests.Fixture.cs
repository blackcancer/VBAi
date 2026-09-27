namespace CodexVBE.Tests.Unit
{
    using System;
    using System.Collections.Generic;
    using System.IO;
    using CodexVBE;
    using Microsoft.VisualStudio.TestTools.UnitTesting;

    public sealed partial class ProjectResolverTests
    {
        public sealed class FakeVbe
        {
            public List<FakeProject> VBProjects { get; set; }
        }

        public sealed class FakeProject
        {
            public string Name { get; set; }
            public string FileName { get; set; }
        }
    }
}

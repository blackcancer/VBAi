namespace VBAi.Tests.Unit
{
    using System;
    using System.Collections.Generic;
    using System.IO;
    using VBAi;
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
        public sealed class HeterogeneousVbe
        {
            public List<object> VBProjects { get; } = new List<object>();
        }
        public sealed class UnsavedProject
        {
            public string Name => "Unsaved";
            public string FileName => throw new InvalidOperationException("Unsaved project has no file name");
        }
    }
}

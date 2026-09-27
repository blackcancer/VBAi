namespace CodexVBE.Tests.Unit
{
    using System;
    using System.Collections.Generic;
    using System.Linq;
    using CodexVBE;
    using Microsoft.VisualStudio.TestTools.UnitTesting;

    public sealed partial class VbeFormsContractTests
    {
        public sealed class FakeVbe
        {
            public List<FakeProject> VBProjects { get; } = new List<FakeProject>();
        }

        public sealed class FakeProject
        {
            public string Name { get; set; }
            public List<FakeComponent> VBComponents { get; } = new List<FakeComponent>();
        }

        public sealed class FakeComponent
        {
            public string Name { get; set; }
            public int Type { get; set; }
            public bool HasOpenDesigner { get; set; }
        }
    }
}

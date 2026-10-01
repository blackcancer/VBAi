using System;
using System.Collections.Generic;
using System.Linq;

namespace VBAi
{
    internal sealed class VbaTestProjectSnapshot
    {
        public string Id { get; set; }
        public string Selector { get; set; }
        public string Name { get; set; }
        public string HostPath { get; set; }
        public string Revision { get; set; }
        public string ReferencesHash { get; set; }
        public VbaTestModuleSnapshot[] Modules { get; set; } = new VbaTestModuleSnapshot[0];
    }

    internal sealed class VbaTestModuleSnapshot
    {
        public string Name { get; set; }
        public string Source { get; set; }
        public string Hash { get; set; }
        public int ComponentType { get; set; }
    }

    internal sealed class VbaTestDescriptor
    {
        public string Id { get; set; }
        public string Module { get; set; }
        public string Procedure { get; set; }
        public string Kind { get; set; }
        public int Line { get; set; }
        public string[] Categories { get; set; } = new string[0];
        public string IgnoreReason { get; set; }
        public string Diagnostic { get; set; }
    }

    internal sealed class VbaTestModule
    {
        public string Name { get; set; }
        public string Diagnostic { get; set; }
        public List<VbaTestDescriptor> Tests { get; set; } = new List<VbaTestDescriptor>();
        public VbaTestDescriptor ModuleInitialize { get; set; }
        public VbaTestDescriptor ModuleCleanup { get; set; }
        public VbaTestDescriptor TestInitialize { get; set; }
        public VbaTestDescriptor TestCleanup { get; set; }
    }

    internal sealed class VbaTestCatalog
    {
        public VbaTestProjectSnapshot Project { get; set; }
        public List<VbaTestModule> Modules { get; set; } = new List<VbaTestModule>();
        public List<string> Diagnostics { get; set; } = new List<string>();
        public IEnumerable<VbaTestDescriptor> Tests => Modules.SelectMany(module => module.Tests);
    }

    internal enum VbaTestOutcome
    {
        NotRun, Passed, Failed, Error, Inconclusive, Skipped, Cancelled, Blocked, OutcomeUnknown
    }

    internal sealed class VbaTestResult
    {
        public VbaTestDescriptor Test { get; set; }
        public VbaTestOutcome Outcome { get; set; }
        public string Message { get; set; }
        public string Phase { get; set; }
        public int ErrorNumber { get; set; }
        public TimeSpan Duration { get; set; }
    }

    internal sealed class VbaTestRun
    {
        public string Id { get; set; }
        public string Project { get; set; }
        public string Revision { get; set; }
        public List<VbaTestResult> Results { get; set; } = new List<VbaTestResult>();
        public bool OutcomeUnknown { get; set; }
        public string Error { get; set; }
        public VbaCoverageReport Coverage { get; set; }
    }
}

using System;
using System.Collections.Generic;
using System.Linq;

namespace VBAi
{

    /// <summary>Owns the vba test project snapshot state and operations.</summary>
    internal sealed class VbaTestProjectSnapshot
    {

        /// <summary>Gets or sets the id.</summary>
        /// <value>Current id exposed by vba test project snapshot.</value>
        public string Id { get; set; }

        /// <summary>Gets or sets the selector.</summary>
        /// <value>Current selector exposed by vba test project snapshot.</value>
        public string Selector { get; set; }

        /// <summary>Gets or sets the name.</summary>
        /// <value>Current name exposed by vba test project snapshot.</value>
        public string Name { get; set; }

        /// <summary>Gets or sets the host path.</summary>
        /// <value>Current host path exposed by vba test project snapshot.</value>
        public string HostPath { get; set; }

        /// <summary>Gets or sets the revision.</summary>
        /// <value>Current revision exposed by vba test project snapshot.</value>
        public string Revision { get; set; }

        /// <summary>Gets or sets the references hash.</summary>
        /// <value>Current references hash exposed by vba test project snapshot.</value>
        public string ReferencesHash { get; set; }

        /// <summary>Gets or sets the modules.</summary>
        /// <value>Current modules exposed by vba test project snapshot.</value>
        public VbaTestModuleSnapshot[] Modules { get; set; } = new VbaTestModuleSnapshot[0];
    }

    /// <summary>Owns the vba test module snapshot state and operations.</summary>
    internal sealed class VbaTestModuleSnapshot
    {

        /// <summary>Gets or sets the name.</summary>
        /// <value>Current name exposed by vba test module snapshot.</value>
        public string Name { get; set; }

        /// <summary>Gets or sets the source.</summary>
        /// <value>Current source exposed by vba test module snapshot.</value>
        public string Source { get; set; }

        /// <summary>Gets or sets the hash.</summary>
        /// <value>Current hash exposed by vba test module snapshot.</value>
        public string Hash { get; set; }

        /// <summary>Gets or sets the component type.</summary>
        /// <value>Current component type exposed by vba test module snapshot.</value>
        public int ComponentType { get; set; }
    }

    /// <summary>Owns the vba test descriptor state and operations.</summary>
    internal sealed class VbaTestDescriptor
    {

        /// <summary>Gets or sets the id.</summary>
        /// <value>Current id exposed by vba test descriptor.</value>
        public string Id { get; set; }

        /// <summary>Gets or sets the module.</summary>
        /// <value>Current module exposed by vba test descriptor.</value>
        public string Module { get; set; }

        /// <summary>Gets or sets the procedure.</summary>
        /// <value>Current procedure exposed by vba test descriptor.</value>
        public string Procedure { get; set; }

        /// <summary>Gets or sets the kind.</summary>
        /// <value>Current kind exposed by vba test descriptor.</value>
        public string Kind { get; set; }

        /// <summary>Gets or sets the line.</summary>
        /// <value>Current line exposed by vba test descriptor.</value>
        public int Line { get; set; }

        /// <summary>Gets or sets the categories.</summary>
        /// <value>Current categories exposed by vba test descriptor.</value>
        public string[] Categories { get; set; } = new string[0];

        /// <summary>Gets or sets the ignore reason.</summary>
        /// <value>Current ignore reason exposed by vba test descriptor.</value>
        public string IgnoreReason { get; set; }

        /// <summary>Gets or sets the diagnostic.</summary>
        /// <value>Current diagnostic exposed by vba test descriptor.</value>
        public string Diagnostic { get; set; }
    }

    /// <summary>Owns the vba test module state and operations.</summary>
    internal sealed class VbaTestModule
    {

        /// <summary>Gets or sets the name.</summary>
        /// <value>Current name exposed by vba test module.</value>
        public string Name { get; set; }

        /// <summary>Gets or sets the diagnostic.</summary>
        /// <value>Current diagnostic exposed by vba test module.</value>
        public string Diagnostic { get; set; }

        /// <summary>Gets or sets the tests.</summary>
        /// <value>Current tests exposed by vba test module.</value>
        public List<VbaTestDescriptor> Tests { get; set; } = new List<VbaTestDescriptor>();

        /// <summary>Gets or sets the module initialize.</summary>
        /// <value>Current module initialize exposed by vba test module.</value>
        public VbaTestDescriptor ModuleInitialize { get; set; }

        /// <summary>Gets or sets the module cleanup.</summary>
        /// <value>Current module cleanup exposed by vba test module.</value>
        public VbaTestDescriptor ModuleCleanup { get; set; }

        /// <summary>Gets or sets the test initialize.</summary>
        /// <value>Current test initialize exposed by vba test module.</value>
        public VbaTestDescriptor TestInitialize { get; set; }

        /// <summary>Gets or sets the test cleanup.</summary>
        /// <value>Current test cleanup exposed by vba test module.</value>
        public VbaTestDescriptor TestCleanup { get; set; }
    }

    /// <summary>Owns the vba test catalog state and operations.</summary>
    internal sealed class VbaTestCatalog
    {

        /// <summary>Gets or sets the project.</summary>
        /// <value>Current project exposed by vba test catalog.</value>
        public VbaTestProjectSnapshot Project { get; set; }

        /// <summary>Gets or sets the modules.</summary>
        /// <value>Current modules exposed by vba test catalog.</value>
        public List<VbaTestModule> Modules { get; set; } = new List<VbaTestModule>();

        /// <summary>Gets or sets the diagnostics.</summary>
        /// <value>Current diagnostics exposed by vba test catalog.</value>
        public List<string> Diagnostics { get; set; } = new List<string>();

        /// <summary>Gets the tests.</summary>
        /// <value>Current tests exposed by vba test catalog.</value>
        public IEnumerable<VbaTestDescriptor> Tests => Modules.SelectMany(module => module.Tests);
    }

    /// <summary>Lists the supported vba test outcome values.</summary>
    internal enum VbaTestOutcome
    {

        /// <summary>Identifies the not run case of vba test outcome.</summary>
        NotRun,

/// <summary>Identifies the passed case of vba test outcome.</summary>
Passed,

/// <summary>Identifies the failed case of vba test outcome.</summary>
Failed,

/// <summary>Identifies the error case of vba test outcome.</summary>
Error,

/// <summary>Identifies the inconclusive case of vba test outcome.</summary>
Inconclusive,

/// <summary>Identifies the skipped case of vba test outcome.</summary>
Skipped,

/// <summary>Identifies the cancelled case of vba test outcome.</summary>
Cancelled,

/// <summary>Identifies the blocked case of vba test outcome.</summary>
Blocked,

/// <summary>Identifies the outcome unknown case of vba test outcome.</summary>
OutcomeUnknown
    }

    /// <summary>Owns the vba test result state and operations.</summary>
    internal sealed class VbaTestResult
    {

        /// <summary>Gets or sets the test.</summary>
        /// <value>Current test exposed by vba test result.</value>
        public VbaTestDescriptor Test { get; set; }

        /// <summary>Gets or sets the outcome.</summary>
        /// <value>Current outcome exposed by vba test result.</value>
        public VbaTestOutcome Outcome { get; set; }

        /// <summary>Gets or sets the message.</summary>
        /// <value>Current message exposed by vba test result.</value>
        public string Message { get; set; }

        /// <summary>Gets or sets the phase.</summary>
        /// <value>Current phase exposed by vba test result.</value>
        public string Phase { get; set; }

        /// <summary>Gets or sets the error number.</summary>
        /// <value>Current error number exposed by vba test result.</value>
        public int ErrorNumber { get; set; }

        /// <summary>Gets or sets the duration.</summary>
        /// <value>Current duration exposed by vba test result.</value>
        public TimeSpan Duration { get; set; }
    }

    /// <summary>Owns the vba test run state and operations.</summary>
    internal sealed class VbaTestRun
    {

        /// <summary>Gets or sets the id.</summary>
        /// <value>Current id exposed by vba test run.</value>
        public string Id { get; set; }

        /// <summary>Gets or sets the project.</summary>
        /// <value>Current project exposed by vba test run.</value>
        public string Project { get; set; }

        /// <summary>Gets or sets the revision.</summary>
        /// <value>Current revision exposed by vba test run.</value>
        public string Revision { get; set; }

        /// <summary>Gets or sets the results.</summary>
        /// <value>Current results exposed by vba test run.</value>
        public List<VbaTestResult> Results { get; set; } = new List<VbaTestResult>();

        /// <summary>Gets or sets the outcome unknown.</summary>
        /// <value>Current outcome unknown exposed by vba test run.</value>
        public bool OutcomeUnknown { get; set; }

        /// <summary>Gets or sets the error.</summary>
        /// <value>Current error exposed by vba test run.</value>
        public string Error { get; set; }

        /// <summary>Gets or sets the coverage.</summary>
        /// <value>Current coverage exposed by vba test run.</value>
        public VbaCoverageReport Coverage { get; set; }
    }
}

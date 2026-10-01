using System;
using System.Collections.Generic;
using System.Threading;
using System.Threading.Tasks;

namespace VBAi
{
    /// <summary>Owning-thread boundary between the test explorer and live project operations.</summary>
    internal interface IVbaTestExplorerService
    {
        IReadOnlyList<VbaTestProjectSnapshot> ReadProjects();
        VbaTestCatalog Discover(string projectId);
        string ExecutionUnavailableReason(VbaTestCatalog catalog);
        Task<VbaTestRun> RunAsync(VbaTestCatalog catalog, IReadOnlyList<VbaTestDescriptor> tests,
            Action<VbaTestResult> onResult, CancellationToken cancellation);
        void Navigate(VbaTestCatalog catalog, VbaTestDescriptor test);
        void InstallSupport(VbaTestCatalog catalog);
    }

    internal interface IVbaTestCoverageExplorerService
    {
        string CoverageUnavailableReason(VbaTestCatalog catalog);
        Task<VbaTestRun> RunCoverageAsync(VbaTestCatalog catalog, IReadOnlyList<VbaTestDescriptor> tests,
            Action<VbaTestResult> onResult, CancellationToken cancellation);
    }
}

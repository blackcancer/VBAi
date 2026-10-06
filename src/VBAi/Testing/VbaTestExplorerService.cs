using System;
using System.Collections.Generic;
using System.Threading;
using System.Threading.Tasks;

namespace VBAi
{

    /// <summary>Owning-thread boundary between the test explorer and live project operations.</summary>
    internal interface IVbaTestExplorerService
    {

        /// <summary>Captures the projects currently available to the explorer on its owning VBE thread.</summary>
        /// <returns>Read-only project snapshots with stable IDs for subsequent catalog and run requests.</returns>
        IReadOnlyList<VbaTestProjectSnapshot> ReadProjects();

        /// <summary>Discovers eligible test procedures in the project identified by a prior project snapshot.</summary>
        /// <param name="projectId">Stable project ID returned by <see cref="ReadProjects"/>.</param>
        /// <returns>Catalog bound to that project's current identity and revision; stale IDs are rejected by implementations.</returns>
        VbaTestCatalog Discover(string projectId);

        /// <summary>Explains why the supplied catalog cannot be run in the current host state.</summary>
        /// <param name="catalog">Catalog previously returned by <see cref="Discover"/>.</param>
        /// <returns>Null when execution is currently available; otherwise a user-displayable reason.</returns>
        string ExecutionUnavailableReason(VbaTestCatalog catalog);

        /// <summary>Runs selected catalog tests and reports each completed test through the callback.</summary>
        /// <param name="catalog">Catalog that supplies project identity, revision, and discovered procedures.</param>
        /// <param name="tests">Subset of that catalog to execute, in requested order.</param>
        /// <param name="onResult">Receives each result as execution completes; implementations marshal native work to the owning thread.</param>
        /// <param name="cancellation">Requests cooperative cancellation between safe test phases.</param>
        /// <returns>Task completing with the aggregate run status and results.</returns>
        Task<VbaTestRun> RunAsync(VbaTestCatalog catalog, IReadOnlyList<VbaTestDescriptor> tests,
            Action<VbaTestResult> onResult, CancellationToken cancellation);

        /// <summary>Opens the source location for one test in the catalog's project.</summary>
        /// <param name="catalog">Catalog used to validate project and revision identity.</param>
        /// <param name="test">Discovered procedure whose module and source location should be selected.</param>
        void Navigate(VbaTestCatalog catalog, VbaTestDescriptor test);

        /// <summary>Installs the test support module after validating the target catalog and project.</summary>
        /// <param name="catalog">Catalog identifying the project that receives the support module.</param>
        void InstallSupport(VbaTestCatalog catalog);
    }

    /// <summary>Defines isolated coverage runs for tests and their source project.</summary>
    internal interface IVbaTestCoverageExplorerService
    {

        /// <summary>Explains why coverage cannot be collected for the supplied project catalog.</summary>
        /// <param name="catalog">Catalog identifying source project and eligible tests.</param>
        /// <returns>Null when coverage is available; otherwise a user-displayable reason.</returns>
        string CoverageUnavailableReason(VbaTestCatalog catalog);

        /// <summary>Runs selected tests in a disposable project clone while collecting source coverage.</summary>
        /// <param name="catalog">Catalog whose source snapshot seeds the isolated coverage project.</param>
        /// <param name="tests">Subset of discovered procedures to run in the clone.</param>
        /// <param name="onResult">Receives each test result as it completes.</param>
        /// <param name="cancellation">Requests cooperative cancellation of the coverage run.</param>
        /// <returns>Aggregate run status and per-test results after clone cleanup.</returns>
        Task<VbaTestRun> RunCoverageAsync(VbaTestCatalog catalog, IReadOnlyList<VbaTestDescriptor> tests,
            Action<VbaTestResult> onResult, CancellationToken cancellation);
    }
}

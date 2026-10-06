using System;
using System.Collections.Generic;
using System.Threading;
using System.Threading.Tasks;

namespace VBAi
{

    /// <summary>Owning-thread boundary between the test explorer and live project operations.</summary>
    internal interface IVbaTestExplorerService
    {

        /// <summary>Reads projects for i vba test explorer service.</summary>
        /// <returns>i read only list&lt;vba test project snapshot&gt; produced by the operation for read projects on i vba test explorer service.</returns>
        IReadOnlyList<VbaTestProjectSnapshot> ReadProjects();

        /// <summary>Handles discover for i vba test explorer service.</summary>
        /// <param name="projectId">Text that supplies the project id value. Use the format required by the calling operation.</param>
        /// <returns>vba test catalog produced by the operation for discover on i vba test explorer service.</returns>
        VbaTestCatalog Discover(string projectId);

        /// <summary>Handles execution unavailable reason for i vba test explorer service.</summary>
        /// <param name="catalog">vba test catalog that supplies the catalog for this operation.</param>
        /// <returns>Text produced by the operation for execution unavailable reason on i vba test explorer service.</returns>
        string ExecutionUnavailableReason(VbaTestCatalog catalog);

        /// <summary>Runs async for i vba test explorer service.</summary>
        /// <param name="catalog">vba test catalog that supplies the catalog for this operation.</param>
        /// <param name="tests">i read only list&lt;vba test descriptor&gt; that supplies the tests for this operation.</param>
        /// <param name="onResult">action&lt;vba test result&gt; that supplies the on result for this operation.</param>
        /// <param name="cancellation">Token used to cancel the operation.</param>
        /// <returns>task&lt;vba test run&gt; produced by the operation for run async on i vba test explorer service.</returns>
        Task<VbaTestRun> RunAsync(VbaTestCatalog catalog, IReadOnlyList<VbaTestDescriptor> tests,
            Action<VbaTestResult> onResult, CancellationToken cancellation);

        /// <summary>Handles navigate for i vba test explorer service.</summary>
        /// <param name="catalog">vba test catalog that supplies the catalog for this operation.</param>
        /// <param name="test">vba test descriptor that supplies the test for this operation.</param>
        void Navigate(VbaTestCatalog catalog, VbaTestDescriptor test);

        /// <summary>Handles install support for i vba test explorer service.</summary>
        /// <param name="catalog">vba test catalog that supplies the catalog for this operation.</param>
        void InstallSupport(VbaTestCatalog catalog);
    }

    /// <summary>Defines the i vba test coverage explorer service contract.</summary>
    internal interface IVbaTestCoverageExplorerService
    {

        /// <summary>Handles coverage unavailable reason for i vba test coverage explorer service.</summary>
        /// <param name="catalog">vba test catalog that supplies the catalog for this operation.</param>
        /// <returns>Text produced by the operation for coverage unavailable reason on i vba test coverage explorer service.</returns>
        string CoverageUnavailableReason(VbaTestCatalog catalog);

        /// <summary>Runs coverage async for i vba test coverage explorer service.</summary>
        /// <param name="catalog">vba test catalog that supplies the catalog for this operation.</param>
        /// <param name="tests">i read only list&lt;vba test descriptor&gt; that supplies the tests for this operation.</param>
        /// <param name="onResult">action&lt;vba test result&gt; that supplies the on result for this operation.</param>
        /// <param name="cancellation">Token used to cancel the operation.</param>
        /// <returns>task&lt;vba test run&gt; produced by the operation for run coverage async on i vba test coverage explorer service.</returns>
        Task<VbaTestRun> RunCoverageAsync(VbaTestCatalog catalog, IReadOnlyList<VbaTestDescriptor> tests,
            Action<VbaTestResult> onResult, CancellationToken cancellation);
    }
}

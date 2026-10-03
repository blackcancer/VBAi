namespace VBAi.Tests.Integration
{
    // Test-local names only. The 2026-09-29 product predates these opt-ins and
    // cannot emit the later product phase trace; an empty trace is not native proof.
    internal static class VbeInspectionTrace
    {
        internal const string EnvironmentName = "VBAi_VBE_INSPECTION_TRACE";
        internal const long MaximumFileBytes = 16 * 1024 * 1024;
    }
    internal static class PathVisibilityDiagnostic
    {
        internal const string EnvironmentName = "VBAi_TEST_PATH_VISIBILITY_MANIFEST";
    }
}

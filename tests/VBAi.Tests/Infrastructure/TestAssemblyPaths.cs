using System;
using System.IO;

namespace VBAi.Tests.Infrastructure
{
    /// <summary>Locates deployed test assets independently of the runner's application domain directory.</summary>
    internal static class TestAssemblyPaths
    {
        /// <summary>Gets the physical directory containing this test assembly and its copied fixtures.</summary>
        internal static string OutputDirectory
        {
            get
            {
                string location = typeof(TestAssemblyPaths).Assembly.Location;
                if (string.IsNullOrEmpty(location))
                    throw new InvalidOperationException("Physical test assembly location is required to resolve fixture assets.");
                return Path.GetDirectoryName(location);
            }
        }
    }
}

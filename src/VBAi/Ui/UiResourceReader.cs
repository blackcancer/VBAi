using System;
using System.Reflection;

namespace VBAi
{
    /// <summary>Loads embedded form resources without requiring binding redirects in the application's COM host.</summary>
    internal static class UiResourceReader
    {
        /// <summary>Resolves only the resource compiler's legacy reader identity to the deployed reader assembly.</summary>
        /// <param name="sender">The application domain requesting resolution.</param>
        /// <param name="request">The requested assembly identity.</param>
        /// <returns>The deployed resource reader for the exact known identity; otherwise null.</returns>
        internal static Assembly Resolve(object sender, ResolveEventArgs request)
        {
            return request.Name == "System.Resources.Extensions, Version=4.0.0.0, Culture=neutral, PublicKeyToken=cc7b13ffcd2ddd51"
                ? typeof(System.Resources.Extensions.DeserializingResourceReader).Assembly : null;
        }

        /// <summary>Runs one resource initialization and removes its resolution handler even when initialization fails.</summary>
        /// <param name="initialize">The original form resource initialization.</param>
        internal static void Apply(Action initialize)
        {
            if (initialize == null) throw new ArgumentNullException(nameof(initialize));
            ResolveEventHandler reader = Resolve;
            AppDomain.CurrentDomain.AssemblyResolve += reader;
            try { initialize(); }
            finally { AppDomain.CurrentDomain.AssemblyResolve -= reader; }
        }
    }
}

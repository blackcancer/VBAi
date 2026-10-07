using System;

namespace VBAi.Tests.Integration
{
    /// <summary>Selects an explicitly approved real desktop or a configured isolated desktop without switching either.</summary>
    internal static class NativeTestDesktop
    {
        internal const string MainEnvironment = "VBAi_RUN_MAIN_DESKTOP_TESTS";

        /// <summary>Checks the selected mode before any native host or observer is created.</summary>
        internal static string Select(string mainOptIn, string configured, string required,
            Action requireMain, Action<string> requirePrivate)
        {
            if (requireMain == null || requirePrivate == null)
                throw new ArgumentNullException("Desktop validation dependencies");
            if (mainOptIn == "1")
            {
                if (!string.IsNullOrEmpty(configured) || !string.IsNullOrEmpty(required))
                    throw new InvalidOperationException("Main qualification refuses inherited private desktop descriptors.");
                requireMain();
                return "Default";
            }
            if (string.IsNullOrEmpty(configured) || configured != required)
                throw new InvalidOperationException("Select the real desktop explicitly or configure an exact private desktop pair.");
            requirePrivate(configured);
            return configured;
        }

        /// <summary>Revalidates actual placement and the unchanged environment on the owning or observer thread.</summary>
        internal static void RequireCurrent(string expected)
        {
            string selected = Current();
            if (!string.Equals(selected, expected, StringComparison.Ordinal))
                throw new InvalidOperationException("The selected native test desktop changed; no native action is permitted.");
        }

        /// <summary>Returns the selected desktop only after checking its actual thread and input placement.</summary>
        internal static string Current() => Select(Environment.GetEnvironmentVariable(MainEnvironment),
            Environment.GetEnvironmentVariable("VBAi_TEST_DESKTOP_NAME"),
            Environment.GetEnvironmentVariable("VBAi_QUALIFICATION_DESKTOP"),
            IsolatedTestDesktop.RequireMainCurrent, IsolatedTestDesktop.RequireCurrent);
    }
}

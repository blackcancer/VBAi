using System;
using System.Diagnostics;
using System.IO;
using System.Text;
using System.Web.Script.Serialization;
using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace VBAi.Tests.Integration
{
    /// <summary>Verifies the actual testhost desktop before opted-in isolated campaigns execute any test.</summary>
    [TestClass]
    public sealed class QualificationDesktopGuard
    {
        [AssemblyInitialize]
        public static void Initialize(TestContext context)
        {
            string expected = Environment.GetEnvironmentVariable("VBAi_QUALIFICATION_DESKTOP");
            if (string.IsNullOrEmpty(expected)) return; // Ordinary explicitly interactive invocations retain their contract.
            uint thread = IsolatedTestDesktop.GetCurrentThreadId();
            string actual = IsolatedTestDesktop.DesktopName(thread), input = IsolatedTestDesktop.InputDesktopName();
            IsolatedTestDesktop.RequireObserved(expected, actual, input);
            string evidence = Environment.GetEnvironmentVariable("VBAi_QUALIFICATION_DESKTOP_EVIDENCE");
            RequireEvidence(evidence);
            string receipt = Path.Combine(evidence, "testhost-" + Process.GetCurrentProcess().Id + "-" + Guid.NewGuid().ToString("N") + ".json");
            using (var file = new FileStream(receipt, FileMode.CreateNew, FileAccess.Write, FileShare.Read))
            using (var writer = new StreamWriter(file, new UTF8Encoding(false)))
                writer.Write(new JavaScriptSerializer().Serialize(new { ExpectedDesktop = expected,
                    ActualDesktop = actual, InputDesktop = input, ProcessId = Process.GetCurrentProcess().Id,
                    ThreadId = thread, TestAssemblyMvid = typeof(QualificationDesktopGuard).Module.ModuleVersionId,
                    Utc = DateTime.UtcNow.ToString("o") }));
            context.AddResultFile(receipt);
        }

        internal static void RequireEvidence(string evidence)
        {
            if (string.IsNullOrEmpty(evidence) || !Path.IsPathRooted(evidence) || !Directory.Exists(evidence))
                throw new InvalidOperationException("A campaign-owned desktop evidence directory is required.");
        }
    }
}

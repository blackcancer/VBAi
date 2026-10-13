using Microsoft.VisualStudio.TestTools.UnitTesting;
using System;
using System.IO;
using System.Linq;
using VBAi.Tests.Infrastructure;

namespace VBAi.Tests.Unit
{
    /// <summary>Exercises connection rollback through the actual add-in entry point and disposable host contracts.</summary>
    public sealed partial class AddInCoverageTests
    {
        [STATestMethod]
        [DataRow(false, false, false), DataRow(false, false, true)]
        [DataRow(false, true, false), DataRow(false, true, true)]
        [DataRow(true, false, false), DataRow(true, false, true)]
        [DataRow(true, true, false), DataRow(true, true, true)]
        public void ConnectionRollbackStillReleasesResourcesWhenLoggingOrCleanupFails(bool reporterCreated, bool logFails, bool cleanupFails)
        {
            using (var fixture = new AddInModernEditorFixture())
            {
                var original = new IOException("startup interrupted");
                var logging = new ApplicationException("failure logger unavailable");
                var cleanup = new ApplicationException("update stop unavailable");
                int stopCalls = 0, failureLogs = 0;
                if (reporterCreated) AddIn.StartBridge = bridge => throw original;
                else AddIn.StartUpdateCheck = () => throw original;
                AddIn.WriteLog = message =>
                {
                    fixture.Scope.Logs.Add(message);
                    if (message.StartsWith("OnConnection failed:", StringComparison.Ordinal))
                    { failureLogs++; if (logFails) throw logging; }
                };
                AddIn.StopUpdateCheck = () => { stopCalls++; if (cleanupFails) throw cleanup; };
                Exception observed = null; object[] custom = null;
                try { fixture.Instance.OnConnection(fixture.Scope.Host, 0, fixture.Scope.Host.AddIns.AddIn, ref custom); }
                catch (Exception error) { observed = error; }
                finally { AddIn.StopUpdateCheck = () => { }; }
                Assert.AreEqual(1, stopCalls); Assert.AreEqual(1, failureLogs);
                foreach (var name in new[] { "dispatcher", "server", "crashReporter", "vbe", "addIn" })
                    Assert.IsNull(LlmBoundaryScope.Get<object>(fixture.Instance, name), name);
                var expected = new Exception[] { original, logFails ? logging : null, cleanupFails ? cleanup : null }.Where(error => error != null).ToArray();
                if (expected.Length == 1) Assert.AreSame(original, observed);
                else CollectionAssert.AreEqual(expected, ((AggregateException)observed).InnerExceptions.ToArray());
            }
        }
    }
}

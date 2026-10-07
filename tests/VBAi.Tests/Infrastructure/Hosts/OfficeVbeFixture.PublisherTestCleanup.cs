using Microsoft.VisualStudio.TestTools.UnitTesting;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Threading;

namespace VBAi.Tests.Integration
{
    internal sealed partial class OfficeVbeFixture
    {
        private Dictionary<string, string> publisherTestCleanupSources;
        private string publisherTestCleanupProject, publisherTestCleanupPath;

        /// <summary>Authorizes discard only after the disposable test row has reviewed its complete final source inventory.</summary>
        internal void AllowReviewedPublisherTestCleanup(string moduleName, IDictionary<string, string> expectedSources)
        {
            RequireUsableOwnedHost();
            Assert.AreEqual("Publisher", Kind);
            Assert.AreEqual("VBAiOfficeModule", moduleName, "Only the qualification row's synthetic module may authorize discard.");
            Assert.IsNull(publisherTestCleanupSources, "The reviewed cleanup authority cannot be replaced.");
            Assert.IsTrue(expectedSources != null && expectedSources.ContainsKey(moduleName) &&
                expectedSources.TryGetValue(VbaTestRuntimeSource.ModuleName, out string support) && VbaTestRuntimeSource.IsOwned(support),
                "The complete reviewed synthetic source and generated support are required.");
            RequirePublisherTestCleanupProcess();
            publisherTestCleanupSources = new Dictionary<string, string>(expectedSources, StringComparer.OrdinalIgnoreCase);
            publisherTestCleanupProject = Project;
            publisherTestCleanupPath = DocumentPath;
            steps.Add(new
            {
                PublisherTestCleanup = "ReviewedDiscardAuthorized",
                ProcessId,
                DocumentPath,
                Project,
                Modules = publisherTestCleanupSources.Keys.ToArray(),
                SaveInvoked = false,
                PersistenceQualification = false
            });
            FlushAdapterEvidence();
        }

        private void RequirePublisherTestCleanupProcess()
        {
            RequireUsableOwnedHost();
            Assert.AreEqual(shutdownOwnerThread, Thread.CurrentThread.ManagedThreadId);
            Assert.AreEqual("Publisher", Kind);
            Assert.IsFalse(NativeExecutionUnsettled, "Unsettled test execution forbids Publisher cleanup.");
            Assert.IsTrue(owned && ownedProcess != null && !ownedProcess.HasExited);
            Assert.AreEqual(ProcessId, ownedProcess.Id);
            Assert.IsNotNull(shutdownEvidence);
            Assert.AreEqual(ProcessId, shutdownEvidence.Record["ProcessId"]);
            Assert.AreEqual(shutdownEvidence.Record["ProcessStartedUtc"], ownedProcess.StartTime.ToUniversalTime().ToString("o"));
            Assert.AreEqual(shutdownEvidence.Record["OriginalProcessHandle"], "0x" + unchecked((ulong)ownedProcess.Handle.ToInt64()).ToString("X16"));
            Assert.AreEqual(shutdownEvidence.Record["ProcessImage"], ExcelOwnedProcessImage.Read(ownedProcess.Handle));
            Assert.IsFalse(string.IsNullOrWhiteSpace(Project), "The original Publisher project must already be bound.");
            Assert.AreEqual(Path.GetFullPath(Path.Combine(Root, "Disposable.pub")), Path.GetFullPath(DocumentPath), true);
            Assert.IsTrue(File.Exists(DocumentPath));
        }

        private void RequirePublisherTestCleanupDocument()
        {
            RequirePublisherTestCleanupProcess();
            RequirePublisherPublication("ReviewedTestDiscardIdentity", false);
            object current = null, collection = null, sole = null;
            try
            {
                Assert.IsTrue(SamePublicationPath((string)((dynamic)document).FullName, DocumentPath));
                current = ((dynamic)application).ActiveDocument;
                collection = ((dynamic)application).Documents;
                Assert.AreEqual(1, Convert.ToInt32(((dynamic)collection).Count));
                sole = ((dynamic)collection)[1];
                Assert.IsTrue(SamePublisherTestCleanupIdentity(document, current) && SamePublisherTestCleanupIdentity(document, sole),
                    "Only the original retained synthetic publication may be discarded.");
            }
            finally { BalancePublicationGetter(sole); BalancePublicationGetter(collection); BalancePublicationGetter(current); }
        }

        /// <summary>Revalidates the exact review immediately before the existing single Quit; no Save or Saved setter is used.</summary>
        private void RequireReviewedPublisherTestCleanup()
        {
            RequirePublisherTestCleanupDocument();
            Assert.AreEqual(true, shutdownEvidence.Record["TeardownPrepared"]);
            Assert.AreEqual(0, shutdownEvidence.Record["QuitEntries"]);
            Assert.AreEqual(publisherTestCleanupProject, Project);
            Assert.AreEqual(publisherTestCleanupPath, DocumentPath);
            var state = Data("debug_state", "Project", publisherTestCleanupProject);
            Assert.AreEqual(2, Convert.ToInt32(state["Mode"]));
            var persistence = Data("project_persistence_status", "Project", publisherTestCleanupProject);
            RequirePublisherPersistenceProof(persistence, publisherTestCleanupProject);
            RequirePublisherTestCleanupSources(publisherTestCleanupSources, ReadPublisherTestCleanupSources());
            // A second complete read refuses source/inventory changes caused by any first-pass getter or response.
            RequirePublisherTestCleanupSources(publisherTestCleanupSources, ReadPublisherTestCleanupSources());
            var finalState = Data("debug_state", "Project", publisherTestCleanupProject);
            Assert.AreEqual(2, Convert.ToInt32(finalState["Mode"]));
            RequirePublisherTestCleanupDocument();
            steps.Add(new
            {
                PublisherTestCleanup = "VerifiedBeforeSingleQuit",
                ProcessId,
                DocumentPath,
                Project,
                DiscardReviewedSyntheticTestSources = true,
                SaveInvoked = false,
                SavedSetterInvoked = false,
                PersistenceQualification = false,
                Persistence = persistence,
                NativeExecutionUnsettled = false,
                BridgePending = false,
                BridgeUncertain = false
            });
            FlushAdapterEvidence();
            RequirePublisherTestCleanupProcess();
        }

        private Dictionary<string, string> ReadPublisherTestCleanupSources()
        {
            var actual = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
            foreach (var module in Items("list_modules", "Project", publisherTestCleanupProject))
            {
                string name = (string)module["Name"];
                actual.Add(name, (string)Data("read_module", "Project", publisherTestCleanupProject, "Module", name)["Code"]);
            }
            return actual;
        }

        internal static void RequirePublisherTestCleanupSources(IDictionary<string, string> expected, IDictionary<string, string> actual)
        {
            Assert.AreEqual(expected.Count, actual.Count, "The publication's module inventory changed.");
            foreach (var source in expected)
            {
                Assert.IsTrue(actual.TryGetValue(source.Key, out string text), "The publication's module identity changed: " + source.Key);
                Assert.AreEqual(CanonicalSource(source.Value), CanonicalSource(text), "Unreviewed Publisher source: " + source.Key);
            }
        }

        private static bool SamePublisherTestCleanupIdentity(object first, object second)
        {
            if (first == null || second == null) return false;
            if (!Marshal.IsComObject(first) || !Marshal.IsComObject(second)) return ReferenceEquals(first, second);
            IntPtr left = IntPtr.Zero, right = IntPtr.Zero;
            try { left = Marshal.GetIUnknownForObject(first); right = Marshal.GetIUnknownForObject(second); return left == right; }
            finally { if (right != IntPtr.Zero) Marshal.Release(right); if (left != IntPtr.Zero) Marshal.Release(left); }
        }
    }
}

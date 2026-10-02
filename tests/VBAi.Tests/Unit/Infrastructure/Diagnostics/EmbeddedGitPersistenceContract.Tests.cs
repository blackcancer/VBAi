using System;
using System.Collections.Generic;
using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace VBAi.Tests.Integration
{
    [TestClass, TestCategory("Unit")]
    public sealed class EmbeddedGitPersistenceContractTests
    {
        [TestMethod]
        public void ObservedNormalExitAllowsTheNextHost()
        {
            EmbeddedGitPersistenceContract.RequireNormalExit(Shutdown(true, false, 0));
        }

        [TestMethod]
        public void ReturnedQuitWithoutExitCannotBecomePersistenceAcceptance()
        {
            var shutdown = Shutdown(false, false, 0); shutdown["QuitReturned"] = true;
            Assert.ThrowsException<InvalidOperationException>(() => EmbeddedGitPersistenceContract.RequireNormalExit(shutdown));
        }

        [TestMethod]
        public void ForcedAbnormalAndUnobservedExitsRemainFailures()
        {
            foreach (var shutdown in new[] { Shutdown(true, true, 0), Shutdown(true, false, 1), Shutdown(true, false, null),
                new Dictionary<string, object>(), null })
                Assert.ThrowsException<InvalidOperationException>(() => EmbeddedGitPersistenceContract.RequireNormalExit(shutdown));
        }

        [TestMethod]
        public void LaterOwnedProcessCanReuseThePreviousPid()
        {
            EmbeddedGitPersistenceContract.RequireFreshIdentity("old-guid", "2026-10-02T00:00:00Z", 100,
                "new-guid", "2026-10-02T00:01:00Z", 100);
        }

        [TestMethod]
        public void ChangedPidAloneCannotProveFreshness()
        {
            Assert.ThrowsException<InvalidOperationException>(() => EmbeddedGitPersistenceContract.RequireFreshIdentity(
                "same-root", "2026-10-02T00:00:00Z", 100, "same-root", "2026-10-02T00:01:00Z", 101));
            Assert.ThrowsException<InvalidOperationException>(() => EmbeddedGitPersistenceContract.RequireFreshIdentity(
                "old-root", "2026-10-02T00:00:00Z", 100, "new-root", "2026-10-02T00:00:00Z", 101));
            Assert.ThrowsException<InvalidOperationException>(() => EmbeddedGitPersistenceContract.RequireFreshIdentity(
                "old-root", "invalid", 100, "new-root", "2026-10-02T00:01:00Z", 101));
        }

        [TestMethod]
        public void DiskMutationOrMissingSavedOracleIsRejected()
        {
            string saved = new string('A', 64);
            EmbeddedGitPersistenceContract.RequireDiskHash(saved, saved.ToLowerInvariant());
            Assert.ThrowsException<InvalidOperationException>(() => EmbeddedGitPersistenceContract.RequireDiskHash(saved, new string('B', 64)));
            Assert.ThrowsException<InvalidOperationException>(() => EmbeddedGitPersistenceContract.RequireDiskHash(null, null));
        }

        private static IDictionary<string, object> Shutdown(bool exited, bool forced, object code)
        {
            return new Dictionary<string, object> { ["Exited"] = exited, ["ForcedTermination"] = forced, ["ExitCode"] = code };
        }
    }
}

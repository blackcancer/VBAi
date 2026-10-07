using System;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Text;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using VBAi;

namespace VBAi.Tests.Unit
{
    [TestClass, TestCategory("Unit")]
    public sealed class GitRevisionFixtureTraceTests
    {
        private static string CreateRoot()
        {
            string root = Path.Combine(Path.GetTempPath(), "VBAiGitRevision-" + Guid.NewGuid().ToString("N"));
            Directory.CreateDirectory(root); return root;
        }

        [TestMethod]
        public void CreateNewRejectsReusedOrUnownedRootWithoutOverwritingEvidence()
        {
            string root = CreateRoot();
            try
            {
                var repository = new MacroGitRepository(root, "main");
                Assert.IsNull(GitRevisionFixtureTrace.TryBeginAtRoot(root + Path.DirectorySeparatorChar, repository));
                using (var trace = GitRevisionFixtureTrace.TryBeginAtRoot(root, repository))
                { Assert.IsNotNull(trace); trace.Record(new { Kind = "Retained" }); }
                string path = Path.Combine(root, "diagnostic.jsonl"); byte[] retained = File.ReadAllBytes(path);
                Assert.IsNull(GitRevisionFixtureTrace.TryBeginAtRoot(root, repository));
                CollectionAssert.AreEqual(retained, File.ReadAllBytes(path)); Assert.IsNull(repository.CommandOverride);
                Assert.IsNull(GitRevisionFixtureTrace.TryBeginAtRoot(Path.GetTempPath(), repository));
            }
            finally { Directory.Delete(root, true); }
        }

        [TestMethod]
        public void ForwardingRecordsActualResultHashWithoutRawArgumentsAndRestoresOnPrimaryFailure()
        {
            string root = CreateRoot();
            try
            {
                var repository = new MacroGitRepository(root, "main");
                using (var trace = GitRevisionFixtureTrace.TryBeginAtRoot(root, repository))
                {
                    Assert.IsNotNull(trace); var forwarding = repository.CommandOverride;
                    var result = MacroGitRepositoryTests.Run(repository, new[] { "--version" }, null, false, false);
                    Assert.AreEqual(0, result.ExitCode); Assert.IsTrue(result.Bytes.Length > 0);
                    Assert.AreSame(forwarding, repository.CommandOverride);
                    var primary = new IOException("Synthetic exact primary command failure"); int starts = 0;
                    repository.StartProcess = process => { starts++; throw primary; };
                    var observed = Assert.ThrowsException<IOException>(() => MacroGitRepositoryTests.Run(repository,
                        new[] { "rev-parse", "--verify", "--quiet", "SecretArgumentMustNotAppear" }, null, true, true));
                    Assert.AreSame(primary, observed); Assert.AreEqual(1, starts);
                    Assert.AreSame(forwarding, repository.CommandOverride);
                }
                Assert.IsNull(repository.CommandOverride);
                string journal = File.ReadAllText(Path.Combine(root, "diagnostic.jsonl"));
                StringAssert.Contains(journal, "\"Kind\":\"Command\"");
                StringAssert.Contains(journal, "\"ExitCode\":0");
                StringAssert.Contains(journal, "\"FailureType\":\"System.IO.IOException\"");
                StringAssert.Contains(journal, "\"Kind\":\"Closed\"");
                Assert.IsFalse(journal.Contains("SecretArgumentMustNotAppear"));
                Assert.IsFalse(journal.Contains("git version ")); Assert.IsFalse(journal.Contains(root));
            }
            finally { Directory.Delete(root, true); }
        }

        private sealed class ThrowingPayload
        { public string Value { get { throw new IOException("Synthetic diagnostic serialization failure"); } } }

        [TestMethod]
        public void FailedWriterCannotReplaceTheCommandFailureOrLeakTheForwardingHook()
        {
            string root = CreateRoot();
            try
            {
                var repository = new MacroGitRepository(root, "main");
                using (var trace = GitRevisionFixtureTrace.TryBeginAtRoot(root, repository))
                {
                    Assert.IsNotNull(trace); trace.Record(new ThrowingPayload());
                    var primary = new InvalidOperationException("Native command primary"); int starts = 0;
                    repository.StartProcess = process => { starts++; throw primary; };
                    Assert.AreSame(primary, Assert.ThrowsException<InvalidOperationException>(() =>
                        MacroGitRepositoryTests.Run(repository, new[] { "status" })));
                    Assert.AreEqual(1, starts);
                }
                Assert.IsNull(repository.CommandOverride);
                Assert.AreEqual(1, File.ReadAllLines(Path.Combine(root, "diagnostic.jsonl")).Length);
            }
            finally { Directory.Delete(root, true); }
        }

        [TestMethod]
        public void EventAndByteLimitsBoundPersistentEvidenceWithoutThrowing()
        {
            string root = CreateRoot(); string second = CreateRoot();
            try
            {
                using (var trace = GitRevisionFixtureTrace.TryBeginAtRoot(root, new MacroGitRepository(root, "main")))
                {
                    Assert.IsNotNull(trace);
                    for (int i = 0; i < GitRevisionFixtureTrace.MaximumEvents + 20; i++) trace.Record(new { Kind = "Bounded" });
                }
                string path = Path.Combine(root, "diagnostic.jsonl");
                Assert.AreEqual(GitRevisionFixtureTrace.MaximumEvents, File.ReadAllLines(path).Length);
                Assert.IsTrue(new FileInfo(path).Length <= GitRevisionFixtureTrace.MaximumBytes);
                using (var trace = GitRevisionFixtureTrace.TryBeginAtRoot(second, new MacroGitRepository(second, "main")))
                {
                    Assert.IsNotNull(trace);
                    trace.Record(new { Kind = "TooLarge", Payload = new string('x', GitRevisionFixtureTrace.MaximumBytes) });
                    trace.Record(new { Kind = "AfterBound" });
                }
                path = Path.Combine(second, "diagnostic.jsonl");
                Assert.AreEqual(1, File.ReadAllLines(path).Length);
                Assert.IsTrue(new FileInfo(path).Length <= GitRevisionFixtureTrace.MaximumBytes);
            }
            finally { Directory.Delete(root, true); Directory.Delete(second, true); }
        }
    }
}
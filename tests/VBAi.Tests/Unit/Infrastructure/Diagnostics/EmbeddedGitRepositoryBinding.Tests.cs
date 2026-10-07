using Microsoft.VisualStudio.TestTools.UnitTesting;
using System;
using System.Collections.Generic;
using System.IO;
using System.Text;
using System.Web.Script.Serialization;

namespace VBAi.Tests.Integration
{
    /// <summary>Qualifies selected-path/marker filesystem guards; these synthetic files do not claim native Git or Office acceptance.</summary>
    [TestClass, TestCategory("Unit")]
    public sealed class EmbeddedGitRepositoryBindingTests
    {
        [TestMethod]
        public void ChildRecoveryMarkerRefusesSaveWhenDocumentRootHasNoMarker()
        {
            WithBoundFiles((cache, child, proof) =>
            {
                Assert.IsFalse(File.Exists(Path.Combine(cache, "codex-recovery")));
                File.WriteAllText(Path.Combine(child, "codex-recovery"), "synthetic-backup");
                Assert.ThrowsException<InvalidOperationException>(() => proof.RequireReadyForSave(cache, Remote, Branch));
                Assert.AreEqual("synthetic-backup", File.ReadAllText(Path.Combine(child, "codex-recovery")));
            });
        }

        [TestMethod]
        public void AbsentSelectedMarkerUsesTheProductionSelectedChildObserver()
        {
            WithBoundFiles((cache, child, proof) =>
            {
                var repository = proof.RequireReadyForSave(cache, Remote, Branch);
                Assert.AreEqual(Path.Combine(child, "codex-recovery"), repository.RecoveryFile);
                Assert.IsFalse(repository.RecoveryPending);
            });
        }

        [TestMethod]
        public void MetadataDenialPropagatesUnchangedAndObservesOnlyTheSelectedMarker()
        {
            WithBoundFiles((cache, child, proof) =>
            {
                var denied = new UnauthorizedAccessException("Synthetic selected-marker access denial");
                int calls = 0;
                var observed = Assert.ThrowsException<UnauthorizedAccessException>(() => proof.RequireReadyForSave(cache, Remote, Branch,
                    path => { ++calls; Assert.AreEqual(Path.Combine(child, "codex-recovery"), path); throw denied; }));
                Assert.AreSame(denied, observed); Assert.AreEqual(1, calls);
            });
        }

        [TestMethod]
        public void ExistingMarkerDirectoryAlsoRefusesSaveWithoutDeletingIt()
        {
            WithBoundFiles((cache, child, proof) =>
            {
                string marker = Path.Combine(child, "codex-recovery"); Directory.CreateDirectory(marker);
                Assert.ThrowsException<InvalidOperationException>(() => proof.RequireReadyForSave(cache, Remote, Branch));
                Assert.IsTrue(Directory.Exists(marker));
            });
        }

        [TestMethod]
        public void MissingSelectedRepositoryIsNeverInterpretedAsNoRecovery()
        {
            WithBoundFiles((cache, child, proof) =>
            {
                Directory.Move(child, child + ".retained");
                Assert.ThrowsException<FileNotFoundException>(() => proof.RequireReadyForSave(cache, Remote, Branch));
                Assert.ThrowsException<FileNotFoundException>(() => EmbeddedGitRepositoryBinding.Capture(cache, Remote, Branch));
            });
        }

        [TestMethod]
        public void BindingMustStillSelectTheValidatedRemoteAndBranch()
        {
            WithBoundFiles((cache, child, proof) =>
            {
                WriteBinding(cache, Remote, "different-branch");
                Assert.ThrowsException<InvalidOperationException>(() => proof.RequireReadyForSave(cache, Remote, Branch));
                Assert.ThrowsException<InvalidOperationException>(() => EmbeddedGitRepositoryBinding.Capture(cache, Remote, Branch));
            });
        }

        [TestMethod]
        [DataRow("config")]
        [DataRow("HEAD")]
        [DataRow("binding.json")]
        public void RepositoryIdentityBytesMustNotChangeBeforeSave(string file)
        {
            WithBoundFiles((cache, child, proof) =>
            {
                File.AppendAllText(Path.Combine(file == "binding.json" ? cache : child, file), "\n");
                Assert.ThrowsException<InvalidOperationException>(() => proof.RequireReadyForSave(cache, Remote, Branch));
            });
        }

        [TestMethod]
        public void AnotherBranchCannotReuseAnExistingProofOrWrongRoot()
        {
            WithBoundFiles((cache, child, proof) =>
            {
                Assert.AreNotEqual(child, EmbeddedGitRepositoryBinding.SelectedPath(cache, Remote, "other-branch"));
                Assert.ThrowsException<InvalidOperationException>(() => proof.RequireReadyForSave(cache, Remote, "other-branch"));
                Assert.ThrowsException<InvalidOperationException>(() => proof.RequireReadyForSave(child, Remote, Branch));
            });
        }

        private const string Remote = "https://github.com/blackcancer/vbai-qualification-20260929203712-7267b1e6.git";
        private const string Branch = "qualification-embedded-ui-synthetic-guard";

        private static void WithBoundFiles(Action<string, string, EmbeddedGitRepositoryBinding> action)
        {
            string cache = Path.Combine(Path.GetTempPath(), "VBAi-BoundGitGuard-" + Guid.NewGuid().ToString("N"));
            string child = EmbeddedGitRepositoryBinding.SelectedPath(cache, Remote, Branch);
            Directory.CreateDirectory(child);
            try
            {
                WriteBinding(cache, Remote, Branch);
                File.WriteAllText(Path.Combine(child, "config"), "[core]\n\tbare = true\n[remote \"origin\"]\n\turl = " + Remote + "\n", new UTF8Encoding(false));
                File.WriteAllText(Path.Combine(child, "HEAD"), "ref: refs/heads/" + Branch + "\n", new UTF8Encoding(false));
                action(cache, child, EmbeddedGitRepositoryBinding.Capture(cache, Remote, Branch));
            }
            finally { Directory.Delete(cache, true); }
        }

        private static void WriteBinding(string cache, string remote, string branch)
        {
            File.WriteAllText(Path.Combine(cache, "binding.json"), new JavaScriptSerializer().Serialize(
                new Dictionary<string, string> { ["Remote"] = remote, ["Branch"] = branch }), new UTF8Encoding(false));
        }
    }
}

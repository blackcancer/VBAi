using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Threading;
using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace VBAi.Tests.Integration
{
    internal sealed partial class ExcelVbeFixture
    {
        internal string EmbeddedProcessStartedUtc => Convert.ToString(startupEvidence["HostStartedUtc"]);

        /// <summary>Saves once after a terminal import, retaining an uncertain Save without retry or cleanup.</summary>
        internal void SaveEmbeddedImportedForm(EmbeddedGitScope scope, string remote, string branch, string selectedCommit,
            Action<bool> pending, Action<object> evidence)
        {
            Assert.AreEqual(ApartmentState.STA, Thread.CurrentThread.GetApartmentState());
            RequireEmbeddedProcess(scope);
            try
            {
                if (scope.SelectedRepository == null) throw new InvalidOperationException("The terminal UI did not attest its selected bare repository; Save is refused.");
                var repository = scope.SelectedRepository.RequireReadyForSave(scope.Cache, remote, branch);
                Assert.AreEqual(selectedCommit, repository.Resolve("refs/remotes/origin/selected"), "The selected authorized repository revision changed before Save.");
                evidence(new { Phase = "SelectedRepositoryBeforeSaveVerified", scope.SelectedRepository.RepositoryPath,
                    Remote = remote, Branch = branch, SelectedCommit = selectedCommit,
                    BindingSha256 = EmbeddedRawHash(Path.Combine(scope.Cache, "binding.json")),
                    ConfigSha256 = EmbeddedRawHash(Path.Combine(scope.SelectedRepository.RepositoryPath, "config")),
                    HeadSha256 = EmbeddedRawHash(Path.Combine(scope.SelectedRepository.RepositoryPath, "HEAD")), RecoveryPending = false });
            }
            catch { PreserveMonacoNativeOutcome(); throw; }
            object project = null;
            IntPtr identity = IntPtr.Zero;
            pending(true);
            try
            {
                project = OwnGitProjectRcw(); identity = Marshal.GetIUnknownForObject(project);
                Assert.AreEqual(embeddedGitProjectIdentity, identity, "Save must target the exact retained imported project.");
                Assert.AreEqual(scope.Path, Convert.ToString(((dynamic)workbook).FullName), true);
                Assert.IsFalse(Convert.ToBoolean(((dynamic)workbook).ReadOnly));
                ((dynamic)application).EnableEvents = false;
                ((dynamic)application).AutomationSecurity = 3;
                evidence(new { Phase = "PostImportSaveIntent", ProcessId, Path = scope.Path, ProjectIdentity = identity.ToInt64(),
                    SaveAttempts = 1, MacroExecutions = 0, OwnerSta = Thread.CurrentThread.ManagedThreadId });
                ((dynamic)workbook).Save();
                Assert.IsTrue(Convert.ToBoolean(((dynamic)workbook).Saved), "The one Save must report a saved workbook.");
                Assert.AreEqual(scope.Path, Convert.ToString(((dynamic)workbook).FullName), true);
                object afterProject = null; IntPtr afterIdentity = IntPtr.Zero;
                try
                {
                    afterProject = OwnGitProjectRcw(); afterIdentity = Marshal.GetIUnknownForObject(afterProject);
                    Assert.AreEqual(identity, afterIdentity, "The saved workbook must still own the imported native project.");
                }
                finally { if (afterIdentity != IntPtr.Zero) Marshal.Release(afterIdentity); Release(afterProject); }
                pending(false);
            }
            finally { if (identity != IntPtr.Zero) Marshal.Release(identity); Release(project); }
            VerifyEmbeddedSavedForm(scope, pending, evidence, false);
        }

        /// <summary>Attests the exact repository actually linked by the terminal owner UI; no initialization or remote command occurs.</summary>
        internal void AttestEmbeddedSelectedRepository(EmbeddedGitScope scope, string remote, string branch, string selectedCommit,
            Action<object> evidence)
        {
            var proof = EmbeddedGitRepositoryBinding.Capture(scope.Cache, remote, branch);
            var repository = new MacroGitRepository(proof.RepositoryPath, branch);
            Assert.AreEqual(selectedCommit, repository.Resolve("refs/remotes/origin/selected"));
            scope.SelectedRepository = proof;
            evidence(new { Phase = "SelectedRepositoryTerminalUiProof", proof.RepositoryPath, Remote = remote, Branch = branch,
                SelectedCommit = selectedCommit, BindingSha256 = EmbeddedRawHash(Path.Combine(scope.Cache, "binding.json")),
                ConfigSha256 = EmbeddedRawHash(Path.Combine(proof.RepositoryPath, "config")),
                HeadSha256 = EmbeddedRawHash(Path.Combine(proof.RepositoryPath, "HEAD")) });
        }

        /// <summary>Attests the actual installed bytes and host PID before reading a reopened synthetic project.</summary>
        internal void VerifyEmbeddedPersistenceCandidate(Guid expected, string hash, Action<bool> pending, Action<object> evidence)
        {
            var data = EmbeddedData(new Request { Command = "status" }, pending, evidence);
            RequireMonacoCandidate(expected, typeof(VbeSession).Module.ModuleVersionId, ProcessId, data);
            Assert.AreEqual(hash, EmbeddedRawHash(Convert.ToString(data["AssemblyPath"])), true);
            evidence(new { Phase = "FreshLoadedCandidateVerified", ProcessId, Root, StartedUtc = EmbeddedProcessStartedUtc,
                AssemblyMvid = expected.ToString("D"), AssemblySha256 = hash, OwnerSta = Thread.CurrentThread.ManagedThreadId });
        }

        /// <summary>Reads the actual saved sources/resources and native designer, without import, Save or font assignment.</summary>
        internal void VerifyEmbeddedReopenedForm(EmbeddedGitScope expected, Action<bool> pending, Action<object> evidence)
        {
            VerifyEmbeddedSavedForm(expected, pending, evidence, true);
        }

        private void VerifyEmbeddedSavedForm(EmbeddedGitScope expected, Action<bool> pending, Action<object> evidence, bool readOnly)
        {
            Assert.AreEqual(ApartmentState.STA, Thread.CurrentThread.GetApartmentState());
            object project = null, components = null, references = null;
            var code = new Dictionary<string, string>(StringComparer.Ordinal);
            var types = new Dictionary<string, int>(StringComparer.Ordinal);
            var rows = new List<string>();
            pending(true);
            try
            {
                Assert.AreEqual(expected.Path, Convert.ToString(((dynamic)workbook).FullName), true);
                Assert.AreEqual(readOnly, Convert.ToBoolean(((dynamic)workbook).ReadOnly));
                Assert.IsTrue(Convert.ToBoolean(((dynamic)workbook).Saved));
                project = OwnGitProjectRcw(); components = ((dynamic)project).VBComponents;
                Assert.AreEqual(2, Convert.ToInt32(((dynamic)project).Mode));
                Assert.AreEqual(0, Convert.ToInt32(((dynamic)project).Protection));
                foreach (object component in (dynamic)components)
                {
                    object module = null;
                    try
                    {
                        string name = Convert.ToString(((dynamic)component).Name);
                        module = ((dynamic)component).CodeModule;
                        int count = Convert.ToInt32(((dynamic)module).CountOfLines);
                        code.Add(name, count == 0 ? "" : Convert.ToString(((dynamic)module).Lines[1, count]));
                        types.Add(name, Convert.ToInt32(((dynamic)component).Type));
                    }
                    finally { Release(module); Release(component); }
                }
                references = ((dynamic)project).References;
                foreach (object reference in (dynamic)references)
                    try
                    {
                        Assert.IsFalse(Convert.ToBoolean(((dynamic)reference).IsBroken));
                        rows.Add(Convert.ToString(((dynamic)reference).GUID) + ":" + ((dynamic)reference).Major + ":" + ((dynamic)reference).Minor);
                    }
                    finally { Release(reference); }
                pending(false);
            }
            finally { Release(references); Release(components); Release(project); }
            CollectionAssert.AreEquivalent(expected.Code.Keys.ToArray(), code.Keys.ToArray());
            foreach (var item in expected.Code) Assert.AreEqual(item.Value, code[item.Key], "Persisted source: " + item.Key);
            CollectionAssert.AreEquivalent(expected.Types.Keys.ToArray(), types.Keys.ToArray());
            foreach (var item in expected.Types) Assert.AreEqual(item.Value, types[item.Key], "Persisted component type: " + item.Key);
            string revision = string.Join(";", rows.OrderBy(x => x, StringComparer.Ordinal).Select(x => x.ToUpperInvariant()));
            Assert.AreEqual(expected.References, revision);
            var actualScope = new EmbeddedGitScope { Path = expected.Path, Layout = expected.Layout, Code = code, Types = types, References = revision };
            var snapshot = ExportEmbeddedBaseline(actualScope, pending, evidence, "owner-bridge-persistence-" + Guid.NewGuid().ToString("N"));
            evidence(new { Phase = readOnly ? "FreshProcessSnapshotReadback" : "SavedSnapshotReadback",
                Exact = expected.Baseline.SameAs(snapshot), Changes = snapshot.Changes(expected.Baseline),
                Files = EmbeddedGitSnapshotOracle.Describe(snapshot) });
            Assert.IsTrue(expected.Baseline.SameAs(snapshot), "Saved/reopened complete sources and resources must remain exact.");
            pending(true);
            var layout = ReadGitLayout("EmbeddedForm", expected.Layout);
            var fonts = ReadGitLayoutFonts("EmbeddedForm", expected.Layout);
            pending(false);
            evidence(new { Phase = readOnly ? "FreshProcessNativeReadback" : "SavedNativeReadback", Properties = layout, Fonts = fonts });
            CollectionAssert.AreEquivalent(expected.NativeLayout.Keys.ToArray(), layout.Keys.ToArray());
            foreach (var item in expected.NativeLayout) Assert.AreEqual(item.Value, layout[item.Key], "Persisted property: " + item.Key);
            CollectionAssert.AreEquivalent(expected.NativeFonts.Keys.ToArray(), fonts.Keys.ToArray());
            foreach (var item in expected.NativeFonts) Assert.AreEqual(item.Value, fonts[item.Key], "Persisted font: " + item.Key);
        }
    }
}

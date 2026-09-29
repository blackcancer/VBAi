using System;
using System.IO;
using System.Linq;
using VBAi;
using VBAi.Tests.Infrastructure;
using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace VBAi.Tests.Unit
{
    [TestClass, TestCategory("Unit")]
    public sealed class LlmSettingsConcurrencyTests
    {
        private static void Isolated(Action<string> test)
        {
            using (var scope = new LlmBoundaryScope())
            {
                string previous = LlmSettings.StoragePathOverride;
                try
                {
                    string path = Path.Combine(scope.Root, "settings.json");
                    LlmSettings.StoragePathOverride = path;
                    test(path);
                }
                finally { LlmSettings.StoragePathOverride = previous; }
            }
        }

        [TestMethod]
        public void IndependentHostEditsMergeIncludingKeysWithinTheSameDictionary()
        {
            Isolated(path =>
            {
                new LlmSettings().Save();
                var first = LlmSettings.Load(); var second = LlmSettings.Load();
                first.NativeVbeDarkTheme = true;
                first.ProviderModels["one"] = "model-one";
                first.Save();
                second.VbeEditApproval = "ReadOnly";
                second.ProviderModels["two"] = "model-two";
                second.Save();
                var stored = LlmSettings.Load();
                Assert.IsTrue(stored.NativeVbeDarkTheme);
                Assert.AreEqual("ReadOnly", stored.VbeEditApproval);
                Assert.AreEqual("model-one", stored.ProviderModels["one"]);
                Assert.AreEqual("model-two", stored.ProviderModels["two"]);
                Assert.IsTrue(second.NativeVbeDarkTheme, "The saving instance must receive merged fields.");
                second.GitHubAccount = "local-fixture"; second.Save();
                Assert.IsTrue(LlmSettings.Load().NativeVbeDarkTheme);
                Assert.AreEqual(0, Directory.GetFiles(Path.GetDirectoryName(path), "*.tmp").Length);
            });
        }

        [TestMethod]
        public void ConflictingFieldOrDictionaryRemovalDoesNotOverwriteTheOtherHost()
        {
            Isolated(path =>
            {
                var seed = new LlmSettings(); seed.ProviderModels["one"] = "base"; seed.Save();
                var a = LlmSettings.Load(); var b = LlmSettings.Load();
                a.VbeEditApproval = "ReadOnly"; a.Save();
                b.VbeEditApproval = "AskEachTime";
                byte[] before = File.ReadAllBytes(path);
                StringAssert.Contains(Assert.ThrowsException<IOException>(() => b.Save()).Message, "VbeEditApproval");
                CollectionAssert.AreEqual(before, File.ReadAllBytes(path));
                a = LlmSettings.Load(); b = LlmSettings.Load();
                a.ProviderModels["one"] = "changed"; a.Save();
                b.ProviderModels.Remove("one");
                Assert.ThrowsException<IOException>(() => b.Save());
                Assert.AreEqual("changed", LlmSettings.Load().ProviderModels["one"]);
            });
        }

        [TestMethod]
        public void UnknownFieldsAndLegacyApprovalSurviveWhileCorruptStorageIsPreserved()
        {
            Isolated(path =>
            {
                Directory.CreateDirectory(Path.GetDirectoryName(path));
                File.WriteAllText(path, "{\"FutureSetting\":{\"value\":17}}");
                var settings = LlmSettings.Load(); settings.NativeVbeDarkTheme = true; settings.Save();
                StringAssert.Contains(File.ReadAllText(path), "FutureSetting");
                Assert.AreEqual("AskEachTime", LlmSettings.Load().VbeEditApproval);
                File.WriteAllText(path, "{broken");
                settings.NativeVbeDarkTheme = false;
                Assert.ThrowsException<ArgumentException>(() => settings.Save());
                Assert.AreEqual("{broken", File.ReadAllText(path));
            });
        }

        [TestMethod]
        public void FirstRunSnapshotsMergeAndLockedOrRemovedStorageFailsWithoutRewriting()
        {
            Isolated(path =>
            {
                var a = LlmSettings.Load(); var b = LlmSettings.Load();
                a.NativeVbeDarkTheme = true; a.Save();
                b.GitHubAccount = "fixture"; b.Save();
                Assert.IsTrue(LlmSettings.Load().NativeVbeDarkTheme);
                byte[] before = File.ReadAllBytes(path);
                using (var held = new FileStream(path + ".lock", FileMode.OpenOrCreate, FileAccess.ReadWrite, FileShare.None))
                {
                    b.GitHubAccount = "pending";
                    Assert.ThrowsException<IOException>(() => b.Save());
                }
                CollectionAssert.AreEqual(before, File.ReadAllBytes(path));
                File.Delete(path);
                Assert.ThrowsException<IOException>(() => b.Save());
                Assert.IsFalse(File.Exists(path));
            });
        }
    }
}

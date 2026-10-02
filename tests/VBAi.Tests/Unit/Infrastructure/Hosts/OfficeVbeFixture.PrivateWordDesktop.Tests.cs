using System;
using System.IO;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using VBAi.Tests.Integration;

namespace VBAi.Tests.Unit
{
    [TestClass, TestCategory("Unit")]
    public sealed class OfficeVbeFixturePrivateWordDesktopTests
    {
        private const string Executable = @"C:\Program Files\Microsoft Office\root\Office16\WINWORD.EXE";
        private static readonly string Hash = new string('A', 64);

        [TestMethod]
        public void PrivateWordLaunchRequiresExactReviewedExecutableAndMacroFreeSeedArguments()
        {
            int existsCalls = 0, hashCalls = 0;
            Assert.AreEqual(Path.GetFullPath(Executable), OfficeVbeFixture.RequireExactWordExecutable(
                Executable, Hash, path => { existsCalls++; return path == Executable; },
                path => { hashCalls++; return Hash.ToLowerInvariant(); }));
            Assert.AreEqual(1, existsCalls); Assert.AreEqual(1, hashCalls);
            CollectionAssert.AreEqual(new[] { "/n", "/q", "/m", @"C:\Owned\NativeObjectModelSeed.docx" },
                OfficeVbeFixture.WordPrivateArguments(@"C:\Owned\NativeObjectModelSeed.docx"));
        }

        [TestMethod]
        public void PrivateWordLaunchRejectsWrongPathHashOrSeedBeforeAnyLauncherCall()
        {
            foreach (string path in new[] { "WINWORD.EXE", @"C:\Owned\Other.exe", @"C:\Owned\WINWORD.EXE\child", "" })
                Assert.ThrowsException<InvalidOperationException>(() => OfficeVbeFixture.RequireExactWordExecutable(
                    path, Hash, unused => { Assert.Fail("No file lookup after an invalid path."); return false; }, unused => Hash));
            foreach (string hash in new[] { null, "", new string('G', 64), new string('A', 63) })
                Assert.ThrowsException<InvalidOperationException>(() => OfficeVbeFixture.RequireExactWordExecutable(
                    Executable, hash, unused => { Assert.Fail("No file lookup after an invalid hash."); return false; }, unused => Hash));
            Assert.ThrowsException<InvalidOperationException>(() => OfficeVbeFixture.RequireExactWordExecutable(
                Executable, Hash, unused => false, unused => { Assert.Fail("No hash after a missing executable."); return Hash; }));
            Assert.ThrowsException<InvalidOperationException>(() => OfficeVbeFixture.RequireExactWordExecutable(
                Executable, Hash, unused => true, unused => new string('B', 64)));
            foreach (string seed in new[] { "seed.docx", @"C:\Owned\Seed.docm", "" })
                Assert.ThrowsException<InvalidOperationException>(() => OfficeVbeFixture.WordPrivateArguments(seed));
        }
    }
}

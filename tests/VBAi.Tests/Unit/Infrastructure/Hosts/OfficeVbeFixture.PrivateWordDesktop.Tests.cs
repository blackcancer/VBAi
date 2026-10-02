using System;
using System.IO;
using System.IO.Packaging;
using System.Linq;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Text;
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
        public void DisabledOfficeOptInIsInconclusiveBeforeAnyPrivateDesktopObservation()
        {
            string name = "VBAiTests_" + Guid.NewGuid().ToString("N");
            int observations = 0;
            Action<string> observe = unused => observations++;
            foreach (string kind in new[] { "Word", "Access", "PowerPoint", "Publisher" })
                foreach (string optIn in new[] { null, "", "0", "true", "2" })
                    Assert.ThrowsException<AssertInconclusiveException>(() =>
                        OfficeVbeFixture.RequireEnabledOfficeDesktop(kind, optIn, name, null, observe));
            Assert.AreEqual(0, observations,
                "Disabled Office tests must skip before desktop comparison, native inventory or activation.");
        }

        [TestMethod]
        public void EnabledOfficeOptInValidatesPrivateWordDesktopBeforeAnyHostInventory()
        {
            string name = "VBAiTests_" + Guid.NewGuid().ToString("N");
            int observations = 0;
            Action<string> observe = actual => { Assert.AreEqual(name, actual); observations++; };
            Assert.IsNull(OfficeVbeFixture.RequireEnabledOfficeDesktop("Word", "1", null, null, observe));
            Assert.AreEqual(0, observations);
            Assert.AreEqual(name, OfficeVbeFixture.RequireEnabledOfficeDesktop("Word", "1", name, name, observe));
            Assert.AreEqual(1, observations);
            foreach (string configured in new[] { null, "", "wrong" })
                Assert.ThrowsException<InvalidOperationException>(() =>
                    OfficeVbeFixture.RequireEnabledOfficeDesktop("Word", "1", name, configured, observe));
            Assert.ThrowsException<InvalidOperationException>(() =>
                OfficeVbeFixture.RequireEnabledOfficeDesktop("Access", "1", name, name, observe));
            Assert.AreEqual(1, observations, "Pair and host refusals must precede desktop/native access.");
        }

        [TestMethod]
        public void PrivateWordDesktopRequiresWorkerAndTestNamesToMatchBeforeHostActivation()
        {
            string name = "VBAiTests_" + Guid.NewGuid().ToString("N");
            int observations = 0;
            Action<string> observe = actual => { Assert.AreEqual(name, actual); observations++; };
            Assert.IsNull(OfficeVbeFixture.RequirePrivateWordDesktop("Word", null, null, observe));
            Assert.AreEqual(0, observations);
            Assert.AreEqual(name, OfficeVbeFixture.RequirePrivateWordDesktop("Word", name, name, observe));
            Assert.AreEqual(1, observations);
            Assert.AreEqual(name, OfficeVbeFixture.RequirePrivateWordDesktop("Word", null, name, observe));
            Assert.AreEqual(2, observations);
            foreach (string configured in new[] { null, "", name.ToLowerInvariant(), "another" })
                Assert.ThrowsException<InvalidOperationException>(() =>
                    OfficeVbeFixture.RequirePrivateWordDesktop("Word", name, configured, observe));
            Assert.ThrowsException<InvalidOperationException>(() =>
                OfficeVbeFixture.RequirePrivateWordDesktop("Excel", name, name, observe));
            Assert.ThrowsException<InvalidOperationException>(() =>
                OfficeVbeFixture.RequirePrivateWordDesktop("Access", null, name, observe));
            Assert.AreEqual(2, observations, "Refusals must precede desktop and host access.");
            Assert.ThrowsException<InvalidOperationException>(() =>
                OfficeVbeFixture.RequirePrivateWordDesktop("Word", name, name,
                    unused => throw new InvalidOperationException("Inactive desktop not established.")));
        }

        [TestMethod]
        public void NativeWordObjectModelPInvokeMarshalsRequestedIDispatchAsInterface()
        {
            MethodInfo method = typeof(OfficeVbeFixture).GetMethod("NativeAccessibleObjectFromWindow",
                BindingFlags.NonPublic | BindingFlags.Static);
            Assert.IsNotNull(method);
            var parameters = method.GetParameters();
            Assert.AreEqual(4, parameters.Length);
            Assert.AreEqual(typeof(Guid).MakeByRefType(), parameters[2].ParameterType);
            Assert.AreEqual(typeof(object).MakeByRefType(), parameters[3].ParameterType);
            var marshaling = parameters[3].GetCustomAttribute<MarshalAsAttribute>();
            Assert.IsNotNull(marshaling);
            Assert.AreEqual(UnmanagedType.Interface, marshaling.Value);
        }

        [TestMethod]
        public void MacroFreeWordSeedHasOneDocumentPartAndExactRootRelationshipAndRefusesOverwrite()
        {
            string directory = Path.Combine(Path.GetTempPath(), "VBAi-PrivateWordSeed-" + Guid.NewGuid().ToString("N"));
            Directory.CreateDirectory(directory);
            string path = Path.Combine(directory, "NativeObjectModelSeed.docx");
            MethodInfo create = typeof(OfficeVbeFixture).GetMethod("WriteMacroFreeSeed", BindingFlags.NonPublic | BindingFlags.Static);
            Assert.IsNotNull(create);
            try
            {
                create.Invoke(null, new object[] { path });
                byte[] original = File.ReadAllBytes(path);
                using (var package = Package.Open(path, FileMode.Open, FileAccess.Read))
                {
                    var parts = package.GetParts().Where(part => !PackUriHelper.IsRelationshipPartUri(part.Uri)).ToArray();
                    Assert.AreEqual(1, parts.Length);
                    Assert.AreEqual("/word/document.xml", parts[0].Uri.ToString());
                    Assert.AreEqual("application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml", parts[0].ContentType);
                    var relations = package.GetRelationships().ToArray();
                    Assert.AreEqual(1, relations.Length);
                    Assert.AreEqual("http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument", relations[0].RelationshipType);
                    Assert.AreEqual(TargetMode.Internal, relations[0].TargetMode);
                    Assert.AreEqual(parts[0].Uri, relations[0].TargetUri);
                    using (var reader = new StreamReader(parts[0].GetStream(), Encoding.UTF8))
                    {
                        string xml = reader.ReadToEnd();
                        StringAssert.Contains(xml, "<w:document");
                        StringAssert.Contains(xml, "<w:body><w:p/></w:body>");
                    }
                }
                var exception = Assert.ThrowsException<TargetInvocationException>(() => create.Invoke(null, new object[] { path }));
                Assert.IsInstanceOfType(exception.InnerException, typeof(AssertFailedException));
                CollectionAssert.AreEqual(original, File.ReadAllBytes(path));
            }
            finally { Directory.Delete(directory, true); }
        }

        [TestMethod]
        public void PrivateWordLaunchRequiresExactReviewedExecutableAndMacroFreeSeedArguments()
        {
            int existsCalls = 0, hashCalls = 0;
            Assert.AreEqual(Path.GetFullPath(Executable), OfficeVbeFixture.RequireExactWordExecutable(
                Executable, Hash, path => { existsCalls++; return path == Executable; },
                path => { hashCalls++; return Hash.ToLowerInvariant(); }));
            Assert.AreEqual(1, existsCalls); Assert.AreEqual(1, hashCalls);
            CollectionAssert.AreEqual(new[] { "/a", @"C:\Owned\NativeObjectModelSeed.docx" },
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

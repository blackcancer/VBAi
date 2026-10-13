using Microsoft.VisualStudio.TestTools.UnitTesting;
using System;
using System.Reflection;
using System.IO;
using System.IO.Packaging;
using System.Linq;
using VBAi.Tests.Integration;

namespace VBAi.Tests.Unit
{
    [TestClass, TestCategory("Unit"), DoNotParallelize]
    public sealed class OfficeVbeFixtureWordLifecycleControlTests
    {
        [TestMethod]
        public void SourceHashReadsAlongsideWordWriteHandleWithoutChangingBytesAndDetectsLaterWrites()
        {
            WithRoot(root =>
            {
                string path = Path.Combine(root, "open-source.docm");
                byte[] original = { 1, 2, 3, 4 };
                File.WriteAllBytes(path, original);
                string expected = OfficeVbeFixture.ReadWordLifecycleSourceHash(path);
                using (var writer = new FileStream(path, FileMode.Open, FileAccess.ReadWrite, FileShare.Read))
                {
                    Assert.AreEqual(expected, OfficeVbeFixture.ReadWordLifecycleSourceHash(path));
                    Assert.AreEqual(expected, OfficeVbeFixture.RequireUnchangedOpenWordSeed(path, expected));
                    writer.WriteByte(9); writer.Flush(true);
                    Assert.AreNotEqual(expected, OfficeVbeFixture.ReadWordLifecycleSourceHash(path));
                    Assert.ThrowsException<InvalidOperationException>(() => OfficeVbeFixture.RequireUnchangedOpenWordSeed(path, expected));
                }
                File.WriteAllBytes(path, original);
                using (var denied = new FileStream(path, FileMode.Open, FileAccess.ReadWrite, FileShare.None))
                    Assert.ThrowsException<IOException>(() => OfficeVbeFixture.ReadWordLifecycleSourceHash(path));
                CollectionAssert.AreEqual(original, File.ReadAllBytes(path));
            });
        }

        [DataTestMethod, DataRow(null), DataRow(""), DataRow("relative.docm")]
        public void SourceHashRefusesMissingOrRelativeIdentityBeforeFileAccess(string path)
        {
            Assert.ThrowsException<ArgumentException>(() => OfficeVbeFixture.ReadWordLifecycleSourceHash(path));
        }

        [TestMethod]
        public void OwnedTemplatePackageContainsNoMacrosOrExternalDependenciesAndPreservesAnExistingFile()
        {
            WithRoot(root =>
            {
                string path = Path.Combine(root, "WordLifecycleTemplate.dotx");
                OfficeVbeFixture.WriteWordLifecycleTemplateSeed(path);
                byte[] before = File.ReadAllBytes(path);
                using (var package = Package.Open(path, FileMode.Open, FileAccess.Read))
                {
                    var parts = package.GetParts().Where(part => !PackUriHelper.IsRelationshipPartUri(part.Uri)).ToArray();
                    Assert.AreEqual(1, parts.Length);
                    Assert.AreEqual("application/vnd.openxmlformats-officedocument.wordprocessingml.template.main+xml", parts[0].ContentType);
                    Assert.AreEqual(0, parts[0].GetRelationships().Count());
                    var link = package.GetRelationships().Single();
                    Assert.AreEqual(TargetMode.Internal, link.TargetMode);
                    Assert.AreEqual(parts[0].Uri, link.TargetUri);
                }
                Assert.ThrowsException<AssertFailedException>(() => OfficeVbeFixture.WriteWordLifecycleTemplateSeed(path));
                CollectionAssert.AreEqual(before, File.ReadAllBytes(path));
            });
        }

        [TestMethod]
        public void MacroFreeControlReferencesExactlyItsOwnedTemplateIncludingUnicodePaths()
        {
            WithRoot(root =>
            {
                var fixture = (OfficeVbeFixture)Activator.CreateInstance(typeof(OfficeVbeFixture), true);
                typeof(OfficeVbeFixture).GetField("<Root>k__BackingField", BindingFlags.Instance | BindingFlags.NonPublic).SetValue(fixture, root);
                string path = Path.Combine(root, "control.docx"), template = Path.Combine(root, "WordLifecycleTemplate.dotx");
                OfficeVbeFixture.WriteWordLifecycleTemplateSeed(template);
                var writer = typeof(OfficeVbeFixture).GetMethod("WriteWordLifecycleControlSeed", BindingFlags.Instance | BindingFlags.NonPublic);
                writer.Invoke(fixture, new object[] { path });
                using (var package = Package.Open(path, FileMode.Open, FileAccess.Read))
                {
                    var parts = package.GetParts().Where(part => !PackUriHelper.IsRelationshipPartUri(part.Uri)).ToArray();
                    Assert.AreEqual(2, parts.Length);
                    Assert.IsFalse(parts.Any(part => part.Uri.ToString().IndexOf("vba", StringComparison.OrdinalIgnoreCase) >= 0));
                    var document = package.GetPart(new Uri("/word/document.xml", UriKind.Relative));
                    var settings = package.GetPart(new Uri("/word/settings.xml", UriKind.Relative));
                    Assert.AreEqual(settings.Uri, document.GetRelationships().Single().TargetUri);
                    var attached = settings.GetRelationships().Single();
                    Assert.AreEqual(TargetMode.External, attached.TargetMode);
                    Assert.AreEqual(new Uri(template), attached.TargetUri);
                    Assert.AreEqual("rTemplate", attached.Id);
                    using (var text = new StreamReader(settings.GetStream())) StringAssert.Contains(text.ReadToEnd(), "r:id=\"rTemplate\"");
                }
                byte[] before = File.ReadAllBytes(path);
                var failure = Assert.ThrowsException<TargetInvocationException>(() => writer.Invoke(fixture, new object[] { path }));
                Assert.IsInstanceOfType(failure.InnerException, typeof(AssertFailedException));
                CollectionAssert.AreEqual(before, File.ReadAllBytes(path));
            });
        }

        private static void WithRoot(Action<string> test)
        {
            string root = Path.Combine(Path.GetTempPath(), "VBAi-WordLifecycle-\u00E9-" + Guid.NewGuid().ToString("N"));
            Directory.CreateDirectory(root);
            try { test(root); }
            finally
            {
                foreach (string name in new[] { "open-source.docm", "control.docx", "WordLifecycleTemplate.dotx" })
                    File.Delete(Path.Combine(root, name));
                Directory.Delete(root);
            }
        }

        [DataTestMethod, DataRow(null), DataRow("0"), DataRow("1")]
        public void CalibrationIsDisabledByDefaultAndCannotAdoptAnUnownedWordInstance(string value)
        {
            const string variable = "VBAi_RUN_WORD_LIFECYCLE_CALIBRATION";
            string previous = Environment.GetEnvironmentVariable(variable);
            try
            {
                Environment.SetEnvironmentVariable(variable, value);
                var fixture = (OfficeVbeFixture)Activator.CreateInstance(typeof(OfficeVbeFixture), true);
                typeof(OfficeVbeFixture).GetField("<Kind>k__BackingField", BindingFlags.Instance | BindingFlags.NonPublic).SetValue(fixture, "Word");
                if (value == "1") Assert.ThrowsException<AssertFailedException>(() => fixture.CalibrateWordLifecycle());
                else Assert.IsNull(fixture.CalibrateWordLifecycle());
            }
            finally { Environment.SetEnvironmentVariable(variable, previous); }
        }
    }
}

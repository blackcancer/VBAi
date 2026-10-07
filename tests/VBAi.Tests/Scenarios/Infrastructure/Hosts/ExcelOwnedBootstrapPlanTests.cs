using Microsoft.VisualStudio.TestTools.UnitTesting;
using System;
using System.Diagnostics;
using System.IO;
using System.IO.Packaging;
using System.Linq;
using System.Xml;
using VBAi.Tests.Integration;

namespace VBAi.Tests.Unit
{
    /// <summary>Checks explicit-launch preflight without starting Excel, COM, or a UI thread.</summary>
    [TestClass, TestCategory("Unit")]
    public sealed class ExcelOwnedBootstrapPlanTests
    {
        [TestMethod]
        public void SeedContainsOnlyEmptyWorkbookAndWorksheetWithInternalRelationships()
        {
            InScratch(root =>
            {
                string seed = Path.Combine(root, "Seed.xlsx");
                ExcelOwnedBootstrapPlan.WriteSeed(seed);
                using (var package = Package.Open(seed, FileMode.Open, FileAccess.Read))
                {
                    var parts = package.GetParts().Where(part => !part.Uri.ToString().Contains("/_rels/")).ToArray();
                    CollectionAssert.AreEquivalent(new[] { "/xl/workbook.xml", "/xl/worksheets/sheet1.xml" }, parts.Select(part => part.Uri.ToString()).ToArray());
                    Assert.AreEqual(1, package.GetRelationships().Count());
                    Assert.IsTrue(package.GetRelationships().All(relation => relation.TargetMode == TargetMode.Internal));
                    foreach (var part in parts)
                    {
                        Assert.IsTrue(part.GetRelationships().All(relation => relation.TargetMode == TargetMode.Internal));
                        using (var input = part.GetStream(FileMode.Open, FileAccess.Read))
                        {
                            var xml = new XmlDocument { XmlResolver = null };
                            xml.Load(input);
                            Assert.AreEqual(0, xml.SelectNodes("//*[local-name()='f' or local-name()='c' or local-name()='externalReference']").Count);
                        }
                        Assert.IsFalse(part.ContentType.Contains("macroEnabled"));
                        Assert.IsFalse(part.ContentType.Contains("vbaProject"));
                    }
                    var workbook = package.GetPart(new Uri("/xl/workbook.xml", UriKind.Relative));
                    var worksheet = workbook.GetRelationships().Single();
                    Assert.AreEqual(new Uri("/xl/worksheets/sheet1.xml", UriKind.Relative), PackUriHelper.ResolvePartUri(workbook.Uri, worksheet.TargetUri));
                }
                byte[] original = File.ReadAllBytes(seed);
                Assert.ThrowsException<IOException>(() => ExcelOwnedBootstrapPlan.WriteSeed(seed));
                CollectionAssert.AreEqual(original, File.ReadAllBytes(seed), "Seed creation must never overwrite an existing workbook.");
            });
        }

        [TestMethod]
        public void LaunchUsesSeparateProcessAutomationOwnedSeedAndExplicitTraceWithoutChangingParentEnvironment()
        {
            InScratch(root =>
            {
                string seed = Path.Combine(root, "Owned seed.xlsx"), trace = Path.Combine(root, "phases.jsonl");
                ExcelOwnedBootstrapPlan.WriteSeed(seed);
                string inherited = Environment.GetEnvironmentVariable(VbeInspectionTrace.EnvironmentName);
                var info = ExcelOwnedBootstrapPlan.CreateStartInfo(Path.Combine(root, "EXCEL.EXE"), seed, trace);
                Assert.AreEqual("/x /automation \"" + seed + "\"", info.Arguments);
                Assert.IsFalse(info.UseShellExecute);
                Assert.AreEqual(root, info.WorkingDirectory);
                Assert.AreEqual(ProcessWindowStyle.Hidden, info.WindowStyle);
                Assert.AreEqual(trace, info.EnvironmentVariables[VbeInspectionTrace.EnvironmentName]);
                Assert.AreEqual(inherited, Environment.GetEnvironmentVariable(VbeInspectionTrace.EnvironmentName));
            });
        }

        [DataTestMethod]
        [DataRow(null)]
        [DataRow("relative.xlsx")]
        [DataRow("C:relative.xlsx")]
        [DataRow("\\\\server\\share\\seed.xlsx")]
        [DataRow("C:\\seed.xlsx:stream")]
        public void NonlocalOrAmbiguousPathsAreRejectedBeforeLaunch(string path)
        {
            Assert.ThrowsException<ArgumentException>(() => ExcelOwnedBootstrapPlan.RequireLocalAbsolutePath(path));
        }

        [TestMethod]
        public void MissingSeedWrongFormatAndMissingTraceParentRefuseLaunchPlan()
        {
            InScratch(root =>
            {
                string exe = Path.Combine(root, "EXCEL.EXE"), trace = Path.Combine(root, "phases.jsonl");
                Assert.ThrowsException<ArgumentException>(() => ExcelOwnedBootstrapPlan.CreateStartInfo(exe, Path.Combine(root, "missing.xlsx"), trace));
                string seed = Path.Combine(root, "seed.xlsm");
                File.WriteAllText(seed, "Not a workbook");
                Assert.ThrowsException<ArgumentException>(() => ExcelOwnedBootstrapPlan.CreateStartInfo(exe, seed, trace));
                seed = Path.Combine(root, "seed.xlsx");
                ExcelOwnedBootstrapPlan.WriteSeed(seed);
                Assert.ThrowsException<ArgumentException>(() => ExcelOwnedBootstrapPlan.CreateStartInfo(exe, seed, Path.Combine(root, "absent", "phases.jsonl")));
            });
        }

        [TestMethod]
        public void ExistingExcelAndUnverifiedProcessIdentitiesCannotAuthorizeAttachment()
        {
            ExcelOwnedBootstrapPlan.RequireFreshLaunch(new int[0]);
            Assert.ThrowsException<InvalidOperationException>(() => ExcelOwnedBootstrapPlan.RequireFreshLaunch(null));
            Assert.ThrowsException<InvalidOperationException>(() => ExcelOwnedBootstrapPlan.RequireFreshLaunch(new[] { 42 }));
            string exe = Path.Combine(Path.GetTempPath(), "EXCEL.EXE"), start = "2026-09-30T20:00:00.0000000Z";
            ExcelOwnedBootstrapPlan.VerifyAttachedIdentity(42, exe, start, 42, exe.ToUpperInvariant(), start);
            Assert.ThrowsException<InvalidOperationException>(() => ExcelOwnedBootstrapPlan.VerifyAttachedIdentity(0, exe, start, 0, exe, start));
            Assert.ThrowsException<InvalidOperationException>(() => ExcelOwnedBootstrapPlan.VerifyAttachedIdentity(42, exe, start, 43, exe, start));
            Assert.ThrowsException<InvalidOperationException>(() => ExcelOwnedBootstrapPlan.VerifyAttachedIdentity(42, exe, start, 42, Path.Combine(Path.GetTempPath(), "other", "EXCEL.EXE"), start));
            Assert.ThrowsException<InvalidOperationException>(() => ExcelOwnedBootstrapPlan.VerifyAttachedIdentity(42, exe, start, 42, exe, start + "1"));
            Assert.ThrowsException<InvalidOperationException>(() => ExcelOwnedBootstrapPlan.VerifyAttachedIdentity(42, exe, "", 42, exe, ""));
        }

        [TestMethod]
        public void WrongNameMissingImageInvalidPeAndNonX64ExecutablesAreRejectedWithoutExecution()
        {
            InScratch(root =>
            {
                string exe = Path.Combine(root, "EXCEL.EXE");
                Assert.ThrowsException<InvalidOperationException>(() => ExcelOwnedBootstrapPlan.ValidateExecutable(exe));
                File.WriteAllText(Path.Combine(root, "other.exe"), "Synthetic");
                Assert.ThrowsException<InvalidOperationException>(() => ExcelOwnedBootstrapPlan.ValidateExecutable(Path.Combine(root, "other.exe")));
                File.WriteAllText(exe, "Not a PE image");
                Assert.ThrowsException<InvalidOperationException>(() => ExcelOwnedBootstrapPlan.ValidateExecutable(exe));
                WritePe(exe, -1, 0x00004550, 0x8664);
                Assert.ThrowsException<InvalidOperationException>(() => ExcelOwnedBootstrapPlan.ValidateExecutable(exe));
                WritePe(exe, 64, 0, 0x8664);
                Assert.ThrowsException<InvalidOperationException>(() => ExcelOwnedBootstrapPlan.ValidateExecutable(exe));
                WritePe(exe, 64, 0x00004550, 0x014C);
                Assert.ThrowsException<InvalidOperationException>(() => ExcelOwnedBootstrapPlan.ValidateExecutable(exe));
                WritePe(exe, 64, 0x00004550, 0x8664);
                Assert.ThrowsException<InvalidOperationException>(() => ExcelOwnedBootstrapPlan.ValidateExecutable(exe), "A header alone cannot impersonate the installed Excel image's version identity.");
            });
        }

        private static void WritePe(string path, int pe, uint signature, ushort machine)
        {
            using (var file = new FileStream(path, FileMode.Create, FileAccess.Write))
            using (var writer = new BinaryWriter(file))
            {
                file.SetLength(128);
                writer.Write((ushort)0x5A4D);
                file.Position = 0x3C;
                writer.Write(pe);
                file.Position = 64;
                writer.Write(signature);
                writer.Write(machine);
            }
        }

        private static void InScratch(Action<string> test)
        {
            string root = Path.Combine(Path.GetTempPath(), "VBAi-OwnedBootstrap-" + Guid.NewGuid().ToString("N"));
            Directory.CreateDirectory(root);
            try { test(root); }
            finally { Directory.Delete(root, true); }
        }
    }
}

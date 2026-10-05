using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Reflection;
using System.Security.Cryptography;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using VBAi.Tests.Integration;

namespace VBAi.Tests.Unit
{
    /// <summary>Mirrors independent native-product pins with the actual status field names and a separate local test-reference copy.</summary>
    [TestClass, TestCategory("Unit"), DoNotParallelize]
    public sealed class ExcelVbeFixtureAddInShutdownProductTests
    {
        private static readonly Guid Mvid = new Guid("65b26db8-f7fb-4ba3-8494-f9196a95fb83");
        private static readonly string NativeProduct = Path.Combine(Path.GetTempPath(), "VBAi-product-pin", "VBAi", "Debug", "net48", "VBAi.dll");
        private static readonly string LocalProduct = Path.Combine(Path.GetTempPath(), "VBAi-product-pin", "VBAi.Tests", "Debug", "net48", "VBAi.dll");
        private static readonly string Hash = new string('A', 64);
        private static Dictionary<string, object> Status()
        {
            // These are the producer's status Data fields, also observed at Ready on
            // candidate 2bcabe61 / owned PID 89492. The local test copy has another path.
            return new Dictionary<string, object> {
                ["Version"] = "0.1.0", ["Connected"] = true, ["AssemblyPath"] = NativeProduct,
                ["AssemblyModuleVersionId"] = Mvid.ToString("D"), ["HostProcessId"] = 89492, ["ProcessBitness"] = 64
            };
        }

        [DataTestMethod, DataRow(false), DataRow(true)]
        public void DifferentProductAndTestCopyPathsRequireBothExactHashesAndTheIndependentNativePath(bool caseVariants)
        {
            var status = Status();
            if (caseVariants) { status["AssemblyPath"] = NativeProduct.ToUpperInvariant(); status["AssemblyModuleVersionId"] = Mvid.ToString("D").ToUpperInvariant(); }
            var reads = new List<string>();
            string result = ExcelVbeFixture.RequireAddInShutdownProduct(status, NativeProduct, Mvid.ToString("D"),
                caseVariants ? Hash.ToLowerInvariant() : Hash, LocalProduct, Mvid, path => { reads.Add(path); return Hash; });
            Assert.AreEqual(NativeProduct, result);
            CollectionAssert.AreEqual(new[] { NativeProduct, LocalProduct }, reads);
        }

        [TestMethod]
        public void AReferenceLoadedFromTheNativeProductPathStillRequiresTheSameIndependentPins()
        {
            int reads = 0;
            Assert.AreEqual(NativeProduct, ExcelVbeFixture.RequireAddInShutdownProduct(Status(), NativeProduct, Mvid.ToString("D"), Hash,
                NativeProduct, Mvid, path => { Assert.AreEqual(NativeProduct, path); reads++; return Hash; }));
            Assert.AreEqual(1, reads);
        }

        [DataTestMethod]
        [DataRow("missing product")] [DataRow("empty product")] [DataRow("relative product")] [DataRow("alternate stream")]
        [DataRow("unc product")] [DataRow("noncanonical product")]
        [DataRow("missing mvid")] [DataRow("invalid mvid")] [DataRow("foreign pinned mvid")]
        [DataRow("missing hash")] [DataRow("short hash")] [DataRow("nonhex hash")]
        [DataRow("missing status")] [DataRow("missing native path")] [DataRow("nonstr native path")] [DataRow("foreign native path")]
        [DataRow("native reports local copy")] [DataRow("missing native mvid")] [DataRow("nonstr native mvid")] [DataRow("foreign native mvid")]
        [DataRow("missing local path")] [DataRow("relative local path")] [DataRow("noncanonical local path")] [DataRow("foreign local mvid")]
        [DataRow("foreign native bytes")] [DataRow("foreign local bytes")]
        public void MissingInvalidOrForeignIndependentPinsCannotBeReplacedByAnEchoOfLoadedStatus(string kind)
        {
            Dictionary<string, object> status = Status();
            string product = NativeProduct, local = LocalProduct, mvid = Mvid.ToString("D"), hash = Hash;
            Guid localMvid = Mvid;
            if (kind == "missing product") product = null;
            if (kind == "empty product") product = "";
            if (kind == "relative product") product = "VBAi.dll";
            if (kind == "alternate stream") product += ":foreign";
            if (kind == "unc product") product = @"\\localhost\C$\VBAi.dll";
            if (kind == "noncanonical product") product = Path.Combine(Path.GetDirectoryName(NativeProduct), "nested", "..", "VBAi.dll");
            if (kind == "missing mvid") mvid = null;
            if (kind == "invalid mvid") mvid = "invalid";
            if (kind == "foreign pinned mvid") mvid = Guid.NewGuid().ToString("D");
            if (kind == "missing hash") hash = null;
            if (kind == "short hash") hash = "A";
            if (kind == "nonhex hash") hash = new string('Z', 64);
            if (kind == "missing status") status = null;
            if (kind == "missing native path") status.Remove("AssemblyPath");
            if (kind == "nonstr native path") status["AssemblyPath"] = 71;
            if (kind == "foreign native path") status["AssemblyPath"] = Path.Combine(Path.GetTempPath(), "Foreign.dll");
            if (kind == "native reports local copy") status["AssemblyPath"] = LocalProduct;
            if (kind == "missing native mvid") status.Remove("AssemblyModuleVersionId");
            if (kind == "nonstr native mvid") status["AssemblyModuleVersionId"] = 71;
            if (kind == "foreign native mvid") status["AssemblyModuleVersionId"] = Guid.NewGuid().ToString("D");
            if (kind == "missing local path") local = null;
            if (kind == "relative local path") local = "VBAi.dll";
            if (kind == "noncanonical local path") local = Path.Combine(Path.GetDirectoryName(LocalProduct), "nested", "..", "VBAi.dll");
            if (kind == "foreign local mvid") localMvid = Guid.NewGuid();
            var reads = new List<string>(); Exception failure = null;
            try
            {
                ExcelVbeFixture.RequireAddInShutdownProduct(status, product, mvid, hash, local, localMvid, path =>
                {
                    reads.Add(path);
                    return kind == "foreign native bytes" && path == NativeProduct || kind == "foreign local bytes" && path == LocalProduct ? new string('B', 64) : Hash;
                });
            }
            catch (Exception error) { failure = error; }
            Assert.IsNotNull(failure);
            if (kind == "foreign native bytes" || kind == "foreign local bytes") CollectionAssert.AreEqual(new[] { NativeProduct, LocalProduct }, reads);
            else Assert.AreEqual(0, reads.Count, "Invalid syntax, status or MVID must refuse before file hash reads.");
        }

        [DataTestMethod, DataRow(false), DataRow(true)]
        public void HashReadFailureRemainsTheSameExceptionAndIsNeverRetried(bool localFailure)
        {
            var primary = new IOException("product hash read failed"); var reads = new List<string>();
            var actual = Assert.ThrowsException<IOException>(() => ExcelVbeFixture.RequireAddInShutdownProduct(Status(), NativeProduct,
                Mvid.ToString("D"), Hash, LocalProduct, Mvid, path =>
                {
                    reads.Add(path);
                    if (path == (localFailure ? LocalProduct : NativeProduct)) throw primary;
                    return Hash;
                }));
            Assert.AreSame(primary, actual);
            CollectionAssert.AreEqual(localFailure ? new[] { NativeProduct, LocalProduct } : new[] { NativeProduct }, reads);
        }

        [DataTestMethod]
        [DataRow("distinct copy")] [DataRow("case status")] [DataRow("disabled")]
        [DataRow("missing product")] [DataRow("foreign native status")]
        public void ActualFixtureArmingUsesIndependentCandidatePinsAndTheNativeStatusContract(string kind)
        {
            const BindingFlags fields = BindingFlags.Instance | BindingFlags.NonPublic;
            string[] names = { AddInShutdownDiagnostic.EnvironmentName, ExcelVbeFixture.AddInShutdownProductEnvironmentName,
                "VBAi_TEST_EMBEDDED_GIT_MVID", "VBAi_TEST_EMBEDDED_GIT_SHA256" };
            var prior = new Dictionary<string, string>();
            foreach (string name in names) prior[name] = Environment.GetEnvironmentVariable(name);
            try
            {
                using (var directory = new ShutdownDiagnosticTestData.DirectoryScope())
                using (var descriptor = Process.GetCurrentProcess())
                {
                    string productRoot = Path.Combine(directory.Root, "candidate"); Directory.CreateDirectory(productRoot);
                    string product = Path.Combine(productRoot, "VBAi.dll"), local = typeof(VbeSession).Assembly.Location;
                    File.Copy(local, product);
                    string hash;
                    using (var input = File.OpenRead(product))
                    using (var sha = SHA256.Create()) hash = BitConverter.ToString(sha.ComputeHash(input)).Replace("-", "");
                    string mvid = typeof(VbeSession).Module.ModuleVersionId.ToString("D");
                    Environment.SetEnvironmentVariable(names[0], kind == "disabled" ? null : directory.Root);
                    Environment.SetEnvironmentVariable(names[1], kind == "missing product" || kind == "disabled" ? null : product);
                    Environment.SetEnvironmentVariable(names[2], mvid);
                    Environment.SetEnvironmentVariable(names[3], hash);
                    string fixtureRoot = Path.Combine(directory.Root, "fixture"); Directory.CreateDirectory(fixtureRoot);
                    var fixture = (ExcelVbeFixture)Activator.CreateInstance(typeof(ExcelVbeFixture), true);
                    typeof(ExcelVbeFixture).GetProperty("Root", fields | BindingFlags.Public).SetValue(fixture, fixtureRoot);
                    typeof(ExcelVbeFixture).GetProperty("ProcessId", fields | BindingFlags.Public).SetValue(fixture, descriptor.Id);
                    typeof(ExcelVbeFixture).GetField("owned", fields).SetValue(fixture, true);
                    typeof(ExcelVbeFixture).GetField("ownedProcess", fields).SetValue(fixture, descriptor);
                    typeof(ExcelVbeFixture).GetField("retainEvidence", fields).SetValue(fixture, true);
                    typeof(ExcelVbeFixture).GetField("ownedImagePath", fields).SetValue(fixture, (Func<string>)(() => @"C:\Qualification\EXCEL.EXE"));
                    var status = Status(); status["HostProcessId"] = descriptor.Id; status["AssemblyModuleVersionId"] = mvid;
                    status["AssemblyPath"] = kind == "foreign native status" ? local : kind == "case status" ? product.ToUpperInvariant() : product;
                    var arm = typeof(ExcelVbeFixture).GetMethod("ArmAddInShutdownObservation", fields);
                    if (kind == "missing product" || kind == "foreign native status")
                    {
                        Assert.ThrowsException<TargetInvocationException>(() => arm.Invoke(fixture, new object[] { 73u, status }));
                        Assert.AreEqual(0, Directory.GetFiles(directory.Root, "request-*.json").Length);
                        Assert.IsNull(typeof(ExcelVbeFixture).GetField("addInShutdownIdentity", fields).GetValue(fixture));
                    }
                    else
                    {
                        arm.Invoke(fixture, new object[] { 73u, status });
                        bool enabled = kind != "disabled";
                        Assert.AreEqual(enabled ? 1 : 0, Directory.GetFiles(directory.Root, "request-*.json").Length);
                        Assert.AreEqual(enabled, File.Exists(Path.Combine(fixtureRoot, "addin-shutdown-request.json")));
                        if (enabled)
                        {
                            var identity = AddInShutdownDiagnostic.DecodeRequest(File.ReadAllText(Path.Combine(directory.Root, "request-" + descriptor.Id + ".json")), Path.GetFileName(directory.Root));
                            Assert.AreEqual(product, identity.ProductPath); Assert.AreNotEqual(local, identity.ProductPath);
                            Assert.AreEqual(mvid, identity.ProductMvid); Assert.AreEqual(hash, identity.ProductSha256);
                            Assert.AreEqual(descriptor.Id, identity.ProcessId); Assert.AreEqual(73u, identity.ThreadId);
                            Assert.AreEqual(descriptor.StartTime.ToUniversalTime().ToString("o"), identity.ProcessStartedUtc);
                        }
                    }
                }
            }
            finally { foreach (var entry in prior) Environment.SetEnvironmentVariable(entry.Key, entry.Value); }
        }
    }
}

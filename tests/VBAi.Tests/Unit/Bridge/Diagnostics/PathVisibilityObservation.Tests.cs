using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using VBAi;

namespace VBAi.Tests.Unit
{
    [TestClass, TestCategory("Unit")]
    public sealed class PathVisibilityObservationTests
    {
        [TestMethod]
        public void CallingThreadReadDistinguishesMissingAttributesAndNeverChangesSyntheticBytes()
        {
            string parent = Path.Combine(Path.GetTempPath(), "VBAi-path-observer-" + Guid.NewGuid().ToString("N"));
            string first = Path.Combine(parent, Guid.NewGuid().ToString("N")), missing = Path.Combine(parent, Guid.NewGuid().ToString("N"));
            Directory.CreateDirectory(first);
            string file = Path.Combine(first, PathVisibilityDiagnostic.SyntheticName);
            byte[] original = { 86, 66, 65, 105 };
            File.WriteAllBytes(file, original);
            try
            {
                var report = PathVisibilityObservation.Read(new[] { first, file, missing, Path.Combine(missing, PathVisibilityDiagnostic.SyntheticName) });
                Assert.AreEqual(report["ProcessId"], report["FinalProcessId"]);
                Assert.AreEqual(report["NativeThreadId"], report["FinalNativeThreadId"]);
                var rows = ((IEnumerable<object>)report["Paths"]).Cast<IDictionary<string, object>>().ToArray();
                Assert.AreEqual(true, rows[0]["DirectoryExists"]); Assert.AreEqual(false, rows[1]["DirectoryExists"]);
                Assert.AreEqual(true, rows[0]["NativeSucceeded"]); Assert.AreEqual(true, rows[1]["NativeSucceeded"]);
                Assert.AreEqual(false, rows[0]["NativeErrorMeaningful"]);
                Assert.AreEqual(false, rows[2]["NativeSucceeded"]); Assert.IsTrue(rows[2].ContainsKey("ManagedAttributesError"));
                Assert.AreEqual(uint.MaxValue, rows[2]["NativeAttributes"]);
                Assert.IsTrue((int)rows[2]["NativeLastError"] == 2 || (int)rows[2]["NativeLastError"] == 3);
                CollectionAssert.AreEqual(original, File.ReadAllBytes(file));
                Assert.IsFalse(Directory.Exists(missing));
            }
            finally { Directory.Delete(parent, true); }
        }
    }
}

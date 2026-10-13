using Microsoft.VisualStudio.TestTools.UnitTesting;
using System;
using System.Linq;
using VBAi;

namespace VBAi.Tests.Unit
{
    public sealed partial class MacroGitRepositoryTests
    {
        /// <summary>Checks the actual Git process budget and exact package bytes for small and larger module sets.</summary>
        [TestMethod]
        [DataRow(1), DataRow(12)]
        public void CompleteSnapshotBatchReadUsesThreeProcessesAndPreservesSourceBytes(int modules)
        {
            using (var fixture = new MacroGitOperationsTests.Fixture())
            {
                fixture.Host.VBComponents.Item("Module1").CodeModule.Text = "Attribute VB_Name = \"Module1\"\r\nOption Explicit\r\nPublic Sub Smoke()\r\n Debug.Print \"caf\u00e9\"\r\nEnd Sub\r\n";
                for (int index = 1; index < modules; index++)
                {
                    string name = "Extra" + index;
                    fixture.Host.VBComponents.Add(new global::FakeComponent(name, 1,
                        "Attribute VB_Name = \"" + name + "\"\r\nOption Explicit\r\nPublic Function Value() As Long\r\n Value = " + index + "\r\nEnd Function\r\n"));
                }
                var expected = fixture.Project.Capture();
                string commit = fixture.Repository.Commit(expected, null, "Batch source qualification");
                int starts = 0;
                var start = fixture.Repository.StartProcess;
                fixture.Repository.StartProcess = process => { starts++; return start(process); };
                var actual = fixture.Repository.Read(commit);
                Assert.AreEqual(3, starts, "One root tree, one sized blob tree and one binary batch are sufficient.");
                Assert.IsTrue(actual.SameAs(expected));
                var files = expected.Serialize(); var observed = actual.Serialize();
                CollectionAssert.AreEquivalent(files.Keys.ToArray(), observed.Keys.ToArray());
                foreach (var item in files) CollectionAssert.AreEqual(item.Value, observed[item.Key], item.Key);
            }
        }

        /// <summary>Uses sized, NUL-delimited trees and refuses malformed metadata before any batch body read.</summary>
        [TestMethod]
        [DataRow("no-tab"), DataRow("missing-size"), DataRow("non-numeric-size")]
        [DataRow("negative-size"), DataRow("invalid-id"), DataRow("total-too-large")]
        public void InvalidSizedTreesNeverLaunchABlobRead(string fault)
        {
            using (var fixture = new MacroGitOperationsTests.Fixture())
            {
                string id = new string('a', 40), entry = "100644 blob " + id + " 1\tModule1.bas\0";
                switch (fault)
                {
                    case "no-tab": entry = "100644 blob " + id + " 1\0"; break;
                    case "missing-size": entry = "100644 blob " + id + "\tModule1.bas\0"; break;
                    case "non-numeric-size": entry = "100644 blob " + id + " invalid\tModule1.bas\0"; break;
                    case "negative-size": entry = "100644 blob " + id + " -1\tModule1.bas\0"; break;
                    case "invalid-id": entry = "100644 blob invalid 1\tModule1.bas\0"; break;
                    case "total-too-large": entry = "100644 blob " + id + " " + VbaGitSnapshot.MaxBytes + "\tModule1.bas\0" + "100644 blob " + id + " 1\tExtra.bas\0"; break;
                }
                bool blobRead = false;
                fixture.Repository.CommandOverride = (args, input, use, allow) =>
                {
                    if (args[0] == "ls-tree") return Reply(args.Last() == "root" ? "040000 tree subtree\tvba\0" : entry);
                    blobRead = true; return Reply("");
                };
                Assert.ThrowsException<InvalidOperationException>(() => fixture.Repository.Read("root"));
                Assert.IsFalse(blobRead, "The entire metadata request must be admitted before opening the blob batch.");
            }
        }
    }
}

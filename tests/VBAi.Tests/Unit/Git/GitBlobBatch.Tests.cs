using Microsoft.VisualStudio.TestTools.UnitTesting;
using System;
using System.IO;
using System.Linq;
using System.Text;
using VBAi;

namespace VBAi.Tests.Unit
{
    /// <summary>Checks exact binary framing, request identity and resource limits for Git batch reads.</summary>
    [TestClass, TestCategory("Unit")]
    public sealed class GitBlobBatchTests
    {
        [TestMethod]
        [DataRow(40), DataRow(64)]
        public void OrderedBinaryAndEmptyBlobsKeepTheirExactBytes(int identityLength)
        {
            string first = new string('a', identityLength), second = new string('b', identityLength);
            var binary = new byte[] { 0, 10, 13, 255, 128, 32, 0 };
            byte[] reply = Record(first, binary).Concat(Record(second, new byte[0])).ToArray();
            var actual = GitBlobBatch.Read(reply, new[] { new GitBlobBatch.Entry(first, binary.Length), new GitBlobBatch.Entry(second, 0) });
            CollectionAssert.AreEqual(binary, actual[0]); Assert.AreEqual(0, actual[1].Length);
        }

        [TestMethod]
        public void RepeatedObjectIdsProduceIndependentBodies()
        {
            string id = new string('a', 40); byte[] record = Record(id, new byte[] { 7 });
            var actual = GitBlobBatch.Read(record.Concat(record).ToArray(), new[] { new GitBlobBatch.Entry(id, 1), new GitBlobBatch.Entry(id, 1) });
            actual[0][0] = 9; Assert.AreEqual((byte)7, actual[1][0]);
        }

        [TestMethod]
        public void AnEmptyRequestRequiresAnEmptyReply()
        {
            Assert.AreEqual(0, GitBlobBatch.Read(new byte[0], new GitBlobBatch.Entry[0]).Length);
            Assert.ThrowsException<InvalidOperationException>(() => GitBlobBatch.Read(new byte[] { 10 }, new GitBlobBatch.Entry[0]));
        }

        [TestMethod]
        [DataRow("missing"), DataRow("wrong-id"), DataRow("wrong-type"), DataRow("wrong-size")]
        [DataRow("header-crlf"), DataRow("short-header"), DataRow("short-body"), DataRow("missing-delimiter")]
        [DataRow("wrong-delimiter"), DataRow("trailing"), DataRow("swapped"), DataRow("leading-size-zero")]
        public void MalformedOrMismatchedRepliesCannotBecomeAValidSnapshot(string fault)
        {
            string id = new string('a', 40), other = new string('b', 40);
            byte[] reply = Record(id, new byte[] { 1, 2 });
            var entries = new[] { new GitBlobBatch.Entry(id, 2) };
            switch (fault)
            {
                case "missing": reply = Encoding.ASCII.GetBytes(id + " missing\n"); break;
                case "wrong-id": reply[0] = (byte)'b'; break;
                case "wrong-type": reply = Encoding.ASCII.GetBytes(id + " tree 2\n12\n"); break;
                case "wrong-size": reply = Encoding.ASCII.GetBytes(id + " blob 3\n123\n"); break;
                case "header-crlf": reply = Encoding.ASCII.GetBytes(id + " blob 2\r\n12\n"); break;
                case "short-header": reply = reply.Take(4).ToArray(); break;
                case "short-body": reply = reply.Take(reply.Length - 2).ToArray(); break;
                case "missing-delimiter": reply = reply.Take(reply.Length - 1).ToArray(); break;
                case "wrong-delimiter": reply[reply.Length - 1] = 0; break;
                case "trailing": reply = reply.Concat(new byte[] { 3 }).ToArray(); break;
                case "swapped": reply = Record(other, new byte[] { 1, 2 }).Concat(reply).ToArray(); entries = new[] { new GitBlobBatch.Entry(id, 2), new GitBlobBatch.Entry(other, 2) }; break;
                case "leading-size-zero": reply = Encoding.ASCII.GetBytes(id + " blob 02\n12\n"); break;
            }
            Assert.ThrowsException<InvalidOperationException>(() => GitBlobBatch.Read(reply, entries));
        }

        [TestMethod]
        [DataRow("null-entry"), DataRow("null-id"), DataRow("short-id"), DataRow("expression")]
        [DataRow("newline"), DataRow("negative"), DataRow("single-too-large"), DataRow("aggregate-too-large"), DataRow("too-many")]
        public void InvalidRequestsAreRefusedBeforeGitIsLaunched(string fault)
        {
            string id = new string('a', 40);
            var entries = new[] { new GitBlobBatch.Entry(id, 0) };
            switch (fault)
            {
                case "null-entry": entries[0] = null; break;
                case "null-id": entries[0] = new GitBlobBatch.Entry(null, 0); break;
                case "short-id": entries[0] = new GitBlobBatch.Entry("abc", 0); break;
                case "expression": entries[0] = new GitBlobBatch.Entry("HEAD:path", 0); break;
                case "newline": entries[0] = new GitBlobBatch.Entry(id + "\n", 0); break;
                case "negative": entries[0] = new GitBlobBatch.Entry(id, -1); break;
                case "single-too-large": entries[0] = new GitBlobBatch.Entry(id, VbaGitSnapshot.MaxBytes + 1L); break;
                case "aggregate-too-large": entries = new[] { new GitBlobBatch.Entry(id, VbaGitSnapshot.MaxBytes), new GitBlobBatch.Entry(id, 1) }; break;
                case "too-many": entries = Enumerable.Repeat(new GitBlobBatch.Entry(id, 0), 2050).ToArray(); break;
            }
            Assert.ThrowsException<InvalidOperationException>(() => GitBlobBatch.Validate(entries));
        }

        [TestMethod]
        public void ExactAdmissionBoundsRemainValid()
        {
            string id = new string('a', 40);
            GitBlobBatch.Validate(new[] { new GitBlobBatch.Entry(id, VbaGitSnapshot.MaxBytes) });
            GitBlobBatch.Validate(Enumerable.Repeat(new GitBlobBatch.Entry(id, 0), 2049).ToArray());
        }

        [TestMethod]
        public void NullBuffersAndRequestsRetainArgumentFailures()
        {
            Assert.ThrowsException<ArgumentNullException>(() => GitBlobBatch.Read(null, new GitBlobBatch.Entry[0]));
            Assert.ThrowsException<ArgumentNullException>(() => GitBlobBatch.Read(new byte[0], null));
        }

        /// <summary>Builds the documented raw Git batch record, including the delimiter after an empty body.</summary>
        private static byte[] Record(string id, byte[] body)
        {
            using (var stream = new MemoryStream())
            {
                byte[] header = Encoding.ASCII.GetBytes(id + " blob " + body.Length + "\n");
                stream.Write(header, 0, header.Length); stream.Write(body, 0, body.Length); stream.WriteByte(10);
                return stream.ToArray();
            }
        }
    }
}

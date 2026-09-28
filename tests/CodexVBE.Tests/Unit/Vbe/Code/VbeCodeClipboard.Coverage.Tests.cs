using System;
using CodexVBE;
using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace CodexVBE.Tests.Unit
{
    [TestClass]
    public sealed class VbeCodeClipboardFaultTests
    {
        public sealed class Source { public string Code { get; set; } = "abc\r\ndef"; }
        public sealed class ModeSnapshot { public int Mode { get; set; } = 2; }
        private sealed class Clipboard : ICodeClipboard
        {
            internal CodeClipboardSnapshot State = new CodeClipboardSnapshot { HasText = true, Text = "replace", Version = "revision" };
            public CodeClipboardSnapshot Read() { return State; }
            public CodeClipboardSnapshot Write(string text) { return State; }
        }
        private static Request Request()
        { return new Request { Project = "P", Module = "M", ExpectedSha256 = VbeCodeClipboard.Hash("abc\r\ndef"),
            StartLine = 1, EndLine = 1, StartColumn = 1, EndColumn = 2, ExpectedClipboardVersion = "revision" }; }

        [TestMethod]
        public void ClipboardEditingValidatesCommandsReadsVersionsAndNonemptyRanges()
        {
            var clipboard = new Clipboard();
            Func<Request, Response> execute = r => r.Command == "debug_state" ? Response.Success(new ModeSnapshot()) : Response.Success(new Source());
            var tool = new VbeCodeClipboard(execute, clipboard);
            Assert.AreSame(clipboard.State, tool.Read());
            Assert.ThrowsException<ArgumentException>(() => tool.Edit(Request(), "invalid"));
            tool = new VbeCodeClipboard(r => Response.Failure("unavailable"), clipboard);
            Assert.ThrowsException<InvalidOperationException>(() => tool.Edit(Request(), "cut"));
            Assert.ThrowsException<InvalidOperationException>(() => tool.Edit(Request(), "copy"));
            tool = new VbeCodeClipboard(execute, clipboard);
            var request = Request(); request.ExpectedSha256 = " ";
            Assert.ThrowsException<InvalidOperationException>(() => tool.Edit(request, "copy"));
            request = Request(); request.EndColumn = request.StartColumn;
            Assert.ThrowsException<ArgumentException>(() => tool.Edit(request, "copy"));
            request = Request(); request.ExpectedClipboardVersion = " ";
            Assert.ThrowsException<InvalidOperationException>(() => tool.Edit(request, "paste"));
        }

        [TestMethod]
        public void ClipboardEditsPreserveFailureDetailsAndIndependentlyVerifyNativeSource()
        {
            foreach (string action in new[] { "cut", "paste" })
            foreach (int outcome in new[] { 0, 1, 2, 3 })
            {
                int reads = 0; string requested = null;
                var tool = new VbeCodeClipboard(r => {
                    if (r.Command == "debug_state") return Response.Success(new ModeSnapshot());
                    if (r.Command == "replace_lines") { requested = r.Text; return outcome == 0 ? Response.Failure("write failure") : Response.Success(null); }
                    reads++;
                    if (reads == 1) return Response.Success(new Source());
                    return outcome == 1 ? Response.Failure("readback failure") : Response.Success(new Source { Code = outcome == 2 ? requested : "native normalization" });
                }, new Clipboard());
                if (outcome == 0)
                {
                    string error = Assert.ThrowsException<InvalidOperationException>(() => tool.Edit(Request(), action)).Message;
                    StringAssert.Contains(error, "write failure");
                    Assert.AreEqual(action == "cut", error.Contains("copied to the clipboard"));
                }
                else
                {
                    dynamic result = tool.Edit(Request(), action);
                    Assert.IsTrue((bool)result.Applied);
                    Assert.AreEqual(outcome == 2, (bool)result.Verified);
                }
            }
        }

        [TestMethod]
        public void ClipboardRangesRejectEveryInvalidEndpointAndAcceptEmptyInsertion()
        {
            foreach (int[] range in new[] { new[] { 0, 1, 1, 1 }, new[] { 2, 1, 1, 1 }, new[] { 1, 3, 1, 1 },
                new[] { 1, 1, 0, 1 }, new[] { 1, 1, 1, 0 }, new[] { 1, 1, 5, 1 }, new[] { 1, 1, 1, 5 }, new[] { 1, 1, 2, 1 } })
            {
                var request = Request(); request.StartLine = range[0]; request.EndLine = range[1]; request.StartColumn = range[2]; request.EndColumn = range[3];
                Assert.ThrowsException<ArgumentException>(() => VbeCodeClipboard.Range("abc\r\ndef", request, out _, out _));
            }
            var empty = Request(); empty.EndColumn = 1;
            VbeCodeClipboard.Range("", empty, out int start, out int length);
            Assert.AreEqual(0, start); Assert.AreEqual(0, length);
            VbeCodeClipboard.Range("abc\r\ndef", new Request { StartLine = 2, EndLine = 2, StartColumn = 1, EndColumn = 4 }, out start, out length);
            Assert.AreEqual(5, start); Assert.AreEqual(3, length);
        }
    }
}

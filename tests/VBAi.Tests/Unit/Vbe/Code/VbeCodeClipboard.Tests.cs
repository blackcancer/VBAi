using Microsoft.VisualStudio.TestTools.UnitTesting;
using System;
namespace VBAi.Tests.Unit
{
    [TestClass, TestCategory("Unit")]
    public sealed class VbeCodeClipboardTests
    {
        public sealed class CodeData { public string Code { get; set; } }
        public sealed class ModeData { public int Mode { get; set; } }
        private sealed class Clipboard : ICodeClipboard
        {
            public CodeClipboardSnapshot Snapshot = new CodeClipboardSnapshot { HasText = true, Text = "pasted", Version = "1" };
            public bool FailWrite; public int Writes;
            public CodeClipboardSnapshot Read() => Snapshot;
            public CodeClipboardSnapshot Write(string text)
            { Writes++; if (FailWrite) throw new InvalidOperationException("clipboard busy"); return Snapshot = new CodeClipboardSnapshot { HasText = true, Text = text, Version = "2" }; }
        }
        private sealed class Host
        {
            public string Code = "abc\r\ndef"; public int Mode = 2, Writes;
            public Response Execute(Request request)
            {
                if (request.Command == "debug_state") return Response.Success(new ModeData { Mode = Mode });
                if (request.Command == "read_module") return Response.Success(new CodeData { Code = Code });
                if (request.Command == "replace_lines") { Writes++; Code = request.Text; return Response.Success(null); }
                throw new InvalidOperationException(request.Command);
            }
            public Request Request() => new Request
            {
                Project = "P",
                Module = "M",
                ExpectedSha256 = VbeCodeClipboard.Hash(Code),
                StartLine = 1,
                StartColumn = 2,
                EndLine = 2,
                EndColumn = 2,
                ExpectedClipboardVersion = "1"
            };
        }
        [TestMethod]
        public void CopyUsesExclusiveMultilineRangeWithoutEditing()
        {
            var host = new Host(); var clipboard = new Clipboard();
            new VbeCodeClipboard(host.Execute, clipboard).Edit(host.Request(), "copy");
            Assert.AreEqual("bc\r\nd", clipboard.Snapshot.Text); Assert.AreEqual(0, host.Writes);
        }
        [TestMethod]
        public void CutCopiesBeforeDeletingAndDoesNotEditWhenCopyFails()
        {
            var host = new Host(); var clipboard = new Clipboard { FailWrite = true };
            var tool = new VbeCodeClipboard(host.Execute, clipboard);
            Assert.ThrowsException<InvalidOperationException>(() => tool.Edit(host.Request(), "cut"));
            Assert.AreEqual(0, host.Writes); clipboard.FailWrite = false;
            tool.Edit(host.Request(), "cut"); Assert.AreEqual("aef", host.Code); Assert.AreEqual("bc\r\nd", clipboard.Snapshot.Text);
        }
        [TestMethod]
        public void ChangedClipboardOrSourceIsRefusedWithoutSideEffects()
        {
            var host = new Host(); var clipboard = new Clipboard(); var tool = new VbeCodeClipboard(host.Execute, clipboard);
            var request = host.Request(); request.ExpectedClipboardVersion = "old";
            Assert.ThrowsException<InvalidOperationException>(() => tool.Edit(request, "paste"));
            request = host.Request(); request.ExpectedSha256 = "old";
            Assert.ThrowsException<InvalidOperationException>(() => tool.Edit(request, "cut"));
            Assert.AreEqual(0, host.Writes); Assert.AreEqual(0, clipboard.Writes);
        }
        [TestMethod]
        public void NonDesignModeDoesNotChangeClipboardOrCode()
        {
            var host = new Host { Mode = 1 }; var clipboard = new Clipboard();
            Assert.ThrowsException<InvalidOperationException>(() => new VbeCodeClipboard(host.Execute, clipboard).Edit(host.Request(), "cut"));
            Assert.AreEqual(0, clipboard.Writes); Assert.AreEqual(0, host.Writes);
        }
        [TestMethod]
        public void PasteNormalizesNewlinesAndSupportsEmptyModules()
        {
            var host = new Host { Code = "" }; var clipboard = new Clipboard(); clipboard.Snapshot.Text = "' é\n' Ω";
            var request = host.Request(); request.StartColumn = request.EndColumn = request.StartLine = request.EndLine = 1;
            new VbeCodeClipboard(host.Execute, clipboard).Edit(request, "paste");
            Assert.AreEqual("' é\r\n' Ω", host.Code);
        }
        [TestMethod]
        public void InvalidRangeAndNonTextClipboardNeverEdit()
        {
            var host = new Host(); var clipboard = new Clipboard(); var tool = new VbeCodeClipboard(host.Execute, clipboard);
            var request = host.Request(); request.EndColumn = 9;
            Assert.ThrowsException<ArgumentException>(() => tool.Edit(request, "copy"));
            clipboard.Snapshot.HasText = false;
            Assert.ThrowsException<InvalidOperationException>(() => tool.Edit(host.Request(), "paste"));
            Assert.AreEqual(0, host.Writes); Assert.AreEqual(0, clipboard.Writes);
        }
        [TestMethod]
        public void IdenticalPasteDoesNotTouchNativeCodeOrHistory()
        {
            var host = new Host(); var clipboard = new Clipboard(); clipboard.Snapshot.Text = "bc\r\nd";
            new VbeCodeClipboard(host.Execute, clipboard).Edit(host.Request(), "paste");
            Assert.AreEqual(0, host.Writes);
        }
        [TestMethod]
        public void UnsupportedControlCharactersAreRefused()
        {
            var host = new Host(); var clipboard = new Clipboard(); clipboard.Snapshot.Text = "a\0b";
            Assert.ThrowsException<ArgumentException>(() => new VbeCodeClipboard(host.Execute, clipboard).Edit(host.Request(), "paste"));
            Assert.AreEqual(0, host.Writes);
        }
    }
}

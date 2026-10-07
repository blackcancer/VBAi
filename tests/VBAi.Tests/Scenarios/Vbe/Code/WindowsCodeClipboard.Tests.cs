using Microsoft.VisualStudio.TestTools.UnitTesting;
using System;
using System.Threading;
using System.Windows.Forms;

namespace VBAi.Tests.Unit
{
    [TestClass]
    [TestCategory("Unit")]
    public sealed class WindowsCodeClipboardTests
    {
        [STATestMethod]
        public void UnicodeClipboardReadsRequireStableRevisionAndBoundedText()
        {
            string text = "é Ω\r\ntext";
            bool present = true;
            var clipboard = new WindowsCodeClipboard
            {
                SequenceNative = () => 42,
                ContainsNative = format => { Assert.AreEqual(TextDataFormat.UnicodeText, format); return present; },
                GetNative = format => { Assert.AreEqual(TextDataFormat.UnicodeText, format); return text; }
            };
            var state = clipboard.Read();
            Assert.IsTrue(state.HasText); Assert.AreEqual(text, state.Text);
            Assert.AreEqual("42:" + VbeCodeClipboard.Hash(text), state.Version);
            present = false; state = clipboard.Read();
            Assert.IsFalse(state.HasText); Assert.IsNull(state.Text);
            present = true; text = null; Assert.IsNull(clipboard.Read().Text);
            text = new string('x', 1024 * 1024); Assert.AreEqual(text.Length, clipboard.Read().Text.Length);
            text += "x"; Assert.ThrowsException<InvalidOperationException>(() => clipboard.Read());
            int revision = 0; clipboard.SequenceNative = () => (uint)revision++;
            Assert.ThrowsException<InvalidOperationException>(() => clipboard.Read());
        }

        [STATestMethod]
        public void UnicodeClipboardWritesValidateBoundsAndVerifyTheNativeReadback()
        {
            string current = null; bool present = true, ignore = false;
            var clipboard = new WindowsCodeClipboard
            {
                SequenceNative = () => 1,
                ContainsNative = format => present,
                GetNative = format => current,
                SetNative = (text, format) => { Assert.AreEqual(TextDataFormat.UnicodeText, format); if (!ignore) current = text; }
            };
            foreach (string invalid in new[] { null, "", new string('x', 1024 * 1024 + 1) })
                Assert.ThrowsException<ArgumentException>(() => clipboard.Write(invalid));
            Assert.AreEqual("é", clipboard.Write("é").Text);
            Assert.AreEqual(1024 * 1024, clipboard.Write(new string('x', 1024 * 1024)).Text.Length);
            ignore = true; Assert.ThrowsException<InvalidOperationException>(() => clipboard.Write("changed"));
            present = false; Assert.ThrowsException<InvalidOperationException>(() => clipboard.Write("changed"));
            clipboard.SetNative = (text, format) => { throw new InvalidOperationException("clipboard busy"); };
            StringAssert.Contains(Assert.ThrowsException<InvalidOperationException>(() => clipboard.Write("changed")).Message, "clipboard busy");
        }

        [TestMethod]
        public void WindowsClipboardRejectsMtaBeforeNativeAccess()
        {
            Exception failure = null;
            var thread = new Thread(() =>
            {
                try
                {
                    var clipboard = new WindowsCodeClipboard { SequenceNative = () => { Assert.Fail("No native access on MTA"); return 0; } };
                    Assert.ThrowsException<InvalidOperationException>(() => clipboard.Read());
                    Assert.ThrowsException<InvalidOperationException>(() => clipboard.Write("text"));
                }
                catch (Exception error) { failure = error; }
            });
            thread.SetApartmentState(ApartmentState.MTA); thread.Start();
            Assert.IsTrue(thread.Join(TimeSpan.FromSeconds(10)));
            if (failure != null) throw failure;
        }
    }
}

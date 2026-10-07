using Microsoft.VisualStudio.TestTools.UnitTesting;
using System;
using System.IO;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace VBAi.Tests.Unit
{
    [TestClass]
    [TestCategory("Unit")]
    public sealed class VbeDebugClipboardTests
    {
        [TestMethod]
        public void SnapshotCopiesAllSupportedFormatsWithoutHoldingMutableStreams()
        {
            var sourceBytes = new MemoryStream(new byte[] { 1, 2, 3 });
            var source = new DataObject();
            source.SetData(DataFormats.UnicodeText, false, "original text");
            source.SetData("VBAi.Binary", false, sourceBytes);
            var snapshot = VbeDebugClipboard.Snapshot.Capture(source);
            sourceBytes.Position = 0;
            sourceBytes.WriteByte(9);
            var restored = snapshot.CreateDataObject();
            Assert.AreEqual("original text", restored.GetData(DataFormats.UnicodeText, false));
            CollectionAssert.AreEqual(new byte[] { 1, 2, 3 },
                ((MemoryStream)restored.GetData("VBAi.Binary", false)).ToArray());
            Assert.IsTrue(snapshot.Matches(restored));
        }

        [TestMethod]
        public void SnapshotRejectsExtraFormatAfterRestoration()
        {
            var snapshot = VbeDebugClipboard.Snapshot.Capture(Text("original"));
            var changed = snapshot.CreateDataObject();
            changed.SetData("VBAi.Unexpected", false, "extra data");
            Assert.IsFalse(snapshot.Matches(changed), "Extra formats mean the clipboard was not restored exactly.");
        }

        [WinFormsTestMethod]
        public void UnsupportedAndOversizedClipboardRefuseBeforeSelectionOrCopy()
        {
            foreach (var original in new IDataObject[] { Unsupported(), Oversized(), new MissingFormatData() })
            {
                int selections = 0, copies = 0;
                var clipboard = new VbeDebugClipboard
                {
                    Sequence = () => 4,
                    YieldNative = () => Task.CompletedTask,
                    ReadData = () => original,
                    WriteData = _ => Assert.Fail("Clipboard must remain untouched.")
                };
                Assert.ThrowsException<InvalidOperationException>(() => clipboard.ReadAsync(
                    () => selections++, () => copies++).GetAwaiter().GetResult());
                Assert.AreEqual(0, selections);
                Assert.AreEqual(0, copies);
            }
        }

        [WinFormsTestMethod]
        public void ClipboardChangeBeforeCopyRefusesCopy()
        {
            uint sequence = 1;
            int copies = 0;
            var clipboard = new VbeDebugClipboard
            {
                Sequence = () => sequence,
                YieldNative = () => Task.CompletedTask,
                ReadData = () => Text("original"),
                WriteData = _ => Assert.Fail("A concurrent clipboard change must not be overwritten.")
            };
            Assert.ThrowsException<InvalidOperationException>(() => clipboard.ReadAsync(
                () => sequence++, () => copies++).GetAwaiter().GetResult());
            Assert.AreEqual(0, copies);
        }

        [WinFormsTestMethod]
        public void NativeCopyReadsFreshTextAndRestoresOriginalFormats()
        {
            uint sequence = 1;
            IDataObject current = Text("original");
            int restorations = 0;
            var clipboard = new VbeDebugClipboard
            {
                Sequence = () => sequence,
                YieldNative = () => Task.CompletedTask,
                ReadData = () => current,
                IsHostOwner = () => true,
                WriteData = data => { restorations++; current = data; sequence++; }
            };
            string read = clipboard.ReadAsync(() => { }, () => { current = Text("fresh output"); sequence++; }).GetAwaiter().GetResult();
            Assert.AreEqual("fresh output", read);
            Assert.AreEqual("original", current.GetData(DataFormats.UnicodeText, false));
            Assert.AreEqual(1, restorations);
        }

        [WinFormsTestMethod]
        public void DelayedNativeCopyUpdateIsObservedWithoutRepeatingCopy()
        {
            uint sequence = 1;
            IDataObject current = Text("original");
            bool copyIssued = false;
            int yields = 0, copies = 0;
            var clipboard = new VbeDebugClipboard
            {
                Sequence = () => sequence,
                YieldNative = () =>
                {
                    if (++yields == 3 && copyIssued) { current = Text("delayed output"); sequence++; }
                    return Task.CompletedTask;
                },
                ReadData = () => current,
                IsHostOwner = () => true,
                WriteData = data => { current = data; sequence++; }
            };
            string result = clipboard.ReadAsync(() => { }, () => { copyIssued = true; copies++; }).GetAwaiter().GetResult();
            Assert.AreEqual("delayed output", result);
            Assert.AreEqual(1, copies);
            Assert.IsTrue(yields >= 3);
            Assert.AreEqual("original", current.GetData(DataFormats.UnicodeText, false));
        }

        [WinFormsTestMethod]
        public void CopyExceptionStillRestoresWhenHostOwnsItsClipboardUpdate()
        {
            uint sequence = 1;
            IDataObject current = Text("original");
            int restorations = 0;
            var clipboard = new VbeDebugClipboard
            {
                Sequence = () => sequence,
                YieldNative = () => Task.CompletedTask,
                ReadData = () => current,
                IsHostOwner = () => true,
                WriteData = data => { restorations++; current = data; sequence++; }
            };
            var error = Assert.ThrowsException<InvalidOperationException>(() => clipboard.ReadAsync(() => { }, () =>
            {
                current = Text("partial copy"); sequence++;
                throw new InvalidOperationException("Copy failed after changing clipboard.");
            }).GetAwaiter().GetResult());
            StringAssert.Contains(error.Message, "Native Copy failed");
            Assert.AreEqual("original", current.GetData(DataFormats.UnicodeText, false));
            Assert.AreEqual(1, restorations);
        }

        [WinFormsTestMethod]
        public void ConcurrentChangeAfterCopyIsPreservedAndNoStaleOutputReturned()
        {
            uint sequence = 1;
            IDataObject current = Text("original");
            int restorations = 0;
            var clipboard = new VbeDebugClipboard
            {
                Sequence = () => sequence,
                YieldNative = () => Task.CompletedTask,
                ReadData = () => { if (sequence == 2) { current = Text("new owner"); sequence++; } return current; },
                IsHostOwner = () => sequence == 2,
                WriteData = _ => restorations++
            };
            Assert.ThrowsException<InvalidOperationException>(() => clipboard.ReadAsync(
                () => { }, () => { current = Text("copied output"); sequence++; }).GetAwaiter().GetResult());
            Assert.AreEqual("new owner", current.GetData(DataFormats.UnicodeText, false));
            Assert.AreEqual(0, restorations);
        }

        [WinFormsTestMethod]
        public void MissingClipboardUpdateDoesNotReturnOldText()
        {
            int restorations = 0;
            var clipboard = new VbeDebugClipboard
            {
                Sequence = () => 1,
                YieldNative = () => Task.CompletedTask,
                ReadData = () => Text("old text"),
                IsHostOwner = () => true,
                WriteData = _ => restorations++
            };
            Assert.ThrowsException<InvalidOperationException>(() => clipboard.ReadAsync(() => { }, () => { }).GetAwaiter().GetResult());
            Assert.AreEqual(0, restorations);
        }

        private static DataObject Text(string value)
        {
            var data = new DataObject();
            data.SetData(DataFormats.UnicodeText, false, value);
            return data;
        }

        private static DataObject Unsupported()
        {
            var data = new DataObject();
            data.SetData("VBAi.Unsupported", false, new object());
            return data;
        }

        private static DataObject Oversized()
        {
            var data = new DataObject();
            data.SetData("VBAi.Large", false, new MemoryStream(new byte[VbeDebugClipboard.Snapshot.MaximumBytes + 1]));
            return data;
        }

        private sealed class MissingFormatData : IDataObject
        {
            public object GetData(string format, bool autoConvert) => null;
            public object GetData(string format) => null;
            public object GetData(Type format) => null;
            public bool GetDataPresent(string format, bool autoConvert) => false;
            public bool GetDataPresent(string format) => false;
            public bool GetDataPresent(Type format) => false;
            public string[] GetFormats(bool autoConvert) => new[] { "VBAi.Missing" };
            public string[] GetFormats() => GetFormats(false);
            public void SetData(string format, bool autoConvert, object data) => throw new NotSupportedException();
            public void SetData(string format, object data) => throw new NotSupportedException();
            public void SetData(Type format, object data) => throw new NotSupportedException();
            public void SetData(object data) => throw new NotSupportedException();
        }
    }
}

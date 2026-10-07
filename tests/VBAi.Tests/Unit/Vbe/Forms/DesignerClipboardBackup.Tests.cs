using Microsoft.VisualStudio.TestTools.UnitTesting;
using System;
using System.IO;
using System.Windows.Forms;
namespace VBAi.Tests.Unit
{
    [TestClass, TestCategory("Unit")]
    public sealed class DesignerClipboardBackupTests
    {
        [TestMethod]
        public void BackupRejectsNullUnsupportedEmptyAndNonbinaryBags()
        {
            Assert.ThrowsException<InvalidOperationException>(() => DesignerClipboardBackup.Capture(null));
            var data = new DesignerDataFixture(); data.Values["MS Forms Bag"] = new object();
            Assert.ThrowsException<InvalidOperationException>(() => DesignerClipboardBackup.Capture(data));
            data.Values["MS Forms Bag"] = "not binary";
            Assert.ThrowsException<InvalidOperationException>(() => DesignerClipboardBackup.Capture(data));
            data.Values["MS Forms Bag"] = new MemoryStream();
            Assert.ThrowsException<InvalidOperationException>(() => DesignerClipboardBackup.Capture(data));
        }
        [TestMethod]
        public void BackupRetainsStringsOmitsNullAndMatchesEachSerializedFormat()
        {
            var data = new DesignerDataFixture(); data.Values["MS Forms Bag"] = new MemoryStream(new byte[] { 1, 2 });
            data.Values["text"] = "label"; data.Values["unreadable"] = null;
            var backup = DesignerClipboardBackup.Capture(data);
            Assert.AreEqual(12, backup.ByteCount); CollectionAssert.AreEqual(new[] { "unreadable" }, backup.OmittedFormats);
            Assert.IsTrue(backup.Matches(backup.CreateDataObject())); Assert.IsFalse(backup.Matches(null));
            data.Values["text"] = "other"; Assert.IsFalse(backup.Matches(data));
            data.Values["text"] = "label"; data.Values["MS Forms Bag"] = "wrong type"; Assert.IsFalse(backup.Matches(data));
            data.Values["MS Forms Bag"] = null; Assert.IsFalse(backup.Matches(data));
            data.Values["MS Forms Bag"] = new MemoryStream(new byte[] { 1, 2 }); Assert.IsTrue(backup.Matches(data));
            data.Values["text"] = new string('a', DesignerClipboardBackup.MaximumBytes / 2);
            Assert.ThrowsException<InvalidOperationException>(() => DesignerClipboardBackup.Capture(data));
        }
        [TestMethod]
        public void BackupOwnsBytesAndDetectsDifferentPayload()
        {
            var data = new DataObject(); var bytes = new byte[] { 1, 2, 3 };
            data.SetData("MS Forms Bag", false, new MemoryStream(bytes));
            var backup = DesignerClipboardBackup.Capture(data);
            bytes[0] = 9;
            Assert.IsTrue(backup.Matches(backup.CreateDataObject()));
            Assert.IsFalse(backup.Matches(data));
            Assert.AreEqual(3, backup.ByteCount);
        }
        [TestMethod]
        public void MissingOrOversizedBagCannotBecomeRecovery()
        {
            Assert.ThrowsException<InvalidOperationException>(() => DesignerClipboardBackup.Capture(new DataObject()));
            var data = new DataObject();
            data.SetData("MS Forms Bag", false, new MemoryStream(new byte[DesignerClipboardBackup.MaximumBytes + 1]));
            Assert.ThrowsException<InvalidOperationException>(() => DesignerClipboardBackup.Capture(data));
        }
    }
}

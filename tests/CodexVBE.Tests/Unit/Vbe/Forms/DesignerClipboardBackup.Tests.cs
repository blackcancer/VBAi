using System;
using System.IO;
using System.Windows.Forms;
using CodexVBE;
using Microsoft.VisualStudio.TestTools.UnitTesting;
namespace CodexVBE.Tests.Unit
{
    [TestClass, TestCategory("Unit")]
    public sealed class DesignerClipboardBackupTests
    {
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

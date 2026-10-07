namespace VBAi.Tests.Unit
{
    using Microsoft.VisualStudio.TestTools.UnitTesting;
    using System;
    using VBAi;

    public sealed partial class VbeProjectExcelHostTests
    {
        [TestMethod]
        public void DeferredSignatureKeepsOriginalIdentityAndIgnoresOnlySavedFlag()
        {
            var f = Create();
            var component = new VbeProjectComponentsTests.FakeComponent("Module1", 1);
            f.Project.VBComponents.Add(component);
            Action guard = f.Service.CaptureSignaturePersistence(f.Project.Name);
            f.Project.Saved = false;
            guard();
            component.CodeModule.Source = "Option Explicit\r\nSub Changed()\r\nEnd Sub";
            Assert.ThrowsException<InvalidOperationException>(() => guard());
            component.CodeModule.Source = "Option Explicit";
            f.Project.Mode = 1;
            Assert.ThrowsException<InvalidOperationException>(() => guard());
            f.Project.Mode = 2;
            f.Project.Description = "changed";
            Assert.ThrowsException<InvalidOperationException>(() => guard());
            f.Project.Description = "Original";
            f.Project.FileName = @"C:\fixture\Another.xlsm";
            Assert.ThrowsException<InvalidOperationException>(() => guard());
            f.Project.FileName = f.Workbook.FullName;
            guard();
            f.Vbe.VBProjects.Clear();
            f.Vbe.VBProjects.Add(new VbeProjectComponentsTests.FakeProject { Name = f.Project.Name, FileName = f.Project.FileName });
            Assert.ThrowsException<InvalidOperationException>(() => guard());
            Assert.AreEqual(0, f.Workbook.SaveAttempts);
        }

        [TestMethod]
        public void DeferredSignatureRefusesAWorkerThread()
        {
            var f = Create(); Action guard = f.Service.CaptureSignaturePersistence(f.Project.Name);
            Exception failure = null;
            var worker = new System.Threading.Thread(() => { try { guard(); } catch (Exception ex) { failure = ex; } });
            worker.Start(); Assert.IsTrue(worker.Join(5000));
            Assert.IsInstanceOfType(failure, typeof(InvalidOperationException));
            Assert.AreEqual(0, f.Workbook.SaveAttempts);
        }

        [TestMethod]
        public void SignatureFinalAuthorizationRunsAfterPreparatoryReadsBeforeSave()
        {
            var f = Create(); f.Workbook.VBASigned = true; int checks = 0;
            Assert.ThrowsException<InvalidOperationException>(() => f.Service.PersistExcelSignature(f.Project.Name, () =>
            {
                checks++;
                throw new InvalidOperationException("authority revoked during native preparation");
            }));
            Assert.AreEqual(1, checks); Assert.AreEqual(0, f.Workbook.SaveAttempts);
        }

        [TestMethod]
        public void SignatureSavePostMutationBusyFailureIsNeverReplayed()
        {
            var f = Create(); f.Workbook.VBASigned = true;
            f.Workbook.AfterSave = () => { throw new System.Runtime.InteropServices.COMException("post-save observation busy", unchecked((int)0x800AC472)); };
            Assert.ThrowsException<System.Runtime.InteropServices.COMException>(() => f.Service.PersistExcelSignature(f.Project.Name));
            Assert.AreEqual(1, f.Workbook.SaveAttempts);
            Assert.IsTrue(f.Workbook.Saved);
        }
    }
}

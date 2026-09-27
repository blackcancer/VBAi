using System;
using System.Collections.Generic;
using CodexVBE;
using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace CodexVBE.Tests.Unit
{
    [TestClass]
    [TestCategory("Unit")]
    public sealed class VbeDebugWindowsSignatureProbeTests
    {
        [TestMethod]
        public void ReadSignatureRequiresDialogAndCancelControl()
        {
            var fake = new SignatureFake { Open = false };
            Assert.ThrowsException<InvalidOperationException>(() => VbeDebugWindows.ReadSignatureDialog("Book", fake));
            Assert.AreEqual(60, fake.Pauses);
            fake.Open = true;
            Assert.ThrowsException<InvalidOperationException>(() => VbeDebugWindows.ReadSignatureDialog("Book", fake));
            Assert.AreEqual(1, fake.Closes);
        }

        [TestMethod]
        public void ReadSignatureReturnsLabelsAndClosesByCancel()
        {
            var fake = new SignatureFake { CloseAfterCancel = true };
            fake.Items.Add(new VbeDebugWindows.SignatureChild { Index = 1, Role = 41, Name = "[No certificate]" });
            fake.Items.Add(new VbeDebugWindows.SignatureChild { Index = 2, Role = 41, Name = "Certificate name" });
            fake.Items.Add(new VbeDebugWindows.SignatureChild { Index = 3, Role = 41,
                Name = "The VBA project is currently signed as" });
            fake.Items.Add(new VbeDebugWindows.SignatureChild { Index = 4, Role = 41, Name = "Cert A" });
            fake.Items.Add(new VbeDebugWindows.SignatureChild { Index = 5, Role = 41, Name = "Certificate name" });
            fake.Items.Add(new VbeDebugWindows.SignatureChild { Index = 6, Role = 41, Name = "Sign as" });
            fake.Items.Add(new VbeDebugWindows.SignatureChild { Index = 7, Role = 43, Name = "Cancel" });
            fake.Items.Add(new VbeDebugWindows.SignatureChild { Index = 8, Role = 99, Name = "ignored" });
            dynamic result = VbeDebugWindows.ReadSignatureDialog("Book", fake);
            Assert.AreEqual("Book", (string)result.Project);
            Assert.AreEqual("[No certificate]", (string)result.CurrentCertificate);
            Assert.AreEqual("Cert A", (string)result.SignAsCertificate);
            Assert.AreEqual(7, fake.CancelIndex);
            Assert.AreEqual(0, fake.Closes);
            Assert.IsTrue((bool)result.DialogClosed);
        }

        [TestMethod]
        public void ReadSignatureClosesOwnDialogWhenAccessibilityFailsAndRejectsStuckDialog()
        {
            var fake = new SignatureFake { FailChildren = true };
            Assert.ThrowsException<InvalidOperationException>(() => VbeDebugWindows.ReadSignatureDialog("Book", fake));
            Assert.AreEqual(1, fake.Closes);
            fake = new SignatureFake();
            fake.Items.Add(new VbeDebugWindows.SignatureChild { Index = 1, Role = 43, Name = "Annuler" });
            Assert.ThrowsException<InvalidOperationException>(() => VbeDebugWindows.ReadSignatureDialog("Book", fake));
            Assert.AreEqual(20, fake.ClosePolls);
        }

        private sealed class SignatureFake : VbeDebugWindows.ISignatureProbe
        {
            public readonly List<VbeDebugWindows.SignatureChild> Items =
                new List<VbeDebugWindows.SignatureChild>();
            public bool Open = true, CloseAfterCancel, FailChildren;
            public int Pauses, Closes, CancelIndex, ClosePolls;
            private bool cancelled;
            public IntPtr Dialog()
            {
                if (cancelled) ClosePolls++;
                return Open && !(cancelled && CloseAfterCancel) ? new IntPtr(2) : IntPtr.Zero;
            }
            public IList<VbeDebugWindows.SignatureChild> Children(IntPtr dialog)
            {
                if (FailChildren) throw new InvalidOperationException("MSAA unavailable");
                return Items;
            }
            public void Cancel(IntPtr dialog, int index) { CancelIndex = index; cancelled = true; }
            public void Close(IntPtr dialog) { Closes++; Open = false; }
            public void Pause(int milliseconds) { Pauses++; }
        }
    }
}

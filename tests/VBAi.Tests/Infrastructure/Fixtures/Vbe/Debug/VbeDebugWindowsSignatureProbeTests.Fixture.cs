namespace VBAi.Tests.Unit
{
    using System;
    using System.Collections.Generic;
    using VBAi;
    using Microsoft.VisualStudio.TestTools.UnitTesting;

    public sealed partial class VbeDebugWindowsSignatureProbeTests
    {
        private sealed class SignatureFake : VbeDebugWindows.ISignatureProbe
        {
            public readonly List<VbeDebugWindows.SignatureChild> Items = new List<VbeDebugWindows.SignatureChild>();
            public bool Open = true, CloseAfterCancel, FailChildren;
            public int Pauses, Closes, CancelIndex, ClosePolls;
            private bool cancelled;
            public IntPtr Dialog()
            {
                if (cancelled)
                    ClosePolls++;
                return Open && !(cancelled && CloseAfterCancel) ? new IntPtr(2) : IntPtr.Zero;
            }

            public IList<VbeDebugWindows.SignatureChild> Children(IntPtr dialog)
            {
                if (FailChildren)
                    throw new InvalidOperationException("MSAA unavailable");
                return Items;
            }

            public void Cancel(IntPtr dialog, int index)
            {
                CancelIndex = index;
                cancelled = true;
            }

            public void Close(IntPtr dialog)
            {
                Closes++;
                Open = false;
            }

            public void Pause(int milliseconds)
            {
                Pauses++;
            }
        }
    }
}

using System;
using System.Runtime.InteropServices;
using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace VBAi.Tests.Unit
{
    [TestClass, TestCategory("Unit")]
    public sealed class OwnedTeardownTraceHelperTests
    {
        [TestMethod]
        public void X64ContextLayoutAndAlignmentMatchTheSdkContract()
        {
            Assert.AreEqual(8, IntPtr.Size);
            Assert.AreEqual(1232, Marshal.SizeOf(typeof(OwnedTeardownTraceHelper.Context64)));
            Assert.AreEqual(48, Marshal.OffsetOf(typeof(OwnedTeardownTraceHelper.Context64), "Flags").ToInt32());
            Assert.AreEqual(152, Marshal.OffsetOf(typeof(OwnedTeardownTraceHelper.Context64), "StackPointer").ToInt32());
            Assert.AreEqual(248, Marshal.OffsetOf(typeof(OwnedTeardownTraceHelper.Context64), "InstructionPointer").ToInt32());
            Assert.AreEqual(152, Marshal.SizeOf(typeof(OwnedTeardownTraceHelper.ExceptionRecord)));
            Assert.AreEqual(16, Marshal.OffsetOf(typeof(OwnedTeardownTraceHelper.ExceptionRecord), "Address").ToInt32());
            for (int offset = 0; offset < 16; offset++)
            {
                var allocation = new IntPtr(4096 + offset);
                long aligned = OwnedTeardownTraceHelper.AlignContext(allocation).ToInt64();
                Assert.AreEqual(0L, aligned & 15);
                Assert.IsTrue(aligned >= allocation.ToInt64() && aligned - allocation.ToInt64() <= 15);
            }
            Assert.ThrowsException<InvalidOperationException>(() => OwnedTeardownTraceHelper.AlignContext(IntPtr.Zero));
        }

        [DataTestMethod]
        [DataRow("Complete")]
        [DataRow("AvComplete")]
        [DataRow("NoFlags")]
        [DataRow("NoArchitecture")]
        [DataRow("NoControl")]
        [DataRow("NoInteger")]
        [DataRow("ZeroRip")]
        [DataRow("ZeroRsp")]
        [DataRow("Unaligned")]
        public void SyntheticContextMustBeCompleteBeforeConstructingTheFatalRecord(string variation)
        {
            // Pure memory/layout checks only: never RtlCaptureContext or RaiseFailFastException.
            IntPtr allocation = Marshal.AllocHGlobal(1232 + 15);
            try
            {
                IntPtr context = OwnedTeardownTraceHelper.AlignContext(allocation);
                uint flags = variation == "NoFlags" ? 0U : variation == "NoArchitecture" ? 3U :
                    variation == "NoControl" ? 0x100002U : variation == "NoInteger" ? 0x100001U : 0x10000bU;
                var captured = new OwnedTeardownTraceHelper.Context64 {
                    Flags = flags, InstructionPointer = variation == "ZeroRip" ? 0UL : 0x102030UL,
                    StackPointer = variation == "ZeroRsp" ? 0UL : 0x405060UL };
                Marshal.StructureToPtr(captured, context, false);
                if (variation != "Complete" && variation != "AvComplete")
                {
                    IntPtr invalid = variation == "Unaligned" ? IntPtr.Add(context, 1) : context;
                    Assert.ThrowsException<InvalidOperationException>(() => OwnedTeardownTraceHelper.CreateRecord(invalid));
                    return;
                }
                uint code = variation == "AvComplete" ? 0xc0000005U : 0xc0000409U;
                var record = OwnedTeardownTraceHelper.CreateRecord(context, code);
                Assert.AreEqual(code, record.Code);
                Assert.AreEqual(1U, record.Flags);
                Assert.AreEqual(new IntPtr(0x102030), record.Address);
                Assert.AreEqual(variation == "AvComplete" ? 2U : 1U, record.NumberParameters);
                Assert.AreEqual(15, record.Information.Length);
                Assert.AreEqual(variation == "AvComplete" ? 0UL : 7UL, record.Information[0].ToUInt64());
                Assert.AreEqual(variation == "AvComplete" ? 1UL : 0UL, record.Information[1].ToUInt64());
                Assert.ThrowsException<InvalidOperationException>(() => OwnedTeardownTraceHelper.CreateRecord(context, 0xe0000002));
                Assert.AreEqual(IntPtr.Zero, record.Record);
                Assert.ThrowsException<InvalidOperationException>(() => OwnedTeardownTraceHelper.CreateRecord(IntPtr.Zero));
            }
            finally { Marshal.FreeHGlobal(allocation); }
        }
    }
}

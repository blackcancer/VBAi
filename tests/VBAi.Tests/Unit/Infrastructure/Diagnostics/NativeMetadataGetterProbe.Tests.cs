using Microsoft.VisualStudio.TestTools.UnitTesting;
using System;
using System.Runtime.InteropServices;
using System.Text;
using VBAi.Tests.Integration;

namespace VBAi.Tests.Unit
{
    /// <summary>Checks actual x64 OLE result allocation and byte-preserving BSTR observation without any Office host.</summary>
    [TestClass, TestCategory("Unit")]
    public sealed class NativeMetadataGetterProbeTests
    {
        [TestMethod]
        public void FullX64VariantUnionAndI4ResultPreserveCanary()
        {
            var row = NativeMetadataGetterProbe.ReadVariant(pointer =>
            {
                // A native callee may write all 24 ABI bytes, even for an I4 result.
                Marshal.Copy(new byte[24], 0, pointer, 24);
                Marshal.GetNativeVariantForObject(321, pointer);
                return 0;
            });
            Assert.AreEqual(24, row["BufferBytes"]);
            Assert.AreEqual(true, row["CanaryIntact"]);
            Assert.AreEqual((ushort)VarEnum.VT_I4, row["VariantType"]);
            Assert.AreEqual(321, row["Value"]);
            Assert.AreEqual("READ", row["State"]);
            Assert.AreEqual("0x00000000", row["VariantClearHResult"]);
        }

        [TestMethod]
        public void WideBstrAndPackedAnsiObservationKeepExactOriginalBytesWithoutCorrection()
        {
            string expected = "E:\\D\u00e9veloppement\\OwnedMetadataHelp.chm";
            var wide = NativeMetadataGetterProbe.ReadVariant(pointer => { Marshal.GetNativeVariantForObject(expected, pointer); return 0; });
            Assert.AreEqual(expected, wide["Value"]);
            Assert.AreEqual((uint)(expected.Length * 2), wide["BstrByteLength"]);
            Assert.AreEqual((uint)expected.Length, wide["BstrCharLength"]);
            Assert.AreEqual(BitConverter.ToString(Encoding.Unicode.GetBytes(expected)).Replace("-", ""), wide["BstrBytesHex"]);
            // Mirror the observed byte-packing defect as evidence, never repair it.
            byte[] ansi = Encoding.GetEncoding(1252).GetBytes("E:\\OwnedHelp.chm");
            if ((ansi.Length & 1) != 0) Array.Resize(ref ansi, ansi.Length + 1);
            string packed = Encoding.Unicode.GetString(ansi);
            var row = NativeMetadataGetterProbe.ReadVariant(pointer => { Marshal.GetNativeVariantForObject(packed, pointer); return 0; });
            Assert.AreEqual(packed, row["Value"]);
            Assert.AreNotEqual("E:\\OwnedHelp.chm", row["Value"]);
            Assert.AreEqual(BitConverter.ToString(ansi).Replace("-", ""), row["BstrBytesHex"]);
            Assert.AreEqual(true, row["CanaryIntact"]);
        }

        [TestMethod]
        public void FailedInvokePreservesNativeHresultAndStillClearsAllocatedBstr()
        {
            var row = NativeMetadataGetterProbe.ReadVariant(pointer =>
            {
                Marshal.GetNativeVariantForObject("Owned failure marker", pointer);
                return unchecked((int)0x80020009);
            });
            Assert.AreEqual("INVOKE_FAILED", row["State"]);
            Assert.AreEqual("0x80020009", row["HResult"]);
            Assert.AreEqual("0x00000000", row["VariantClearHResult"]);
            Assert.IsFalse(row.ContainsKey("Value"), "A failed native getter must not become a successful value read.");
        }

        [TestMethod]
        public void BoundaryOverwriteAndUnsupportedByrefRemainExplicitWithoutLosingInvokeStatus()
        {
            var overrun = NativeMetadataGetterProbe.ReadVariant(pointer => { Marshal.WriteInt64(pointer, 24, 0); return 0; });
            Assert.AreEqual(false, overrun["CanaryIntact"]);
            Assert.AreEqual("ERROR", overrun["State"]);
            Assert.AreEqual("0x00000000", overrun["HResult"], "Diagnostic failure must not replace the original Invoke HRESULT.");
            var byref = NativeMetadataGetterProbe.ReadVariant(pointer =>
            {
                Marshal.WriteInt16(pointer, unchecked((short)((int)VarEnum.VT_BYREF | (int)VarEnum.VT_I4)));
                return 0;
            });
            Assert.AreEqual("ERROR", byref["State"]);
            Assert.IsFalse(byref.ContainsKey("Value"), "Unsupported pointer results must not be dereferenced.");
            Assert.AreEqual("0x00000000", byref["VariantClearHResult"]);
        }
    }
}

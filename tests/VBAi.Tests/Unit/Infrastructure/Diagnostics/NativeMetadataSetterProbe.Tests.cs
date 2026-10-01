using System;
using System.Collections.Generic;
using System.Globalization;
using System.Runtime.InteropServices;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using VBAi.Tests.Integration;

namespace VBAi.Tests.Unit
{
    /// <summary>Checks x64 input/result ownership and exact named PROPERTYPUT arguments without an Office host.</summary>
    [TestClass, TestCategory("Unit")]
    public sealed class NativeMetadataSetterProbeTests
    {
        [TestMethod]
        public void I4PutUsesOneNamedArgumentFullBuffersInvariantLocaleAndOneInvoke()
        {
            int entries = 0;
            var row = NativeMetadataSetterProbe.PutRawCore("HelpContextID", 321, (parameters, result) => {
                entries++;
                Assert.AreEqual(1, parameters.cArgs); Assert.AreEqual(1, parameters.cNamedArgs);
                Assert.AreEqual(-3, Marshal.ReadInt32(parameters.rgdispidNamedArgs));
                Assert.AreEqual((short)VarEnum.VT_I4, Marshal.ReadInt16(parameters.rgvarg));
                Assert.AreEqual(321, Marshal.GetObjectForNativeVariant(parameters.rgvarg));
                Marshal.WriteInt64(parameters.rgvarg, 16, 0x18374625); // Exercise the tail of a complete x64 argument.
                Marshal.Copy(new byte[24], 0, result, 24); // A setter may initialize the entire ignored result union.
                return 0;
            });
            Assert.AreEqual(1, entries); Assert.AreEqual(1, row["InvokeEntries"]);
            Assert.AreEqual(117, row["Dispid"]); Assert.AreEqual(CultureInfo.InvariantCulture.LCID, row["Locale"]);
            Assert.AreEqual("DISPATCH_PROPERTYPUT only", row["Flags"]); Assert.AreEqual(false, row["MutationRetryAllowed"]);
            var argument = (IDictionary<string, object>)row["ArgumentVariant"];
            var returned = (IDictionary<string, object>)row["ReturnVariant"];
            Assert.AreEqual(321, argument["Value"]); Assert.AreEqual(24, argument["BufferBytes"]);
            Assert.AreEqual(true, argument["CanaryIntact"]); Assert.AreEqual(true, returned["CanaryIntact"]);
            Assert.AreEqual("0x00000000", argument["VariantClearHResult"]); Assert.AreEqual("0x00000000", returned["VariantClearHResult"]);
            Assert.AreEqual(-3, row["NamedDispidAfterInvoke"]);
        }

        [TestMethod]
        public void BstrInputAndFailedSetterKeepOriginalHresultAndReleaseBothBuffersWithoutRetry()
        {
            int entries = 0;
            string expected = "E:\\D\u00e9veloppement\\OwnedMetadataHelp.chm";
            var row = NativeMetadataSetterProbe.PutRawCore("HelpFile", expected, (parameters, result) => {
                entries++; Assert.AreEqual(expected, Marshal.GetObjectForNativeVariant(parameters.rgvarg));
                return unchecked((int)0xD09072B8); // Preserve the observed failure class; do not interpret its apparent pointer bits.
            });
            var argument = (IDictionary<string, object>)row["ArgumentVariant"];
            var returned = (IDictionary<string, object>)row["ReturnVariant"];
            Assert.AreEqual(1, entries); Assert.AreEqual(116, row["Dispid"]);
            Assert.AreEqual(expected, argument["Value"]); Assert.AreEqual((ushort)VarEnum.VT_BSTR, argument["VariantType"]);
            Assert.AreEqual((uint)(expected.Length * 2), argument["BstrByteLength"]);
            Assert.AreEqual("INVOKE_FAILED", returned["State"]); Assert.AreEqual("0xD09072B8", returned["HResult"]);
            Assert.AreEqual("0x00000000", argument["VariantClearHResult"]); Assert.AreEqual("0x00000000", returned["VariantClearHResult"]);
        }

        [TestMethod]
        public void UnsupportedPropertyTypeOrCoercionIsRejectedBeforeInvoke()
        {
            int entries = 0;
            foreach (var item in new[] { new { Name = "HelpContextID", Value = (object)"321" },
                new { Name = "HelpFile", Value = (object)321 }, new { Name = "Name", Value = (object)"Changed" },
                new { Name = "HelpFile", Value = (object)new string('x', 4097) } })
                Assert.ThrowsException<ArgumentException>(() => NativeMetadataSetterProbe.PutRawCore(item.Name, item.Value,
                    (parameters, result) => { entries++; return 0; }));
            Assert.AreEqual(0, entries);
        }
    }
}

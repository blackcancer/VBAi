using Microsoft.VisualStudio.TestTools.UnitTesting;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;

namespace VBAi.Tests.Unit
{
    public sealed partial class FormResourcePreflightTests
    {
        // Direct private-helper qualification, not additional malformed-public-file reachability.
        // The real reader/state are initialized by the production Read+Validate methods; fields are never rewritten.
        [DataTestMethod]
        [DataRow("U16", 2, 0x1122UL)]
        [DataRow("U32", 4, 0x11223344UL)]
        [DataRow("U64", 8, 0x1122334455667788UL)]
        public void PrivateFieldReadersCannotEscapeTheCfbRegionIntoExistingEnvelopeOrTrailingBytes(string method, int width, ulong lastValue)
        {
            int length; byte[] bytes = PrivateContractResource(out length), original = (byte[])bytes.Clone();
            object compound = ReadValidatedCompound(bytes);
            Assert.AreEqual(lastValue, Convert.ToUInt64(CallCompound(compound, method, length - width)), "Exact last valid field must be readable.");
            Assert.AreEqual(BitConverterValue(original, 24, width), Convert.ToUInt64(CallCompound(compound, method, 0)),
                "The first field is the actual CFB signature, not the preceding LB envelope.");
            foreach (int invalid in new[] { -1, length - width + 1, length, int.MaxValue })
                AssertPrivateCompoundRefusal(() => CallCompound(compound, method, invalid));
            // -1 and one-byte-overrun are valid offsets in the surrounding byte[]: only the CFB region guard excludes them.
            Assert.IsTrue(24 - 1 >= 0 && 24 + length - width + 1 + width <= bytes.Length);
            CollectionAssert.AreEqual(original, bytes, "Header, envelope, logical bytes and trailing sentinels must remain exact.");
        }

        [DataTestMethod]
        [DataRow(0u, 512)]
        [DataRow(4u, 2560)]
        [DataRow(5u, -1)]
        [DataRow(uint.MaxValue, -1)]
        public void PrivateSectorAddressRequiresAnActualSectorWithinTheDeclaredCfbRegion(uint sector, int expected)
        {
            int length; byte[] bytes = PrivateContractResource(out length), original = (byte[])bytes.Clone();
            object compound = ReadValidatedCompound(bytes);
            if (expected >= 0)
            {
                Assert.AreEqual(expected, CallCompound(compound, "Sector", sector));
                Assert.IsTrue(expected >= 512 && expected + 512 <= length);
            }
            else AssertPrivateCompoundRefusal(() => CallCompound(compound, "Sector", sector));
            CollectionAssert.AreEqual(original, bytes);
        }

        [DataTestMethod]
        [DataRow(-1L)]
        [DataRow(long.MinValue)]
        [DataRow(3073L)]
        [DataRow(long.MaxValue)]
        public void PrivateChainRejectsInvalidSizeBeforeReservingAnOtherwiseFreeSector(long size)
        {
            foreach (uint start in new[] { 0xfffffffeu, 4u })
            {
                int length; byte[] bytes = PrivateContractResource(out length), original = (byte[])bytes.Clone();
                Assert.AreEqual(3072, length); Assert.AreEqual(uint.MaxValue, BitConverter.ToUInt32(bytes, 24 + 512 + 4 * 4));
                object compound = ReadValidatedCompound(bytes);
                AssertPrivateCompoundRefusal(() => CallCompound(compound, "Chain", start, (long?)size));
                // No field/state injection: a real subsequent claim proves the failed size check did not reserve sector4.
                Assert.IsNull(CallCompound(compound, "Claim", 4u));
                AssertPrivateCompoundRefusal(() => CallCompound(compound, "Claim", 4u));
                CollectionAssert.AreEqual(original, bytes, "Failed chain and claims must not write resource or sentinel bytes.");
            }
        }

        [DataTestMethod]
        [DataRow(false)]
        [DataRow(true)]
        public void PrivateChainAcceptsTheEndMarkerForAnUnsizedOrExactlyEmptyChain(bool sized)
        {
            int length; byte[] bytes = PrivateContractResource(out length), original = (byte[])bytes.Clone();
            object compound = ReadValidatedCompound(bytes);
            var result = (IEnumerable<uint>)CallCompound(compound, "Chain", 0xfffffffeu, sized ? (long?)0 : null);
            Assert.AreEqual(0, result.Count());
            Assert.IsNull(CallCompound(compound, "Claim", 4u));
            AssertPrivateCompoundRefusal(() => CallCompound(compound, "Claim", 4u));
            CollectionAssert.AreEqual(original, bytes);
        }

        private static byte[] PrivateContractResource(out int length)
        {
            byte[] resource = ExtendResource(5); length = (int)BitConverter.ToUInt32(resource, 4);
            Buffer.BlockCopy(BitConverter.GetBytes(0x1122334455667788UL), 0, resource, resource.Length - 8, 8);
            var bytes = new byte[resource.Length + 16]; Buffer.BlockCopy(resource, 0, bytes, 0, resource.Length);
            for (int i = resource.Length; i < bytes.Length; i++) bytes[i] = (byte)(0xa0 + i - resource.Length);
            return bytes;
        }

        private static object ReadValidatedCompound(byte[] bytes)
        {
            var read = typeof(FormResourcePreflight).GetMethod("Read", BindingFlags.Static | BindingFlags.NonPublic);
            Assert.IsNotNull(read); object compound = read.Invoke(null, new object[] { bytes, 0 });
            CallCompound(compound, "Validate"); return compound;
        }

        private static object CallCompound(object compound, string name, params object[] values)
        {
            var method = compound.GetType().GetMethod(name, BindingFlags.Instance | BindingFlags.NonPublic);
            Assert.IsNotNull(method, "The production helper is absent: " + name); return method.Invoke(compound, values);
        }

        private static void AssertPrivateCompoundRefusal(Action call)
        {
            var wrapped = Assert.ThrowsException<TargetInvocationException>(call);
            Assert.IsNotNull(wrapped.InnerException);
            Assert.AreEqual(typeof(InvalidOperationException), wrapped.InnerException.GetType(),
                "Boundary violations require the controlled CFB refusal, not an index, overflow or null-reference exception.");
            Assert.AreEqual("Invalid or truncated UserForm OLE resource container.", wrapped.InnerException.Message);
        }

        private static ulong BitConverterValue(byte[] bytes, int at, int width)
        { return width == 2 ? BitConverter.ToUInt16(bytes, at) : width == 4 ? BitConverter.ToUInt32(bytes, at) : BitConverter.ToUInt64(bytes, at); }
    }
}

using System;
using System.Linq;
using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace VBAi.Tests.Unit
{
    [TestClass, TestCategory("Unit")]
    public sealed class FormFontRestorationTests
    {
        [TestMethod]
        public void ExactNativeDescriptorsAndDocumentedBoundariesValidateWithoutFontActivation()
        {
            var bindings = FormStreamPadding.ReadFontBindings(FormStreamPaddingTests.ContainerStreamsBefore(), FormStreamPaddingTests.ContainerMetadata());
            foreach (var binding in bindings) FormFontRestoration.ValidateDescriptor(binding.Descriptor);
            var emptyFace = new byte[] { 1, 255, 255, 14, 232, 3, 1, 0, 0, 0, 0 };
            FormFontRestoration.ValidateDescriptor(emptyFace);
            Array.Copy(BitConverter.GetBytes(655350000u), 0, emptyFace, 6, 4);
            FormFontRestoration.ValidateDescriptor(emptyFace);
        }

        [TestMethod]
        public void DescriptorTruncationTrailingDataUnknownFlagsAndInvalidMetricFieldsAreRefused()
        {
            byte[] valid = FormStreamPadding.ReadFontBindings(FormStreamPaddingTests.ContainerStreamsBefore(), FormStreamPaddingTests.ContainerMetadata())[0].Descriptor;
            Assert.ThrowsException<InvalidOperationException>(() => FormFontRestoration.ValidateDescriptor(null));
            for (int length = 0; length < valid.Length; length++)
                Assert.ThrowsException<InvalidOperationException>(() => FormFontRestoration.ValidateDescriptor(valid.Take(length).ToArray()));
            Assert.ThrowsException<InvalidOperationException>(() => FormFontRestoration.ValidateDescriptor(valid.Concat(new byte[] { 0 }).ToArray()));
            foreach (int position in new[] { 0, 3, 4, 6, 10, 11 })
            {
                byte[] data = (byte[])valid.Clone();
                if (position == 0) data[0] = 2;
                else if (position == 3) data[3] = 1;
                else if (position == 4) Array.Copy(BitConverter.GetBytes((ushort)1001), 0, data, 4, 2);
                else if (position == 6) Array.Clear(data, 6, 4);
                else if (position == 10) data[10] = 32;
                else data[position] = 128;
                Assert.ThrowsException<InvalidOperationException>(() => FormFontRestoration.ValidateDescriptor(data));
            }
            byte[] high = (byte[])valid.Clone(); Array.Copy(BitConverter.GetBytes(655350001u), 0, high, 6, 4);
            Assert.ThrowsException<InvalidOperationException>(() => FormFontRestoration.ValidateDescriptor(high));
        }
    }
}

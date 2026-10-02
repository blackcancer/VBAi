using System;
using System.Linq;
using System.Runtime.InteropServices;
using System.Runtime.InteropServices.ComTypes;
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

        [TestMethod]
        public void RootBindingUsesComponentFontPropertyObjectOnceAndPreservesExactDescriptor()
        {
            var component = new FakeComponent(); int checks = 0;
            var root = RootBinding();

            FormFontRestoration.Restore(component, new[] { root }, () => checks++);

            Assert.IsTrue(checks >= 2, "Identity must be rechecked immediately before the native setter.");
            Assert.AreEqual(1, component.Properties.ItemCalls);
            Assert.AreEqual(1, component.Properties.FontProperty.ObjectSetterCalls);
            Assert.AreEqual(0, component.Designer.FontSetterCalls);
            CollectionAssert.AreEqual(root.Descriptor, component.Properties.FontProperty.SavedDescriptor);
        }

        [TestMethod]
        public void RootBindingRefusesWrongTypeNameAndIndexedPropertyBeforeAnySetter()
        {
            foreach (var invalid in new[] { "type", "name", "index" })
            {
                var component = new FakeComponent();
                var binding = RootBinding();
                if (invalid == "type") binding = new FormStreamPadding.FormFontBinding("", binding.Descriptor, 14);
                if (invalid == "name") component.Properties.FontProperty.Name = "Caption";
                if (invalid == "index") component.Properties.FontProperty.NumIndices = 1;

                Assert.ThrowsException<InvalidOperationException>(() => FormFontRestoration.Restore(component,
                    new[] { binding }, () => { }), invalid);
                Assert.AreEqual(0, component.Properties.FontProperty.ObjectSetterCalls, invalid);
                Assert.AreEqual(0, component.Designer.FontSetterCalls, invalid);
            }
        }

        [TestMethod]
        public void LaterMalformedBindingRefusesBeforeEarlierRootSetter()
        {
            var component = new FakeComponent();
            var root = RootBinding();
            var malformed = new FormStreamPadding.FormFontBinding("", root.Descriptor, 14);

            Assert.ThrowsException<InvalidOperationException>(() => FormFontRestoration.Restore(component,
                new[] { root, malformed }, () => { }));

            Assert.AreEqual(0, component.Properties.FontProperty.ObjectSetterCalls);
            Assert.AreEqual(0, component.Designer.FontSetterCalls);
        }

        [TestMethod]
        public void RootPropertyObjectFailureIsOneDeliveryWithoutDesignerFallback()
        {
            var component = new FakeComponent();
            component.Properties.FontProperty.ThrowOnSet = true;

            Assert.ThrowsException<InvalidOperationException>(() => FormFontRestoration.Restore(component,
                new[] { RootBinding() }, () => { }));

            Assert.AreEqual(1, component.Properties.FontProperty.ObjectSetterCalls);
            Assert.AreEqual(0, component.Designer.FontSetterCalls);
        }

        [TestMethod]
        public void RootRevalidationFailureBeforeAssignmentLeavesBothFontRoutesUntouched()
        {
            var component = new FakeComponent(); int checks = 0;

            Assert.ThrowsException<InvalidOperationException>(() => FormFontRestoration.Restore(component,
                new[] { RootBinding() }, () => { if (++checks == 2) throw new InvalidOperationException("Identity changed."); }));

            Assert.AreEqual(2, checks);
            Assert.AreEqual(0, component.Properties.FontProperty.ObjectSetterCalls);
            Assert.AreEqual(0, component.Designer.FontSetterCalls);
        }

        private static FormStreamPadding.FormFontBinding RootBinding()
        {
            return FormStreamPadding.ReadFontBindings(FormStreamPaddingTests.ContainerStreamsBefore(),
                FormStreamPaddingTests.ContainerMetadata()).Single(binding => binding.OwnerPath == "");
        }

        // Public fake members are required for the production dynamic COM boundary.
        public sealed class FakeComponent
        {
            public FakeDesigner Designer { get; } = new FakeDesigner();
            public FakeProperties Properties { get; } = new FakeProperties();
        }

        public sealed class FakeDesigner
        {
            public int FontSetterCalls { get; private set; }
            public object Font { set { FontSetterCalls++; } }
        }

        public sealed class FakeProperties
        {
            public int ItemCalls { get; private set; }
            public FakeFontProperty FontProperty { get; } = new FakeFontProperty();
            public FakeFontProperty Item(string name)
            {
                ItemCalls++;
                Assert.AreEqual("Font", name);
                return FontProperty;
            }
        }

        public sealed class FakeFontProperty
        {
            public string Name { get; set; } = "Font";
            public int NumIndices { get; set; }
            public int ObjectSetterCalls { get; private set; }
            public bool ThrowOnSet { get; set; }
            public byte[] SavedDescriptor { get; private set; }
            public object Object
            {
                set
                {
                    ObjectSetterCalls++;
                    if (ThrowOnSet) throw new InvalidOperationException("The native property setter failed.");
                    // Save while the production method owns the font RCW. Do not release that borrowed RCW.
                    IStream stream = null; IntPtr readCount = IntPtr.Zero;
                    try
                    {
                        Marshal.ThrowExceptionForHR(CreateStreamOnHGlobal(IntPtr.Zero, true, out stream));
                        ((PersistStream)value).Save(stream, false);
                        stream.Seek(0, 0, IntPtr.Zero);
                        SavedDescriptor = new byte[RootBinding().Descriptor.Length];
                        readCount = Marshal.AllocHGlobal(sizeof(int));
                        stream.Read(SavedDescriptor, SavedDescriptor.Length, readCount);
                        Assert.AreEqual(SavedDescriptor.Length, Marshal.ReadInt32(readCount));
                    }
                    finally
                    {
                        if (readCount != IntPtr.Zero) Marshal.FreeHGlobal(readCount);
                        if (stream != null) Marshal.ReleaseComObject(stream);
                    }
                }
            }
        }

        [DllImport("ole32.dll", ExactSpelling = true)]
        private static extern int CreateStreamOnHGlobal(IntPtr handle, [MarshalAs(UnmanagedType.Bool)] bool free, out IStream stream);

        [ComImport, Guid("00000109-0000-0000-C000-000000000046"), InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
        private interface PersistStream
        {
            void GetClassID(out Guid clsid);
            [PreserveSig] int IsDirty();
            void Load(IStream stream);
            void Save(IStream stream, [MarshalAs(UnmanagedType.Bool)] bool clearDirty);
            void GetSizeMax(out long size);
        }
    }
}

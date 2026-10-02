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
        public void RootPropertyObjectComFailureReportsExactSetterAndNeverFallsBack()
        {
            var component = new FakeComponent();
            var failure = InjectedComFailure();
            component.Properties.FontProperty.ObjectSetterFailure = failure;

            var error = Assert.ThrowsException<InvalidOperationException>(() => FormFontRestoration.Restore(component,
                new[] { RootBinding() }, () => { }));

            AssertComProvenance(error, "VBIDE.Property.Object.set", failure);
            Assert.AreEqual(1, component.Properties.FontProperty.ObjectSetterCalls);
            Assert.AreEqual(0, component.Designer.FontSetterCalls);
        }

        [TestMethod]
        public void RootPropertyObjectMappedUnsupportedFailureReportsStageAndRetainsOriginalWithoutRetry()
        {
            var component = new FakeComponent();
            var failure = new NotSupportedException("Mapped unsupported native operation.");
            component.Properties.FontProperty.ObjectSetterFailure = failure;

            var error = Assert.ThrowsException<InvalidOperationException>(() => FormFontRestoration.Restore(component,
                new[] { RootBinding() }, () => { }));

            Assert.AreSame(failure, error.InnerException);
            StringAssert.Contains(error.Message, "VBIDE.Property.Object.set");
            StringAssert.Contains(error.Message, "HRESULT 0x80131515");
            Assert.AreEqual(1, component.Properties.FontProperty.ObjectSetterCalls);
            Assert.AreEqual(0, component.Designer.FontSetterCalls);
        }

        [TestMethod]
        public void RootComGetterFailuresReportExactOperationBeforeAnyFontSetter()
        {
            foreach (string stage in new[] { "Designer", "Properties", "Item", "Name", "NumIndices" })
            {
                var component = new FakeComponent();
                var failure = InjectedComFailure();
                string operation;
                switch (stage)
                {
                    case "Designer": component.DesignerGetFailure = failure; operation = "VBComponent.Designer.get"; break;
                    case "Properties": component.PropertiesGetFailure = failure; operation = "VBComponent.Properties.get"; break;
                    case "Item": component.Properties.ItemFailure = failure; operation = "VBIDE.Properties.Item(Font)"; break;
                    case "Name": component.Properties.FontProperty.NameGetFailure = failure; operation = "VBIDE.Property.Name.get"; break;
                    default: component.Properties.FontProperty.NumIndicesGetFailure = failure; operation = "VBIDE.Property.NumIndices.get"; break;
                }

                var error = Assert.ThrowsException<InvalidOperationException>(() => FormFontRestoration.Restore(component,
                    new[] { RootBinding() }, () => { }), stage);

                AssertComProvenance(error, operation, failure);
                Assert.AreEqual(0, component.RootFontProperty.ObjectSetterCalls, stage);
                Assert.AreEqual(0, component.DesignerFontSetterCalls, stage);
            }
        }

        [TestMethod]
        public void RootRevalidationFailureBeforeAssignmentLeavesBothFontRoutesUntouched()
        {
            var component = new FakeComponent(); int checks = 0;

            var error = Assert.ThrowsException<InvalidOperationException>(() => FormFontRestoration.Restore(component,
                new[] { RootBinding() }, () => { if (++checks == 2) throw new InvalidOperationException("Identity changed."); }));

            Assert.AreEqual("Identity changed.", error.Message);
            Assert.IsNull(error.InnerException);
            Assert.AreEqual(2, checks);
            Assert.AreEqual(0, component.Properties.FontProperty.ObjectSetterCalls);
            Assert.AreEqual(0, component.Designer.FontSetterCalls);
        }

        private static FormStreamPadding.FormFontBinding RootBinding()
        {
            return FormStreamPadding.ReadFontBindings(FormStreamPaddingTests.ContainerStreamsBefore(),
                FormStreamPaddingTests.ContainerMetadata()).Single(binding => binding.OwnerPath == "");
        }

        private static COMException InjectedComFailure()
        {
            return new COMException("Injected native COM failure.", unchecked((int)0x80020003));
        }

        private static void AssertComProvenance(InvalidOperationException error, string operation, COMException failure)
        {
            StringAssert.Contains(error.Message, operation);
            StringAssert.Contains(error.Message, "0x80020003");
            Assert.AreSame(failure, error.InnerException, "Preserve the exact native COM exception for diagnosis.");
            Assert.AreEqual(unchecked((int)0x80020003), ((COMException)error.InnerException).ErrorCode);
        }

        // Public fake members are required for the production dynamic COM boundary.
        public sealed class FakeComponent
        {
            private readonly FakeDesigner designer = new FakeDesigner();
            private readonly FakeProperties properties = new FakeProperties();
            public COMException DesignerGetFailure { get; set; }
            public COMException PropertiesGetFailure { get; set; }
            public FakeDesigner Designer
            {
                get { if (DesignerGetFailure != null) throw DesignerGetFailure; return designer; }
            }
            public FakeProperties Properties
            {
                get { if (PropertiesGetFailure != null) throw PropertiesGetFailure; return properties; }
            }
            public FakeFontProperty RootFontProperty { get { return properties.FontProperty; } }
            public int DesignerFontSetterCalls { get { return designer.FontSetterCalls; } }
        }

        public sealed class FakeDesigner
        {
            public int FontSetterCalls { get; private set; }
            public object Font { set { FontSetterCalls++; } }
        }

        public sealed class FakeProperties
        {
            public int ItemCalls { get; private set; }
            public COMException ItemFailure { get; set; }
            public FakeFontProperty FontProperty { get; } = new FakeFontProperty();
            public FakeFontProperty Item(string name)
            {
                ItemCalls++;
                if (ItemFailure != null) throw ItemFailure;
                Assert.AreEqual("Font", name);
                return FontProperty;
            }
        }

        public sealed class FakeFontProperty
        {
            private string name = "Font";
            private int numIndices;
            public COMException NameGetFailure { get; set; }
            public COMException NumIndicesGetFailure { get; set; }
            public Exception ObjectSetterFailure { get; set; }
            public string Name
            {
                get { if (NameGetFailure != null) throw NameGetFailure; return name; }
                set { name = value; }
            }
            public int NumIndices
            {
                get { if (NumIndicesGetFailure != null) throw NumIndicesGetFailure; return numIndices; }
                set { numIndices = value; }
            }
            public int ObjectSetterCalls { get; private set; }
            public bool ThrowOnSet { get; set; }
            public byte[] SavedDescriptor { get; private set; }
            public object Object
            {
                set
                {
                    ObjectSetterCalls++;
                    if (ObjectSetterFailure != null) throw ObjectSetterFailure;
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

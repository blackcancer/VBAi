using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
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
            byte[] valid = RootBinding().Descriptor;
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
        public void RootBindingWritesSevenExactScalarsWithWeightLastAndNeverUsesObjectOrBold()
        {
            var component = new FakeComponent(); int checks = 0;
            var binding = RootBinding();
            byte[] descriptor = (byte[])binding.Descriptor.Clone();
            Array.Copy(BitConverter.GetBytes(82700u), 0, descriptor, 6, 4);
            Array.Copy(BitConverter.GetBytes((ushort)525), 0, descriptor, 4, 2);
            Array.Copy(BitConverter.GetBytes((short)204), 0, descriptor, 1, 2);
            descriptor[3] = 14;

            FormFontRestoration.Restore(component,
                new[] { new FormStreamPadding.FormFontBinding("", descriptor, binding.Type) }, () => checks++);

            CollectionAssert.AreEqual(new[] { "Name", "Size", "Charset", "Italic", "Underline", "Strikethrough", "Weight" },
                component.RootFontProperty.Value.Writes.ToArray());
            Assert.AreEqual(8, checks, "Identity is checked before preflight and before every scalar delivery.");
            Assert.AreEqual(8.27m, component.RootFontProperty.Value.Child("Size").Value);
            Assert.AreEqual((short)525, component.RootFontProperty.Value.Child("Weight").Value);
            Assert.AreEqual((short)204, component.RootFontProperty.Value.Child("Charset").Value);
            Assert.AreEqual(true, component.RootFontProperty.Value.Child("Italic").Value);
            Assert.AreEqual(true, component.RootFontProperty.Value.Child("Underline").Value);
            Assert.AreEqual(true, component.RootFontProperty.Value.Child("Strikethrough").Value);
            Assert.AreEqual(0, component.RootFontProperty.Value.Child("Bold").SetterCalls);
            Assert.AreEqual(0, component.RootFontProperty.ObjectSetterCalls);
            Assert.AreEqual(0, component.DesignerFontSetterCalls);
        }

        [TestMethod]
        public void RootFixtureSizeUsesExactDecimal825WithoutBinaryFloatRounding()
        {
            var component = new FakeComponent();
            FormFontRestoration.Restore(component, new[] { RootBinding() }, () => { });
            Assert.AreEqual(8.25m, component.RootFontProperty.Value.Child("Size").Value);
            Assert.IsInstanceOfType(component.RootFontProperty.Value.Child("Size").Value, typeof(decimal));
        }

        [TestMethod]
        public void LaterMalformedBindingAndInvalidChildRefuseBeforeAnySetter()
        {
            var root = RootBinding();
            foreach (string invalid in new[] { "later", "missing", "name", "index", "type" })
            {
                var component = new FakeComponent();
                FormStreamPadding.FormFontBinding[] bindings = { root };
                var child = component.RootFontProperty.Value.Child("Weight");
                if (invalid == "later") bindings = new[] { root, new FormStreamPadding.FormFontBinding("", root.Descriptor, 14) };
                if (invalid == "missing") component.RootFontProperty.Value.Remove("Weight");
                if (invalid == "name") child.Name = "Caption";
                if (invalid == "index") child.NumIndices = 1;
                if (invalid == "type") child.SetObservedValue(525m);
                Assert.ThrowsException<InvalidOperationException>(() => FormFontRestoration.Restore(component, bindings, () => { }), invalid);
                Assert.AreEqual(0, component.RootFontProperty.Value.Writes.Count, invalid);
                Assert.AreEqual(0, component.RootFontProperty.ObjectSetterCalls, invalid);
                Assert.AreEqual(0, component.DesignerFontSetterCalls, invalid);
            }
        }

        [TestMethod]
        public void RootPropertyIdentityRefusesBeforeAnyChildSetter()
        {
            foreach (string invalid in new[] { "type", "name", "index" })
            {
                var component = new FakeComponent();
                var binding = RootBinding();
                if (invalid == "type") binding = new FormStreamPadding.FormFontBinding("", binding.Descriptor, 14);
                if (invalid == "name") component.RootFontProperty.Name = "Caption";
                if (invalid == "index") component.RootFontProperty.NumIndices = 1;
                Assert.ThrowsException<InvalidOperationException>(() =>
                    FormFontRestoration.Restore(component, new[] { binding }, () => { }), invalid);
                Assert.AreEqual(0, component.RootFontProperty.Value.Writes.Count, invalid);
                Assert.AreEqual(0, component.RootFontProperty.ObjectSetterCalls, invalid);
            }
        }

        [TestMethod]
        public void RootNativeGetterFailureReportsExactStageBeforeAnyChildSetter()
        {
            foreach (string stage in new[] { "Designer", "Properties", "Item", "Name", "NumIndices", "Value" })
            {
                var component = new FakeComponent();
                var failure = InjectedComFailure();
                string operation;
                switch (stage)
                {
                    case "Designer": component.DesignerGetFailure = failure; operation = "VBComponent.Designer.get"; break;
                    case "Properties": component.PropertiesGetFailure = failure; operation = "VBComponent.Properties.get"; break;
                    case "Item": component.Properties.ItemFailure = failure; operation = "VBIDE.Properties.Item(Font)"; break;
                    case "Name": component.RootFontProperty.NameGetFailure = failure; operation = "VBIDE.Property.Name.get"; break;
                    case "NumIndices": component.RootFontProperty.NumIndicesGetFailure = failure; operation = "VBIDE.Property.NumIndices.get"; break;
                    default: component.RootFontProperty.ValueGetFailure = failure; operation = "VBIDE.Property.Value.get(Font)"; break;
                }
                var error = Assert.ThrowsException<InvalidOperationException>(() =>
                    FormFontRestoration.Restore(component, new[] { RootBinding() }, () => { }), stage);
                AssertComProvenance(error, operation, failure);
                Assert.AreEqual(0, component.RootFontProperty.ObservedChildren.Writes.Count, stage);
            }
        }

        [TestMethod]
        public void ChildGetterFailureReportsExactStageBeforeAnySetter()
        {
            var component = new FakeComponent();
            var failure = InjectedComFailure();
            component.RootFontProperty.Value.Child("Size").GetFailure = failure;
            var error = Assert.ThrowsException<InvalidOperationException>(() =>
                FormFontRestoration.Restore(component, new[] { RootBinding() }, () => { }));
            AssertComProvenance(error, "VBIDE.Property.Value.get(Font.Size)", failure);
            Assert.AreEqual(0, component.RootFontProperty.Value.Writes.Count);
        }

        [TestMethod]
        public void ChildComSetterFailureIsOneDeliveryWithExactProvenanceAndNoFallback()
        {
            var component = new FakeComponent();
            var failure = InjectedComFailure();
            component.RootFontProperty.Value.Child("Size").SetterFailure = failure;
            var error = Assert.ThrowsException<InvalidOperationException>(() =>
                FormFontRestoration.Restore(component, new[] { RootBinding() }, () => { }));
            AssertComProvenance(error, "VBIDE.Property.Value.set(Font.Size)", failure);
            CollectionAssert.AreEqual(new[] { "Name", "Size" }, component.RootFontProperty.Value.Writes.ToArray());
            Assert.AreEqual(1, component.RootFontProperty.Value.Child("Size").SetterCalls);
            Assert.AreEqual(0, component.RootFontProperty.ObjectSetterCalls);
            Assert.AreEqual(0, component.DesignerFontSetterCalls);
        }

        [TestMethod]
        public void MappedUnsupportedChildSetterRetainsOriginalFailureWithoutRetry()
        {
            var component = new FakeComponent();
            var failure = new NotSupportedException("Mapped unsupported native operation.");
            component.RootFontProperty.Value.Child("Name").SetterFailure = failure;
            var error = Assert.ThrowsException<InvalidOperationException>(() =>
                FormFontRestoration.Restore(component, new[] { RootBinding() }, () => { }));
            Assert.AreSame(failure, error.InnerException);
            StringAssert.Contains(error.Message, "VBIDE.Property.Value.set(Font.Name)");
            StringAssert.Contains(error.Message, "HRESULT 0x80131515");
            CollectionAssert.AreEqual(new[] { "Name" }, component.RootFontProperty.Value.Writes.ToArray());
        }

        [TestMethod]
        public void RevalidationBeforeNextChildStopsRemainingDeliveriesWithoutRetry()
        {
            var component = new FakeComponent(); int checks = 0;
            var error = Assert.ThrowsException<InvalidOperationException>(() =>
                FormFontRestoration.Restore(component, new[] { RootBinding() }, () =>
                { if (++checks == 3) throw new InvalidOperationException("Identity changed."); }));
            Assert.AreEqual("Identity changed.", error.Message);
            Assert.AreEqual(3, checks);
            CollectionAssert.AreEqual(new[] { "Name" }, component.RootFontProperty.Value.Writes.ToArray());
            Assert.AreEqual(0, component.RootFontProperty.ObjectSetterCalls);
            Assert.AreEqual(0, component.DesignerFontSetterCalls);
        }

        [TestMethod]
        public void PlannedDistinctChildNameDeliversTemporaryThenTargetOnceAndConfirmsEachTerminal()
        {
            var children = new FakeChildren();
            object[] properties = RootChildren(children);
            var receipts = new List<string>(); int checks = 0;
            FormFontRestoration.AssignRootCore(properties, RootBinding().Descriptor, () => checks++, "Arial",
                (field, value) => receipts.Add("before:" + field),
                (field, value, property) => {
                    Assert.AreEqual(value, ((FakeChild)property).Value, field);
                    receipts.Add("after:" + field);
                }, () => receipts.Add("after-temporary-snapshot"));
            CollectionAssert.AreEqual(new[] { "Name", "Name", "Size", "Charset", "Italic", "Underline", "Strikethrough", "Weight" },
                children.Writes.ToArray());
            CollectionAssert.AreEqual(new[] { "before:Name.temporary", "after:Name.temporary", "after-temporary-snapshot",
                "before:Name", "after:Name", "before:Size", "after:Size", "before:Charset", "after:Charset",
                "before:Italic", "after:Italic", "before:Underline", "after:Underline",
                "before:Strikethrough", "after:Strikethrough", "before:Weight", "after:Weight" }, receipts.ToArray());
            Assert.AreEqual(8, checks);
            Assert.AreEqual(0, children.Child("Bold").SetterCalls);
        }

        [TestMethod]
        public void FailedOrUnconfirmedTemporaryNameStopsBeforeTargetWithoutFallback()
        {
            foreach (bool setterFails in new[] { true, false })
            {
                var children = new FakeChildren();
                object[] properties = RootChildren(children);
                var failure = InjectedComFailure();
                if (setterFails) children.Child("Name").SetterFailure = failure;
                Exception error = Assert.ThrowsException<InvalidOperationException>(() =>
                    FormFontRestoration.AssignRootCore(properties, RootBinding().Descriptor, () => { }, "Arial",
                        (field, value) => { },
                        (field, value, property) => { if (!setterFails) throw new InvalidOperationException("Temporary readback failed."); },
                        () => Assert.Fail("No snapshot after uncertain temporary setter.")));
                if (setterFails) AssertComProvenance((InvalidOperationException)error,
                    "VBIDE.Property.Value.set(Font.Name.temporary)", failure);
                CollectionAssert.AreEqual(new[] { "Name" }, children.Writes.ToArray());
                Assert.AreEqual(1, children.Child("Name").SetterCalls);
            }
        }

        [TestMethod]
        public void CleanupReleasesEveryOwnedReferenceAndKeepsPrimaryBeforeAllReleaseFailures()
        {
            var references = new object[] { new object(), new object(), new object() };
            var native = new COMException("Native getter failed.", unchecked((int)0x80020003));
            var firstCleanup = new InvalidOperationException("Second acquired reference failed to release.");
            var secondCleanup = new InvalidOperationException("First acquired reference failed to release.");
            var visited = new List<object>();
            var error = Assert.ThrowsException<AggregateException>(() =>
                FormFontRestoration.ReleaseOwnedReferences(references, item => {
                    visited.Add(item);
                    if (ReferenceEquals(item, references[1])) throw firstCleanup;
                    if (ReferenceEquals(item, references[0])) throw secondCleanup;
                }, native));
            CollectionAssert.AreEqual(new[] { references[2], references[1], references[0] }, visited.ToArray());
            CollectionAssert.AreEqual(new Exception[] { native, firstCleanup, secondCleanup }, error.InnerExceptions.ToArray());
        }

        [TestMethod]
        public void SuccessfulCleanupDoesNotReplacePrimaryAndStandaloneCleanupStillReportsFailure()
        {
            var native = new InvalidOperationException("Original native read failed.");
            FormFontRestoration.ReleaseOwnedReferences(new object[] { new object() }, item => { }, native);
            var cleanup = new InvalidOperationException("Release failed.");
            var error = Assert.ThrowsException<AggregateException>(() =>
                FormFontRestoration.ReleaseOwnedReferences(new object[] { new object() }, item => { throw cleanup; }, null));
            Assert.AreSame(cleanup, error.InnerExceptions.Single());
        }

        private static object[] RootChildren(FakeChildren children)
        {
            return new[] { "Name", "Size", "Bold", "Italic", "Underline", "Strikethrough", "Weight", "Charset" }
                .Select(name => (object)children.Child(name)).ToArray();
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
            Assert.AreSame(failure, error.InnerException);
            Assert.AreEqual(unchecked((int)0x80020003), ((COMException)error.InnerException).ErrorCode);
        }

        // Public fake members are needed at the production dynamic COM boundary.
        public sealed class FakeComponent
        {
            private readonly FakeDesigner designer = new FakeDesigner();
            private readonly FakeProperties properties = new FakeProperties();
            public COMException DesignerGetFailure { get; set; }
            public COMException PropertiesGetFailure { get; set; }
            public FakeDesigner Designer { get { if (DesignerGetFailure != null) throw DesignerGetFailure; return designer; } }
            public FakeProperties Properties { get { if (PropertiesGetFailure != null) throw PropertiesGetFailure; return properties; } }
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
            public FakeFontProperty FontProperty { get; } = new FakeFontProperty();
            public COMException ItemFailure { get; set; }
            public FakeFontProperty Item(string name)
            {
                if (ItemFailure != null) throw ItemFailure;
                Assert.AreEqual("Font", name);
                return FontProperty;
            }
        }

        public sealed class FakeFontProperty
        {
            private string name = "Font";
            private int numIndices;
            private readonly FakeChildren children = new FakeChildren();
            public COMException NameGetFailure { get; set; }
            public COMException NumIndicesGetFailure { get; set; }
            public COMException ValueGetFailure { get; set; }
            public FakeChildren ObservedChildren { get { return children; } }
            public string Name { get { if (NameGetFailure != null) throw NameGetFailure; return name; } set { name = value; } }
            public int NumIndices { get { if (NumIndicesGetFailure != null) throw NumIndicesGetFailure; return numIndices; } set { numIndices = value; } }
            public FakeChildren Value { get { if (ValueGetFailure != null) throw ValueGetFailure; return children; } }
            public int ObjectSetterCalls { get; private set; }
            public object Object { set { ObjectSetterCalls++; } }
        }

        public sealed class FakeChildren
        {
            private readonly Dictionary<string, FakeChild> children = new Dictionary<string, FakeChild>(StringComparer.Ordinal);
            public List<string> Writes { get; } = new List<string>();
            public FakeChildren()
            {
                Add("Name", "Tahoma");
                Add("Size", 8.25m);
                Add("Bold", false);
                Add("Italic", false);
                Add("Underline", false);
                Add("Strikethrough", false);
                Add("Weight", (short)400);
                Add("Charset", (short)0);
            }
            private void Add(string name, object value) { children.Add(name, new FakeChild(this, name, value)); }
            public FakeChild Item(string name) { return Child(name); }
            public FakeChild Child(string name)
            {
                if (!children.TryGetValue(name, out FakeChild child))
                    throw new InvalidOperationException("Missing child property " + name);
                return child;
            }
            public void Remove(string name) { children.Remove(name); }
            internal void Record(string name) { Writes.Add(name); }
        }

        public sealed class FakeChild
        {
            private readonly FakeChildren owner;
            private object value;
            public FakeChild(FakeChildren owner, string name, object initialValue)
            { this.owner = owner; Name = name; value = initialValue; }
            public string Name { get; set; }
            public int NumIndices { get; set; }
            public COMException GetFailure { get; set; }
            public Exception SetterFailure { get; set; }
            public int SetterCalls { get; private set; }
            public object Value
            {
                get { if (GetFailure != null) throw GetFailure; return value; }
                set
                {
                    SetterCalls++;
                    owner.Record(Name);
                    if (SetterFailure != null) throw SetterFailure;
                    this.value = value;
                }
            }
            public void SetObservedValue(object observed) { value = observed; }
        }
    }
}

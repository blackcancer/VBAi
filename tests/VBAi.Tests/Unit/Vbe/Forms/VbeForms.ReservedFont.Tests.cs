namespace VBAi.Tests.Unit
{
    using System;
    using System.Collections.Generic;
    using System.ComponentModel;
    using System.Drawing;
    using System.Linq;
    using System.Reflection;
    using System.Runtime.InteropServices.ComTypes;
    using System.Web.Script.Serialization;
    using Microsoft.VisualStudio.TestTools.UnitTesting;
    using VBAi;

    [TestClass]
    [TestCategory("Unit")]
    [DoNotParallelize]
    public sealed class VbeFormsReservedFontTests
    {
        [TestMethod]
        public void CapturedMsFormsWriteOnlyAliasNeverInvokesFormOrControlGetters()
        {
            using (var typeInfo = new ComTypeInfoFixture())
            using (var font = new Font("Tahoma", 8.25f))
            {
                typeInfo.Functions.Add(new ComTypeInfoFixture.Function {
                    Name = "_Font_Reserved", MemberId = 0x7ffffdff,
                    Kind = INVOKEKIND.INVOKE_PROPERTYPUT, Parameters = 1 });
                var alias = new CountingDescriptor("_Font_Reserved", typeof(Font)) { Fail = true };
                var ordinary = new CountingDescriptor("Font", typeof(Font)) { Value = font };
                var owner = new MetadataOwner { Info = typeInfo, Font = font,
                    Metadata = new PropertyDescriptorCollection(new PropertyDescriptor[] { alias, ordinary }) };
                var reservedProperty = new NativeProperty { Name = "_Font_Reserved", Fail = true };
                var form = new NativeForm { Designer = owner, Properties = new[] {
                    reservedProperty, new NativeProperty { Name = "Font", Stored = font } } };
                var properties = (List<VbePropertyInfo>)Read("DescribeProperties", form);
                var controls = (List<VbePropertyInfo>)Read("ReadObjectProperties", owner);
                foreach (var rows in new[] { properties, controls })
                {
                    var reserved = rows.Single(p => p.Name == "_Font_Reserved");
                    Assert.AreEqual(typeof(Font).FullName, reserved.Type);
                    Assert.AreEqual(false, reserved.ReadOnly);
                    Assert.AreEqual("writeOnly", reserved.Kind);
                    Assert.AreEqual("GetterUnavailable", reserved.SetterStatus);
                    Assert.IsNull(reserved.Error); Assert.IsNull(reserved.Value);
                    Assert.IsNull(reserved.Digest); Assert.IsNull(reserved.Members);
                    var actualFont = rows.Single(p => p.Name == "Font");
                    Assert.AreEqual("object", actualFont.Kind); Assert.IsNull(actualFont.Error);
                    Assert.IsNotNull(actualFont.Members);
                }
                Assert.AreEqual(0, alias.Reads);
                Assert.AreEqual(0, reservedProperty.Reads);
                Assert.AreEqual(2, owner.MetadataReads);
                Assert.AreEqual(2, typeInfo.AttributeReleases);
                Assert.AreEqual(2, typeInfo.FunctionReleases);
            }
        }

        [TestMethod]
        public void MissingAmbiguousOrReadableMetadataDoesNotSuppressGetterErrors()
        {
            foreach (int mode in Enumerable.Range(0, 7))
            using (var typeInfo = new ComTypeInfoFixture())
            {
                var function = new ComTypeInfoFixture.Function { Name = "_Font_Reserved",
                    MemberId = 0x7ffffdff, Kind = INVOKEKIND.INVOKE_PROPERTYPUT };
                typeInfo.Functions.Add(function);
                if (mode == 0) function.MemberId = 42;
                if (mode == 1) function.Kind = INVOKEKIND.INVOKE_PROPERTYGET;
                if (mode == 2) function.Kind = INVOKEKIND.INVOKE_FUNC;
                if (mode == 3) typeInfo.Functions.Add(new ComTypeInfoFixture.Function {
                    Name = "_Font_Reserved", MemberId = 0x7ffffdff,
                    Kind = INVOKEKIND.INVOKE_PROPERTYGET });
                if (mode == 4) function.NamesFailure = true;
                var descriptor = new CountingDescriptor("_Font_Reserved",
                    mode == 5 ? typeof(object) : typeof(Font)) { Fail = true };
                var owner = new MetadataOwner { Info = mode == 6 ? null : typeInfo,
                    Metadata = new PropertyDescriptorCollection(new PropertyDescriptor[] { descriptor }) };
                var rows = (List<VbePropertyInfo>)Read("ReadObjectProperties", owner);
                Assert.AreEqual(1, descriptor.Reads, "Mode " + mode);
                Assert.IsNotNull(rows.Single().Error, "Mode " + mode);
                Assert.AreNotEqual("writeOnly", rows.Single().Kind);
            }
        }

        [TestMethod]
        public void OrdinaryFontAndUnknownResourceErrorsRemainPublicationBlockers()
        {
            foreach (string name in new[] { "Font", "Picture", "UnrecognizedResource" })
            {
                var descriptor = new CountingDescriptor(name, typeof(Font)) { Fail = true };
                var owner = new MetadataOwner { Metadata = new PropertyDescriptorCollection(new PropertyDescriptor[] { descriptor }) };
                var properties = (List<VbePropertyInfo>)Read("ReadObjectProperties", owner);
                Assert.AreEqual(1, descriptor.Reads); Assert.IsNotNull(properties.Single().Error);
                var serializer = new JavaScriptSerializer();
                var tree = serializer.DeserializeObject(serializer.Serialize(new { Properties = properties }));
                Assert.ThrowsException<InvalidOperationException>(() => VbeProjectComponents.RequirePublicationDesignerReadable(tree));
            }
        }

        private static object Read(string name, object value)
        {
            try { return typeof(VbeForms).GetMethod(name, BindingFlags.Static | BindingFlags.NonPublic).Invoke(null, new[] { value }); }
            catch (TargetInvocationException e) { System.Runtime.ExceptionServices.ExceptionDispatchInfo.Capture(e.InnerException).Throw(); throw; }
        }
        public sealed class NativeForm { public MetadataOwner Designer; public NativeProperty[] Properties; }
        public sealed class NativeProperty
        {
            public string Name; public int NumIndices => 0; public object Stored; public bool Fail; public int Reads;
            public object Value { get { Reads++; if (Fail) throw new InvalidOperationException("Reserved property has no getter"); return Stored; } }
        }
        public sealed class CountingDescriptor : PropertyDescriptor
        {
            readonly Type type;
            public CountingDescriptor(string name, Type type) : base(name, null) { this.type = type; }
            public object Value; public bool Fail; public int Reads;
            public override Type ComponentType => typeof(MetadataOwner);
            public override Type PropertyType => type;
            public override bool IsReadOnly => false;
            public override object GetValue(object component) { Reads++; if (Fail) throw new InvalidOperationException("Native getter unavailable"); return Value; }
            public override void SetValue(object component, object value) { throw new InvalidOperationException("No setter permitted in inspection test"); }
            public override bool CanResetValue(object component) => false;
            public override bool ShouldSerializeValue(object component) => false;
            public override void ResetValue(object component) { }
        }
        public sealed class MetadataOwner : ICustomTypeDescriptor, IProvideClassInfo
        {
            public ITypeInfo Info; public object Font; public int MetadataReads;
            public PropertyDescriptorCollection Metadata;
            public void GetClassInfo(out ITypeInfo value) { MetadataReads++; value = Info; }
            public AttributeCollection GetAttributes() => AttributeCollection.Empty;
            public string GetClassName() => "Label";
            public string GetComponentName() => "Q020Label";
            public TypeConverter GetConverter() => new TypeConverter();
            public EventDescriptor GetDefaultEvent() => null;
            public PropertyDescriptor GetDefaultProperty() => null;
            public object GetEditor(Type type) => null;
            public EventDescriptorCollection GetEvents() => EventDescriptorCollection.Empty;
            public EventDescriptorCollection GetEvents(Attribute[] attributes) => GetEvents();
            public PropertyDescriptorCollection GetProperties() => Metadata;
            public PropertyDescriptorCollection GetProperties(Attribute[] attributes) => Metadata;
            public object GetPropertyOwner(PropertyDescriptor descriptor) => this;
        }
    }
}

namespace CodexVBE.Tests.Unit
{
    using System;
    using System.ComponentModel;
    using System.Reflection;
    using CodexVBE;

    public sealed partial class VbeProjectComponentsTests
    {
        public enum Choice { First = 1, Second = 2 }
        public sealed class ProjectReference
        {
            public string GUID { get; set; }
            public int Major { get; set; }
            public int Minor { get; set; }
            public bool IsBroken { get; set; }
            public bool BuiltIn { get; set; }
        }

        public sealed class MetadataComponent : FakeComponent, ICustomTypeDescriptor
        {
            public MetadataComponent() : base("Module1", 1) { }
            public PropertyDescriptorCollection Metadata { get; set; } = new PropertyDescriptorCollection(new PropertyDescriptor[0]);
            public AttributeCollection GetAttributes() => AttributeCollection.Empty;
            public string GetClassName() => GetType().FullName;
            public string GetComponentName() => Name;
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

        public sealed class MetadataProperty : PropertyDescriptor
        {
            private readonly Type type;
            public MetadataProperty(string name, Type type, object value = null) : base(name, null) { this.type = type; Value = value; }
            public object Value { get; set; }
            public bool ReadOnly { get; set; }
            public bool FailRead { get; set; }
            public bool IgnoreWrite { get; set; }
            public bool NormalizeString { get; set; }
            public override Type ComponentType => typeof(MetadataComponent);
            public override Type PropertyType => type;
            public override bool IsReadOnly => ReadOnly;
            public override bool CanResetValue(object component) => false;
            public override bool ShouldSerializeValue(object component) => false;
            public override void ResetValue(object component) { }
            public override object GetValue(object component) { if (FailRead) throw new InvalidOperationException("Unreadable metadata"); return Value; }
            public override void SetValue(object component, object value) { if (!IgnoreWrite) Value = NormalizeString ? Convert.ToString(value).ToUpperInvariant() : value; }
        }

        private static object PrivateProjectMethod(string method, params object[] arguments)
        {
            try { return typeof(VbeProjectComponents).GetMethod(method, BindingFlags.Static | BindingFlags.NonPublic).Invoke(null, arguments); }
            catch (TargetInvocationException error) { System.Runtime.ExceptionServices.ExceptionDispatchInfo.Capture(error.InnerException).Throw(); throw; }
        }
    }
}

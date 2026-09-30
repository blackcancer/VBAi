using System;
using System.ComponentModel;
using System.Reflection;
using System.Runtime.ExceptionServices;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using VBAi;

namespace VBAi.Tests.Unit
{
    [TestClass, TestCategory("Unit"), DoNotParallelize]
    public sealed class VbeScalarPropertyTests
    {
        [TestMethod]
        public void NativeComponentScalarKeepsConversionValidationAndReadbackWithoutDescriptorWrites()
        {
            var previous = VbeScalarProperty.NativeObject;
            try
            {
                VbeScalarProperty.NativeObject = _ => true;
                var target = new ScalarTarget();
                SetComponentScalar(target, "Number", "12.5");
                SetComponentScalar(target, "Custom", "été");
                SetComponentScalar(target, "Choice", "Friday");
                Assert.AreEqual(12.5d, target.Number);
                Assert.AreEqual("été", target.Custom);
                Assert.AreEqual(DayOfWeek.Friday, target.Choice);
                Assert.AreEqual(0, target.DescriptorWrites);
                Assert.ThrowsException<ArgumentException>(() => SetComponentScalar(target, "Custom", 42));
                Assert.ThrowsException<InvalidOperationException>(() => SetComponentScalar(target, "Unknown", 1));
                target.ReadOnly = true;
                Assert.ThrowsException<InvalidOperationException>(() => SetComponentScalar(target, "Number", 2));
                Assert.AreEqual(12.5d, target.Number);
                target.ReadOnly = false; target.Ignore = true;
                Assert.ThrowsException<InvalidOperationException>(() => SetComponentScalar(target, "Number", 2));
                Assert.AreEqual(12.5d, target.Number, "Failed retention must still fail readback.");
                target.Ignore = false; target.Fail = true;
                var error = Assert.ThrowsException<InvalidOperationException>(() => SetComponentScalar(target, "Number", 3));
                Assert.AreEqual("native setter failure", error.Message);
                Assert.AreEqual(0, target.DescriptorWrites);
            }
            finally { VbeScalarProperty.NativeObject = previous; }
        }

        [TestMethod]
        public void ManagedDescriptorsRemainSupportedAndMissingNativeSetterPropagates()
        {
            var previous = VbeScalarProperty.NativeObject;
            try
            {
                VbeScalarProperty.NativeObject = _ => false;
                int writes = 0;
                var descriptor = new VbeFormsCoverageTests.LiveProperty("Virtual", typeof(double), () => 0d, _ => writes++);
                VbeScalarProperty.Set(new object(), descriptor, 2d);
                Assert.AreEqual(1, writes, "Managed virtual properties must retain their descriptor implementation.");
                Assert.ThrowsException<MissingMethodException>(() => VbeScalarProperty.SetNative(new ScalarTarget(), "Absent", 2d));
            }
            finally { VbeScalarProperty.NativeObject = previous; }
        }

        private static void SetComponentScalar(object target, string name, object value)
        {
            try { typeof(VbeProjectComponents).GetMethod("SetScalar", BindingFlags.Static | BindingFlags.NonPublic).Invoke(null, new[] { target, name, value }); }
            catch (TargetInvocationException error) { ExceptionDispatchInfo.Capture(error.InnerException).Throw(); throw; }
        }

        public sealed class ScalarTarget : CustomTypeDescriptor
        {
            private double number;
            public bool ReadOnly, Ignore, Fail;
            public int DescriptorWrites;
            public double Number { get => number; set { if (Fail) throw new InvalidOperationException("native setter failure"); if (!Ignore) number = value; } }
            public string Custom { get; set; }
            public DayOfWeek Choice { get; set; }
            public override PropertyDescriptorCollection GetProperties()
            {
                var properties = new PropertyDescriptor[3];
                int index = 0;
                foreach (string name in new[] { "Number", "Custom", "Choice" })
                {
                    var property = GetType().GetProperty(name);
                    properties[index++] = new VbeFormsCoverageTests.LiveProperty(name, property.PropertyType,
                        () => property.GetValue(this), ReadOnly ? (Action<object>)null : _ => {
                            DescriptorWrites++; throw new InvalidOperationException("Unsafe native descriptor setter invoked.");
                        });
                }
                return new PropertyDescriptorCollection(properties);
            }
            public override PropertyDescriptorCollection GetProperties(Attribute[] attributes) => GetProperties();
        }
    }
}

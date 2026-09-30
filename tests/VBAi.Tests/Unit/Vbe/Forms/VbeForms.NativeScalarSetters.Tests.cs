using System;
using System.ComponentModel;
using System.Reflection;
using VBAi;
using VBAi.Tests.Infrastructure;
using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace VBAi.Tests.Unit
{
    [TestClass, TestCategory("Unit"), DoNotParallelize]
    public sealed class NativeDesignerScalarTests
    {
        [TestMethod]
        public void NativeFallbackUsesRuntimeDispatchInsteadOfComPropertyDescriptor()
        {
            var previous = VbeForms.NativeDesignerObject;
            var setter = typeof(VbeForms).GetMethod("SetDesignerScalar", BindingFlags.Static | BindingFlags.NonPublic);
            try
            {
                VbeForms.NativeDesignerObject = _ => true;
                var target = new FallbackScalarTarget();
                foreach (var pair in new[] {
                    new { Name = "Number", Value = (object)12.5d },
                    new { Name = "Custom", Value = (object)"été" },
                    new { Name = "Choice", Value = (object)DayOfWeek.Friday }
                })
                {
                    var descriptor = new VbeFormsCoverageTests.LiveProperty(pair.Name, pair.Value.GetType(),
                        () => target.GetType().GetProperty(pair.Name).GetValue(target),
                        _ => { throw new InvalidOperationException("Unsafe native descriptor setter invoked."); });
                    setter.Invoke(null, new[] { (object)target, descriptor, pair.Value });
                    Assert.AreEqual(pair.Value, descriptor.GetValue(target), pair.Name);
                }
            }
            finally { VbeForms.NativeDesignerObject = previous; }
        }

        public sealed class FallbackScalarTarget
        {
            public double Number { get; set; }
            public string Custom { get; set; }
            public DayOfWeek Choice { get; set; }
        }

        [TestMethod]
        public void TypedDesignerScalarDispatchCoversEveryNativePropertyAndDescriptorFallback()
        {
            var previous = VbeForms.NativeDesignerObject;
            var setter = typeof(VbeForms).GetMethod("SetDesignerScalar", BindingFlags.Static | BindingFlags.NonPublic);
            try
            {
                Assert.IsFalse(previous(new NativeDesignerScalarTarget()));
                foreach (bool native in new[] { false, true })
                {
                    VbeForms.NativeDesignerObject = candidate => native;
                    var target = new NativeDesignerScalarTarget();
                    foreach (PropertyDescriptor descriptor in TypeDescriptor.GetProperties(target))
                    {
                        object value = descriptor.PropertyType == typeof(float) ? (object)12.5f : descriptor.PropertyType == typeof(bool) ? true : (object)"été";
                        setter.Invoke(null, new[] { (object)target, descriptor, value });
                        Assert.AreEqual(value, descriptor.GetValue(target), descriptor.Name);
                    }
                }
                VbeForms.NativeDesignerObject = candidate => true;
                var number = TypeDescriptor.GetProperties(new NativeDesignerScalarTarget())["Left"];
                var error = Assert.ThrowsException<TargetInvocationException>(() => setter.Invoke(null, new object[] { new NativeDesignerScalarTarget(), number, "invalid number" }));
                Assert.IsInstanceOfType<FormatException>(error.InnerException);
            }
            finally { VbeForms.NativeDesignerObject = previous; }
        }
    }
}

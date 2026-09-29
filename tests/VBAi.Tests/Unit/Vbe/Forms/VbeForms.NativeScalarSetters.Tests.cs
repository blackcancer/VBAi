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

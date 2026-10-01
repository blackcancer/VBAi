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

        [TestMethod]
        public void FailureFormattingPreservesProtocolShapeAndWrappedOriginalDetails()
        {
            var original = new System.Runtime.InteropServices.COMException("synthetic original message", unchecked((int)0x9CFD3148));
            VbeScalarProperty.AnnotateFailure(original, VbeScalarProperty.FailurePhase.SetterInvocation);
            var wrapped = new TargetInvocationException(new TargetInvocationException(original));
            string message = VbeScalarProperty.FormatFailure(wrapped);
            StringAssert.StartsWith(message, original.Message);
            StringAssert.Contains(message, "SetterInvocation");
            StringAssert.Contains(message, "System.Runtime.InteropServices.COMException");
            StringAssert.Contains(message, "0x9CFD3148");
            StringAssert.Contains(message, "do not retry automatically");
            Assert.IsTrue(message.Length - original.Message.Length < 400, "Added diagnostic metadata must remain bounded.");
            var response = Response.Failure(message);
            Assert.IsFalse(response.Ok);
            Assert.IsNull(response.Data);
            Assert.AreEqual(message, response.Error);
            Assert.AreEqual("synthetic original message", original.Message);
            Assert.AreEqual(unchecked((int)0x9CFD3148), original.HResult);
        }

        [TestMethod]
        public void UnannotatedAndInvalidPhaseErrorsKeepExistingBoundaryMessage()
        {
            var original = new InvalidOperationException("original unchanged");
            var wrapped = new TargetInvocationException(original);
            Assert.AreEqual(wrapped.Message, VbeScalarProperty.FormatFailure(wrapped));
            VbeScalarProperty.AnnotateFailure(original, (VbeScalarProperty.FailurePhase)42);
            Assert.AreEqual(original.Message, VbeScalarProperty.FormatFailure(original));
            VbeScalarProperty.AnnotateFailure(original, VbeScalarProperty.FailurePhase.RetentionReadback);
            StringAssert.Contains(VbeScalarProperty.FormatFailure(original), "RetentionReadback");
        }

        [TestMethod]
        public void DescriptorThrownReflectionErrorKeepsItsOwnAnnotationAndIdentity()
        {
            var error = new TargetInvocationException(new InvalidOperationException("synthetic inner failure"));
            VbeScalarProperty.AnnotateFailure(error, VbeScalarProperty.FailurePhase.RetentionReadback);
            string message = VbeScalarProperty.FormatFailure(new TargetInvocationException(error));
            StringAssert.StartsWith(message, error.Message);
            StringAssert.Contains(message, "RetentionReadback");
            StringAssert.Contains(message, typeof(TargetInvocationException).FullName);
            StringAssert.Contains(message, "0x" + unchecked((uint)error.HResult).ToString("X8"));
        }

        [TestMethod]
        public void UnavailableDiagnosticDataCannotMaskTheOriginalFailure()
        {
            var error = new UnavailableDataException();
            VbeScalarProperty.AnnotateFailure(error, VbeScalarProperty.FailurePhase.SetterInvocation);
            Assert.AreEqual("original diagnostic-resistant failure", VbeScalarProperty.FormatFailure(error));
            Assert.AreEqual(unchecked((int)0x80004005), error.HResult);
        }

        private sealed class UnavailableDataException : Exception
        {
            internal UnavailableDataException() : base("original diagnostic-resistant failure") { HResult = unchecked((int)0x80004005); }
            public override System.Collections.IDictionary Data => throw new InvalidOperationException("annotation unavailable");
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

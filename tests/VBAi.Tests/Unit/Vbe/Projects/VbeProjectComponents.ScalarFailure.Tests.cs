using Microsoft.VisualStudio.TestTools.UnitTesting;
using System;
using System.ComponentModel;
using System.Reflection;
using System.Runtime.ExceptionServices;
using System.Runtime.InteropServices;

namespace VBAi.Tests.Unit
{
    [TestClass, TestCategory("Unit"), DoNotParallelize]
    public sealed class VbeProjectScalarFailureTests
    {
        [DataTestMethod, DataRow(false), DataRow(true)]
        public void SetterThatChangesValueThenThrowsKeepsOriginalErrorAndNeverReadsBack(bool native)
        {
            var error = new COMException("synthetic native setter failure", unchecked((int)0x9CFD3148));
            var target = new ScalarFailureTarget { SetterError = error };
            var observed = Assert.ThrowsException<COMException>(() => SetScalar(target, 321, native));
            Assert.AreSame(error, observed);
            Assert.AreEqual(unchecked((int)0x9CFD3148), observed.HResult);
            Assert.AreEqual("synthetic native setter failure", observed.Message);
            Assert.AreEqual(321, target.Retained);
            Assert.AreEqual(1, target.SetterCalls);
            Assert.AreEqual(0, target.GetterCalls, "A throwing setter must not trigger native readback or another write.");
            Assert.AreEqual(native ? 0 : 1, target.DescriptorWrites);
            StringAssert.Contains(VbeScalarProperty.FormatFailure(observed), "SetterInvocation");
        }

        [DataTestMethod, DataRow(false), DataRow(true)]
        public void SuccessfulSetterAndThrowingReadbackKeepOriginalErrorAndDistinctPhase(bool native)
        {
            var error = new COMException("synthetic retention getter failure", unchecked((int)0x80004005));
            var target = new ScalarFailureTarget { GetterError = error };
            var observed = Assert.ThrowsException<COMException>(() => SetScalar(target, 321, native));
            Assert.AreSame(error, observed);
            Assert.AreEqual(error.HResult, observed.HResult);
            Assert.AreEqual(error.Message, observed.Message);
            Assert.AreEqual(321, target.Retained);
            Assert.AreEqual(1, target.SetterCalls);
            Assert.AreEqual(1, target.GetterCalls);
            Assert.AreEqual(native ? 0 : 1, target.DescriptorWrites);
            StringAssert.Contains(VbeScalarProperty.FormatFailure(observed), "RetentionReadback");
        }

        [TestMethod]
        public void RefusedRetentionIsAnnotatedWithoutAnotherSetter()
        {
            var target = new ScalarFailureTarget { IgnoreWrite = true };
            var error = Assert.ThrowsException<InvalidOperationException>(() => SetScalar(target, 321, false));
            StringAssert.Contains(error.Message, "did not retain property HelpContextID");
            StringAssert.Contains(VbeScalarProperty.FormatFailure(error), "RetentionReadback");
            Assert.AreEqual(1, target.SetterCalls);
            Assert.AreEqual(1, target.GetterCalls);
            Assert.AreEqual(0, target.Retained);
        }

        [TestMethod]
        public void ValidationAndConversionFailuresRemainUnannotatedBeforeAnyWrite()
        {
            var target = new ScalarFailureTarget { ReadOnly = true };
            var readOnly = Assert.ThrowsException<InvalidOperationException>(() => SetScalar(target, 321, false));
            Assert.AreEqual(readOnly.Message, VbeScalarProperty.FormatFailure(readOnly));
            target.ReadOnly = false;
            var conversion = Assert.ThrowsException<FormatException>(() => SetScalar(target, "invalid-number", false));
            Assert.AreEqual(conversion.Message, VbeScalarProperty.FormatFailure(conversion));
            Assert.AreEqual(0, target.SetterCalls);
            Assert.AreEqual(0, target.GetterCalls);
        }

        private static void SetScalar(ScalarFailureTarget target, object value, bool native)
        {
            var previous = VbeScalarProperty.NativeObject;
            try
            {
                VbeScalarProperty.NativeObject = _ => native;
                try
                {
                    typeof(VbeProjectComponents).GetMethod("SetScalar", BindingFlags.Static | BindingFlags.NonPublic)
                        .Invoke(null, new[] { (object)target, "HelpContextID", value });
                }
                catch (TargetInvocationException error)
                {
                    ExceptionDispatchInfo.Capture(error.InnerException).Throw();
                    throw;
                }
            }
            finally { VbeScalarProperty.NativeObject = previous; }
        }

        public sealed class ScalarFailureTarget : CustomTypeDescriptor
        {
            public Exception SetterError, GetterError;
            public int SetterCalls, GetterCalls, DescriptorWrites, Retained;
            public bool ReadOnly, IgnoreWrite;
            public int HelpContextID
            {
                get { GetterCalls++; if (GetterError != null) throw GetterError; return Retained; }
                set { SetterCalls++; if (!IgnoreWrite) Retained = value; if (SetterError != null) throw SetterError; }
            }
            public override PropertyDescriptorCollection GetProperties() => new PropertyDescriptorCollection(new[] {
                new VbeFormsCoverageTests.LiveProperty("HelpContextID", typeof(int), () => HelpContextID,
                    ReadOnly ? (Action<object>)null : value => { DescriptorWrites++; HelpContextID = (int)value; }) });
            public override PropertyDescriptorCollection GetProperties(Attribute[] attributes) => GetProperties();
        }
    }
}

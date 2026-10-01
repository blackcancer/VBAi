using System;
using System.Runtime.InteropServices;
using System.Threading;
using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace VBAi.Tests.Unit
{
    [TestClass, TestCategory("Unit")]
    public sealed class VbaTestRuntimeTests
    {
        private static void Arm(VbaTestResultSink sink)
        {
            sink.Arm(new object(), "project", "2", "revision", "run",
                new VbaTestDescriptor { Id = "test", Module = "Tests", Procedure = "Check" }, "Test", "signature");
            sink.BeginNative();
        }

        [TestMethod]
        public void FactoryHasNoWorkOutsideTheOwningExposedAttempt()
        {
            Assert.ThrowsException<InvalidOperationException>(() => new VbaTestRuntime());
            using (var sink = new VbaTestResultSink(ReferenceEquals))
            {
                Arm(sink);
                using (VbaTestRuntime.Expose(sink))
                {
                    Exception caught = null;
                    var other = new Thread(() => { try { new VbaTestRuntime(); } catch (Exception error) { caught = error; } });
                    other.Start(); other.Join();
                    Assert.IsInstanceOfType(caught, typeof(InvalidOperationException));
                    var runtime = new VbaTestRuntime();
                    var job = (object[])runtime.Request("2", "signature");
                    Assert.IsTrue(runtime.Publish((string)job[0], "revision", "run", "test", "Passed", "", 0));
                    sink.CompleteNative(true);
                }
                Assert.ThrowsException<InvalidOperationException>(() => new VbaTestRuntime());
            }
        }

        [DataTestMethod]
        [DataRow("signature")]
        [DataRow("version")]
        [DataRow("run")]
        [DataRow("test")]
        public void FactoryClaimsAndVerdictsCannotCrossThePendingContract(string failure)
        {
            using (var sink = new VbaTestResultSink(ReferenceEquals))
            {
                Arm(sink);
                using (VbaTestRuntime.Expose(sink))
                {
                    var runtime = new VbaTestRuntime();
                    if (failure == "signature" || failure == "version")
                        Assert.ThrowsException<InvalidOperationException>(() => runtime.Request(failure == "version" ? "1" : "2", failure == "signature" ? "foreign" : "signature"));
                    else
                    {
                        var job = (object[])runtime.Request("2", "signature");
                        Assert.ThrowsException<InvalidOperationException>(() => runtime.Publish((string)job[0], "revision", failure == "run" ? "foreign" : "run", failure == "test" ? "foreign" : "test", "Passed", "", 0));
                    }
                    Assert.IsTrue(Assert.ThrowsException<VbaTestInvocationException>(() => sink.CompleteNative(true)).Uncertain);
                }
            }
        }

        [TestMethod]
        public void RetainedRuntimeCannotClaimTheNextAttempt()
        {
            using (var sink = new VbaTestResultSink(ReferenceEquals))
            {
                Arm(sink); VbaTestRuntime old;
                using (VbaTestRuntime.Expose(sink))
                {
                    old = new VbaTestRuntime(); var job = (object[])old.Request("2", "signature");
                    old.Publish((string)job[0], "revision", "run", "test", "Passed", "", 0); sink.CompleteNative(true);
                }
                Arm(sink);
                using (VbaTestRuntime.Expose(sink))
                {
                    Assert.ThrowsException<InvalidOperationException>(() => old.Request("2", "signature"));
                    Assert.IsTrue(Assert.ThrowsException<VbaTestInvocationException>(() => sink.CompleteNative(true)).Uncertain);
                }
            }
        }

        [TestMethod]
        public void RuntimeExposesOnlyTheVersionedCallbackContract()
        {
            Assert.AreEqual("VBAi.TestRuntime", ((ProgIdAttribute)Attribute.GetCustomAttribute(typeof(VbaTestRuntime), typeof(ProgIdAttribute))).Value);
            Assert.AreEqual(ClassInterfaceType.None, ((ClassInterfaceAttribute)Attribute.GetCustomAttribute(typeof(VbaTestRuntime), typeof(ClassInterfaceAttribute))).Value);
            Assert.AreEqual(2, typeof(IVbaTestRuntime).GetMethods().Length);
            Assert.AreEqual(1, ((DispIdAttribute)Attribute.GetCustomAttribute(typeof(IVbaTestRuntime).GetMethod("Request"), typeof(DispIdAttribute))).Value);
            Assert.AreEqual(2, ((DispIdAttribute)Attribute.GetCustomAttribute(typeof(IVbaTestRuntime).GetMethod("Publish"), typeof(DispIdAttribute))).Value);
        }
    }
}

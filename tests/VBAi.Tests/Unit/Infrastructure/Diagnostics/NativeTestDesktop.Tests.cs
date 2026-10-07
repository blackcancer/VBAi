using Microsoft.VisualStudio.TestTools.UnitTesting;
using System;
using VBAi.Tests.Integration;

namespace VBAi.Tests.Unit
{
    /// <summary>Checks explicit desktop selection without launching Office or switching the input desktop.</summary>
    [TestClass, TestCategory("Unit")]
    public sealed class NativeTestDesktopTests
    {
        [TestMethod]
        public void ExplicitMainValidatesRealPlacementOnly()
        {
            int calls = 0;
            Assert.AreEqual("Default", NativeTestDesktop.Select("1", null, null, () => calls++, _ => Assert.Fail()));
            Assert.AreEqual(1, calls);
        }

        [DataTestMethod]
        [DataRow("private", null)]
        [DataRow(null, "private")]
        [DataRow("private", "private")]
        public void MainRefusesAnyInheritedPrivateDescriptor(string configured, string required)
        {
            Assert.ThrowsException<InvalidOperationException>(() => NativeTestDesktop.Select("1", configured, required,
                () => Assert.Fail(), _ => Assert.Fail()));
        }

        [TestMethod]
        public void PrivateSelectionRetainsExactPairAndPlacementGuard()
        {
            string observed = null;
            Assert.AreEqual("private", NativeTestDesktop.Select(null, "private", "private",
                () => Assert.Fail(), desktop => observed = desktop));
            Assert.AreEqual("private", observed);
        }

        [DataTestMethod]
        [DataRow(null, null)]
        [DataRow("private", null)]
        [DataRow(null, "private")]
        [DataRow("private", "other")]
        public void UnselectedOrMismatchedDesktopRefusesBeforeNativeWork(string configured, string required)
        {
            Assert.ThrowsException<InvalidOperationException>(() => NativeTestDesktop.Select(null, configured, required,
                () => Assert.Fail(), _ => Assert.Fail()));
        }

        [TestMethod]
        public void PlacementFailureIsPreservedWithoutFallback()
        {
            var error = new InvalidOperationException("Original placement failure");
            Assert.AreSame(error, Assert.ThrowsException<InvalidOperationException>(() => NativeTestDesktop.Select("1", null, null,
                () => throw error, _ => Assert.Fail())));
        }

        [TestMethod]
        public void MissingValidationDependenciesAreRejected()
        {
            Assert.ThrowsException<ArgumentNullException>(() => NativeTestDesktop.Select("1", null, null, null, _ => { }));
            Assert.ThrowsException<ArgumentNullException>(() => NativeTestDesktop.Select("1", null, null, () => { }, null));
        }
    }
}

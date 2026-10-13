using Microsoft.VisualStudio.TestTools.UnitTesting;
using System;

namespace VBAi.Tests.Unit
{
    /// <summary>Checks the narrowly scoped reader resolution used by forms in COM hosts.</summary>
    [TestClass, TestCategory("Unit")]
    public sealed class UiResourceReaderTests
    {
        [TestMethod]
        public void LegacyReaderResolvesWithoutApplicationRedirects()
        {
            var reader = UiResourceReader.Resolve(null, new ResolveEventArgs(
                "System.Resources.Extensions, Version=4.0.0.0, Culture=neutral, PublicKeyToken=cc7b13ffcd2ddd51"));
            Assert.AreEqual("System.Resources.Extensions", reader.GetName().Name);
            Assert.IsNotNull(reader.GetType("System.Resources.Extensions.RuntimeResourceSet"));
        }

        [TestMethod]
        [DataRow("Unrelated.Assembly")]
        [DataRow("System.Resources.Extensions, Version=8.0.0.0, Culture=neutral, PublicKeyToken=cc7b13ffcd2ddd51")]
        [DataRow("System.Resources.Extensions, Version=4.0.0.0, Culture=fr-FR, PublicKeyToken=cc7b13ffcd2ddd51")]
        [DataRow("System.Resources.Extensions, Version=4.0.0.0, Culture=neutral, PublicKeyToken=0000000000000000")]
        public void OtherIdentitiesRetainTheHostResolutionPath(string identity)
        {
            Assert.IsNull(UiResourceReader.Resolve(null, new ResolveEventArgs(identity)));
        }

        [TestMethod]
        public void InitializationFailureIsPreservedWithoutRetry()
        {
            int calls = 0;
            var expected = new InvalidOperationException("Resource initialization failed.");
            var actual = Assert.ThrowsException<InvalidOperationException>(() =>
                UiResourceReader.Apply(() => { calls++; throw expected; }));
            Assert.AreSame(expected, actual);
            Assert.AreEqual(1, calls);
            Assert.ThrowsException<ArgumentNullException>(() => UiResourceReader.Apply(null));
        }
    }
}

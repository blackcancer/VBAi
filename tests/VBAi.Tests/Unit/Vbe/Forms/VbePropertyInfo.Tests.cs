namespace VBAi.Tests.Unit
{
    using Microsoft.VisualStudio.TestTools.UnitTesting;
    using VBAi;

    [TestClass, TestCategory("Unit")]
    public sealed class VbePropertyInfoTests
    {
        [TestMethod]
        public void SetterStatusUsesExplicitOverrideOrEachNullableDescriptorState()
        {
            foreach (bool? readOnly in new bool?[] { true, false, null })
            {
                var property = new VbePropertyInfo { ReadOnly = readOnly };
                string expected = readOnly == true ? "DescriptorReadOnly" :
                    readOnly == false ? "DescriptorCandidateUnverified" : "Unknown";
                Assert.AreEqual(expected, property.SetterStatus);
                property.SetterStatus = "NativeSetterRejected";
                Assert.AreEqual("NativeSetterRejected", property.SetterStatus);
                property.SetterStatus = null;
                Assert.AreEqual(expected, property.SetterStatus);
            }
        }
    }
}

namespace VBAi.Tests.Unit
{
    using System;
    using System.Collections.Generic;
    using System.Linq;
    using System.Runtime.InteropServices.ComTypes;
    using VBAi;
    using Microsoft.VisualStudio.TestTools.UnitTesting;

    [TestClass]
    [TestCategory("Unit")]
    public sealed class VbeComPropertyAccessorsTests
    {
        [TestMethod]
        public void InputProvidersAndAccessorInvocationKindsMatrix()
        {
            Assert.ThrowsException<ArgumentException>(() => VbeComPropertyAccessors.Inspect(null, "Caption"));
            Assert.ThrowsException<ArgumentException>(() => VbeComPropertyAccessors.Inspect(new object(), null));
            dynamic absent = VbeComPropertyAccessors.Inspect(new object(), "Caption"); Assert.IsNull((object)absent.SetterDeclared);
            foreach (bool fail in new[] { false, true })
            {
                dynamic result = VbeComPropertyAccessors.Inspect(new ComClassInfoFixture { Fail = fail }, "Caption");
                Assert.IsFalse((bool)result.MetadataComplete); Assert.AreEqual(1, result.Errors.Count);
            }
            foreach (var kind in new[] { INVOKEKIND.INVOKE_FUNC, INVOKEKIND.INVOKE_PROPERTYGET, INVOKEKIND.INVOKE_PROPERTYPUT, INVOKEKIND.INVOKE_PROPERTYPUTREF })
            using (var info = new ComTypeInfoFixture())
            {
                info.Functions.Add(new ComTypeInfoFixture.Function { Kind = kind });
                info.Functions.Add(new ComTypeInfoFixture.Function { Name = "Other" });
                info.Functions.Add(new ComTypeInfoFixture.Function { NameCount = 0 });
                dynamic result = VbeComPropertyAccessors.Inspect(new ComClassInfoFixture { Info = info }, "cApTiOn");
                Assert.IsTrue((bool)result.MetadataComplete); Assert.AreEqual(1, result.Accessors.Count);
                Assert.AreEqual(kind == INVOKEKIND.INVOKE_PROPERTYPUT || kind == INVOKEKIND.INVOKE_PROPERTYPUTREF, (bool)result.SetterDeclared);
                Assert.IsFalse((bool)result.ActualSetterVerified); Assert.AreEqual(3, info.FunctionReleases);
            }
        }

        [TestMethod]
        public void AttributeFunctionInheritanceFailuresAndLimitsPreserveDescriptorOwnership()
        {
            foreach (int mode in Enumerable.Range(0, 12))
            using (var info = new ComTypeInfoFixture())
            using (var inherited = new ComTypeInfoFixture())
            {
                var function = new ComTypeInfoFixture.Function(); info.Functions.Add(function);
                var parent = new ComTypeInfoFixture.Inherited { Info = inherited }; info.Parents.Add(parent);
                switch (mode)
                {
                    case 0: info.AttributeFailure = true; break;
                    case 1: info.AttributeFailureAfterAllocation = true; break;
                    case 2: info.FunctionCount = 2049; break;
                    case 3: info.ParentCount = 65; break;
                    case 4: info.DocumentationFailure = true; break;
                    case 5: function.DescriptorFailure = true; break;
                    case 6: function.NamesFailure = true; break;
                    case 7: parent.FlagsFailure = true; break;
                    case 8: parent.ReferenceFailure = true; break;
                    case 9: parent.InfoFailure = true; break;
                    case 10: parent.Info = null; break;
                    case 11: parent.Flags = IMPLTYPEFLAGS.IMPLTYPEFLAG_FSOURCE; break;
                }
                dynamic result = VbeComPropertyAccessors.Inspect(new ComClassInfoFixture { Info = info }, "Caption");
                Assert.AreEqual(mode == 11, (bool)result.MetadataComplete);
                Assert.AreEqual(mode == 11 ? 0 : 1, result.Errors.Count);
                Assert.AreEqual(mode == 0 ? 0 : 1, info.AttributeReleases);
            }
        }

        [TestMethod]
        public void CyclicAndDeepTypeGraphsTerminateAtTheDeclaredLimits()
        {
            using (var cycle = new ComTypeInfoFixture())
            {
                cycle.Parents.Add(new ComTypeInfoFixture.Inherited { Info = cycle });
                dynamic result = VbeComPropertyAccessors.Inspect(new ComClassInfoFixture { Info = cycle }, "Caption");
                Assert.IsTrue((bool)result.MetadataComplete); Assert.IsFalse((bool)result.SetterDeclared);
                Assert.AreEqual(1, result.Interfaces.Count); Assert.AreEqual(2, cycle.AttributeReleases);
            }
            var types = Enumerable.Range(0, 10).Select(_ => new ComTypeInfoFixture()).ToArray();
            try
            {
                for (int i = 0; i < types.Length - 1; i++) types[i].Parents.Add(new ComTypeInfoFixture.Inherited { Info = types[i + 1] });
                dynamic result = VbeComPropertyAccessors.Inspect(new ComClassInfoFixture { Info = types[0] }, "Caption");
                Assert.IsFalse((bool)result.MetadataComplete); Assert.AreEqual(9, result.Interfaces.Count);
                Assert.AreEqual(1, result.Errors.Count); StringAssert.Contains((string)result.Errors[0], "depth 8");
            }
            finally { foreach (var type in types) type.Dispose(); }
        }
    }
}

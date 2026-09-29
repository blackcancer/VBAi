namespace VBAi.Tests.Unit
{
    using System;
    using System.Linq;
    using System.Runtime.InteropServices.ComTypes;
    using VBAi;
    using Microsoft.VisualStudio.TestTools.UnitTesting;

    [TestClass]
    [TestCategory("Unit")]
    public sealed class VbeComEventsTests
    {
        [TestMethod]
        public void ProvidersCoclassShapeAndAttributeFailuresAreReported()
        {
            dynamic absent = VbeComEvents.Read(new object()); Assert.IsFalse((bool)absent.SourceInterfacesComplete);
            foreach (bool fail in new[] { false, true })
            {
                dynamic result = VbeComEvents.Read(new ComClassInfoFixture { Fail = fail });
                Assert.IsFalse((bool)result.SourceInterfacesComplete); Assert.AreEqual(1, result.Errors.Count);
            }
            foreach (int mode in Enumerable.Range(0, 6))
            using (var info = new ComTypeInfoFixture { Kind = TYPEKIND.TKIND_COCLASS })
            {
                if (mode == 0) info.Kind = TYPEKIND.TKIND_INTERFACE;
                if (mode == 1) info.ParentCount = 65;
                if (mode == 2) info.AttributeFailure = true;
                if (mode == 3) info.AttributeFailureAfterAllocation = true;
                if (mode == 5) info.Parents.Add(new ComTypeInfoFixture.Inherited { Info = info });
                dynamic result = VbeComEvents.Read(new ComClassInfoFixture { Info = info });
                Assert.IsFalse((bool)result.SourceInterfacesComplete); Assert.IsFalse((bool)result.VbeEventCatalogComplete);
                Assert.AreEqual(mode < 4 ? 1 : 0, result.Errors.Count);
                Assert.AreEqual(mode == 2 ? 0 : 1, info.AttributeReleases);
            }
        }

        [TestMethod]
        public void SourceEventsNamesParametersLimitsAndDescriptorFailuresMatrix()
        {
            foreach (int mode in Enumerable.Range(0, 12))
            using (var source = new ComTypeInfoFixture())
            using (var coclass = new ComTypeInfoFixture { Kind = TYPEKIND.TKIND_COCLASS })
            {
                var function = new ComTypeInfoFixture.Function { Name = "Click", Parameters = 2 };
                source.Functions.Add(function);
                var inherited = new ComTypeInfoFixture.Inherited { Info = source, Flags = IMPLTYPEFLAGS.IMPLTYPEFLAG_FSOURCE };
                coclass.Parents.Add(inherited);
                switch (mode)
                {
                    case 1: inherited.FlagsFailure = true; break;
                    case 2: inherited.ReferenceFailure = true; break;
                    case 3: inherited.InfoFailure = true; break;
                    case 4: source.AttributeFailure = true; break;
                    case 5: source.AttributeFailureAfterAllocation = true; break;
                    case 6: source.DocumentationFailure = true; break;
                    case 7: function.DescriptorFailure = true; break;
                    case 8: function.NamesFailure = true; break;
                    case 9: function.NameCount = 0; break;
                    case 10: function.Name = " "; break;
                    case 11: source.FunctionCount = 257; break;
                }
                dynamic result = VbeComEvents.Read(new ComClassInfoFixture { Info = coclass });
                Assert.AreEqual(mode == 0, (bool)result.SourceInterfacesComplete);
                Assert.AreEqual(mode == 0 ? 0 : 1, result.Errors.Count);
                Assert.AreEqual(mode == 0 ? 1 : 0, result.Events.Count);
                Assert.AreEqual(1, coclass.AttributeReleases);
                Assert.AreEqual(mode >= 7 && mode <= 10 && mode != 7 || mode == 0 ? 1 : 0, source.FunctionReleases);
            }
            using (var coclass = new ComTypeInfoFixture { Kind = TYPEKIND.TKIND_COCLASS })
            using (var first = new ComTypeInfoFixture())
            using (var second = new ComTypeInfoFixture())
            using (var third = new ComTypeInfoFixture())
            {
                for (int i = 0; i < 256; i++)
                {
                    first.Functions.Add(new ComTypeInfoFixture.Function { Name = "Event" + i, Parameters = 70 });
                    second.Functions.Add(new ComTypeInfoFixture.Function { Name = "Other" + i });
                }
                third.Functions.Add(new ComTypeInfoFixture.Function { Name = "Overflow" });
                foreach (var source in new[] { first, second, third }) coclass.Parents.Add(new ComTypeInfoFixture.Inherited { Info = source, Flags = IMPLTYPEFLAGS.IMPLTYPEFLAG_FSOURCE });
                dynamic result = VbeComEvents.Read(new ComClassInfoFixture { Info = coclass });
                Assert.AreEqual(512, result.Events.Count); Assert.AreEqual(1, result.Errors.Count); Assert.IsFalse((bool)result.SourceInterfacesComplete);
                Assert.AreEqual(256, first.FunctionReleases); Assert.AreEqual(256, second.FunctionReleases);
            }
        }
    }
}

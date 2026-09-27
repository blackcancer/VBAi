namespace CodexVBE.Tests.Unit
{
    using System;
    using System.Collections.Generic;
    using System.Linq;
    using CodexVBE;
    using Microsoft.VisualStudio.TestTools.UnitTesting;

    [TestClass]
    [TestCategory("Unit")]
    public sealed partial class VbeFormsContractTests
    {
        [TestMethod]
        public void NativeControlsAndMissingComMetadataAreReportedHonestly()
        {
            var forms = new VbeForms(new FakeVbe());
            var types = ((IEnumerable<object>)forms.ControlTypes()).Cast<dynamic>().ToArray();
            Assert.IsTrue(types.Any(item => (string)item.ProgId == "Forms.Label.1" && (string)item.Source == "MSForms native"));
            Assert.IsTrue(types.Any(item => (string)item.ProgId == "Forms.TextBox.1"));
            Assert.IsTrue(VbeControlCatalog.IsCandidate("Forms.Label.1", new HashSet<string>(StringComparer.OrdinalIgnoreCase) { "Forms.Label.1" }));
            dynamic events = VbeComEvents.Read(new object ());
            Assert.IsFalse((bool)events.SourceInterfacesComplete);
            Assert.IsFalse((bool)events.VbeEventCatalogComplete);
            Assert.ThrowsException<ArgumentException>(() => VbeComPropertyAccessors.Inspect(null, "Caption"));
            Assert.ThrowsException<ArgumentException>(() => VbeComPropertyAccessors.Inspect(new object (), " "));
            dynamic metadata = VbeComPropertyAccessors.Inspect(new object (), "Caption");
            Assert.IsFalse((bool)metadata.MetadataComplete);
            Assert.IsNull((object)metadata.SetterDeclared);
            Assert.IsFalse((bool)metadata.ActualSetterVerified);
        }
    }
}

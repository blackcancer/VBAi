using Microsoft.VisualStudio.TestTools.UnitTesting;
using System;
using System.Globalization;
using System.IO;

namespace VBAi.Tests.Integration
{
    /// <summary>Separately qualifies writable project metadata in fresh disposable Access/Publisher files.</summary>
    [TestClass, TestCategory("Office"), TestCategory("OfficeAdapterOnly"), TestCategory("OfficeAdapterOnlyMetadata"), DoNotParallelize]
    public sealed class OfficeAdapterOnlyMetadataQualificationTests
    {
        public TestContext TestContext { get; set; }

        /// <summary>Preserves an exact accented description through Access adapter Save and process restart.</summary>
        [STATestMethod] public void Access16DescriptionAdapterSaveReopen() { Qualify("Access", "Description"); }
        /// <summary>Preserves an exact accented description through Publisher adapter Save and process restart.</summary>
        [STATestMethod] public void PublisherDescriptionAdapterSaveReopen() { Qualify("Publisher", "Description"); }
        /// <summary>Qualifies a numeric help-context metadata value without opening Help or executing code.</summary>
        [STATestMethod] public void Access16HelpContextIdAdapterSaveReopen() { Qualify("Access", "HelpContextID"); }
        /// <summary>Qualifies Publisher help-context metadata independently from description persistence.</summary>
        [STATestMethod] public void PublisherHelpContextIdAdapterSaveReopen() { Qualify("Publisher", "HelpContextID"); }
        /// <summary>Stores only an owned HelpFile path marker; no help file is fabricated or opened.</summary>
        [STATestMethod] public void Access16HelpFilePathAdapterSaveReopen() { Qualify("Access", "HelpFile"); }
        /// <summary>Checks an owned Publisher HelpFile path marker without changing any installed help library.</summary>
        [STATestMethod] public void PublisherHelpFilePathAdapterSaveReopen() { Qualify("Publisher", "HelpFile"); }

        /// <summary>Qualifies Description persistence from an audited existing serialized Publisher publication.</summary>
        [STATestMethod] public void PublisherSerializedDescriptionAdapterSaveReopen() { Qualify("Publisher", "Description", true); }

        private void Qualify(string host, string property, bool serializedPublisherSeed = false)
        {
            object expected = null, before = null;
            OfficeAdapterOnlyProjectQualification.Run(host, "Metadata." + property, null, fixture =>
            {
                before = OfficeAdapterOnlyProjectQualification.ReadProperty(fixture, property);
                expected = property == "Description" ? (object)("VBAi adapter metadata été " + Guid.NewGuid().ToString("N")) :
                    property == "HelpFile" ? Path.Combine(fixture.Root, "OwnedMetadataHelp.chm") :
                    (object)(Convert.ToInt32(before, CultureInfo.InvariantCulture) == 321 ? 322 : 321);
                if (property == "HelpContextID" && before is string) expected = Convert.ToString(expected, CultureInfo.InvariantCulture);
                Assert.AreNotEqual(Convert.ToString(before, CultureInfo.InvariantCulture), Convert.ToString(expected, CultureInfo.InvariantCulture));
                fixture.RecordAdapterStage("MetadataExpectedReadback", new
                {
                    Property = property,
                    Before = before,
                    Expected = expected,
                    HelpFileContentCreated = false,
                    HelpInvoked = false
                });
                OfficeAdapterOnlyProjectQualification.SetProperty(fixture, property, expected);
            }, fixture =>
            {
                Assert.AreEqual(Convert.ToString(expected, CultureInfo.InvariantCulture),
                    Convert.ToString(OfficeAdapterOnlyProjectQualification.ReadProperty(fixture, property), CultureInfo.InvariantCulture),
                    "Project metadata was not retained exactly: " + property);
            }, TestContext, serializedPublisherSeed);
        }
    }
}

using Microsoft.VisualStudio.TestTools.UnitTesting;
using System;
using System.Globalization;
using System.IO;

namespace VBAi.Tests.Integration
{
    /// <summary>Separates CLR/raw metadata dispatch on fresh owned Access files with one setter per case and exact disk readback.</summary>
    [TestClass, TestCategory("Office"), TestCategory("OfficeMetadataSetterProbe"), DoNotParallelize]
    public sealed class OfficeMetadataSetterProbeTests
    {
        public TestContext TestContext { get; set; }

        [STATestMethod] public void AccessHelpFileExternalClrOneShotSaveReopen() { Qualify("HelpFile", false); }
        [STATestMethod] public void AccessHelpFileExternalRawOneShotSaveReopen() { Qualify("HelpFile", true); }
        [STATestMethod] public void AccessHelpContextExternalClrOneShotSaveReopen() { Qualify("HelpContextID", false); }
        [STATestMethod] public void AccessHelpContextExternalRawOneShotSaveReopen() { Qualify("HelpContextID", true); }

        private void Qualify(string property, bool rawDispatch)
        {
            if (Environment.GetEnvironmentVariable("VBAi_RUN_OFFICE_METADATA_SETTER_PROBE") != "1" ||
                Environment.GetEnvironmentVariable("VBAi_RUN_OFFICE_METADATA_GETTER_PROBE") != "1")
                Assert.Inconclusive("Set VBAi_RUN_OFFICE_METADATA_SETTER_PROBE=1 and VBAi_RUN_OFFICE_METADATA_GETTER_PROBE=1; every case owns a new disposable database and issues one metadata setter only.");
            object expected = null;
            OfficeAdapterOnlyProjectQualification.Run("Access", "Metadata." + property + ".External" + (rawDispatch ? "Raw" : "Clr"), null, fixture =>
            {
                object before = OfficeAdapterOnlyProjectQualification.ReadProperty(fixture, property);
                expected = property == "HelpFile" ? (object)Path.Combine(fixture.Root, "OwnedMetadataHelp.chm") :
                    (object)(Convert.ToInt32(before, CultureInfo.InvariantCulture) == 321 ? 322 : 321);
                Assert.AreNotEqual(before, expected);
                fixture.RecordAdapterStage("MetadataExpectedReadback", new
                {
                    Property = property,
                    Before = before,
                    Expected = expected,
                    SetterContext = "External fixture STA",
                    RawDispatch = rawDispatch,
                    HelpFileContentCreated = false,
                    HelpInvoked = false
                });
                fixture.InvokeOwnedMetadataSetter(property, expected, rawDispatch);
            }, fixture =>
            {
                Assert.AreEqual(expected, OfficeAdapterOnlyProjectQualification.ReadProperty(fixture, property), "Exact metadata value/type was not retained: " + property);
            }, TestContext);
        }
    }
}

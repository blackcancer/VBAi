using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using Microsoft.Win32;

namespace VBAi.Tests.Integration
{
    /// <summary>Qualifies addition/removal of only the installed optional Microsoft Scripting Runtime reference.</summary>
    [TestClass, TestCategory("Office"), TestCategory("OfficeAdapterOnly"), TestCategory("OfficeAdapterOnlyReferences"), DoNotParallelize]
    public sealed class OfficeAdapterOnlyReferenceQualificationTests
    {
        private const string ScriptingGuid = "{420B2830-E718-11CF-893D-00A0C9054228}";
        public TestContext TestContext { get; set; }

        /// <summary>Adds the installed non-built-in Scripting reference only in an owned Access database.</summary>
        [STATestMethod] public void Access16ReferenceAdditionAdapterSaveReopen() { Qualify("Access", false); }
        /// <summary>Qualifies the installed library-file addition route in a separate disposable Access database.</summary>
        [STATestMethod] public void Access16ReferenceFileAdditionAdapterSaveReopen() { Qualify("Access", false, true); }
        /// <summary>Removes only the Scripting reference explicitly added to the disposable Access baseline.</summary>
        [STATestMethod] public void Access16ReferenceRemovalAdapterSaveReopen() { Qualify("Access", true); }
        /// <summary>Adds the installed optional Scripting reference to an owned Publisher publication.</summary>
        [STATestMethod] public void PublisherReferenceAdditionAdapterSaveReopen() { Qualify("Publisher", false); }
        /// <summary>Qualifies Publisher library-file addition without loading or calling a Scripting object.</summary>
        [STATestMethod] public void PublisherReferenceFileAdditionAdapterSaveReopen() { Qualify("Publisher", false, true); }
        /// <summary>Removes only the optional Scripting reference from a separately owned Publisher baseline.</summary>
        [STATestMethod] public void PublisherReferenceRemovalAdapterSaveReopen() { Qualify("Publisher", true); }

        private void Qualify(string host, bool remove, bool fromFile = false)
        {
            string installedPath = RequireInstalledScripting();
            IDictionary<string, object> baseline = null, expected = null;
            OfficeAdapterOnlyProjectQualification.Run(host, remove ? "References.RemoveScripting" :
                fromFile ? "References.AddScriptingFile" : "References.AddScriptingGuid", fixture => {
                fixture.RecordAdapterStage("InstalledReferencePreflight", new { Guid = ScriptingGuid, Major = 1, Minor = 0,
                    RegisteredWin64Path = installedPath, LibraryObjectInstantiationAllowed = false });
                var original = fixture.Data("list_references");
                Assert.IsFalse(References(original).Any(IsScripting), "Fresh disposable baseline must not already contain Scripting.");
                if (remove) AddScripting(fixture, original);
            }, fixture => {
                baseline = fixture.Data("list_references");
                if (remove)
                {
                    var added = References(baseline).Single(IsScripting);
                    Assert.AreEqual(false, added["BuiltIn"]); Assert.AreEqual(false, added["IsBroken"]);
                    fixture.RecordAdapterStage("OptionalReferenceRemovalStarting", added);
                    fixture.Data("remove_reference", "Guid", ScriptingGuid, "Major", 1, "Minor", 0,
                        "ExpectedReferencesVersion", baseline["Version"]);
                }
                else AddScripting(fixture, baseline, fromFile ? installedPath : null);
                expected = fixture.Data("list_references");
                CollectionAssert.AreEqual(UnrelatedIdentities(baseline), UnrelatedIdentities(expected), "All unrelated native reference identities/order must remain unchanged.");
                Assert.AreNotEqual(baseline["Version"], expected["Version"]);
                fixture.RecordAdapterStage("ReferenceExpectedReadback", new { Removed = remove, Baseline = baseline, Expected = expected });
            }, fixture => {
                var actual = fixture.Data("list_references");
                OfficeAdapterOnlyProjectQualification.AssertReferencesEqual(expected, actual);
                Assert.AreEqual(remove ? 0 : 1, References(actual).Count(IsScripting), "Optional reference presence differs after persistence.");
                if (!remove)
                {
                    var scripting = References(actual).Single(IsScripting);
                    Assert.AreEqual(false, scripting["BuiltIn"]); Assert.AreEqual(false, scripting["IsBroken"]);
                    Assert.AreEqual("Scripting", scripting["Name"]);
                    Assert.IsTrue(File.Exists(Convert.ToString(scripting["FullPath"])), "The referenced installed library must still exist.");
                    Assert.IsTrue(string.Equals(installedPath, Path.GetFullPath(Convert.ToString(scripting["FullPath"])),
                        StringComparison.OrdinalIgnoreCase), "The live reference must identify the exact registered installed library.");
                }
            }, TestContext);
        }

        private static void AddScripting(OfficeVbeFixture fixture, IDictionary<string, object> before, string installedPath = null)
        {
            Assert.IsFalse(References(before).Any(IsScripting));
            fixture.RecordAdapterStage("OptionalReferenceAdditionStarting", new { Guid = ScriptingGuid, Major = 1, Minor = 0,
                NativeCommand = installedPath == null ? "add_reference_guid" : "add_reference_file", InstalledPath = installedPath });
            if (installedPath == null)
                fixture.Data("add_reference_guid", "Guid", ScriptingGuid, "Major", 1, "Minor", 0, "ExpectedReferencesVersion", before["Version"]);
            else
                fixture.Data("add_reference_file", "Path", installedPath, "ExpectedReferencesVersion", before["Version"]);
        }

        private static IEnumerable<IDictionary<string, object>> References(IDictionary<string, object> state) =>
            ((object[])state["References"]).Select(VbeBridgeClient.Object);
        private static bool IsScripting(IDictionary<string, object> reference) =>
            string.Equals(Convert.ToString(reference["Guid"]), ScriptingGuid, StringComparison.OrdinalIgnoreCase) &&
            Convert.ToInt32(reference["Major"]) == 1 && Convert.ToInt32(reference["Minor"]) == 0;
        private static string[] UnrelatedIdentities(IDictionary<string, object> state) => References(state).Where(reference => !IsScripting(reference))
            .Select(reference => new System.Web.Script.Serialization.JavaScriptSerializer().Serialize(reference)).ToArray();

        /// <summary>Reads registration only; unavailable libraries remain a skipped prerequisite, never an installation action.</summary>
        private static string RequireInstalledScripting()
        {
            using (var classes = RegistryKey.OpenBaseKey(RegistryHive.ClassesRoot, RegistryView.Registry64))
            using (var library = classes.OpenSubKey(@"TypeLib\" + ScriptingGuid + @"\1.0\0\win64", false))
            {
                string path = library?.GetValue(null) as string;
                if (string.IsNullOrWhiteSpace(path) || !Path.IsPathRooted(path) || !File.Exists(path))
                    Assert.Inconclusive("Installed Microsoft Scripting Runtime 1.0 win64 registration is required; no library is installed by this test.");
                return Path.GetFullPath(path);
            }
        }
    }
}

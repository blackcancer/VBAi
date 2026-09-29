using System;
using System.IO;
using System.Linq;
using System.Text.RegularExpressions;
using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace VBAi.Tests.Integration
{
    /// <summary>Qualifies installed Office hosts through their live VBAi bridge and disposable documents.</summary>
    [TestClass, TestCategory("Office")]
    public sealed class OfficeHostTests
    {
        /// <summary>Checks Word without changing Normal or an existing document.</summary>
        [STATestMethod] public void WordDocumentRoundTrip() { Qualify("Word"); }
        /// <summary>Checks a disposable macro-enabled PowerPoint presentation.</summary>
        [STATestMethod] public void PowerPointDocumentRoundTrip() { Qualify("PowerPoint"); }
        /// <summary>Checks VBA modules and MSForms in a disposable Access database.</summary>
        [STATestMethod] public void AccessDatabaseRoundTrip() { Qualify("Access"); }
        /// <summary>Checks the VBA project of a disposable Publisher publication.</summary>
        [STATestMethod] public void PublisherDocumentRoundTrip() { Qualify("Publisher"); }

        private static void Qualify(string host)
        {
            using (var fixture = OfficeVbeFixture.Start(host))
            {
                fixture.Scenario("IDE inventory", () => {
                    fixture.Data("project_properties"); fixture.Data("list_references");
                    fixture.Data("debug_state"); fixture.Data("vbe_environment");
                });
                fixture.Scenario("Host library inspection and dynamic references", () => {
                    var snapshot = fixture.Data("list_references");
                    var references = ((object[])snapshot["References"]).Select(VbeBridgeClient.Object).ToArray();
                    var library = references.Single(r => (string)r["Name"] == host);
                    var types = fixture.Data("list_reference_types", "Guid", library["Guid"], "Major", library["Major"], "Minor", library["Minor"], "Limit", 10);
                    Assert.IsTrue(Convert.ToInt32(types["TotalTypes"]) > 0);
                    const string scripting = "{420B2830-E718-11CF-893D-00A0C9054228}";
                    Assert.IsFalse(references.Any(r => string.Equals((string)r["Guid"], scripting, StringComparison.OrdinalIgnoreCase)), "The disposable project's initial references unexpectedly include Scripting.");
                    var added = fixture.Data("add_reference_guid", "Guid", scripting, "Major", 1, "Minor", 0,
                        "ExpectedReferencesVersion", snapshot["Version"]);
                    try
                    {
                        Assert.IsTrue(((object[])added["References"]).Select(VbeBridgeClient.Object).Any(r => (string)r["Name"] == "Scripting"));
                        Assert.IsTrue(Convert.ToInt32(fixture.Data("list_reference_types", "Guid", scripting, "Major", 1, "Minor", 0, "Limit", 10)["TotalTypes"]) > 0);
                    }
                    finally
                    {
                        var restored = fixture.Data("remove_reference", "Guid", scripting, "Major", 1, "Minor", 0,
                            "ExpectedReferencesVersion", fixture.Data("list_references")["Version"]);
                        Assert.AreEqual(snapshot["Version"], restored["Version"], "Original references must be restored exactly.");
                    }
                });
                fixture.Scenario("Module and class editing, exports and stale-write guard", () => {
                    foreach (var name in new[] { "VBAiOfficeModule", "VBAiOfficeClass" })
                    {
                        bool isClass = name.EndsWith("Class", StringComparison.Ordinal);
                        fixture.Data(isClass ? "create_class" : "create_module", "Module", name, "ExpectedMode", 2);
                        var before = fixture.Data("read_module", "Module", name);
                        string original = (string)before["Code"];
                        string code = original.TrimEnd() + "\r\n" +
                            (Regex.IsMatch(original, @"^\s*Option Explicit\s*$", RegexOptions.Multiline | RegexOptions.IgnoreCase) ? "" : "Option Explicit\r\n") +
                            "Public Function OfficeMarker() As String\r\n    OfficeMarker = \"VBAi office été\"\r\nEnd Function\r\n";
                        int originalLines = Convert.ToInt32(fixture.Data("component_properties", "Module", name)["CodeLines"]);
                        fixture.Data("replace_lines", "Module", name, "StartLine", 1, "Count", originalLines,
                            "ExpectedSha256", before["Sha256"], "Text", code);
                        var after = fixture.Data("read_module", "Module", name);
                        StringAssert.Contains((string)after["Code"], "VBAi office été");
                        var stale = fixture.Response("replace_lines", "Module", name, "StartLine", 1, "Count", 0,
                            "ExpectedSha256", before["Sha256"], "Text", "' must not be inserted");
                        Assert.AreEqual(false, stale["Ok"]);
                        Assert.AreEqual(after["Sha256"], fixture.Data("read_module", "Module", name)["Sha256"]);
                        string path = Path.Combine(fixture.Root, name + (isClass ? ".cls" : ".bas"));
                        var state = fixture.Data("component_properties", "Module", name);
                        fixture.Data("export_component", "Module", name, "Path", path, "ExpectedComponentVersion", state["Version"]);
                        Assert.IsTrue(File.Exists(path));
                        fixture.Data("remove_component", "Module", name, "ExpectedComponentVersion", state["Version"],
                            "ExpectedProjectVersion", fixture.Data("project_properties")["Version"]);
                        fixture.Data("import_component", "Path", path,
                            "ExpectedProjectVersion", fixture.Data("project_properties")["Version"]);
                        StringAssert.Contains((string)fixture.Data("read_module", "Module", name)["Code"], "VBAi office été");
                    }
                });
                fixture.Scenario("Native form capability", () => {
                    var before = fixture.Items("list_modules").Select(m => (string)m["Name"]).OrderBy(n => n).ToArray();
                    var created = fixture.Response("create_form", "Form", "VBAiOfficeForm");
                    if (host == "Access" && !Convert.ToBoolean(created["Ok"]))
                    {
                        Assert.AreEqual(false, created["Ok"], "Access native forms must not be mistaken for MSForms.");
                        CollectionAssert.AreEqual(before, fixture.Items("list_modules").Select(m => (string)m["Name"]).OrderBy(n => n).ToArray(), "Rejected creation must leave no component behind.");
                        return;
                    }
                    Assert.AreEqual(true, created["Ok"], Convert.ToString(created["Error"]));
                    fixture.Data("add_form_control", "Form", "VBAiOfficeForm", "Control", "OfficeLabel", "ControlType", "Forms.Label.1",
                        "Left", 12, "Top", 12, "Width", 140, "Height", 24, "Caption", "VBAi office été",
                        "ExpectedFormVersion", fixture.Data("form_state", "Form", "VBAiOfficeForm")["Version"]);
                    fixture.Data("set_form_control_font", "Form", "VBAiOfficeForm", "Control", "OfficeLabel",
                        "FontName", "Arial", "FontSize", 12, "FontBold", true,
                        "ExpectedFormVersion", fixture.Data("form_state", "Form", "VBAiOfficeForm")["Version"]);
                    AssertLabel(fixture);
                });
                fixture.Scenario("Host save capability and policy guards", () => {
                    var persistence = fixture.Data("project_persistence_status");
                    var saved = fixture.Response("save_host_document", "ExpectedHostPath", fixture.DocumentPath,
                        "ExpectedProjectVersion", fixture.Data("project_properties")["Version"]);
                    if (Convert.ToBoolean(persistence["HostAvailable"]))
                        Assert.AreEqual(true, saved["Ok"], Convert.ToString(saved["Error"]));
                    else
                    {
                        Assert.AreEqual(false, saved["Ok"], "Unavailable host saving must fail closed.");
                        Assert.IsFalse(string.IsNullOrWhiteSpace(Convert.ToString(saved["Error"])));
                    }
                });
                fixture.Scenario("Code navigation and native compilation", () => {
                    var module = fixture.Data("read_module", "Module", "VBAiOfficeModule");
                    fixture.Data("select_code", "Module", "VBAiOfficeModule", "StartLine", 2, "ExpectedSha256", module["Sha256"]);
                    Assert.AreEqual(true, fixture.Data("compile_project", "ExpectedMode", 2)["Compiled"]);
                    Assert.AreEqual(2, Convert.ToInt32(fixture.Data("debug_state")["Mode"]));
                });
                fixture.Scenario("Native save and reopen preserves VBA", () => {
                    fixture.Reopen();
                    foreach (var name in new[] { "VBAiOfficeModule", "VBAiOfficeClass" })
                        StringAssert.Contains((string)fixture.Data("read_module", "Module", name)["Code"], "VBAi office été");
                    if (fixture.Items("list_modules").Any(m => (string)m["Name"] == "VBAiOfficeForm")) AssertLabel(fixture);
                });
                Assert.AreEqual(0, fixture.Failures.Count, string.Join(Environment.NewLine, fixture.Failures));
            }
        }

        private static void AssertLabel(OfficeVbeFixture fixture)
        {
            var form = fixture.Data("form_state", "Form", "VBAiOfficeForm");
            var label = ((object[])form["Controls"]).Select(VbeBridgeClient.Object).Single(c => (string)c["Name"] == "OfficeLabel");
            Assert.AreEqual("VBAi office été", label["Caption"]);
            Assert.AreEqual("Arial", label["FontName"]);
            Assert.AreEqual(12d, Convert.ToDouble(label["FontSize"]), 0.1d);
            Assert.AreEqual(true, label["FontBold"]);
            Assert.AreEqual(140d, Convert.ToDouble(label["Width"]), 0.1d);
        }
    }
}

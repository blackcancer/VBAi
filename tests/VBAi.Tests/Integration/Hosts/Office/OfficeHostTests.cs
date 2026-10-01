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
        /// <summary>Checks VBA modules/classes in a disposable Access database; native Access forms are separate from MSForms.</summary>
        [STATestMethod] public void AccessDatabaseRoundTrip() { Qualify("Access"); }
        /// <summary>Checks the VBA project of a disposable Publisher publication.</summary>
        [STATestMethod] public void PublisherDocumentRoundTrip() { Qualify("Publisher"); }

        private static void Qualify(string host)
        {
            using (var fixture = OfficeVbeFixture.Start(host))
            {
                bool formCreated = false;
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
                    formCreated = true;
                    fixture.Data("add_form_control", "Form", "VBAiOfficeForm", "Control", "OfficeLabel", "ControlType", "Forms.Label.1",
                        "Left", 12, "Top", 12, "Width", 140, "Height", 24, "Caption", "VBAi office été",
                        "ExpectedFormVersion", fixture.Data("form_state", "Form", "VBAiOfficeForm")["Version"]);
                    fixture.Data("set_form_control_font", "Form", "VBAiOfficeForm", "Control", "OfficeLabel",
                        "FontName", "Arial", "FontSize", 12, "FontBold", true,
                        "ExpectedFormVersion", fixture.Data("form_state", "Form", "VBAiOfficeForm")["Version"]);
                    AssertLabel(fixture);
                });
                fixture.Scenario("Host save capability and policy guards", () => {
                    bool accessOrPublisher = host == "Access" || host == "Publisher";
                    const string description = "VBAi owned adapter persistence été";
                    if (accessOrPublisher)
                    {
                        fixture.Data("set_project_property", "Property", "Description", "Value", description,
                            "ExpectedProjectVersion", fixture.Data("project_properties")["Version"]);
                        fixture.Data("add_reference_guid", "Guid", "{420B2830-E718-11CF-893D-00A0C9054228}", "Major", 1, "Minor", 0,
                            "ExpectedReferencesVersion", fixture.Data("list_references")["Version"]);
                        var selected = fixture.Data("read_module", "Module", "VBAiOfficeModule");
                        fixture.Data("select_code", "Module", "VBAiOfficeModule", "StartLine", 1, "ExpectedSha256", selected["Sha256"]);
                    }
                    fixture.StopAccessSaveDialogHandler();
                    foreach (string name in new[] { "VBAiOfficeModule", "VBAiOfficeClass" })
                    {
                        var beforeSaveEdit = fixture.Data("read_module", "Module", name);
                        int lines = Convert.ToInt32(fixture.Data("component_properties", "Module", name)["CodeLines"]);
                        string freshCode = ((string)beforeSaveEdit["Code"]).TrimEnd() + "\r\n' Adapter-only pending edit " + Guid.NewGuid().ToString("N") + "\r\n";
                        fixture.Data("replace_lines", "Module", name, "StartLine", 1, "Count", lines,
                            "ExpectedSha256", beforeSaveEdit["Sha256"], "Text", freshCode);
                        Assert.AreNotEqual(beforeSaveEdit["Sha256"], fixture.Data("read_module", "Module", name)["Sha256"],
                            "A fresh source mutation must precede the adapter's save.");
                    }
                    var sourceHashes = new System.Collections.Generic.Dictionary<string, string>();
                    foreach (string name in new[] { "VBAiOfficeModule", "VBAiOfficeClass" })
                        sourceHashes[name] = (string)fixture.Data("read_module", "Module", name)["Sha256"];
                    string referencesVersion = (string)fixture.Data("list_references")["Version"];
                    var persistence = fixture.Data("project_persistence_status");
                    if (accessOrPublisher)
                    {
                        Assert.AreEqual(true, persistence["HostAvailable"], "Q-012 requires an available native adapter; safe refusal does not qualify save/reopen.");
                        Assert.AreEqual(false, persistence["ProjectSaved"], "The fresh VBA edit must remain unsaved before the adapter invocation.");
                    }
                    fixture.RecordNativePersistence("BeforeAdapterSave");
                    var saved = fixture.Response("save_host_document", "ExpectedHostPath", fixture.DocumentPath,
                        "ExpectedProjectVersion", fixture.Data("project_properties")["Version"]);
                    fixture.RecordNativePersistence("ImmediatelyAfterAdapterSave");
                    if (host == "Access" && Convert.ToBoolean(saved["Ok"]))
                    {
                        var observation = VbeBridgeClient.Object(saved["Data"]);
                        if (observation.ContainsKey("Uncertain") && Convert.ToBoolean(observation["Uncertain"]))
                        {
                            // Observe the owning host after it regains its message loop. No save,
                            // compile or helper reopen may intervene or turn the failed result green.
                            int elapsed = 0;
                            foreach (int delay in new[] { 100, 250, 1000 })
                            {
                                System.Threading.Thread.Sleep(delay); elapsed += delay;
                                fixture.RecordNativePersistence("UncertainAdapterReadOnlyAfter" + elapsed + "ms");
                            }
                        }
                    }
                    if (Convert.ToBoolean(persistence["HostAvailable"]))
                    {
                        Assert.AreEqual(true, saved["Ok"], Convert.ToString(saved["Error"]));
                        var result = VbeBridgeClient.Object(saved["Data"]);
                        Assert.AreEqual(true, result["Verified"], "The adapter must verify its native save; protocol success alone is insufficient.");
                        Assert.AreEqual(false, result["Uncertain"], "An uncertain native save is not a qualified persistence result.");
                        Assert.AreEqual(true, result["MutationInvoked"]);
                        Assert.AreEqual(false, result["PersistenceReopenVerified"], "The adapter cannot claim fixture reopen evidence before reopening.");
                        if (host == "Access") Assert.IsNull(result["HostSaved"], "Access must not fabricate a document Saved property.");
                        fixture.ReopenFromDisk();
                        AssertPersistedContent(fixture, formCreated);
                        foreach (var source in sourceHashes)
                            Assert.AreEqual(source.Value, fixture.Data("read_module", "Module", source.Key)["Sha256"], "Exact VBA source changed after adapter-only reopen: " + source.Key);
                        Assert.AreEqual(referencesVersion, fixture.Data("list_references")["Version"], "References changed after adapter-only reopen.");
                        if (accessOrPublisher)
                        {
                            var properties = ((object[])fixture.Data("project_properties")["Properties"]).Select(VbeBridgeClient.Object);
                            Assert.AreEqual(description, properties.Single(property => (string)property["Name"] == "Description")["Value"],
                                "The adapter did not persist the project description.");
                        }
                        fixture.RecordNativePersistence("AfterAdapterOnlyReopen");
                    }
                    else
                    {
                        Assert.AreEqual(false, saved["Ok"], "Unavailable host saving must fail closed.");
                        Assert.IsFalse(string.IsNullOrWhiteSpace(Convert.ToString(saved["Error"])));
                        fixture.CompatibilityGap("save_host_document", Convert.ToString(persistence["Reason"]));
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
                    AssertPersistedContent(fixture, formCreated);
                });
                Assert.AreEqual(0, fixture.Failures.Count, string.Join(Environment.NewLine, fixture.Failures));
            }
        }

        private static void AssertPersistedContent(OfficeVbeFixture fixture, bool expectForm)
        {
            foreach (var name in new[] { "VBAiOfficeModule", "VBAiOfficeClass" })
                StringAssert.Contains((string)fixture.Data("read_module", "Module", name)["Code"], "VBAi office été");
            if (expectForm) AssertLabel(fixture);
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

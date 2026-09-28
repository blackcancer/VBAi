using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading;
using System.Web.Script.Serialization;
using System.IO;
using System.Text;
using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace CodexVBE.Tests.Integration
{
    /// <summary>Qualifie les nouvelles opérations IDE dans un Excel visible et jetable.</summary>
    [TestClass, TestCategory("Excel")]
    public sealed class ExcelIdeSurfaceTests
    {
        /// <summary>Relit la page Toolbox réelle et refuse une mutation dont le contexte natif n'est pas qualifié.</summary>
        [STATestMethod]
        public void NativeToolboxPagesAreObservedWithoutClaimingMutationSupport()
        {
            using (var host = ExcelVbeFixture.Start())
            {
                string project = Convert.ToString(VbeBridgeClient.Object(((object[])host.Command("list_projects")["Data"]).Single())["Name"]);
                Data(host.Command(new { Command = "create_form", Project = project, Form = "ToolboxProbe" }));
                host.FocusFormToolbox("ToolboxProbe");
                var before = Data(host.Command(new { Command = "read_navigation_surface", Pane = "toolbox" }));
                Assert.AreEqual(true, before["Available"], new JavaScriptSerializer().Serialize(before));
                var pages = ((object[])before["Nodes"]).Select(VbeBridgeClient.Object).ToArray();
                var selectedPages = pages.Where(item => (Convert.ToString(item["Name"]) == "Contrôles" || Convert.ToString(item["Name"]) == "Controls") &&
                    Equals(item["ObservedSelected"], true)).ToArray();
                Assert.AreEqual(1, selectedPages.Length, new JavaScriptSerializer().Serialize(before));
                var page = selectedPages[0];
                Assert.IsNull(page["Selected"], "Observation is distinct from qualified mutation support.");
                Assert.AreEqual(true, page["ObservedSelected"]);
                Assert.IsFalse(string.IsNullOrWhiteSpace(Convert.ToString(page["ActionUnavailableReason"])));
                var refused = host.Command(new { Command = "change_navigation_surface", Pane = "toolbox", Control = page["Token"],
                    Action = "select", ExpectedWindowVersion = before["WindowVersion"] });
                Assert.AreEqual(false, refused["Ok"], "No MSAA mutation is delivered.");
                var after = Data(host.Command(new { Command = "read_navigation_surface", Pane = "toolbox" }));
                Assert.AreEqual(before["WindowVersion"], after["WindowVersion"], "The refusal retains the observed surface unchanged.");
            }
        }
        /// <summary>Configure un secret local jetable et confirme le verrouillage après sauvegarde et réouverture réelles.</summary>
        [STATestMethod]
        public void ProjectProtectionPersistsAfterNativeSaveAndReopen()
        {
            using (var host = ExcelVbeFixture.Start())
            {
                string project = Convert.ToString(VbeBridgeClient.Object(((object[])host.Command("list_projects")["Data"]).Single())["Name"]);
                WriteModule(host, project, "ProtectionProbe", "Option Explicit\r\nPublic Sub Probe()\r\nEnd Sub");
                host.FocusCodeModule("ProtectionProbe");
                var properties = Data(host.Command(new { Command = "project_properties", Project = project }));
                var commands = (object[])host.Command(new { Command = "list_commands", Limit = 500 })["Data"];
                var command = commands.Select(VbeBridgeClient.Object).First(item => Convert.ToInt32(item["Id"]) == 2578 && Convert.ToBoolean(item["Enabled"]));
                var protection = Data(host.Command(new { Command = "read_project_protection", Project = project, ExpectedMode = 2,
                    ExpectedProjectVersion = properties["Version"], ControlCaption = command["Caption"] }));
                string path = host.File("protected.xlsm"), secretFile = host.File("secret.txt");
                File.WriteAllText(secretFile, Guid.NewGuid().ToString("N"), new UTF8Encoding(false));
                try
                {
                    var locked = Data(host.Command(new { Command = "set_project_protection", Project = project, ExpectedMode = 2,
                        ExpectedProjectVersion = properties["Version"], ControlCaption = command["Caption"], Action = "lock",
                        ExpectedOptionsVersion = protection["OptionsVersion"], Path = secretFile }));
                    Assert.AreEqual(true, locked["Available"], new JavaScriptSerializer().Serialize(locked));
                    Assert.AreEqual(true, locked["ControlValueVerified"]); Assert.AreEqual(true, locked["DialogClosed"]);
                    Assert.AreEqual(false, locked["PersistenceVerified"]);
                    properties = Data(host.Command(new { Command = "project_properties", Project = project }));
                    Data(host.Command(new { Command = "save_host_document_as", Project = project, Path = path, ExpectedProjectVersion = properties["Version"] }));
                    Assert.AreEqual(1, host.ReopenAndReadProjectProtection(path), "Independent COM readback after reopening must report vbext_pp_locked.");
                }
                finally { File.Delete(secretFile); }
            }
        }
        /// <summary>Modifie puis restaure des préférences de Format/Ancrage et vérifie l'état complet après réouverture.</summary>
        [STATestMethod]
        public void ExtendedOptionsRoundTripAndRestoreTheCompleteNativeState()
        {
            using (var host = ExcelVbeFixture.Start())
            {
                var baseline = Data(host.Command(new { Command = "read_vbe_options" }));
                var tabs = ((object[])baseline["Tabs"]).Select(VbeBridgeClient.Object).ToArray();
                foreach (string fragment in new[] { "format", "ancrage" })
                {
                    var tab = tabs.Single(item => Convert.ToString(item["Tab"]).ToLowerInvariant().Contains(fragment));
                    var checkbox = ((object[])tab["Controls"]).Select(VbeBridgeClient.Object).First(item =>
                        Convert.ToString(item["Type"]) == "ControlType.CheckBox" && item["Error"] == null);
                    bool original = Convert.ToString(checkbox["Value"]) == "On";
                    try
                    {
                        var before = Data(host.Command(new { Command = "read_vbe_options" }));
                        var applied = Data(host.Command(new { Command = "set_vbe_option", Pane = tab["Tab"], Property = checkbox["Name"],
                            Value = !original, ExpectedOptionsVersion = before["OptionsVersion"] }));
                        Assert.AreEqual(true, applied["ControlValueVerified"]);
                        var after = Data(host.Command(new { Command = "read_vbe_options" }));
                        var selectedTab = ((object[])after["Tabs"]).Select(VbeBridgeClient.Object).Single(item => Equals(item["Tab"], tab["Tab"]));
                        var selected = ((object[])selectedTab["Controls"]).Select(VbeBridgeClient.Object).Single(item =>
                            Equals(item["Name"], checkbox["Name"]) && Equals(item["Type"], checkbox["Type"]));
                        Assert.AreEqual(original ? "Off" : "On", selected["Value"]);
                    }
                    finally
                    {
                        var current = Data(host.Command(new { Command = "read_vbe_options" }));
                        Data(host.Command(new { Command = "set_vbe_option", Pane = tab["Tab"], Property = checkbox["Name"],
                            Value = original, ExpectedOptionsVersion = current["OptionsVersion"] }));
                        var restored = Data(host.Command(new { Command = "read_vbe_options" }));
                        Assert.AreEqual(baseline["OptionsVersion"], restored["OptionsVersion"], "The complete native preferences must be restored.");
                    }
                }
            }
        }
        /// <summary>Lit l'arbre natif par son HWND et inspecte la protection sans changer le projet.</summary>
        [STATestMethod]
        public void NativeProjectTreeAndProtectionDialogExposeTheirActualState()
        {
            using (var host = ExcelVbeFixture.Start())
            {
                string project = Convert.ToString(VbeBridgeClient.Object(((object[])host.Command("list_projects")["Data"]).Single())["Name"]);
                WriteModule(host, project, "SurfaceProbe", "Option Explicit\r\nPublic Sub Probe()\r\nEnd Sub");
                host.FocusCodeModule("SurfaceProbe");
                var tree = Data(host.Command(new { Command = "read_navigation_surface", Pane = "project" }));
                Assert.AreEqual(true, tree["Available"], new JavaScriptSerializer().Serialize(tree));
                Assert.IsTrue(((object[])tree["Nodes"]).Length > 0);
                var properties = Data(host.Command(new { Command = "project_properties", Project = project }));
                var commands = (object[])host.Command(new { Command = "list_commands", Limit = 500 })["Data"];
                var command = commands.Select(VbeBridgeClient.Object).First(item => Convert.ToInt32(item["Id"]) == 2578 && Convert.ToBoolean(item["Enabled"]));
                var protection = Data(host.Command(new { Command = "read_project_protection", Project = project, ExpectedMode = 2,
                    ExpectedProjectVersion = properties["Version"], ControlCaption = command["Caption"] }));
                Assert.AreEqual(true, protection["Available"], new JavaScriptSerializer().Serialize(protection));
                Assert.AreEqual(true, protection["DialogClosed"]);
                Assert.AreEqual(false, protection["LockedForViewing"]);
                Assert.AreEqual(false, protection["PasswordPresent"]);
                Assert.AreEqual(false, protection["PersistenceVerified"]);
            }
        }
        /// <summary>Un renommage public relit les deux modules, exécute l'appel qualifié puis permet leur annulation séparée.</summary>
        [STATestMethod]
        public void PublicProcedureRenameAcrossModulesPreservesNativeExecutionAndPerModuleUndo()
        {
            using (var host = ExcelVbeFixture.Start())
            {
                string project = Convert.ToString(VbeBridgeClient.Object(((object[])host.Command("list_projects")["Data"]).Single())["Name"]);
                const string target = "Option Explicit\r\nPublic Function Compute(ByVal value As Long) As Long\r\nCompute = value * 2\r\nEnd Function";
                const string caller = "Option Explicit\r\nPublic Sub Entry()\r\nThisWorkbook.Worksheets(1).Range(\"A1\").Value2 = MathModule.Compute(21)\r\nEnd Sub";
                var source = WriteModule(host, project, "MathModule", target);
                var other = WriteModule(host, project, "Caller", caller);
                var preview = Data(host.Command(new { Command = "preview_procedure_rename", Project = project, Module = "MathModule",
                    Query = "Compute", NewName = "Calculate", ProcKind = 0, StartLine = 2, StartColumn = 17, ExpectedSha256 = source["Sha256"] }));
                Assert.AreEqual(2, ((object[])preview["Edits"]).Length);
                Assert.AreEqual(source["Code"], Data(host.Command(new { Command = "read_module", Project = project, Module = "MathModule" }))["Code"]);
                Data(host.Command(new { Command = "apply_procedure_rename", Project = project, Module = "MathModule", Query = "Compute", NewName = "Calculate",
                    ProcKind = 0, StartLine = 2, StartColumn = 17, ExpectedSha256 = source["Sha256"], ExpectedProjectVersion = preview["ExpectedProjectVersion"], ExpectedMode = 2 }));
                var renamed = Data(host.Command(new { Command = "read_module", Project = project, Module = "Caller" }));
                StringAssert.Contains(Convert.ToString(renamed["Code"]), "MathModule.Calculate(21)");
                var run = Data(host.Command(new { Command = "run_procedure", Project = project, Module = "Caller", Procedure = "Entry",
                    ExpectedSha256 = renamed["Sha256"], ExpectedMode = 2, Arguments = new object[0] }));
                DateTime deadline = DateTime.UtcNow.AddSeconds(15);
                while (Convert.ToBoolean(run["Pending"]) && DateTime.UtcNow < deadline)
                {
                    Thread.Sleep(100);
                    run = Data(host.Command(new { Command = "procedure_run_status", Project = project, Query = run["Query"] }));
                }
                Assert.AreEqual("Delivered", run["State"]);
                Assert.AreEqual(42d, Convert.ToDouble(host.ReadCell("A1")), "Independent Excel result.");
                foreach (var module in new[] { "Caller", "MathModule" })
                {
                    var now = Data(host.Command(new { Command = "read_module", Project = project, Module = module }));
                    Data(host.Command(new { Command = "undo_code_edit", Project = project, Module = module, ExpectedSha256 = now["Sha256"] }));
                    Assert.AreEqual((module == "Caller" ? other : source)["Code"], Data(host.Command(new { Command = "read_module", Project = project, Module = module }))["Code"]);
                }
            }
        }

        /// <summary>Le réglage de taille et d'étendue est confirmé par les métriques COM du classeur possédé.</summary>
        [STATestMethod]
        public void FormContentFittingUsesNativeInsideDimensionsAndRetainsItsChildren()
        {
            using (var host = ExcelVbeFixture.Start())
            {
                string project = Convert.ToString(VbeBridgeClient.Object(((object[])host.Command("list_projects")["Data"]).Single())["Name"]);
                const string form = "ContentProbe";
                Data(host.Command(new { Command = "create_form", Project = project, Form = form }));
                var state = Data(host.Command(new { Command = "form_state", Project = project, Form = form }));
                Data(host.Command(new { Command = "add_form_control", Project = project, Form = form, ExpectedFormVersion = state["Version"],
                    ControlType = "Forms.Label.1", Control = "ContentLabel", Left = 20d, Top = 30d, Width = 90d, Height = 25d, Caption = "Content" }));
                var tree = Data(host.Command(new { Command = "form_tree", Project = project, Form = form }));
                var preview = Data(host.Command(new { Command = "preview_fit_form_content", Project = project, Form = form, ExpectedTreeVersion = tree["TreeVersion"], Action = "fit_container", Left = 10d, Top = 12d }));
                Assert.AreEqual(true, preview["ReadOnly"]);
                var applied = Data(host.Command(new { Command = "apply_fit_form_content", Project = project, Form = form, ExpectedTreeVersion = tree["TreeVersion"], Action = "fit_container", Left = 10d, Top = 12d }));
                Assert.AreEqual(true, applied["Verified"], new JavaScriptSerializer().Serialize(applied));
                double tolerance = Convert.ToDouble(VbeBridgeClient.Object(applied["Plan"])["TolerancePoints"]);
                double insideWidth = host.ReadDesignerMetric(form, "InsideWidth"), insideHeight = host.ReadDesignerMetric(form, "InsideHeight");
                Assert.IsTrue(insideWidth >= 120d && insideWidth <= 120d + tolerance);
                Assert.IsTrue(insideHeight >= 67d && insideHeight <= 67d + tolerance);
                tree = Data(host.Command(new { Command = "form_tree", Project = project, Form = form }));
                Assert.AreEqual(1, ((object[])tree["Controls"]).Length);
                var scroll = Data(host.Command(new { Command = "apply_fit_form_content", Project = project, Form = form, ExpectedTreeVersion = tree["TreeVersion"], Action = "fit_scroll_extent", Left = 30d, Top = 35d }));
                Assert.AreEqual(true, scroll["Verified"]);
                Assert.AreEqual(140d, host.ReadDesignerMetric(form, "ScrollWidth"), 0.1);
                Assert.AreEqual(90d, host.ReadDesignerMetric(form, "ScrollHeight"), 0.1);
            }
        }

        /// <summary>Crée un module de test et relit sa source canonique avant les mutations versionnées.</summary>
        private static IDictionary<string, object> WriteModule(ExcelVbeFixture host, string project, string module, string source)
        {
            var created = Data(host.Command(new { Command = "create_module", Project = project, Module = module, ExpectedMode = 2 }));
            var initial = Data(host.Command(new { Command = "read_module", Project = project, Module = module }));
            Data(host.Command(new { Command = "replace_lines", Project = project, Module = module, ExpectedSha256 = initial["Sha256"], StartLine = 1, Count = created["Lines"], Text = source }));
            return Data(host.Command(new { Command = "read_module", Project = project, Module = module }));
        }

        /// <summary>Vérifie la réponse du véritable pont avant d'en lire les données.</summary>
        private static IDictionary<string, object> Data(IDictionary<string, object> response)
        {
            Assert.IsNotNull(response); Assert.AreEqual(true, response["Ok"], Convert.ToString(response["Error"]));
            return VbeBridgeClient.Object(response["Data"]);
        }
    }
}

using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Web.Script.Serialization;
using VBAi.Tests.Integration;
using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace VBAi.Tests.Integration.Hosts.Excel
{
    /// <summary>Qualification des choix Format via le vrai bridge Excel, avec restauration vérifiée et preuves conservées en cas d'échec.</summary>
    [TestClass, TestCategory("Excel")]
    public sealed class ExcelFormatOptionsTests
    {
        /// <summary>Conserve les instantanés avant mutation même si la restauration échoue.</summary>
        public TestContext TestContext { get; set; }
        private int evidenceSequence;
        private string evidenceDirectory;
        /// <summary>Matrice fixée avant exécution : police, taille native ou refus, palettes, catégorie distincte, marge, garde et restauration.</summary>
        private static readonly string[] Scenarios = { "exact font", "size catalogue or honest refusal", "foreground", "background", "indicator",
            "nondefault category foreground", "margin indicator", "stale version refuses", "full options version restored" };

        /// <summary>Exerce les préférences réelles dans une instance visible possédée; aucun nom/RGB de palette n'est supposé.</summary>
        [TestMethod]
        public void NativeFormatChoicesRoundTripAndRestoreCompleteOptionsVersion()
        {
            Assert.AreEqual(9, Scenarios.Length);
            string output = Environment.GetEnvironmentVariable("VBAi_TEST_FORMAT_OPTIONS_OUTPUT");
            if (!string.IsNullOrEmpty(output) && !Path.IsPathRooted(output))
                throw new ArgumentException("VBAi_TEST_FORMAT_OPTIONS_OUTPUT must be an absolute path.");
            evidenceDirectory = Path.Combine(string.IsNullOrEmpty(output) ? TestContext.TestRunDirectory : output,
                "options-evidence-" + Guid.NewGuid().ToString("N"));
            Directory.CreateDirectory(evidenceDirectory);
            TestContext.WriteLine("Retained options evidence: " + evidenceDirectory);
            using (var excel = ExcelVbeFixture.Start())
            {
                var baseline = Read(excel); var format = Format(baseline); string tab = (string)format["Tab"];
                string baselineFile = Path.Combine(evidenceDirectory, "options-baseline-" + excel.ProcessId + ".json");
                File.WriteAllText(baselineFile, new JavaScriptSerializer { MaxJsonLength = 10 * 1024 * 1024 }.Serialize(baseline), new UTF8Encoding(false));
                TestContext.AddResultFile(baselineFile);
                var restore = new List<Tuple<string, object, string>>();
                Exception primaryFailure = null;
                try
                {
                    var font = Find(format, "Font", "Police :");
                    string alternate = Choices(font).FirstOrDefault(x => x != (string)font["Value"] &&
                        new[] { "Consolas (Occidental)", "Consolas (Western)", "Consolas", "Courier New (Occidental)", "Courier New (Western)", "Courier New" }.Contains(x));
                    Assert.IsNotNull(alternate, "An exact observed alternate font is required for this bounded fixture.");
                    restore.Add(Tuple.Create((string)font["Name"], font["Value"], (string)null));
                    Write(excel, tab, (string)font["Name"], alternate, null);
                    Assert.AreEqual(alternate, Find(Format(Read(excel)), "Font", "Police :")["Value"]);
                    var sizeRead = Read(excel); var size = Find(Format(sizeRead), "Size", "Taille :");
                    var sizes = Choices(size);
                    if (sizes.Length == 0)
                    {
                        Assert.IsNotNull(size["Value"], "Empty native catalogue must still report the real edit value.");
                        var rejected = ObserveCommand(excel, new { Command = "set_vbe_option", Pane = tab, Property = size["Name"], Value = "12", ExpectedOptionsVersion = sizeRead["OptionsVersion"] });
                        Assert.AreEqual(false, rejected["Ok"], "No arbitrary size may be invented outside native choices.");
                        Assert.AreEqual(sizeRead["OptionsVersion"], Read(excel)["OptionsVersion"]);
                    }
                    else
                    {
                        string next = sizes.First(x => x != (string)size["Value"]);
                        restore.Add(Tuple.Create((string)size["Name"], size["Value"], (string)null));
                        Write(excel, tab, (string)size["Name"], next, null);
                        Assert.AreEqual(next, Find(Format(Read(excel)), "Size", "Taille :")["Value"]);
                    }
                    foreach (var names in new[] { new[] { "Foreground", "Premier plan :" }, new[] { "Background", "Arrière-plan :" }, new[] { "Indicator", "Indicateur :" } })
                    {
                        var paletteFormat = Format(Read(excel));
                        string currentCategory = CurrentCategory(paletteFormat);
                        var control = CategoryPalette(paletteFormat, currentCategory, names);
                        string next = Choices(control).First(x => x != (string)control["Value"]);
                        restore.Add(Tuple.Create((string)control["Name"], control["Value"], currentCategory));
                        Write(excel, tab, (string)control["Name"], next, currentCategory);
                        Assert.AreEqual(next, CategoryPalette(Format(Read(excel)), currentCategory, names)["Value"]);
                    }
                    var categoryRead = Read(excel); var categories = (object[])Format(categoryRead)["FormatCategories"];
                    Assert.IsTrue(categories.Length > 1);
                    var other = VbeBridgeClient.Object(categories[1]);
                    var palettes = ((object[])other["Palettes"]).Select(VbeBridgeClient.Object);
                    var foreground = palettes.Single(x => new[] { "Foreground", "Premier plan :" }.Contains((string)x["Name"]));
                    string category = (string)other["Category"];
                    string colour = Choices(foreground).First(x => x != (string)foreground["Value"]);
                    restore.Add(Tuple.Create((string)foreground["Name"], foreground["Value"], category));
                    Write(excel, tab, (string)foreground["Name"], colour, category);
                    var afterOther = ((object[])Format(Read(excel))["FormatCategories"]).Select(VbeBridgeClient.Object).Single(x => (string)x["Category"] == category);
                    Assert.AreEqual(colour, ((object[])afterOther["Palettes"]).Select(VbeBridgeClient.Object).Single(x => (string)x["Name"] == (string)foreground["Name"])["Value"]);
                    var margin = Find(Format(Read(excel)), "Margin Indicator Bar", "Barre des indicateurs en marge");
                    restore.Add(Tuple.Create((string)margin["Name"], (object)((string)margin["Value"] == "On"), (string)null));
                    Write(excel, tab, (string)margin["Name"], (string)margin["Value"] != "On", null);
                    var stale = ObserveCommand(excel, new { Command = "set_vbe_option", Pane = tab, Property = font["Name"], Value = font["Value"], ExpectedOptionsVersion = baseline["OptionsVersion"] });
                    Assert.AreEqual(false, stale["Ok"]);
                }
                catch (Exception error) { primaryFailure = error; throw; }
                finally
                {
                    var restorationErrors = new List<Exception>();
                    // Preserve the earliest value for each preference, then compensate each distinct entry once.
                    // A failed compensation must not prevent independent preferences from being restored.
                    foreach (var item in restore.GroupBy(x => Tuple.Create(x.Item1, x.Item3)).Select(x => x.First()).Reverse())
                    {
                        try { Write(excel, tab, item.Item1, item.Item2, item.Item3, restoring: true); }
                        catch (Exception error)
                        {
                            restorationErrors.Add(new InvalidOperationException("Restoration failed for " + item.Item1 +
                                " (category: " + (item.Item3 ?? "current") + "); this entry was not retried.", error));
                        }
                    }
                    try
                    {
                        var restored = Read(excel);
                        AttachEvidence(excel.ProcessId, "restoration-readback", new { BaselineVersion = baseline["OptionsVersion"],
                            Readback = restored, EntryFailures = restorationErrors.Select(x => x.ToString()).ToArray() });
                        Assert.AreEqual(baseline["OptionsVersion"], restored["OptionsVersion"], "Every category and preference must return to the complete baseline.");
                    }
                    catch (Exception error) { restorationErrors.Add(error); }
                    if (restorationErrors.Count != 0)
                    {
                        if (primaryFailure != null) restorationErrors.Insert(0, primaryFailure);
                        throw new AggregateException("Native scenario/restoration failures; baseline and available diagnostics attached. Each distinct restoration entry was attempted at most once.", restorationErrors);
                    }
                }
            }
        }

        /// <summary>Lit les quatre onglets et exige une réponse native positive.</summary>
        private IDictionary<string, object> Read(ExcelVbeFixture excel) => Data(ObserveCommand(excel, new { Command = "read_vbe_options" }));
        /// <summary>Extrait les données d'une réponse après validation de sa réussite.</summary>
        private static IDictionary<string, object> Data(IDictionary<string, object> response)
        { Assert.IsNotNull(response); Assert.AreEqual(true, response["Ok"], Convert.ToString(response["Error"])); return VbeBridgeClient.Object(response["Data"]); }
        /// <summary>Localise l'onglet Format observé en anglais ou français.</summary>
        private static IDictionary<string, object> Format(IDictionary<string, object> data) => ((object[])data["Tabs"]).Select(VbeBridgeClient.Object)
            .Single(x => new[] { "Editor Format", "Format de l'éditeur", "Format de l’éditeur" }.Contains((string)x["Tab"]));
        /// <summary>Localise un contrôle éditable exact, en ignorant les labels de texte.</summary>
        private static IDictionary<string, object> Find(IDictionary<string, object> format, params string[] names) => ((object[])format["Controls"]).Select(VbeBridgeClient.Object)
            .Single(x => names.Contains((string)x["Name"]) && (string)x["Type"] != "ControlType.Text");
        /// <summary>Reads the exact selected category reported by the native Code Colors control.</summary>
        private static string CurrentCategory(IDictionary<string, object> format)
        {
            var control = ((object[])format["Controls"]).Select(VbeBridgeClient.Object).Single(x =>
                (string)x["Type"] == "ControlType.List" && new[] { "code colors", "couleurs du code", "color text", "texte couleur" }
                    .Contains(((string)x["Name"] ?? "").Replace("&", "").Trim().TrimEnd(':').Trim().ToLowerInvariant()));
            string category = control["Value"] as string;
            Assert.IsFalse(string.IsNullOrWhiteSpace(category), "The observed color category must be readable.");
            Assert.AreEqual(1, Choices(control).Count(x => x == category), "The observed category must be unique.");
            return category;
        }
        /// <summary>Resolves a palette in the exact saved category, independent of the dialog's current selection.</summary>
        private static IDictionary<string, object> CategoryPalette(IDictionary<string, object> format, string category, params string[] names)
        {
            var state = ((object[])format["FormatCategories"]).Select(VbeBridgeClient.Object)
                .Single(x => (string)x["Category"] == category);
            return ((object[])state["Palettes"]).Select(VbeBridgeClient.Object).Single(x => names.Contains((string)x["Name"]));
        }
        /// <summary>Extrait les seules valeurs proposées par le catalogue natif.</summary>
        private static string[] Choices(IDictionary<string, object> control) => ((object[])control["Choices"]).Cast<string>().ToArray();
        /// <summary>Écrit après inspection fraîche; la catégorie optionnelle appartient à cette unique mutation.</summary>
        private void Write(ExcelVbeFixture excel, string tab, string property, object value, string category, bool restoring = false)
        {
            var current = Read(excel);
            if (restoring)
            {
                var format = Format(current);
                IDictionary<string, object> control;
                if (!string.IsNullOrEmpty(category))
                {
                    control = CategoryPalette(format, category, property);
                }
                else control = Find(format, property);
                object expected = value is bool check ? (object)(check ? "On" : "Off") : value;
                if (Equals(control["Value"], expected))
                {
                    AttachEvidence(excel.ProcessId, "restoration-already-matched", new { Before = current,
                        Property = property, Category = category, Expected = expected, MutationRequested = false });
                    return;
                }
            }
            var request = new { Command = "set_vbe_option", Pane = tab, Property = property, Value = value,
                Query = category, ExpectedOptionsVersion = current["OptionsVersion"] };
            IDictionary<string, object> response = null;
            try
            {
                response = ObserveCommand(excel, request);
                var result = Data(response);
                Assert.AreEqual(true, result["ControlValueVerified"]); Assert.AreEqual(true, result["DialogClosed"]);
                var after = Read(excel);
                AttachEvidence(excel.ProcessId, restoring ? "restoration-success" : "write-success",
                    new { Before = current, Request = request, Response = response, IndependentAfterRead = after,
                        ObservedCategoryBefore = CurrentCategory(Format(current)), ObservedCategoryAfter = CurrentCategory(Format(after)),
                        MutationRetried = false });
            }
            catch (Exception failure)
            {
                // This is observation only, including after uncertain transport outcomes; never resend the mutation.
                IDictionary<string, object> after = null;
                string afterReadError = null;
                try { after = ObserveCommand(excel, new { Command = "read_vbe_options" }); }
                catch (Exception error) { afterReadError = error.ToString(); }
                try
                {
                    AttachEvidence(excel.ProcessId, restoring ? "restoration-failure" : "write-failure",
                        new { Before = current, Request = request, Response = response, Failure = failure.ToString(),
                            IndependentAfterRead = after, AfterReadError = afterReadError, MutationRetried = false });
                }
                catch (Exception error) { TestContext.WriteLine("Could not attach options diagnostic: " + error); }
                throw;
            }
        }
        /// <summary>Retains each request before dispatch and every outcome, including successful observations.</summary>
        private IDictionary<string, object> ObserveCommand(ExcelVbeFixture excel, object request)
        {
            int commandSequence = ++evidenceSequence;
            AttachEvidence(excel.ProcessId, "command-request", new { CommandSequence = commandSequence, Request = request });
            IDictionary<string, object> response = null;
            string failure = null;
            try { response = excel.Command(request); return response; }
            catch (Exception error) { failure = error.ToString(); throw; }
            finally
            {
                try
                {
                    AttachEvidence(excel.ProcessId, "command-outcome", new { CommandSequence = commandSequence,
                        Request = request, Response = response, Failure = failure, MutationRetried = false });
                }
                catch (Exception evidenceError)
                {
                    if (failure == null) throw;
                    TestContext.WriteLine("Could not retain command failure diagnostic: " + evidenceError);
                }
            }
        }
        /// <summary>Conserve les preuves hors du processus natif, sans modifier ni remplacer la baseline originale.</summary>
        private void AttachEvidence(int processId, string kind, object evidence)
        {
            string path = Path.Combine(evidenceDirectory, "options-" + kind + "-" + processId + "-" + (++evidenceSequence) + ".json");
            File.WriteAllText(path, new JavaScriptSerializer { MaxJsonLength = 10 * 1024 * 1024 }.Serialize(evidence), new UTF8Encoding(false));
            try { TestContext.AddResultFile(path); }
            catch (Exception error) { TestContext.WriteLine("Evidence retained at " + path + "; result attachment failed: " + error); }
        }
    }
}

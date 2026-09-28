using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Web.Script.Serialization;
using CodexVBE.Tests.Integration;
using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace CodexVBE.Tests.Integration.Hosts.Excel
{
    /// <summary>Qualification reproductible des choix Format via le vrai bridge Excel, avec restauration globale garantie.</summary>
    [TestClass, TestCategory("Excel")]
    public sealed class ExcelFormatOptionsTests
    {
        /// <summary>Conserve les instantanés avant mutation même si la restauration échoue.</summary>
        public TestContext TestContext { get; set; }
        /// <summary>Matrice fixée avant exécution : police, taille native ou refus, palettes, catégorie distincte, marge, garde et restauration.</summary>
        private static readonly string[] Scenarios = { "exact font", "size catalogue or honest refusal", "foreground", "background", "indicator",
            "nondefault category foreground", "margin indicator", "stale version refuses", "full options version restored" };

        /// <summary>Exerce les préférences réelles dans une instance visible possédée; aucun nom/RGB de palette n'est supposé.</summary>
        [TestMethod]
        public void NativeFormatChoicesRoundTripAndRestoreCompleteOptionsVersion()
        {
            Assert.AreEqual(9, Scenarios.Length);
            using (var excel = ExcelVbeFixture.Start())
            {
                var baseline = Read(excel); var format = Format(baseline); string tab = (string)format["Tab"];
                string baselineFile = Path.Combine(TestContext.TestRunDirectory, "options-baseline-" + excel.ProcessId + ".json");
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
                        var rejected = excel.Command(new { Command = "set_vbe_option", Pane = tab, Property = size["Name"], Value = "12", ExpectedOptionsVersion = sizeRead["OptionsVersion"] });
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
                        var control = Find(Format(Read(excel)), names);
                        string next = Choices(control).First(x => x != (string)control["Value"]);
                        restore.Add(Tuple.Create((string)control["Name"], control["Value"], (string)null));
                        Write(excel, tab, (string)control["Name"], next, null);
                        Assert.AreEqual(next, Find(Format(Read(excel)), names)["Value"]);
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
                    var stale = excel.Command(new { Command = "set_vbe_option", Pane = tab, Property = font["Name"], Value = font["Value"], ExpectedOptionsVersion = baseline["OptionsVersion"] });
                    Assert.AreEqual(false, stale["Ok"]);
                }
                catch (Exception error) { primaryFailure = error; throw; }
                finally
                {
                    try
                    {
                        foreach (var item in restore.AsEnumerable().Reverse()) Write(excel, tab, item.Item1, item.Item2, item.Item3);
                        Assert.AreEqual(baseline["OptionsVersion"], Read(excel)["OptionsVersion"], "Every category and preference must return to the complete baseline.");
                    }
                    catch (Exception restorationError)
                    {
                        if (primaryFailure != null) throw new AggregateException("The native scenario and its restoration both failed; baseline attached.", primaryFailure, restorationError);
                        throw;
                    }
                }
            }
        }

        /// <summary>Lit les quatre onglets et exige une réponse native positive.</summary>
        private static IDictionary<string, object> Read(ExcelVbeFixture excel) => Data(excel.Command("read_vbe_options"));
        /// <summary>Extrait les données d'une réponse après validation de sa réussite.</summary>
        private static IDictionary<string, object> Data(IDictionary<string, object> response)
        { Assert.IsNotNull(response); Assert.AreEqual(true, response["Ok"], Convert.ToString(response["Error"])); return VbeBridgeClient.Object(response["Data"]); }
        /// <summary>Localise l'onglet Format observé en anglais ou français.</summary>
        private static IDictionary<string, object> Format(IDictionary<string, object> data) => ((object[])data["Tabs"]).Select(VbeBridgeClient.Object)
            .Single(x => new[] { "Editor Format", "Format de l'éditeur", "Format de l’éditeur" }.Contains((string)x["Tab"]));
        /// <summary>Localise un contrôle éditable exact, en ignorant les labels de texte.</summary>
        private static IDictionary<string, object> Find(IDictionary<string, object> format, params string[] names) => ((object[])format["Controls"]).Select(VbeBridgeClient.Object)
            .Single(x => names.Contains((string)x["Name"]) && (string)x["Type"] != "ControlType.Text");
        /// <summary>Extrait les seules valeurs proposées par le catalogue natif.</summary>
        private static string[] Choices(IDictionary<string, object> control) => ((object[])control["Choices"]).Cast<string>().ToArray();
        /// <summary>Écrit après inspection fraîche; la catégorie optionnelle appartient à cette unique mutation.</summary>
        private static void Write(ExcelVbeFixture excel, string tab, string property, object value, string category)
        {
            var current = Read(excel);
            var result = Data(excel.Command(new { Command = "set_vbe_option", Pane = tab, Property = property, Value = value, Query = category, ExpectedOptionsVersion = current["OptionsVersion"] }));
            Assert.AreEqual(true, result["ControlValueVerified"]); Assert.AreEqual(true, result["DialogClosed"]);
        }
    }
}

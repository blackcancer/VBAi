using Microsoft.VisualStudio.TestTools.UnitTesting;
using System;
using System.Collections.Generic;
using System.Linq;

namespace VBAi.Tests.Integration
{
    internal sealed partial class ExcelFormatOptionsQualification
    {
        /// <summary>Prepares all editable Format values from one snapshot and applies the complete set with one native OK.</summary>
        private void MatrixBatch()
        {
            var format = Format(baseline);
            string tab = (string)format["Tab"];
            var changes = new List<VbeOptionChange>();
            var font = Find(format, "Font", "Police :");
            string alternateFont = Choices(font).FirstOrDefault(choice => choice != (string)font["Value"] &&
                new[] { "Consolas (Occidental)", "Consolas (Western)", "Consolas", "Courier New (Occidental)", "Courier New (Western)", "Courier New" }.Contains(choice));
            Assert.IsNotNull(alternateFont, "An exact observed alternate font is required for this bounded fixture.");
            changes.Add(new VbeOptionChange { Pane = tab, Property = (string)font["Name"],
                Value = alternateFont });
            var size = Find(format, "Size", "Taille :");
            var sizes = Choices(size);
            if (sizes.Length > 0)
                changes.Add(new VbeOptionChange { Pane = tab, Property = (string)size["Name"],
                    Value = sizes.First(choice => choice != (string)size["Value"]) });
            else
                Refusal("EmptySize", new { Command = "set_vbe_option", Pane = tab, Property = size["Name"],
                    Value = "12", ExpectedOptionsVersion = Version(baseline) }, baseline,
                    "The exact native choice is absent, ambiguous or unreadable.");
            // Every supported syntax category is staged while the same dialog remains
            // open. Disabled palettes remain unchanged and are preserved in readback.
            foreach (var category in ((object[])format["FormatCategories"]).Select(VbeBridgeClient.Object))
            {
                foreach (var palette in ((object[])category["Palettes"]).Select(VbeBridgeClient.Object))
                {
                    if (palette.TryGetValue("Enabled", out var enabled) && !Equals(enabled, true)) continue;
                    var choices = Choices(palette);
                    if (choices.Length < 2) continue;
                    changes.Add(new VbeOptionChange { Pane = tab, Property = (string)palette["Name"], Query = (string)category["Category"],
                        Value = choices.First(choice => choice != (string)palette["Value"]) });
                }
            }
            var margin = Find(format, "Margin Indicator Bar", "Barre des indicateurs en marge");
            Assert.IsTrue(Equals(margin["Value"], "On") || Equals(margin["Value"], "Off"));
            changes.Add(new VbeOptionChange { Pane = tab, Property = (string)margin["Name"], Value = !Equals(margin["Value"], "On") });
            var restore = changes.Select(change =>
            {
                object original = Control(format, change.Property, change.Query)["Value"];
                if (change.Value is bool) original = Equals(original, "On");
                return Tuple.Create(change.Pane, change.Property, original, change.Query);
            }).ToArray();
            Send("FormatBatch", new { Command = "set_vbe_option", OptionChanges = changes.ToArray(),
                ExpectedOptionsVersion = Version(baseline) }, true, () => ledger.AddRange(restore));
            var after = Read("FormatBatchReadback");
            foreach (var change in changes)
                Assert.AreEqual(Expected(change.Value), Control(Format(after), change.Property, change.Query)["Value"],
                    change.Property + " in " + change.Query);
            Assert.AreEqual(CurrentCategory(format), CurrentCategory(Format(after)), "Batch navigation must preserve the original viewed category.");
            Assert.AreNotEqual(Version(baseline), Version(after));
            evidence("FormatBatchVerified", new { Changes = changes, Readback = after, NativeCommits = 1, MutationRetried = false });
            Refusal("StaleVersionBatch", new { Command = "set_vbe_option", OptionChanges = changes.ToArray(),
                ExpectedOptionsVersion = Version(baseline) }, after, "VBE options changed since inspection; read them again.");
        }

        /// <summary>Restores every positively committed edit through one guarded batch and one complete independent readback.</summary>
        private void RestoreBatch()
        {
            if (ledger.Count > 0)
            {
                var current = Read("BeforeRestorationBatch");
                var changes = ledger.Where(entry => !Equals(Control(Format(current), entry.Item2, entry.Item4)["Value"], Expected(entry.Item3)))
                    .Select(entry => new VbeOptionChange { Pane = entry.Item1, Property = entry.Item2, Value = entry.Item3, Query = entry.Item4 }).ToArray();
                if (changes.Length > 0)
                    Send("RestorationBatch", new { Command = "set_vbe_option", OptionChanges = changes, ExpectedOptionsVersion = Version(current) }, true, null);
            }
            var restored = Read("CompleteRestorationReadback");
            Assert.AreEqual(Version(baseline), Version(restored));
            var json = new System.Web.Script.Serialization.JavaScriptSerializer();
            Assert.AreEqual(json.Serialize(baseline["Tabs"]), json.Serialize(restored["Tabs"]));
            evidence("BaselineRestored", new { BaselineVersion = Version(baseline), Readback = restored, GroupedRestoration = true });
        }
    }
}

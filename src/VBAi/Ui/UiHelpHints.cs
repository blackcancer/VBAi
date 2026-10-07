using System.Collections.Generic;
using System.ComponentModel;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Runtime.CompilerServices;
using System.Web.Script.Serialization;
using System.Windows.Forms;
#if VBAI_UPDATER
using UiText = VBAi.UpdateText;
#endif

namespace VBAi
{
    /// <summary>Adds usage hints to fixed and dynamically attached WinForms controls without replacing existing contextual hints.</summary>
    internal static class UiHelpHints
    {
        /// <summary>English resource keys indexed by a control name or an owner-type-qualified control name.</summary>
        private static readonly Dictionary<string, string> Rules = ReadRules();

        /// <summary>Fallback tooltip components weakly associated with controls that have no Designer component container.</summary>
        private static readonly ConditionalWeakTable<Control, ToolTip> OwnedTips = new ConditionalWeakTable<Control, ToolTip>();

        /// <summary>Reads the embedded hint catalogue once; no files, accounts or host services are consulted.</summary>
        /// <returns>Usage resource keys indexed with ordinal comparison.</returns>
        private static Dictionary<string, string> ReadRules()
        {
            using (var stream = typeof(UiHelpHints).Assembly.GetManifestResourceStream("VBAi.UiHelpHints.json"))
            using (var reader = new StreamReader(stream))
                return new JavaScriptSerializer().Deserialize<Dictionary<string, string>>(reader.ReadToEnd());
        }

        /// <summary>Finds the most specific hint using ancestor view types, then the shared control-name rule.</summary>
        /// <param name="control">Control on its owning UI thread; a null value has no hint.</param>
        /// <returns>Localized usage text, or null when the catalogue does not describe this control.</returns>
        internal static string ForControl(Control control)
        {
            if (control == null || string.IsNullOrEmpty(control.Name)) return null;
            string key;
            for (var owner = control; owner != null; owner = owner.Parent)
                if (Rules.TryGetValue(owner.GetType().Name + "." + control.Name, out key)) return UiText.Get(key);
            return Rules.TryGetValue(control.Name, out key) ? UiText.Get(key) : null;
        }

        /// <summary>Fills a menu command's missing usage hint using its stable name or untranslated caption.</summary>
        /// <param name="item">Initialized command; separators and existing contextual hints are preserved.</param>
        internal static void Apply(ToolStripItem item)
        {
            if (item == null || item is ToolStripSeparator || !string.IsNullOrEmpty(item.ToolTipText)) return;
            if (Rules.TryGetValue("menu." + item.Name, out string key) || Rules.TryGetValue("menu." + item.Text, out key))
            {
                item.ToolTipText = UiText.Get(key);
                if (string.IsNullOrEmpty(item.AccessibleDescription)) item.AccessibleDescription = item.ToolTipText;
            }
        }

        /// <summary>Fills an absent hint using a supplied or Designer-owned component; existing dynamic hint text wins.</summary>
        /// <param name="control">Live control to update on its owning thread.</param>
        /// <param name="components">Optional Designer container used to discover and own a tooltip component.</param>
        /// <param name="additional">Tooltip components already supplied by the view.</param>
        internal static void Apply(Control control, IContainer components = null, params ToolTip[] additional)
        {
            if (control == null || control.IsDisposed || LicenseManager.UsageMode == LicenseUsageMode.Designtime) return;
            var text = ForControl(control);
            if (string.IsNullOrEmpty(text)) return;
            var tips = new List<ToolTip>(additional ?? new ToolTip[0]);
            if (components != null) tips.AddRange(components.Components.OfType<ToolTip>());
            // Theme traversal also reaches controls created after InitializeComponent.
            // Reuse their nearest view's Designer tooltip instead of adding a competing popup.
            for (var owner = control; owner != null; owner = owner.Parent)
            {
                var local = new List<ToolTip>();
                foreach (var field in owner.GetType().GetFields(BindingFlags.Instance | BindingFlags.NonPublic | BindingFlags.Public))
                    if (typeof(ToolTip).IsAssignableFrom(field.FieldType) && field.GetValue(owner) is ToolTip found) local.Add(found);
                if (local.Count != 0) { tips.InsertRange(0, local); break; }
            }
            if (tips.Any(tip => !string.IsNullOrEmpty(tip.GetToolTip(control)))) return;
            var selected = tips.FirstOrDefault();
            if (selected == null)
            {
                var owner = (Control)control.FindForm() ?? control;
                selected = OwnedTips.GetValue(owner, root =>
                {
                    var value = new ToolTip { ShowAlways = true, AutoPopDelay = 12000 };
                    root.Disposed += (sender, args) => value.Dispose();
                    return value;
                });
            }
            selected.SetToolTip(control, text);
            if (string.IsNullOrEmpty(control.AccessibleDescription)) control.AccessibleDescription = text;
        }
    }
}

using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;

namespace VBAi
{
    internal static partial class VbeDebugWindows
    {
        /// <summary>Validated native target and immutable values prepared before the first staged write.</summary>
        private sealed class PlannedOptionChange
        {
            internal VbeOptionChange Change;
            internal int TabIndex;
            internal string Type;
            internal object Before, WriteValue, Expected;
        }

        /// <summary>Prepares every edit, stages them in the same Options dialog, and requests OK only after complete readback.</summary>
        /// <param name="request">Complete Options revision and a bounded collection of distinct supported edits.</param>
        /// <param name="native">Probe retaining the original owned dialog and control identities.</param>
        /// <returns>Per-edit values and explicit commit/closure evidence; restart persistence still requires independent inspection.</returns>
        internal static object SetVbeOptionBatch(Request request, IWritableOptionsProbe native)
        {
            if (request == null || request.OptionChanges == null || string.IsNullOrWhiteSpace(request.ExpectedOptionsVersion) ||
                !string.IsNullOrEmpty(request.Pane) || !string.IsNullOrEmpty(request.Property) ||
                !string.IsNullOrEmpty(request.Query) || request.Value != null ||
                request.OptionChanges.Length < 1 || request.OptionChanges.Length > 64 ||
                request.OptionChanges.Any(change => change == null || string.IsNullOrWhiteSpace(change.Pane) || string.IsNullOrWhiteSpace(change.Property)))
                throw new ArgumentException("Provide ExpectedOptionsVersion and one to sixty-four OptionChanges, without single-option fields.");
            var changes = request.OptionChanges.Select(change => new VbeOptionChange
            { Pane = change.Pane, Property = change.Property, Query = change.Query, Value = change.Value }).ToArray();
            string expectedVersion = request.ExpectedOptionsVersion;
            IntPtr dialog = IntPtr.Zero;
            for (int attempt = 0; attempt < 60 && dialog == IntPtr.Zero; attempt++) { native.Pause(50); dialog = native.Dialog(); }
            if (dialog == IntPtr.Zero) throw new InvalidOperationException("The native VBE Options dialog did not open.");
            bool commitRequested = false, writePending = false;
            object result = null;
            Exception primary = null;
            try
            {
                var before = CaptureOptionsTabs(native, dialog);
                if (!string.Equals(OptionsRevision(before), expectedVersion, StringComparison.OrdinalIgnoreCase))
                    throw new InvalidOperationException("VBE options changed since inspection; read them again.");
                var tabs = native.Tabs(dialog);
                var planned = new List<PlannedOptionChange>();
                var targets = new HashSet<string>(StringComparer.Ordinal);
                var originalCategories = new Dictionary<int, string>();
                foreach (string pane in changes.Where(change => IsColorPalette(change.Property)).Select(change => change.Pane).Distinct())
                {
                    var indices = Enumerable.Range(0, tabs.Count).Where(index => tabs[index] == pane).ToArray();
                    if (indices.Length != 1) throw new InvalidOperationException("The exact options tab is absent or ambiguous.");
                    var categories = native.Controls(dialog, indices[0]).Where(control => control.Type == "ControlType.List" && IsCodeColorList(control.Name)).ToArray();
                    if (categories.Length != 1 || !(categories[0].Value is string originalCategory))
                        throw new InvalidOperationException("The native Code Colors category is unreadable.");
                    originalCategories.Add(indices[0], originalCategory);
                }
                // Category selection changes only the pending dialog's viewed row.
                // Validate the entire batch before changing any preference value.
                foreach (var change in changes)
                {
                    var indices = Enumerable.Range(0, tabs.Count).Where(tabIndex => tabs[tabIndex] == change.Pane).ToArray();
                    if (indices.Length != 1) throw new InvalidOperationException("The exact options tab is absent or ambiguous.");
                    int index = indices[0];
                    if (IsCodeColorList(change.Property))
                        throw new ArgumentException("Batch category navigation belongs in each palette change's Query, not a separate preference edit.");
                    if (IsColorPalette(change.Property) && string.IsNullOrEmpty(change.Query))
                        change.Query = originalCategories[index];
                    SelectOptionChangeCategory(native, dialog, index, change);
                    var controls = native.Controls(dialog, index);
                    string category = change.Query;
                    if (!targets.Add(change.Pane + "\0" + (category ?? "") + "\0" + change.Property))
                        throw new ArgumentException("The batch contains a duplicate native option target.");
                    var selected = RequireOptionChangeControl(controls, change, "planning", planned.Count + 1);
                    object value = ValidateEditableOption(change.Pane, selected, change.Value);
                    planned.Add(new PlannedOptionChange { Change = change, TabIndex = index, Type = selected.Type,
                        Before = selected.Value, WriteValue = value, Expected = OptionChangeExpected(selected.Type, value) });
                }
                foreach (var plan in planned)
                {
                    SelectOptionChangeCategory(native, dialog, plan.TabIndex, plan.Change);
                    var selected = RequireOptionChangeControl(native.Controls(dialog, plan.TabIndex), plan.Change, "staging", planned.IndexOf(plan) + 1);
                    if (selected.Type != plan.Type || !Equals(selected.Value, plan.Before) ||
                        !Equals(ValidateEditableOption(plan.Change.Pane, selected, plan.WriteValue), plan.WriteValue))
                        throw new InvalidOperationException("A planned native option changed before its staged write; Cancel will be requested.");
                    writePending = true;
                    native.Write(dialog, plan.TabIndex, plan.Change.Property, plan.Type, plan.WriteValue);
                    writePending = false;
                }
                // Verify every final value after all edits, including earlier values
                // that a later font/category change could have affected.
                foreach (var plan in planned)
                {
                    SelectOptionChangeCategory(native, dialog, plan.TabIndex, plan.Change);
                    var controls = native.Controls(dialog, plan.TabIndex);
                    if (!string.IsNullOrEmpty(plan.Change.Query) && controls.Count(control => control.Type == "ControlType.List" &&
                        IsCodeColorList(control.Name) && Equals(control.Value, plan.Change.Query)) != 1)
                        throw new InvalidOperationException("The native Code Colors category changed during batch readback; Cancel will be requested.");
                    var observed = RequireOptionChangeControl(controls, plan.Change, "verification", planned.IndexOf(plan) + 1);
                    if (observed.Type != plan.Type || !Equals(observed.Value, plan.Expected))
                        throw new InvalidOperationException("The complete native Options batch did not retain its requested values; Cancel will be requested.");
                }
                foreach (var original in originalCategories)
                {
                    ((IFormatCategoriesOptionsProbe)native).SelectFormatCategory(dialog, original.Key, original.Value);
                    if (native.Controls(dialog, original.Key).Count(control => control.Type == "ControlType.List" &&
                        IsCodeColorList(control.Name) && Equals(control.Value, original.Value)) != 1)
                        throw new InvalidOperationException("The original viewed Code Colors category was not restored; Cancel will be requested.");
                }
                commitRequested = true;
                native.Accept(dialog);
                bool closed = WaitForOptionsClosure(native, dialog, 40);
                result = new { Changes = planned.Select(plan => new { plan.Change.Pane, plan.Change.Property,
                    Category = plan.Change.Query, plan.Before, After = plan.Expected }).ToArray(),
                    CommitRequested = true, DialogClosed = closed, ControlValueVerified = true, PersistenceVerified = false,
                    NextRead = "read_vbe_options", Limit = "All edits were staged before one OK. Reopen Options once to verify the complete committed state; do not retry automatically." };
            }
            catch (Exception error) { primary = error; }
            Exception cleanup = null;
            if (!commitRequested && !writePending)
                try { CancelOwnedOptionsDialog(native, dialog); } catch (Exception error) { cleanup = error; }
            if (primary != null && cleanup != null)
                throw new AggregateException("VBE Options batch and captured-dialog cancellation both failed. Do not retry automatically.", primary, cleanup);
            if (primary != null) System.Runtime.ExceptionServices.ExceptionDispatchInfo.Capture(primary).Throw();
            if (cleanup != null) System.Runtime.ExceptionServices.ExceptionDispatchInfo.Capture(cleanup).Throw();
            return result;
        }

        /// <summary>Selects an exact category only for supported palette edits in the original dialog.</summary>
        private static void SelectOptionChangeCategory(IWritableOptionsProbe native, IntPtr dialog, int index, VbeOptionChange change)
        {
            if (string.IsNullOrEmpty(change.Query)) return;
            if (!IsColorPalette(change.Property) || !(native is IFormatCategoriesOptionsProbe format))
                throw new InvalidOperationException("Query is supported only for an exact native Code Colors palette category.");
            format.SelectFormatCategory(dialog, index, change.Query);
        }

        /// <summary>Requires one visible, enabled, supported and readable exact target.</summary>
        private static OptionsControl RequireOptionChangeControl(IList<OptionsControl> controls, VbeOptionChange change, string phase, int index)
        {
            var matches = controls.Where(control => control.Visible && control.Enabled && control.Name == change.Property &&
                new[] { "ControlType.CheckBox", "ControlType.RadioButton", "ControlType.Edit", "ControlType.ComboBox", "ControlType.List" }.Contains(control.Type)).ToArray();
            if (matches.Length != 1 || !string.IsNullOrEmpty(matches[0].Error))
            {
                string candidates = string.Join("; ", controls.Where(control => control.Name == change.Property)
                    .Select(control => "Type=" + control.Type + ",Visible=" + control.Visible +
                        ",Enabled=" + control.Enabled + ",Error=" + (control.Error ?? "")));
                throw new InvalidOperationException("The exact option is absent, ambiguous or unreadable. Phase='" + phase +
                    "'; Index=" + index + "; Tab='" + change.Pane + "'; Property='" + change.Property +
                    "'; Category='" + (change.Query ?? "") + "'; Candidates=[" + candidates + "].");
            }
            return matches[0];
        }

        /// <summary>Converts a validated write into the exact native readback representation.</summary>
        private static object OptionChangeExpected(string type, object value) => type == "ControlType.CheckBox"
            ? (object)((bool)value ? "On" : "Off") : type == "ControlType.RadioButton" ? (object)true : Convert.ToString(value, CultureInfo.InvariantCulture);
    }
}

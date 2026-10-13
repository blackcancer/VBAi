// SPDX-License-Identifier: MPL-2.0
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Drawing;
using System.Globalization;
using System.Linq;
using System.Windows.Forms;
using Blackcancer.ChatUI.Presentation;

namespace Blackcancer.ChatUI.Views
{
    /// <summary>Shared reading layout; the host supplies choices and owns their meaning and persistence.</summary>
    public partial class ChatReadingOptionsForm : Form
    {
        private Font ownedFont;
        private Icon ownedIcon;

        public ChatReadingOptionsForm()
        {
            InitializeComponent();
            modePicker.SelectedIndexChanged += SelectionChanged;
            filterPicker.SelectedIndexChanged += SelectionChanged;
            SelectionChanged(this, EventArgs.Empty);
        }

        [Browsable(false), DesignerSerializationVisibility(DesignerSerializationVisibility.Hidden)]
        public string SelectedModeId => (modePicker.SelectedItem as ChatReadingChoice)?.Id;
        [Browsable(false), DesignerSerializationVisibility(DesignerSerializationVisibility.Hidden)]
        public string SelectedFilterId => (filterPicker.SelectedItem as ChatReadingChoice)?.Id;
        [Browsable(false), DesignerSerializationVisibility(DesignerSerializationVisibility.Hidden)]
        public string ContextQuery => contextQuery.Text;

        public void BindChoices(IEnumerable<ChatReadingChoice> modes, IEnumerable<ChatReadingChoice> filters,
            string modeId, string filterId, string query, ChatServices services = null,
            SharedChatStyle style = null, CultureInfo culture = null)
        {
            ChatReadingChoice[] modeChoices = ValidateChoices(modes, modeId, nameof(modes));
            ChatReadingChoice[] filterChoices = ValidateChoices(filters, filterId, nameof(filters));
            query = query ?? string.Empty;
            if (query.Length > contextQuery.MaxLength || query.Any(char.IsControl))
                throw new ArgumentException("The reading query must contain at most 256 printable characters.", nameof(query));
            services = services ?? new ChatServices();
            var resources = new ComponentResourceManager(typeof(ChatReadingOptionsForm));
            culture = culture ?? CultureInfo.CurrentUICulture;
            Text = services.Text(resources.GetString("$this.Text", culture) ?? "Reading options");
            foreach (Control control in new Control[] { modeLabel, filterLabel, queryLabel, explanation, saveButton, cancelButton })
                control.Text = services.Text(resources.GetString(control.Name + ".Text", culture) ?? control.Text);
            modePicker.BeginUpdate(); filterPicker.BeginUpdate();
            try
            {
                modePicker.Items.Clear(); filterPicker.Items.Clear();
                modePicker.Items.AddRange(modeChoices); filterPicker.Items.AddRange(filterChoices);
                modePicker.SelectedItem = modeChoices.Single(choice => choice.Id == modeId);
                filterPicker.SelectedItem = filterChoices.Single(choice => choice.Id == filterId);
                contextQuery.Text = query;
            }
            finally { modePicker.EndUpdate(); filterPicker.EndUpdate(); }
            if (style != null)
            {
                ApplyPalette(this, style);
                var next = (Font)(style.BodyFont ?? SystemFonts.MessageBoxFont).Clone();
                if (ownedFont != null && ownedFont.Equals(next)) next.Dispose();
                else { Font previous = ownedFont; Font = ownedFont = next; previous?.Dispose(); }
                RightToLeft = style.RightToLeft ? RightToLeft.Yes : RightToLeft.No;
                RightToLeftLayout = style.RightToLeft;
            }
            Icon icon = services.GetWindowIcon?.Invoke();
            if (icon != null)
            {
                var next = (Icon)icon.Clone();
                Icon previous = ownedIcon; Icon = ownedIcon = next; previous?.Dispose();
            }
        }

        /// <summary>Domain wording is supplied by the adapter, rather than embedded in the shared module.</summary>
        public void SetReadingLabels(string title, string mode, string filter, string query, string note)
        {
            Text = title ?? Text; modeLabel.Text = mode ?? modeLabel.Text;
            filterLabel.Text = filter ?? filterLabel.Text; queryLabel.Text = query ?? queryLabel.Text;
            explanation.Text = note ?? explanation.Text;
        }

        private static ChatReadingChoice[] ValidateChoices(IEnumerable<ChatReadingChoice> choices, string selected, string argument)
        {
            var values = (choices ?? throw new ArgumentNullException(argument)).ToArray();
            if (values.Length == 0 || values.Any(choice => choice == null)
                || values.Select(choice => choice.Id).Distinct(StringComparer.Ordinal).Count() != values.Length
                || !values.Any(choice => string.Equals(choice.Id, selected, StringComparison.Ordinal)))
                throw new ArgumentException("Reading choices need unique identifiers and an existing selection.", argument);
            return values;
        }
        private void SelectionChanged(object sender, EventArgs args)
            => saveButton.Enabled = SelectedModeId != null && SelectedFilterId != null;
        private static void ApplyPalette(Control control, SharedChatStyle style)
        {
            control.BackColor = style.Background; control.ForeColor = style.Foreground;
            foreach (Control child in control.Controls) ApplyPalette(child, style);
        }
        private void DisposeOwnedAppearance()
        {
            ownedFont?.Dispose(); ownedFont = null;
            ownedIcon?.Dispose(); ownedIcon = null;
        }
    }
}

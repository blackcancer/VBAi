// SPDX-License-Identifier: MPL-2.0
using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.ComponentModel;
using System.Drawing;
using System.Globalization;
using System.Linq;
using System.Windows.Forms;
using Blackcancer.ChatUI.Presentation;

namespace Blackcancer.ChatUI.Views
{
    /// <summary>Modeless picker of detached choices. All source reads and attachment decisions belong to its owner.</summary>
    public partial class ChatReferencePickerForm : Form
    {
        private IReadOnlyList<ChatReferencePickerChoice> choices = new ReadOnlyCollection<ChatReferencePickerChoice>(new ChatReferencePickerChoice[0]);
        private readonly Dictionary<string, ListViewItem> rows = new Dictionary<string, ListViewItem>(StringComparer.Ordinal);
        private ChatServices services = new ChatServices();
        private Action<string> searchAction, reportError = value => System.Diagnostics.Debug.WriteLine(value);
        private Action<ChatReferencePickerChoice> inspectAction, attachAction;
        private CultureInfo resourceCulture;
        private bool busy, dispatching, bindingQuery, queryDirty;
        private string lastSearchedQuery;
        private long requestVersion, boundVersion = -1;
        private Font ownedFont;
        private Icon ownedIcon;

        public ChatReferencePickerForm()
        {
            InitializeComponent();
            search.Click += (sender, args) => RequestSearch();
            searchText.TextChanged += QueryChanged;
            searchText.KeyDown += (sender, args) =>
            {
                if (args.KeyCode != Keys.Enter) return;
                args.Handled = args.SuppressKeyPress = true; RequestSearch();
            };
            debounceTimer.Tick += Debounce_Tick;
            choiceList.SelectedIndexChanged += (sender, args) => UpdateActions();
            choiceList.ItemActivate += (sender, args) => RequestChoice(inspectAction);
            inspect.Click += (sender, args) => RequestChoice(inspectAction);
            attach.Click += (sender, args) => RequestChoice(attachAction);
            close.Click += (sender, args) => Close();
        }

        [Browsable(false), DesignerSerializationVisibility(DesignerSerializationVisibility.Hidden)]
        public string Query { get => searchText.Text; set => SetQuery(value); }
        [Browsable(false), DesignerSerializationVisibility(DesignerSerializationVisibility.Hidden)]
        public bool PreserveComposerFocus { get; set; }
        [Browsable(false), DesignerSerializationVisibility(DesignerSerializationVisibility.Hidden)]
        public TextBox SearchBox => searchText;
        [Browsable(false), DesignerSerializationVisibility(DesignerSerializationVisibility.Hidden)]
        public IReadOnlyList<ChatReferencePickerChoice> Choices => choices;
        [Browsable(false), DesignerSerializationVisibility(DesignerSerializationVisibility.Hidden)]
        public ChatReferencePickerChoice SelectedChoice => choiceList.SelectedItems.Count == 1
            ? choiceList.SelectedItems[0].Tag as ChatReferencePickerChoice : null;
        [Browsable(false), DesignerSerializationVisibility(DesignerSerializationVisibility.Hidden)]
        public long RequestVersion => requestVersion;
        [Browsable(false), DesignerSerializationVisibility(DesignerSerializationVisibility.Hidden)]
        public long BoundRequestVersion => boundVersion;
        [Browsable(false), DesignerSerializationVisibility(DesignerSerializationVisibility.Hidden)]
        public bool IsBusy => busy;
        [Browsable(false), DesignerSerializationVisibility(DesignerSerializationVisibility.Hidden)]
        public string Status => statusNote.Text;

        public void Configure(Action<string> onSearch = null, Action<ChatReferencePickerChoice> onInspect = null,
            Action<ChatReferencePickerChoice> onAttach = null, ChatServices services = null, SharedChatStyle style = null,
            Action<string> onError = null, CultureInfo culture = null, string query = null)
        {
            if (IsDisposed) throw new ObjectDisposedException(nameof(ChatReferencePickerForm));
            this.services = services ?? new ChatServices();
            searchAction = onSearch; inspectAction = onInspect; attachAction = onAttach;
            reportError = onError ?? (message => System.Diagnostics.Debug.WriteLine(message));
            resourceCulture = culture ?? CultureInfo.CurrentUICulture;
            Text = Localized("$this.Text", "Choose references");
            foreach (Control control in new Control[] { queryLabel, search, inspect, attach, close, explanation })
                control.Text = Localized(control.Name + ".Text", control.Text);
            labelColumn.Text = Localized("labelColumn.Text", "Reference");
            kindColumn.Text = Localized("kindColumn.Text", "Kind");
            detailColumn.Text = Localized("detailColumn.Text", "Details");
            if (query != null) SetQuery(query);
            ApplyAppearance(style ?? new SharedChatStyle());
            Icon icon = this.services.GetWindowIcon?.Invoke();
            if (icon != null) UseBorrowedIcon(icon);
            UpdateActions();
        }

        /// <summary>Programmatic initialization never starts a search unless explicitly requested.</summary>
        public void SetQuery(string query, bool requestSearch = false)
        {
            bool changed = searchText.Text != (query ?? string.Empty);
            debounceTimer.Stop(); bindingQuery = true;
            try { searchText.Text = query ?? string.Empty; }
            finally { bindingQuery = false; }
            lastSearchedQuery = null;
            if (changed) { busy = true; queryDirty = true; UpdateActions(); }
            if (requestSearch) RequestSearch();
        }

        /// <summary>Ignore late results before enumerating them. Same IDs retain their selection and rows.</summary>
        public bool BindChoices(long requestVersion, IEnumerable<ChatReferencePickerChoice> values, string status = null)
        {
            ValidateVersion(requestVersion);
            if (IsDisposed || requestVersion < this.requestVersion || (requestVersion == this.requestVersion && queryDirty)) return false;
            var next = (values ?? throw new ArgumentNullException(nameof(values))).ToArray();
            if (next.Any(value => value == null) || next.Select(value => value.Id).Distinct(StringComparer.Ordinal).Count() != next.Length)
                throw new ArgumentException("Choices require unique non-null identifiers.", nameof(values));
            string selectedId = SelectedChoice?.Id, topId = choiceList.TopItem?.Name;
            this.requestVersion = boundVersion = requestVersion;
            queryDirty = false;
            var nextIds = new HashSet<string>(next.Select(value => value.Id), StringComparer.Ordinal);
            choiceList.BeginUpdate();
            try
            {
                foreach (string id in rows.Keys.Where(id => !nextIds.Contains(id)).ToArray())
                { choiceList.Items.Remove(rows[id]); rows.Remove(id); }
                for (int index = 0; index < next.Length; index++)
                {
                    ChatReferencePickerChoice value = next[index]; ListViewItem row;
                    if (!rows.TryGetValue(value.Id, out row))
                    {
                        row = new ListViewItem(new[] { value.Label, value.Kind, value.Detail }) { Name = value.Id };
                        rows.Add(value.Id, row);
                    }
                    else
                    {
                        if (row.Text != value.Label) row.Text = value.Label;
                        if (row.SubItems[1].Text != value.Kind) row.SubItems[1].Text = value.Kind;
                        if (row.SubItems[2].Text != value.Detail) row.SubItems[2].Text = value.Detail;
                    }
                    row.Tag = value;
                    if (row.Index != index)
                    { if (row.ListView != null) choiceList.Items.Remove(row); choiceList.Items.Insert(index, row); }
                    row.Selected = string.Equals(value.Id, selectedId, StringComparison.Ordinal);
                }
                choices = new ReadOnlyCollection<ChatReferencePickerChoice>(next);
                if (topId != null && rows.ContainsKey(topId) && choiceList.IsHandleCreated) choiceList.TopItem = rows[topId];
            }
            finally { choiceList.EndUpdate(); }
            busy = false; SetStatus(status); UpdateActions(); return true;
        }

        /// <summary>An announced newer version immediately makes old visible choices non-actionable.</summary>
        public bool SetBusy(bool isBusy, string status = null, long? requestVersion = null)
        {
            if (requestVersion.HasValue)
            {
                ValidateVersion(requestVersion.Value);
                if (requestVersion.Value < this.requestVersion) return false;
                this.requestVersion = requestVersion.Value;
                if (isBusy) queryDirty = false;
            }
            if (IsDisposed) return false;
            busy = isBusy; if (status != null) SetStatus(status); UpdateActions(); return true;
        }

        public void SetStatus(string status)
        {
            if (!IsDisposed) statusNote.Text = status ?? string.Empty;
        }

        public void RequestSearch()
        {
            debounceTimer.Stop(); DispatchSearch(true);
        }

        private void QueryChanged(object sender, EventArgs args)
        {
            if (bindingQuery || IsDisposed) return;
            // The visible choices belong to the previous query until the owner binds new results.
            busy = true; queryDirty = true; UpdateActions(); debounceTimer.Stop(); debounceTimer.Start();
        }
        private void Debounce_Tick(object sender, EventArgs args)
        { debounceTimer.Stop(); DispatchSearch(false); }
        private void DispatchSearch(bool force)
        {
            if (IsDisposed || dispatching || searchAction == null || (!force && lastSearchedQuery == Query)) return;
            string query = Query; lastSearchedQuery = query; queryDirty = false; busy = true; dispatching = true; UpdateActions();
            try { searchAction(query); }
            catch (Exception error) { reportError(error.Message); }
            finally { dispatching = false; if (!IsDisposed) UpdateActions(); }
        }
        private void RequestChoice(Action<ChatReferencePickerChoice> action)
        {
            ChatReferencePickerChoice selected = SelectedChoice;
            if (IsDisposed || busy || dispatching || boundVersion != requestVersion || selected == null || action == null) return;
            busy = true; dispatching = true; UpdateActions();
            try { action(selected); }
            catch (Exception error) { reportError(error.Message); }
            finally { dispatching = false; if (!IsDisposed) UpdateActions(); }
        }
        private void UpdateActions()
        {
            bool available = !busy && !dispatching && boundVersion == requestVersion && SelectedChoice != null;
            inspect.Enabled = available && inspectAction != null;
            attach.Enabled = available && attachAction != null;
            search.Enabled = !dispatching && searchAction != null;
        }
        private static void ValidateVersion(long version)
        { if (version < 0) throw new ArgumentOutOfRangeException(nameof(version)); }
        private string Localized(string key, string fallback)
            => services.Text(new ComponentResourceManager(typeof(ChatReferencePickerForm)).GetString(key, resourceCulture ?? CultureInfo.CurrentUICulture) ?? fallback);
        private void ApplyAppearance(SharedChatStyle style)
        {
            ApplyPalette(this, style);
            var next = (Font)(style.BodyFont ?? SystemFonts.MessageBoxFont).Clone();
            if (ownedFont != null && ownedFont.Equals(next)) next.Dispose();
            else { Font previous = ownedFont; Font = ownedFont = next; previous?.Dispose(); }
            RightToLeft = style.RightToLeft ? RightToLeft.Yes : RightToLeft.No; RightToLeftLayout = style.RightToLeft;
            choiceList.RightToLeftLayout = style.RightToLeft;
        }
        private static void ApplyPalette(Control control, SharedChatStyle style)
        {
            control.BackColor = style.Background; control.ForeColor = style.Foreground;
            foreach (Control child in control.Controls) ApplyPalette(child, style);
        }
        protected override void OnShown(EventArgs args)
        { if (ownedIcon == null && Owner?.Icon != null) UseBorrowedIcon(Owner.Icon); base.OnShown(args); }
        protected override bool ShowWithoutActivation => PreserveComposerFocus;
        private void UseBorrowedIcon(Icon icon)
        { var next = (Icon)icon.Clone(); Icon previous = ownedIcon; Icon = ownedIcon = next; previous?.Dispose(); }
        private void DisposeOwnedAppearance()
        { ownedFont?.Dispose(); ownedFont = null; ownedIcon?.Dispose(); ownedIcon = null; }
    }
}

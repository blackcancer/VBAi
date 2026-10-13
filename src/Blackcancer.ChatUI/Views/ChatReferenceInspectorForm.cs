// SPDX-License-Identifier: MPL-2.0
using System;
using System.ComponentModel;
using System.Drawing;
using System.Globalization;
using System.Linq;
using System.Windows.Forms;
using Blackcancer.ChatUI.Presentation;

namespace Blackcancer.ChatUI.Views
{
    /// <summary>Modeless reader for an immutable capture. Navigation is an explicit optional host callback.</summary>
    public partial class ChatReferenceInspectorForm : Form
    {
        private ChatReferenceSnapshot snapshot;
        private ChatServices services = new ChatServices();
        private Action<ChatReferenceSnapshot> navigateAction;
        private Action compareAction;
        private bool comparisonEnabled, comparing;
        private Action<string> reportError = value => System.Diagnostics.Debug.WriteLine(value);
        private CultureInfo resourceCulture;
        private ChatReferenceNavigationState navigationState;
        private string navigationStatus = string.Empty;
        private Font ownedFont, ownedTextFont;
        private Icon ownedIcon;

        public ChatReferenceInspectorForm()
        {
            InitializeComponent();
            copy.Click += Copy_Click;
            navigate.Click += Navigate_Click;
            compare.Click += Compare_Click;
            close.Click += (sender, args) => Close();
        }

        [Browsable(false), DesignerSerializationVisibility(DesignerSerializationVisibility.Hidden)]
        public ChatReferenceSnapshot Snapshot => snapshot;
        [Browsable(false), DesignerSerializationVisibility(DesignerSerializationVisibility.Hidden)]
        public string ContentText => snapshot?.ExactText ?? string.Empty;
        [Browsable(false), DesignerSerializationVisibility(DesignerSerializationVisibility.Hidden)]
        public TextBox ContentBox => content;
        [Browsable(false), DesignerSerializationVisibility(DesignerSerializationVisibility.Hidden)]
        public TextBox MetadataBox => metadataText;
        [Browsable(false), DesignerSerializationVisibility(DesignerSerializationVisibility.Hidden)]
        public ChatReferenceNavigationState NavigationState => navigationState;

        public void ShowSnapshot(ChatReferenceSnapshot value, ChatServices callbacks = null, SharedChatStyle style = null,
            Action<ChatReferenceSnapshot> onNavigate = null, Action<string> onError = null, CultureInfo culture = null)
        {
            if (value == null) throw new ArgumentNullException(nameof(value));
            if (!ReferenceEquals(snapshot, value)) ConfigureComparison(null, false);
            snapshot = value;
            services = callbacks ?? new ChatServices(); navigateAction = onNavigate;
            reportError = onError ?? (message => System.Diagnostics.Debug.WriteLine(message));
            resourceCulture = culture ?? CultureInfo.CurrentUICulture;
            style = style ?? new SharedChatStyle();
            Text = Localized("$this.Text", "Captured reference") + (snapshot.Label.Length == 0 ? "" : " — " + snapshot.Label);
            foreach (Control control in new Control[] { copy, navigate, compare, close, labelLabel, idLabel, originLabel, metadataLabel, contentLabel })
                control.Text = Localized(control.Name + ".Text", control.Text);
            labelText.Text = snapshot.Label; idText.Text = snapshot.Id; originText.Text = snapshot.Origin;
            metadataText.Text = string.Join(Environment.NewLine, snapshot.Metadata.Select(entry => entry.Key + ": " + entry.Value));
            content.Text = snapshot.ExactText;
            note.Text = snapshot.Truncated
                ? Localized("truncatedNote", "The captured text is truncated. Copy preserves only this captured fragment.")
                : Localized("completeNote", "Captured text only. This window does not reread or refresh the source.");
            ApplyPalette(this, style);
            Font nextFont = (Font)(style.BodyFont ?? SystemFonts.MessageBoxFont).Clone();
            // WinForms keeps the previous instance when Font.Equals(newFont).
            // Reuse our owned font in that case rather than disposing a live one.
            if (ownedFont != null && ownedFont.Equals(nextFont)) nextFont.Dispose();
            else { Font previousFont = ownedFont; Font = ownedFont = nextFont; previousFont?.Dispose(); }
            Font previousTextFont = ownedTextFont;
            Font nextTextFont = style.CodeFont == null ? null : (Font)style.CodeFont.Clone();
            if (ownedTextFont != null && nextTextFont != null && ownedTextFont.Equals(nextTextFont))
            { nextTextFont.Dispose(); previousTextFont = null; }
            else ownedTextFont = nextTextFont;
            if (ownedTextFont == null) { content.ResetFont(); metadataText.ResetFont(); }
            else content.Font = metadataText.Font = ownedTextFont;
            previousTextFont?.Dispose();
            RightToLeft = style.RightToLeft ? RightToLeft.Yes : RightToLeft.No; RightToLeftLayout = style.RightToLeft;
            Icon icon = services.GetWindowIcon?.Invoke();
            if (icon != null) UseBorrowedIcon(icon);
            copy.Enabled = snapshot.ExactText.Length > 0;
            SetNavigationState(snapshot.NavigationState, snapshot.Status);
        }

        /// <summary>Optional explicit host comparison request; configuring it never rereads the source.</summary>
        public void ConfigureComparison(Action callback, bool enabled)
        {
            if (IsDisposed) return;
            compareAction = callback; comparisonEnabled = enabled;
            compare.Visible = callback != null;
            compare.Enabled = callback != null && enabled && !comparing;
        }

        private void Compare_Click(object sender, EventArgs args)
        {
            if (snapshot == null || compareAction == null || !comparisonEnabled || comparing) return;
            comparing = true; compare.Enabled = false;
            try { compareAction(); }
            catch (Exception error) { reportError(error.Message); }
            finally { comparing = false; if (!IsDisposed) compare.Enabled = compareAction != null && comparisonEnabled; }
        }

        /// <summary>Explicit host invalidation changes navigation, never the captured content or its selection.</summary>
        public void SetNavigationState(ChatReferenceNavigationState state, string status = null)
        {
            if (!Enum.IsDefined(typeof(ChatReferenceNavigationState), state)) throw new ArgumentOutOfRangeException(nameof(state));
            navigationState = state; navigationStatus = status ?? string.Empty;
            string label = state == ChatReferenceNavigationState.Stale ? Localized("navigationStale", "Navigation: stale reference")
                : state == ChatReferenceNavigationState.Available && navigateAction != null ? Localized("navigationAvailable", "Navigation: available on request")
                : Localized("navigationUnavailable", "Navigation: unavailable");
            navigationNote.Text = label + (navigationStatus.Length == 0 ? "" : " · " + navigationStatus);
            navigate.Enabled = snapshot != null && state == ChatReferenceNavigationState.Available && navigateAction != null;
        }

        protected override void OnShown(EventArgs args)
        {
            if (ownedIcon == null && Owner?.Icon != null) UseBorrowedIcon(Owner.Icon);
            base.OnShown(args);
        }
        private string Localized(string key, string fallback)
        {
            var resources = new ComponentResourceManager(typeof(ChatReferenceInspectorForm));
            return services.Text(resources.GetString(key, resourceCulture ?? CultureInfo.CurrentUICulture) ?? fallback);
        }
        private void Copy_Click(object sender, EventArgs args)
        {
            if (snapshot == null) return;
            try { services.CopyText?.Invoke(snapshot.ExactText); } catch (Exception error) { reportError(error.Message); }
        }
        private void Navigate_Click(object sender, EventArgs args)
        {
            if (snapshot == null || navigationState != ChatReferenceNavigationState.Available || navigateAction == null) return;
            try { navigateAction(snapshot); } catch (Exception error) { reportError(error.Message); }
        }
        private static void ApplyPalette(Control control, SharedChatStyle style)
        {
            control.BackColor = style.Background; control.ForeColor = style.Foreground;
            foreach (Control child in control.Controls) ApplyPalette(child, style);
        }
        private void UseBorrowedIcon(Icon icon)
        {
            var next = (Icon)icon.Clone(); Icon previous = ownedIcon; Icon = ownedIcon = next; previous?.Dispose();
        }
        private void DisposeOwnedAppearance()
        {
            ownedFont?.Dispose(); ownedFont = null;
            ownedTextFont?.Dispose(); ownedTextFont = null;
            ownedIcon?.Dispose(); ownedIcon = null;
        }
    }
}

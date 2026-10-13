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
    /// <summary>Detached comparison of two captures. Only an explicit host action may replace a draft.</summary>
    public partial class ChatReferenceComparisonForm : Form
    {
        private ChatReferenceSnapshot oldSnapshot, newSnapshot;
        private ChatServices services = new ChatServices();
        private Action<ChatReferenceSnapshot> acceptAction;
        private Action<string> reportError = message => System.Diagnostics.Debug.WriteLine(message);
        private CultureInfo resourceCulture;
        private bool acceptanceEnabled, dispatching;
        private string acceptanceStatus = string.Empty;
        private Font ownedFont, ownedTextFont;
        private Icon ownedIcon;

        public ChatReferenceComparisonForm()
        {
            InitializeComponent();
            copyOld.Click += (sender, args) => Copy(oldSnapshot);
            copyNew.Click += (sender, args) => Copy(newSnapshot);
            accept.Click += Accept_Click;
            cancel.Click += (sender, args) => Close();
        }

        [Browsable(false), DesignerSerializationVisibility(DesignerSerializationVisibility.Hidden)]
        public ChatReferenceSnapshot OldSnapshot => oldSnapshot;
        [Browsable(false), DesignerSerializationVisibility(DesignerSerializationVisibility.Hidden)]
        public ChatReferenceSnapshot NewSnapshot => newSnapshot;
        [Browsable(false), DesignerSerializationVisibility(DesignerSerializationVisibility.Hidden)]
        public TextBox OldContentBox => oldContent;
        [Browsable(false), DesignerSerializationVisibility(DesignerSerializationVisibility.Hidden)]
        public TextBox NewContentBox => newContent;
        [Browsable(false), DesignerSerializationVisibility(DesignerSerializationVisibility.Hidden)]
        public bool CanAccept => !IsDisposed && oldSnapshot != null && newSnapshot != null && acceptanceEnabled && !dispatching && acceptAction != null;
        [Browsable(false), DesignerSerializationVisibility(DesignerSerializationVisibility.Hidden)]
        public string Status => statusNote.Text;

        /// <summary>Binding does not read the source or accept the candidate. Call Show for a modeless window.</summary>
        public void ShowComparison(ChatReferenceSnapshot oldSnapshot, ChatReferenceSnapshot newSnapshot,
            ChatServices services = null, SharedChatStyle style = null, Action<ChatReferenceSnapshot> onAccept = null,
            Action<string> onError = null, CultureInfo culture = null, bool canAccept = true, string status = null)
        {
            if (oldSnapshot == null) throw new ArgumentNullException(nameof(oldSnapshot));
            if (newSnapshot == null) throw new ArgumentNullException(nameof(newSnapshot));
            if (dispatching) throw new InvalidOperationException("A comparison cannot be rebound while its acceptance callback is running.");
            this.oldSnapshot = oldSnapshot; this.newSnapshot = newSnapshot;
            this.services = services ?? new ChatServices(); acceptAction = onAccept;
            reportError = onError ?? (message => System.Diagnostics.Debug.WriteLine(message));
            resourceCulture = culture ?? CultureInfo.CurrentUICulture;
            Text = Localized("$this.Text", "Compare captured references");
            foreach (Control control in new Control[] { copyOld, copyNew, accept, cancel, oldTitle, newTitle,
                oldMetadataLabel, newMetadataLabel, oldContentLabel, newContentLabel, note })
                control.Text = Localized(control.Name + ".Text", control.Text);
            BindCapture(oldSnapshot, oldIdentity, oldMetadata, oldContent, oldCaptureNote);
            BindCapture(newSnapshot, newIdentity, newMetadata, newContent, newCaptureNote);
            copyOld.Enabled = oldSnapshot.ExactText.Length != 0;
            copyNew.Enabled = newSnapshot.ExactText.Length != 0;
            ApplyAppearance(style ?? new SharedChatStyle());
            Icon icon = this.services.GetWindowIcon?.Invoke();
            if (icon != null) UseBorrowedIcon(icon);
            SetAcceptanceState(canAccept, status);
        }

        /// <summary>Changes only acceptance and its status; captured text, selection and scroll remain intact.</summary>
        public void SetAcceptanceState(bool canAccept, string status = null)
        {
            if (IsDisposed) return;
            acceptanceEnabled = canAccept; acceptanceStatus = status ?? string.Empty;
            statusNote.Text = acceptanceStatus.Length == 0
                ? Localized(CanAccept ? "acceptanceAvailable" : "acceptanceUnavailable",
                    CanAccept ? "Apply to the draft only after reviewing both captures." : "Applying this capture to the draft is unavailable.")
                : acceptanceStatus;
            accept.Enabled = CanAccept;
        }

        public void ConfigureAcceptance(bool canAccept, string status = null) => SetAcceptanceState(canAccept, status);

        /// <summary>Neutral host feedback labels; this method does not change or assert any workflow state.</summary>
        public static string GetStatusText(ChatReferenceComparisonStatus state, CultureInfo culture = null, ChatServices services = null)
        {
            if (!Enum.IsDefined(typeof(ChatReferenceComparisonStatus), state)) throw new ArgumentOutOfRangeException(nameof(state));
            var resources = new ComponentResourceManager(typeof(ChatReferenceComparisonForm));
            string value = resources.GetString(state.ToString(), culture ?? CultureInfo.CurrentUICulture) ?? state.ToString();
            return (services ?? new ChatServices()).Text(value);
        }

        private void BindCapture(ChatReferenceSnapshot value, TextBox identity, TextBox metadata, TextBox content, Label captureNote)
        {
            identity.Text = value.Label + Environment.NewLine + Localized("identityId", "ID") + ": " + value.Id
                + Environment.NewLine + Localized("identityOrigin", "Origin") + ": " + value.Origin;
            metadata.Text = string.Join(Environment.NewLine, value.Metadata.Select(entry => entry.Key + ": " + entry.Value));
            content.Text = value.ExactText;
            captureNote.Text = value.Truncated
                ? Localized("truncatedCapture", "Truncated capture: only the captured fragment is available.")
                : Localized("completeCapture", "Captured text; no automatic reread.");
        }

        private void Accept_Click(object sender, EventArgs args)
        {
            if (!CanAccept) return;
            // Async host callbacks must commit before closing. Never report success or close here.
            acceptanceEnabled = false; dispatching = true; accept.Enabled = false;
            statusNote.Text = Localized("acceptanceRequested", "Application requested. Waiting for the host to validate the draft.");
            try { acceptAction(newSnapshot); }
            catch (Exception error) { reportError(error.Message); }
            finally { dispatching = false; if (!IsDisposed) accept.Enabled = CanAccept; }
        }

        private void Copy(ChatReferenceSnapshot value)
        {
            if (value == null) return;
            try { services.CopyText?.Invoke(value.ExactText); }
            catch (Exception error) { reportError(error.Message); }
        }

        private string Localized(string key, string fallback)
        {
            var resources = new ComponentResourceManager(typeof(ChatReferenceComparisonForm));
            return services.Text(resources.GetString(key, resourceCulture ?? CultureInfo.CurrentUICulture) ?? fallback);
        }

        private void ApplyAppearance(SharedChatStyle style)
        {
            ApplyPalette(this, style);
            var nextFont = (Font)(style.BodyFont ?? SystemFonts.MessageBoxFont).Clone();
            if (ownedFont != null && ownedFont.Equals(nextFont)) nextFont.Dispose();
            else { Font previous = ownedFont; Font = ownedFont = nextFont; previous?.Dispose(); }
            Font previousText = ownedTextFont;
            Font nextText = style.CodeFont == null ? null : (Font)style.CodeFont.Clone();
            if (ownedTextFont != null && nextText != null && ownedTextFont.Equals(nextText))
            { nextText.Dispose(); previousText = null; }
            else ownedTextFont = nextText;
            foreach (var box in new[] { oldContent, newContent, oldMetadata, newMetadata })
                if (ownedTextFont == null) box.ResetFont(); else box.Font = ownedTextFont;
            previousText?.Dispose();
            RightToLeft = style.RightToLeft ? RightToLeft.Yes : RightToLeft.No;
            RightToLeftLayout = style.RightToLeft;
        }

        private static void ApplyPalette(Control control, SharedChatStyle style)
        {
            control.BackColor = style.Background; control.ForeColor = style.Foreground;
            foreach (Control child in control.Controls) ApplyPalette(child, style);
        }

        protected override void OnShown(EventArgs args)
        {
            if (ownedIcon == null && Owner?.Icon != null) UseBorrowedIcon(Owner.Icon);
            base.OnShown(args);
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

    public enum ChatReferenceComparisonStatus
    {
        CapturePending, ValidationPending, ComparisonUnavailable, SourceChanged, CandidateStale, Accepted, SavingPending, SavingFailed
    }
}

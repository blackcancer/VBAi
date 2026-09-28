using System;
using System.Drawing;
using System.Windows.Forms;

namespace CodexVBE
{
    /// <summary>First integrated editor prototype with explicit, conflict-checked writes.</summary>
    internal sealed class ModernEditorWindow : Form
    {
        private readonly object vbe;
        private readonly RichTextBox code = new RichTextBox { Dock = DockStyle.Fill, AcceptsTab = true, WordWrap = false, DetectUrls = false, HideSelection = false, ReadOnly = true, Font = new Font("Consolas", 11) };
        private readonly Label status = new Label { Dock = DockStyle.Bottom, Height = 52, Padding = new Padding(8), AutoEllipsis = true };
        private readonly Button apply = new Button { AutoSize = true, Enabled = false };
        private readonly Timer timer = new Timer { Interval = 500 };
        private VbeEditorDocument document;
        private bool loading;

        internal ModernEditorWindow(object vbe)
        {
            this.vbe = vbe;
            Text = UiText.Get("VBAi Editor (prototype)");
            Size = new Size(900, 650);
            var toolbar = new FlowLayoutPanel { Dock = DockStyle.Top, AutoSize = true, Padding = new Padding(6) };
            var open = new Button { Text = UiText.Get("Open active module"), AutoSize = true };
            var reload = new Button { Text = UiText.Get("Reload module"), AutoSize = true };
            apply.Text = UiText.Get("Apply to module");
            toolbar.Controls.AddRange(new Control[] { open, apply, reload });
            Controls.Add(code); Controls.Add(toolbar); Controls.Add(status);
            open.Click += (s, e) => Guard(OpenActive);
            reload.Click += (s, e) => Guard(() => { if (document != null && ConfirmDiscard()) { document.Buffer.Reload(); LoadDraft(); } });
            apply.Click += (s, e) => Guard(() => { document.Buffer.Apply(); LoadDraft(); status.Text = UiText.Get("Applied in memory. Save the project in the host application."); });
            code.TextChanged += (s, e) => { if (!loading && document != null) { document.Buffer.Draft = code.Text; RefreshMode(); } };
            timer.Tick += (s, e) => RefreshMode();
            FormClosing += (s, e) => { if (!ConfirmDiscard()) e.Cancel = true; };
            UiTheme.Apply(this);
            timer.Start();
            status.Text = UiText.Get("Select a code module in the VBE, then open it here.");
        }

        internal void OpenActive()
        {
            if (!ConfirmDiscard()) return;
            dynamic host = vbe;
            dynamic pane = host.ActiveCodePane;
            if (pane == null) throw new InvalidOperationException(UiText.Get("Select a code module in the VBE, then open it here."));
            var next = new VbeEditorDocument(vbe, (object)pane.CodeModule.Parent);
            document = next;
            LoadDraft();
        }

        private void LoadDraft()
        {
            loading = true;
            try { code.Text = document.Buffer.Draft; code.ClearUndo(); }
            finally { loading = false; }
            status.Text = document.Title;
            RefreshMode();
        }

        private void RefreshMode()
        {
            bool writable = document != null && document.CanWrite;
            code.ReadOnly = !writable;
            apply.Enabled = writable && document.Buffer.Dirty;
        }

        private bool ConfirmDiscard() => document == null || !document.Buffer.Dirty ||
            MessageBox.Show(this, UiText.Get("Discard the unapplied draft?"), Text, MessageBoxButtons.YesNo, MessageBoxIcon.Warning, MessageBoxDefaultButton.Button2) == DialogResult.Yes;

        private void Guard(Action action)
        {
            try { action(); }
            catch (Exception error) { status.Text = error.Message; }
        }

        protected override void Dispose(bool disposing)
        {
            if (disposing)
            {
                timer.Dispose();
                if (document != null && document.Buffer.Dirty)
                {
                    try
                    {
                        string directory = System.IO.Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData), "CodexVBE", "editor-drafts");
                        System.IO.Directory.CreateDirectory(directory);
                        string path = System.IO.Path.Combine(directory, DateTime.UtcNow.ToString("yyyyMMdd-HHmmss") + "-" + Guid.NewGuid().ToString("N") + ".json");
                        System.IO.File.WriteAllText(path, new System.Web.Script.Serialization.JavaScriptSerializer().Serialize(new
                        { Baseline = document.Buffer.Baseline, Draft = document.Buffer.Draft }));
                        LoadLog.Write("Unapplied editor draft saved: " + path);
                    }
                    catch (Exception error) { LoadLog.Write("Editor draft recovery failed: " + error.Message); }
                    document = null;
                }
            }
            base.Dispose(disposing);
        }
    }
}

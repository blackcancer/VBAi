// SPDX-License-Identifier: MPL-2.0
namespace Blackcancer.ChatUI.Views
{
    partial class ChatReferencePickerForm
    {
        private System.ComponentModel.IContainer components;
        private System.Windows.Forms.TableLayoutPanel layout, queryLayout;
        private System.Windows.Forms.FlowLayoutPanel toolbar;
        private System.Windows.Forms.Label queryLabel, statusNote, explanation;
        private System.Windows.Forms.TextBox searchText;
        private System.Windows.Forms.Button search, inspect, attach, close;
        private System.Windows.Forms.ListView choiceList;
        private System.Windows.Forms.ColumnHeader labelColumn, kindColumn, detailColumn;
        private System.Windows.Forms.Timer debounceTimer;
        protected override void Dispose(bool disposing)
        {
            if (disposing) components?.Dispose();
            base.Dispose(disposing);
            if (disposing) DisposeOwnedAppearance();
        }
        private void InitializeComponent()
        {
            this.components = new System.ComponentModel.Container();
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(ChatReferencePickerForm));
            this.layout = new System.Windows.Forms.TableLayoutPanel();
            this.queryLayout = new System.Windows.Forms.TableLayoutPanel();
            this.queryLabel = new System.Windows.Forms.Label();
            this.searchText = new System.Windows.Forms.TextBox();
            this.search = new System.Windows.Forms.Button();
            this.statusNote = new System.Windows.Forms.Label();
            this.choiceList = new System.Windows.Forms.ListView();
            this.labelColumn = ((System.Windows.Forms.ColumnHeader)(new System.Windows.Forms.ColumnHeader()));
            this.kindColumn = ((System.Windows.Forms.ColumnHeader)(new System.Windows.Forms.ColumnHeader()));
            this.detailColumn = ((System.Windows.Forms.ColumnHeader)(new System.Windows.Forms.ColumnHeader()));
            this.explanation = new System.Windows.Forms.Label();
            this.toolbar = new System.Windows.Forms.FlowLayoutPanel();
            this.inspect = new System.Windows.Forms.Button();
            this.attach = new System.Windows.Forms.Button();
            this.close = new System.Windows.Forms.Button();
            this.debounceTimer = new System.Windows.Forms.Timer(this.components);
            this.layout.SuspendLayout();
            this.queryLayout.SuspendLayout();
            this.toolbar.SuspendLayout();
            this.SuspendLayout();
            // 
            // layout
            // 
            resources.ApplyResources(this.layout, "layout");
            this.layout.Controls.Add(this.queryLayout, 0, 0);
            this.layout.Controls.Add(this.statusNote, 0, 1);
            this.layout.Controls.Add(this.choiceList, 0, 2);
            this.layout.Controls.Add(this.explanation, 0, 3);
            this.layout.Controls.Add(this.toolbar, 0, 4);
            this.layout.Name = "layout";
            // 
            // queryLayout
            // 
            resources.ApplyResources(this.queryLayout, "queryLayout");
            this.queryLayout.Controls.Add(this.queryLabel, 0, 0);
            this.queryLayout.Controls.Add(this.searchText, 1, 0);
            this.queryLayout.Controls.Add(this.search, 2, 0);
            this.queryLayout.Name = "queryLayout";
            // 
            // queryLabel
            // 
            resources.ApplyResources(this.queryLabel, "queryLabel");
            this.queryLabel.Name = "queryLabel";
            // 
            // searchText
            // 
            resources.ApplyResources(this.searchText, "searchText");
            this.searchText.Name = "searchText";
            // 
            // search
            // 
            resources.ApplyResources(this.search, "search");
            this.search.Name = "search";
            // 
            // statusNote
            // 
            resources.ApplyResources(this.statusNote, "statusNote");
            this.statusNote.Name = "statusNote";
            // 
            // choiceList
            // 
            this.choiceList.Columns.AddRange(new System.Windows.Forms.ColumnHeader[] {
            this.labelColumn,
            this.kindColumn,
            this.detailColumn});
            resources.ApplyResources(this.choiceList, "choiceList");
            this.choiceList.FullRowSelect = true;
            this.choiceList.HideSelection = false;
            this.choiceList.MultiSelect = false;
            this.choiceList.Name = "choiceList";
            this.choiceList.UseCompatibleStateImageBehavior = false;
            this.choiceList.View = System.Windows.Forms.View.Details;
            // 
            // labelColumn
            // 
            resources.ApplyResources(this.labelColumn, "labelColumn");
            // 
            // kindColumn
            // 
            resources.ApplyResources(this.kindColumn, "kindColumn");
            // 
            // detailColumn
            // 
            resources.ApplyResources(this.detailColumn, "detailColumn");
            // 
            // explanation
            // 
            resources.ApplyResources(this.explanation, "explanation");
            this.explanation.Name = "explanation";
            // 
            // toolbar
            // 
            resources.ApplyResources(this.toolbar, "toolbar");
            this.toolbar.Controls.Add(this.inspect);
            this.toolbar.Controls.Add(this.attach);
            this.toolbar.Controls.Add(this.close);
            this.toolbar.Name = "toolbar";
            // 
            // inspect
            // 
            resources.ApplyResources(this.inspect, "inspect");
            this.inspect.Name = "inspect";
            // 
            // attach
            // 
            resources.ApplyResources(this.attach, "attach");
            this.attach.Name = "attach";
            // 
            // close
            // 
            resources.ApplyResources(this.close, "close");
            this.close.Name = "close";
            // 
            // debounceTimer
            // 
            this.debounceTimer.Interval = 250;
            // 
            // ChatReferencePickerForm
            // 
            resources.ApplyResources(this, "$this");
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.Controls.Add(this.layout);
            this.Name = "ChatReferencePickerForm";
            this.ShowInTaskbar = false;
            this.layout.ResumeLayout(false);
            this.layout.PerformLayout();
            this.queryLayout.ResumeLayout(false);
            this.queryLayout.PerformLayout();
            this.toolbar.ResumeLayout(false);
            this.toolbar.PerformLayout();
            this.ResumeLayout(false);

        }
    }
}

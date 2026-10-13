// SPDX-License-Identifier: MPL-2.0
// Ported from SolidAi native chat and VBAi native Designer-backed composer; source-vendored host-independent UI.
namespace Blackcancer.ChatUI.Views
{
    partial class ChatWorkspaceView
    {
        private System.ComponentModel.IContainer components = null;
        public System.Windows.Forms.Panel TranscriptHost;
        public System.Windows.Forms.FlowLayoutPanel AdditionalContextHost;
        public System.Windows.Forms.FlowLayoutPanel PendingMessagesHost;
        public System.Windows.Forms.FlowLayoutPanel HeaderActionsHost;
        public Blackcancer.ChatUI.Views.ChatActionButton windowPinButton;
        [System.ComponentModel.Browsable(false), System.ComponentModel.DesignerSerializationVisibility(System.ComponentModel.DesignerSerializationVisibility.Hidden)]
        public System.Windows.Forms.ToolTip toolTips;
        [System.ComponentModel.Browsable(false), System.ComponentModel.DesignerSerializationVisibility(System.ComponentModel.DesignerSerializationVisibility.Hidden)]
        public Blackcancer.ChatUI.Views.ChatActionButton modelSummary;
        [System.ComponentModel.Browsable(false), System.ComponentModel.DesignerSerializationVisibility(System.ComponentModel.DesignerSerializationVisibility.Hidden)]
        public System.Windows.Forms.ComboBox effortPicker;
        [System.ComponentModel.Browsable(false), System.ComponentModel.DesignerSerializationVisibility(System.ComponentModel.DesignerSerializationVisibility.Hidden)]
        public System.Windows.Forms.ContextMenuStrip optionsMenu;
        [System.ComponentModel.Browsable(false), System.ComponentModel.DesignerSerializationVisibility(System.ComponentModel.DesignerSerializationVisibility.Hidden)]
        public System.Windows.Forms.ToolStripMenuItem configureMenuItem;
        protected override void Dispose(bool disposing)
        {
            if (disposing)
            {
                fileAttachmentDialog?.Dispose();
                if (components != null) components.Dispose();
            }
            base.Dispose(disposing);
        }
        #region Windows Form Designer generated code
        private void InitializeComponent()
        {
            this.components = new System.ComponentModel.Container();
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(ChatWorkspaceView));
            this.TranscriptHost = new System.Windows.Forms.Panel();
            this.transcriptPanel = new Blackcancer.ChatUI.Views.ChatTranscriptView();
            this.AdditionalContextHost = new System.Windows.Forms.FlowLayoutPanel();
            this.PendingMessagesHost = new System.Windows.Forms.FlowLayoutPanel();
            this.HeaderActionsHost = new System.Windows.Forms.FlowLayoutPanel();
            this.windowPinButton = new Blackcancer.ChatUI.Views.ChatActionButton();
            this.fileAttachmentDialog = new System.Windows.Forms.OpenFileDialog();
            this.toolTips = new System.Windows.Forms.ToolTip(this.components);
            this.modelSummary = new Blackcancer.ChatUI.Views.ChatActionButton();
            this.effortPicker = new System.Windows.Forms.ComboBox();
            this.optionsMenu = new System.Windows.Forms.ContextMenuStrip(this.components);
            this.configureMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.rootLayout = new System.Windows.Forms.TableLayoutPanel();
            this.headingLayout = new System.Windows.Forms.TableLayoutPanel();
            this.titleLayout = new System.Windows.Forms.TableLayoutPanel();
            this.sessionTitle = new System.Windows.Forms.Label();
            this.newChatButton = new Blackcancer.ChatUI.Views.ChatActionButton();
            this.settingsButton = new Blackcancer.ChatUI.Views.ChatActionButton();
            this.scopeLayout = new System.Windows.Forms.TableLayoutPanel();
            this.scopePicker = new System.Windows.Forms.ComboBox();
            this.historyButton = new Blackcancer.ChatUI.Views.ChatActionButton();
            this.projectLayout = new System.Windows.Forms.TableLayoutPanel();
            this.projectPicker = new System.Windows.Forms.ComboBox();
            this.createProjectButton = new Blackcancer.ChatUI.Views.ChatActionButton();
            this.conversationPanel = new System.Windows.Forms.Panel();
            this.welcomePanel = new System.Windows.Forms.TableLayoutPanel();
            this.welcomeTitle = new System.Windows.Forms.Label();
            this.welcomeBody = new System.Windows.Forms.Label();
            this.historyPanel = new System.Windows.Forms.Panel();
            this.historyLayout = new System.Windows.Forms.TableLayoutPanel();
            this.historyLabel = new System.Windows.Forms.Label();
            this.historyEmptyLabel = new System.Windows.Forms.Label();
            this.sessionList = new System.Windows.Forms.ListBox();
            this.closeHistoryButton = new Blackcancer.ChatUI.Views.ChatActionButton();
            this.historySearchTextBox = new System.Windows.Forms.TextBox();
            this.showArchivedCheckBox = new System.Windows.Forms.CheckBox();
            this.chatTitleEditor = new System.Windows.Forms.TextBox();
            this.historyActions = new System.Windows.Forms.FlowLayoutPanel();
            this.renameChatButton = new Blackcancer.ChatUI.Views.ChatActionButton();
            this.archiveChatButton = new Blackcancer.ChatUI.Views.ChatActionButton();
            this.pinChatButton = new Blackcancer.ChatUI.Views.ChatActionButton();
            this.exportChatButton = new Blackcancer.ChatUI.Views.ChatActionButton();
            this.deleteChatButton = new Blackcancer.ChatUI.Views.ChatActionButton();
            this.memoryToggleButton = new Blackcancer.ChatUI.Views.ChatActionButton();
            this.memoryPanel = new System.Windows.Forms.GroupBox();
            this.memoryLayout = new System.Windows.Forms.TableLayoutPanel();
            this.memoryEditor = new System.Windows.Forms.TextBox();
            this.attachMemoryCheckBox = new System.Windows.Forms.CheckBox();
            this.saveMemoryButton = new Blackcancer.ChatUI.Views.ChatActionButton();
            this.composerLayout = new System.Windows.Forms.TableLayoutPanel();
            this.contextChips = new System.Windows.Forms.FlowLayoutPanel();
            this.inputTextBox = new Blackcancer.ChatUI.Views.ChatInputView();
            this.actionsLayout = new System.Windows.Forms.TableLayoutPanel();
            this.attachContextButton = new Blackcancer.ChatUI.Views.ChatActionButton();
            this.cadContextMenu = new System.Windows.Forms.ContextMenuStrip(this.components);
            this.attachDocumentMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.attachSelectionMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.attachFileButton = new Blackcancer.ChatUI.Views.ChatActionButton();
            this.sendActionsPanel = new System.Windows.Forms.Panel();
            this.stopButton = new Blackcancer.ChatUI.Views.ChatActionButton();
            this.sendButton = new Blackcancer.ChatUI.Views.ChatActionButton();
            this.workflowLayout = new System.Windows.Forms.TableLayoutPanel();
            this.modePicker = new System.Windows.Forms.ComboBox();
            this.approvalPicker = new System.Windows.Forms.ComboBox();
            this.attachSelectionButton = new Blackcancer.ChatUI.Views.ChatActionButton();
            this.contextButton = new Blackcancer.ChatUI.Views.ChatActionButton();
            this.contextPreview = new System.Windows.Forms.TextBox();
            this.providerLayout = new System.Windows.Forms.TableLayoutPanel();
            this.providerPicker = new System.Windows.Forms.ComboBox();
            this.modelPicker = new System.Windows.Forms.ComboBox();
            this.latestButton = new Blackcancer.ChatUI.Views.ChatActionButton();
            this.footerLayout = new System.Windows.Forms.TableLayoutPanel();
            this.statusLabel = new System.Windows.Forms.Label();
            this.activityBar = new System.Windows.Forms.ProgressBar();
            this.appTitle = new System.Windows.Forms.Label();
            this.contextLabel = new System.Windows.Forms.Label();
            this.TranscriptHost.SuspendLayout();
            this.optionsMenu.SuspendLayout();
            this.rootLayout.SuspendLayout();
            this.headingLayout.SuspendLayout();
            this.titleLayout.SuspendLayout();
            this.scopeLayout.SuspendLayout();
            this.projectLayout.SuspendLayout();
            this.conversationPanel.SuspendLayout();
            this.welcomePanel.SuspendLayout();
            this.historyPanel.SuspendLayout();
            this.historyLayout.SuspendLayout();
            this.historyActions.SuspendLayout();
            this.memoryPanel.SuspendLayout();
            this.memoryLayout.SuspendLayout();
            this.composerLayout.SuspendLayout();
            this.actionsLayout.SuspendLayout();
            this.cadContextMenu.SuspendLayout();
            this.sendActionsPanel.SuspendLayout();
            this.workflowLayout.SuspendLayout();
            this.providerLayout.SuspendLayout();
            this.footerLayout.SuspendLayout();
            this.SuspendLayout();
            // 
            // TranscriptHost
            // 
            this.TranscriptHost.Controls.Add(this.transcriptPanel);
            resources.ApplyResources(this.TranscriptHost, "TranscriptHost");
            this.TranscriptHost.Name = "TranscriptHost";
            // 
            // transcriptPanel
            // 
            resources.ApplyResources(this.transcriptPanel, "transcriptPanel");
            this.transcriptPanel.Name = "transcriptPanel";
            // 
            // AdditionalContextHost
            // 
            resources.ApplyResources(this.AdditionalContextHost, "AdditionalContextHost");
            this.AdditionalContextHost.Name = "AdditionalContextHost";
            // 
            // PendingMessagesHost
            // 
            resources.ApplyResources(this.PendingMessagesHost, "PendingMessagesHost");
            this.PendingMessagesHost.Name = "PendingMessagesHost";
            // 
            // HeaderActionsHost
            // 
            resources.ApplyResources(this.HeaderActionsHost, "HeaderActionsHost");
            this.HeaderActionsHost.Name = "HeaderActionsHost";
            this.HeaderActionsHost.Controls.Add(this.windowPinButton);
            this.windowPinButton.Name = "pinButton";
            this.windowPinButton.Size = new System.Drawing.Size(32, 30);
            this.windowPinButton.Margin = new System.Windows.Forms.Padding(0);
            this.windowPinButton.IconOnly = true;
            this.windowPinButton.Symbol = "\uE718";
            this.windowPinButton.Text = "Pin window";
            this.windowPinButton.Visible = false;
            // 
            // fileAttachmentDialog
            // 
            resources.ApplyResources(this.fileAttachmentDialog, "fileAttachmentDialog");
            this.fileAttachmentDialog.Multiselect = true;
            this.fileAttachmentDialog.RestoreDirectory = true;
            // 
            // modelSummary
            // 
            resources.ApplyResources(this.modelSummary, "modelSummary");
            this.modelSummary.Cursor = System.Windows.Forms.Cursors.Hand;
            this.modelSummary.Name = "modelSummary";
            // 
            // effortPicker
            // 
            resources.ApplyResources(this.effortPicker, "effortPicker");
            this.effortPicker.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.effortPicker.FormattingEnabled = true;
            this.effortPicker.Name = "effortPicker";
            // 
            // optionsMenu
            // 
            this.optionsMenu.Items.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.configureMenuItem});
            this.optionsMenu.Name = "optionsMenu";
            resources.ApplyResources(this.optionsMenu, "optionsMenu");
            // 
            // configureMenuItem
            // 
            this.configureMenuItem.Name = "configureMenuItem";
            resources.ApplyResources(this.configureMenuItem, "configureMenuItem");
            // 
            // rootLayout
            // 
            resources.ApplyResources(this.rootLayout, "rootLayout");
            this.rootLayout.Controls.Add(this.headingLayout, 0, 0);
            this.rootLayout.Controls.Add(this.scopeLayout, 0, 1);
            this.rootLayout.Controls.Add(this.AdditionalContextHost, 0, 2);
            this.rootLayout.Controls.Add(this.projectLayout, 0, 3);
            this.rootLayout.Controls.Add(this.PendingMessagesHost, 0, 4);
            this.rootLayout.Controls.Add(this.conversationPanel, 0, 5);
            this.rootLayout.Controls.Add(this.composerLayout, 0, 6);
            this.rootLayout.Controls.Add(this.providerLayout, 0, 7);
            this.rootLayout.Controls.Add(this.latestButton, 0, 8);
            this.rootLayout.Controls.Add(this.footerLayout, 0, 9);
            this.rootLayout.Name = "rootLayout";
            // 
            // headingLayout
            // 
            resources.ApplyResources(this.headingLayout, "headingLayout");
            this.headingLayout.Controls.Add(this.titleLayout, 0, 0);
            this.headingLayout.Controls.Add(this.newChatButton, 1, 0);
            this.headingLayout.Controls.Add(this.settingsButton, 2, 0);
            this.headingLayout.Controls.Add(this.HeaderActionsHost, 3, 0);
            this.headingLayout.Name = "headingLayout";
            // 
            // titleLayout
            // 
            resources.ApplyResources(this.titleLayout, "titleLayout");
            this.titleLayout.Controls.Add(this.sessionTitle, 0, 0);
            this.titleLayout.Name = "titleLayout";
            // 
            // sessionTitle
            // 
            this.sessionTitle.AutoEllipsis = true;
            resources.ApplyResources(this.sessionTitle, "sessionTitle");
            this.sessionTitle.Name = "sessionTitle";
            // 
            // newChatButton
            // 
            resources.ApplyResources(this.newChatButton, "newChatButton");
            this.newChatButton.Cursor = System.Windows.Forms.Cursors.Hand;
            this.newChatButton.FlatAppearance.BorderSize = 0;
            this.newChatButton.IconOnly = true;
            this.newChatButton.Name = "newChatButton";
            this.newChatButton.Symbol = "+";
            // 
            // settingsButton
            // 
            resources.ApplyResources(this.settingsButton, "settingsButton");
            this.settingsButton.Cursor = System.Windows.Forms.Cursors.Hand;
            this.settingsButton.FlatAppearance.BorderSize = 0;
            this.settingsButton.IconOnly = true;
            this.settingsButton.Name = "settingsButton";
            this.settingsButton.Symbol = "⋯";
            // 
            // scopeLayout
            // 
            resources.ApplyResources(this.scopeLayout, "scopeLayout");
            this.scopeLayout.Controls.Add(this.scopePicker, 1, 0);
            this.scopeLayout.Controls.Add(this.historyButton, 0, 0);
            this.scopeLayout.Name = "scopeLayout";
            // 
            // scopePicker
            // 
            resources.ApplyResources(this.scopePicker, "scopePicker");
            this.scopePicker.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.scopePicker.FormattingEnabled = true;
            this.scopePicker.Name = "scopePicker";
            // 
            // historyButton
            // 
            this.historyButton.Cursor = System.Windows.Forms.Cursors.Hand;
            resources.ApplyResources(this.historyButton, "historyButton");
            this.historyButton.IconOnly = true;
            this.historyButton.Name = "historyButton";
            this.historyButton.Symbol = "";
            // 
            // projectLayout
            // 
            resources.ApplyResources(this.projectLayout, "projectLayout");
            this.projectLayout.Controls.Add(this.projectPicker, 0, 0);
            this.projectLayout.Controls.Add(this.createProjectButton, 1, 0);
            this.projectLayout.Name = "projectLayout";
            // 
            // projectPicker
            // 
            resources.ApplyResources(this.projectPicker, "projectPicker");
            this.projectPicker.DisplayMember = "Name";
            this.projectPicker.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.projectPicker.Name = "projectPicker";
            // 
            // createProjectButton
            // 
            this.createProjectButton.Cursor = System.Windows.Forms.Cursors.Hand;
            resources.ApplyResources(this.createProjectButton, "createProjectButton");
            this.createProjectButton.IconOnly = true;
            this.createProjectButton.Name = "createProjectButton";
            this.createProjectButton.Symbol = "+";
            // 
            // conversationPanel
            // 
            this.conversationPanel.Controls.Add(this.TranscriptHost);
            this.conversationPanel.Controls.Add(this.welcomePanel);
            this.conversationPanel.Controls.Add(this.historyPanel);
            resources.ApplyResources(this.conversationPanel, "conversationPanel");
            this.conversationPanel.Name = "conversationPanel";
            // 
            // welcomePanel
            // 
            resources.ApplyResources(this.welcomePanel, "welcomePanel");
            this.welcomePanel.Controls.Add(this.welcomeTitle, 0, 0);
            this.welcomePanel.Controls.Add(this.welcomeBody, 0, 1);
            this.welcomePanel.Name = "welcomePanel";
            // 
            // welcomeTitle
            // 
            resources.ApplyResources(this.welcomeTitle, "welcomeTitle");
            this.welcomeTitle.Name = "welcomeTitle";
            // 
            // welcomeBody
            // 
            resources.ApplyResources(this.welcomeBody, "welcomeBody");
            this.welcomeBody.Name = "welcomeBody";
            // 
            // historyPanel
            // 
            this.historyPanel.Controls.Add(this.historyLayout);
            resources.ApplyResources(this.historyPanel, "historyPanel");
            this.historyPanel.Name = "historyPanel";
            // 
            // historyLayout
            // 
            resources.ApplyResources(this.historyLayout, "historyLayout");
            this.historyLayout.Controls.Add(this.historyLabel, 0, 0);
            this.historyLayout.Controls.Add(this.historyEmptyLabel, 0, 3);
            this.historyLayout.Controls.Add(this.sessionList, 0, 4);
            this.historyLayout.Controls.Add(this.closeHistoryButton, 1, 0);
            this.historyLayout.Controls.Add(this.historySearchTextBox, 0, 1);
            this.historyLayout.Controls.Add(this.showArchivedCheckBox, 0, 2);
            this.historyLayout.Controls.Add(this.chatTitleEditor, 0, 5);
            this.historyLayout.Controls.Add(this.historyActions, 0, 6);
            this.historyLayout.Controls.Add(this.memoryToggleButton, 0, 7);
            this.historyLayout.Controls.Add(this.memoryPanel, 0, 8);
            this.historyLayout.Name = "historyLayout";
            // 
            // historyLabel
            // 
            resources.ApplyResources(this.historyLabel, "historyLabel");
            this.historyLabel.Name = "historyLabel";
            // 
            // historyEmptyLabel
            // 
            resources.ApplyResources(this.historyEmptyLabel, "historyEmptyLabel");
            this.historyLayout.SetColumnSpan(this.historyEmptyLabel, 2);
            this.historyEmptyLabel.Name = "historyEmptyLabel";
            // 
            // sessionList
            // 
            this.sessionList.BorderStyle = System.Windows.Forms.BorderStyle.None;
            this.historyLayout.SetColumnSpan(this.sessionList, 2);
            resources.ApplyResources(this.sessionList, "sessionList");
            this.sessionList.FormattingEnabled = true;
            this.sessionList.Name = "sessionList";
            // 
            // closeHistoryButton
            // 
            this.closeHistoryButton.Cursor = System.Windows.Forms.Cursors.Hand;
            resources.ApplyResources(this.closeHistoryButton, "closeHistoryButton");
            this.closeHistoryButton.IconOnly = true;
            this.closeHistoryButton.Name = "closeHistoryButton";
            this.closeHistoryButton.Symbol = "×";
            // 
            // historySearchTextBox
            // 
            resources.ApplyResources(this.historySearchTextBox, "historySearchTextBox");
            this.historySearchTextBox.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
            this.historyLayout.SetColumnSpan(this.historySearchTextBox, 2);
            this.historySearchTextBox.Name = "historySearchTextBox";
            // 
            // showArchivedCheckBox
            // 
            resources.ApplyResources(this.showArchivedCheckBox, "showArchivedCheckBox");
            this.historyLayout.SetColumnSpan(this.showArchivedCheckBox, 2);
            this.showArchivedCheckBox.Name = "showArchivedCheckBox";
            // 
            // chatTitleEditor
            // 
            resources.ApplyResources(this.chatTitleEditor, "chatTitleEditor");
            this.chatTitleEditor.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
            this.historyLayout.SetColumnSpan(this.chatTitleEditor, 2);
            this.chatTitleEditor.Name = "chatTitleEditor";
            // 
            // historyActions
            // 
            resources.ApplyResources(this.historyActions, "historyActions");
            this.historyLayout.SetColumnSpan(this.historyActions, 2);
            this.historyActions.Controls.Add(this.renameChatButton);
            this.historyActions.Controls.Add(this.archiveChatButton);
            this.historyActions.Controls.Add(this.pinChatButton);
            this.historyActions.Controls.Add(this.exportChatButton);
            this.historyActions.Controls.Add(this.deleteChatButton);
            this.historyActions.Name = "historyActions";
            // 
            // renameChatButton
            // 
            resources.ApplyResources(this.renameChatButton, "renameChatButton");
            this.renameChatButton.Cursor = System.Windows.Forms.Cursors.Hand;
            this.renameChatButton.IconOnly = true;
            this.renameChatButton.Name = "renameChatButton";
            this.renameChatButton.Symbol = "";
            // 
            // archiveChatButton
            // 
            resources.ApplyResources(this.archiveChatButton, "archiveChatButton");
            this.archiveChatButton.Cursor = System.Windows.Forms.Cursors.Hand;
            this.archiveChatButton.IconOnly = true;
            this.archiveChatButton.Name = "archiveChatButton";
            this.archiveChatButton.Symbol = "";
            // 
            // pinChatButton
            // 
            resources.ApplyResources(this.pinChatButton, "pinChatButton");
            this.pinChatButton.Cursor = System.Windows.Forms.Cursors.Hand;
            this.pinChatButton.IconOnly = true;
            this.pinChatButton.Name = "pinChatButton";
            this.pinChatButton.Symbol = "";
            // 
            // exportChatButton
            // 
            resources.ApplyResources(this.exportChatButton, "exportChatButton");
            this.exportChatButton.Cursor = System.Windows.Forms.Cursors.Hand;
            this.exportChatButton.IconOnly = true;
            this.exportChatButton.Name = "exportChatButton";
            this.exportChatButton.Symbol = "↓";
            // 
            // deleteChatButton
            // 
            resources.ApplyResources(this.deleteChatButton, "deleteChatButton");
            this.deleteChatButton.Cursor = System.Windows.Forms.Cursors.Hand;
            this.deleteChatButton.IconOnly = true;
            this.deleteChatButton.Name = "deleteChatButton";
            this.deleteChatButton.Symbol = "";
            // 
            // memoryToggleButton
            // 
            resources.ApplyResources(this.memoryToggleButton, "memoryToggleButton");
            this.historyLayout.SetColumnSpan(this.memoryToggleButton, 2);
            this.memoryToggleButton.Cursor = System.Windows.Forms.Cursors.Hand;
            this.memoryToggleButton.Name = "memoryToggleButton";
            // 
            // memoryPanel
            // 
            resources.ApplyResources(this.memoryPanel, "memoryPanel");
            this.historyLayout.SetColumnSpan(this.memoryPanel, 2);
            this.memoryPanel.Controls.Add(this.memoryLayout);
            this.memoryPanel.Name = "memoryPanel";
            this.memoryPanel.TabStop = false;
            // 
            // memoryLayout
            // 
            resources.ApplyResources(this.memoryLayout, "memoryLayout");
            this.memoryLayout.Controls.Add(this.memoryEditor, 0, 0);
            this.memoryLayout.Controls.Add(this.attachMemoryCheckBox, 0, 1);
            this.memoryLayout.Controls.Add(this.saveMemoryButton, 0, 2);
            this.memoryLayout.Name = "memoryLayout";
            // 
            // memoryEditor
            // 
            this.memoryEditor.AcceptsReturn = true;
            this.memoryEditor.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
            resources.ApplyResources(this.memoryEditor, "memoryEditor");
            this.memoryEditor.Name = "memoryEditor";
            // 
            // attachMemoryCheckBox
            // 
            resources.ApplyResources(this.attachMemoryCheckBox, "attachMemoryCheckBox");
            this.attachMemoryCheckBox.Name = "attachMemoryCheckBox";
            // 
            // saveMemoryButton
            // 
            resources.ApplyResources(this.saveMemoryButton, "saveMemoryButton");
            this.saveMemoryButton.Cursor = System.Windows.Forms.Cursors.Hand;
            this.saveMemoryButton.Name = "saveMemoryButton";
            // 
            // composerLayout
            // 
            resources.ApplyResources(this.composerLayout, "composerLayout");
            this.composerLayout.BackColor = System.Drawing.Color.White;
            this.composerLayout.Controls.Add(this.contextChips, 0, 0);
            this.composerLayout.Controls.Add(this.inputTextBox, 0, 2);
            this.composerLayout.Controls.Add(this.actionsLayout, 0, 3);
            this.composerLayout.Controls.Add(this.workflowLayout, 0, 4);
            this.composerLayout.Controls.Add(this.contextPreview, 0, 5);
            this.composerLayout.Name = "composerLayout";
            // 
            // contextChips
            // 
            resources.ApplyResources(this.contextChips, "contextChips");
            this.contextChips.Name = "contextChips";
            // 
            // inputTextBox
            // 
            resources.ApplyResources(this.inputTextBox, "inputTextBox");
            this.inputTextBox.Name = "inputTextBox";
            // 
            // actionsLayout
            // 
            resources.ApplyResources(this.actionsLayout, "actionsLayout");
            this.actionsLayout.Controls.Add(this.attachContextButton, 0, 0);
            this.actionsLayout.Controls.Add(this.attachFileButton, 1, 0);
            this.actionsLayout.Controls.Add(this.modelSummary, 2, 0);
            this.actionsLayout.Controls.Add(this.sendActionsPanel, 3, 0);
            this.actionsLayout.Name = "actionsLayout";
            // 
            // attachContextButton
            // 
            resources.ApplyResources(this.attachContextButton, "attachContextButton");
            this.attachContextButton.ContextMenuStrip = this.cadContextMenu;
            this.attachContextButton.Cursor = System.Windows.Forms.Cursors.Hand;
            this.attachContextButton.IconOnly = true;
            this.attachContextButton.Name = "attachContextButton";
            this.attachContextButton.Symbol = "";
            // 
            // cadContextMenu
            // 
            this.cadContextMenu.Items.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.attachDocumentMenuItem,
            this.attachSelectionMenuItem});
            this.cadContextMenu.Name = "cadContextMenu";
            resources.ApplyResources(this.cadContextMenu, "cadContextMenu");
            // 
            // attachDocumentMenuItem
            // 
            this.attachDocumentMenuItem.Name = "attachDocumentMenuItem";
            resources.ApplyResources(this.attachDocumentMenuItem, "attachDocumentMenuItem");
            // 
            // attachSelectionMenuItem
            // 
            this.attachSelectionMenuItem.Name = "attachSelectionMenuItem";
            resources.ApplyResources(this.attachSelectionMenuItem, "attachSelectionMenuItem");
            // 
            // attachFileButton
            // 
            resources.ApplyResources(this.attachFileButton, "attachFileButton");
            this.attachFileButton.Cursor = System.Windows.Forms.Cursors.Hand;
            this.attachFileButton.IconOnly = true;
            this.attachFileButton.Name = "attachFileButton";
            this.attachFileButton.Symbol = "";
            // 
            // sendActionsPanel
            // 
            this.sendActionsPanel.Controls.Add(this.stopButton);
            this.sendActionsPanel.Controls.Add(this.sendButton);
            resources.ApplyResources(this.sendActionsPanel, "sendActionsPanel");
            this.sendActionsPanel.Name = "sendActionsPanel";
            // 
            // stopButton
            // 
            this.stopButton.Cursor = System.Windows.Forms.Cursors.Hand;
            resources.ApplyResources(this.stopButton, "stopButton");
            this.stopButton.FlatAppearance.BorderSize = 0;
            this.stopButton.IconOnly = true;
            this.stopButton.Name = "stopButton";
            this.stopButton.Primary = true;
            this.stopButton.Symbol = "■";
            // 
            // sendButton
            // 
            this.sendButton.Cursor = System.Windows.Forms.Cursors.Hand;
            resources.ApplyResources(this.sendButton, "sendButton");
            this.sendButton.FlatAppearance.BorderSize = 0;
            this.sendButton.IconOnly = true;
            this.sendButton.Name = "sendButton";
            this.sendButton.Primary = true;
            this.sendButton.Symbol = "↑";
            // 
            // workflowLayout
            // 
            resources.ApplyResources(this.workflowLayout, "workflowLayout");
            this.workflowLayout.Controls.Add(this.modePicker, 0, 0);
            this.workflowLayout.Controls.Add(this.approvalPicker, 1, 0);
            this.workflowLayout.Controls.Add(this.attachSelectionButton, 2, 0);
            this.workflowLayout.Controls.Add(this.contextButton, 3, 0);
            this.workflowLayout.Name = "workflowLayout";
            // 
            // modePicker
            // 
            resources.ApplyResources(this.modePicker, "modePicker");
            this.modePicker.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.modePicker.FormattingEnabled = true;
            this.modePicker.Name = "modePicker";
            // 
            // approvalPicker
            // 
            resources.ApplyResources(this.approvalPicker, "approvalPicker");
            this.approvalPicker.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.approvalPicker.DropDownWidth = 260;
            this.approvalPicker.FormattingEnabled = true;
            this.approvalPicker.Name = "approvalPicker";
            // 
            // attachSelectionButton
            // 
            resources.ApplyResources(this.attachSelectionButton, "attachSelectionButton");
            this.attachSelectionButton.Cursor = System.Windows.Forms.Cursors.Hand;
            this.attachSelectionButton.IconOnly = true;
            this.attachSelectionButton.Name = "attachSelectionButton";
            this.attachSelectionButton.Symbol = "#";
            // 
            // contextButton
            // 
            resources.ApplyResources(this.contextButton, "contextButton");
            this.contextButton.Cursor = System.Windows.Forms.Cursors.Hand;
            this.contextButton.FlatAppearance.BorderSize = 0;
            this.contextButton.IconOnly = true;
            this.contextButton.Name = "contextButton";
            this.contextButton.Symbol = "";
            // 
            // contextPreview
            // 
            this.contextPreview.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
            resources.ApplyResources(this.contextPreview, "contextPreview");
            this.contextPreview.Name = "contextPreview";
            this.contextPreview.ReadOnly = true;
            // 
            // providerLayout
            // 
            resources.ApplyResources(this.providerLayout, "providerLayout");
            this.providerLayout.Controls.Add(this.effortPicker, 2, 0);
            this.providerLayout.Controls.Add(this.providerPicker, 0, 0);
            this.providerLayout.Controls.Add(this.modelPicker, 1, 0);
            this.providerLayout.Name = "providerLayout";
            // 
            // providerPicker
            // 
            resources.ApplyResources(this.providerPicker, "providerPicker");
            this.providerPicker.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.providerPicker.FormattingEnabled = true;
            this.providerPicker.Name = "providerPicker";
            // 
            // modelPicker
            // 
            resources.ApplyResources(this.modelPicker, "modelPicker");
            this.modelPicker.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.modelPicker.FormattingEnabled = true;
            this.modelPicker.Name = "modelPicker";
            // 
            // latestButton
            // 
            this.latestButton.Cursor = System.Windows.Forms.Cursors.Hand;
            this.latestButton.FlatAppearance.BorderSize = 0;
            resources.ApplyResources(this.latestButton, "latestButton");
            this.latestButton.IconOnly = true;
            this.latestButton.Name = "latestButton";
            this.latestButton.Symbol = "↓";
            // 
            // footerLayout
            // 
            resources.ApplyResources(this.footerLayout, "footerLayout");
            this.footerLayout.Controls.Add(this.statusLabel, 0, 0);
            this.footerLayout.Controls.Add(this.activityBar, 1, 0);
            this.footerLayout.Name = "footerLayout";
            // 
            // statusLabel
            // 
            resources.ApplyResources(this.statusLabel, "statusLabel");
            this.statusLabel.AutoEllipsis = true;
            this.statusLabel.Name = "statusLabel";
            // 
            // activityBar
            // 
            resources.ApplyResources(this.activityBar, "activityBar");
            this.activityBar.MarqueeAnimationSpeed = 30;
            this.activityBar.Name = "activityBar";
            this.activityBar.Style = System.Windows.Forms.ProgressBarStyle.Marquee;
            // 
            // appTitle
            // 
            resources.ApplyResources(this.appTitle, "appTitle");
            this.appTitle.Name = "appTitle";
            // 
            // contextLabel
            // 
            resources.ApplyResources(this.contextLabel, "contextLabel");
            this.contextLabel.AutoEllipsis = true;
            this.contextLabel.Name = "contextLabel";
            // 
            // ChatWorkspaceView
            // 
            resources.ApplyResources(this, "$this");
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(248)))), ((int)(((byte)(250)))), ((int)(((byte)(252)))));
            this.Controls.Add(this.rootLayout);
            this.ForeColor = System.Drawing.Color.FromArgb(((int)(((byte)(51)))), ((int)(((byte)(65)))), ((int)(((byte)(85)))));
            this.Name = "ChatWorkspaceView";
            this.TranscriptHost.ResumeLayout(false);
            this.optionsMenu.ResumeLayout(false);
            this.rootLayout.ResumeLayout(false);
            this.rootLayout.PerformLayout();
            this.headingLayout.ResumeLayout(false);
            this.headingLayout.PerformLayout();
            this.titleLayout.ResumeLayout(false);
            this.scopeLayout.ResumeLayout(false);
            this.projectLayout.ResumeLayout(false);
            this.conversationPanel.ResumeLayout(false);
            this.welcomePanel.ResumeLayout(false);
            this.welcomePanel.PerformLayout();
            this.historyPanel.ResumeLayout(false);
            this.historyLayout.ResumeLayout(false);
            this.historyLayout.PerformLayout();
            this.historyActions.ResumeLayout(false);
            this.memoryPanel.ResumeLayout(false);
            this.memoryLayout.ResumeLayout(false);
            this.memoryLayout.PerformLayout();
            this.composerLayout.ResumeLayout(false);
            this.composerLayout.PerformLayout();
            this.actionsLayout.ResumeLayout(false);
            this.cadContextMenu.ResumeLayout(false);
            this.sendActionsPanel.ResumeLayout(false);
            this.workflowLayout.ResumeLayout(false);
            this.providerLayout.ResumeLayout(false);
            this.footerLayout.ResumeLayout(false);
            this.ResumeLayout(false);

        }
        #endregion
        [System.ComponentModel.Browsable(false), System.ComponentModel.DesignerSerializationVisibility(System.ComponentModel.DesignerSerializationVisibility.Hidden)]
        public System.Windows.Forms.TableLayoutPanel titleLayout;
        [System.ComponentModel.Browsable(false), System.ComponentModel.DesignerSerializationVisibility(System.ComponentModel.DesignerSerializationVisibility.Hidden)]
        public System.Windows.Forms.Label sessionTitle;
        [System.ComponentModel.Browsable(false), System.ComponentModel.DesignerSerializationVisibility(System.ComponentModel.DesignerSerializationVisibility.Hidden)]
        public System.Windows.Forms.ComboBox scopePicker;
        [System.ComponentModel.Browsable(false), System.ComponentModel.DesignerSerializationVisibility(System.ComponentModel.DesignerSerializationVisibility.Hidden)]
        public Blackcancer.ChatUI.Views.ChatActionButton historyButton;
        [System.ComponentModel.Browsable(false), System.ComponentModel.DesignerSerializationVisibility(System.ComponentModel.DesignerSerializationVisibility.Hidden)]
        public System.Windows.Forms.TableLayoutPanel workflowLayout;
        [System.ComponentModel.Browsable(false), System.ComponentModel.DesignerSerializationVisibility(System.ComponentModel.DesignerSerializationVisibility.Hidden)]
        public System.Windows.Forms.ComboBox approvalPicker;
        [System.ComponentModel.Browsable(false), System.ComponentModel.DesignerSerializationVisibility(System.ComponentModel.DesignerSerializationVisibility.Hidden)]
        public System.Windows.Forms.TableLayoutPanel projectLayout;
        [System.ComponentModel.Browsable(false), System.ComponentModel.DesignerSerializationVisibility(System.ComponentModel.DesignerSerializationVisibility.Hidden)]
        public System.Windows.Forms.ComboBox projectPicker;
        [System.ComponentModel.Browsable(false), System.ComponentModel.DesignerSerializationVisibility(System.ComponentModel.DesignerSerializationVisibility.Hidden)]
        public Blackcancer.ChatUI.Views.ChatActionButton createProjectButton;
        [System.ComponentModel.Browsable(false), System.ComponentModel.DesignerSerializationVisibility(System.ComponentModel.DesignerSerializationVisibility.Hidden)]
        public System.Windows.Forms.Panel historyPanel;
        [System.ComponentModel.Browsable(false), System.ComponentModel.DesignerSerializationVisibility(System.ComponentModel.DesignerSerializationVisibility.Hidden)]
        public System.Windows.Forms.TableLayoutPanel historyLayout;
        [System.ComponentModel.Browsable(false), System.ComponentModel.DesignerSerializationVisibility(System.ComponentModel.DesignerSerializationVisibility.Hidden)]
        public System.Windows.Forms.Label historyLabel;
        [System.ComponentModel.Browsable(false), System.ComponentModel.DesignerSerializationVisibility(System.ComponentModel.DesignerSerializationVisibility.Hidden)]
        public System.Windows.Forms.Label historyEmptyLabel;
        [System.ComponentModel.Browsable(false), System.ComponentModel.DesignerSerializationVisibility(System.ComponentModel.DesignerSerializationVisibility.Hidden)]
        public System.Windows.Forms.ListBox sessionList;
        [System.ComponentModel.Browsable(false), System.ComponentModel.DesignerSerializationVisibility(System.ComponentModel.DesignerSerializationVisibility.Hidden)]
        public Blackcancer.ChatUI.Views.ChatActionButton closeHistoryButton;
        [System.ComponentModel.Browsable(false), System.ComponentModel.DesignerSerializationVisibility(System.ComponentModel.DesignerSerializationVisibility.Hidden)]
        public System.Windows.Forms.TableLayoutPanel rootLayout;
        [System.ComponentModel.Browsable(false), System.ComponentModel.DesignerSerializationVisibility(System.ComponentModel.DesignerSerializationVisibility.Hidden)]
        public System.Windows.Forms.TableLayoutPanel headingLayout;
        [System.ComponentModel.Browsable(false), System.ComponentModel.DesignerSerializationVisibility(System.ComponentModel.DesignerSerializationVisibility.Hidden)]
        public System.Windows.Forms.Label appTitle;
        [System.ComponentModel.Browsable(false), System.ComponentModel.DesignerSerializationVisibility(System.ComponentModel.DesignerSerializationVisibility.Hidden)]
        public Blackcancer.ChatUI.Views.ChatActionButton newChatButton;
        [System.ComponentModel.Browsable(false), System.ComponentModel.DesignerSerializationVisibility(System.ComponentModel.DesignerSerializationVisibility.Hidden)]
        public Blackcancer.ChatUI.Views.ChatActionButton settingsButton;
        [System.ComponentModel.Browsable(false), System.ComponentModel.DesignerSerializationVisibility(System.ComponentModel.DesignerSerializationVisibility.Hidden)]
        public System.Windows.Forms.TableLayoutPanel scopeLayout;
        [System.ComponentModel.Browsable(false), System.ComponentModel.DesignerSerializationVisibility(System.ComponentModel.DesignerSerializationVisibility.Hidden)]
        public System.Windows.Forms.Label contextLabel;
        [System.ComponentModel.Browsable(false), System.ComponentModel.DesignerSerializationVisibility(System.ComponentModel.DesignerSerializationVisibility.Hidden)]
        public Blackcancer.ChatUI.Views.ChatActionButton contextButton;
        [System.ComponentModel.Browsable(false), System.ComponentModel.DesignerSerializationVisibility(System.ComponentModel.DesignerSerializationVisibility.Hidden)]
        public System.Windows.Forms.TextBox contextPreview;
        [System.ComponentModel.Browsable(false), System.ComponentModel.DesignerSerializationVisibility(System.ComponentModel.DesignerSerializationVisibility.Hidden)]
        public System.Windows.Forms.Panel conversationPanel;
        [System.ComponentModel.Browsable(false), System.ComponentModel.DesignerSerializationVisibility(System.ComponentModel.DesignerSerializationVisibility.Hidden)]
        public System.Windows.Forms.TableLayoutPanel welcomePanel;
        [System.ComponentModel.Browsable(false), System.ComponentModel.DesignerSerializationVisibility(System.ComponentModel.DesignerSerializationVisibility.Hidden)]
        public System.Windows.Forms.Label welcomeTitle;
        [System.ComponentModel.Browsable(false), System.ComponentModel.DesignerSerializationVisibility(System.ComponentModel.DesignerSerializationVisibility.Hidden)]
        public System.Windows.Forms.Label welcomeBody;
        [System.ComponentModel.Browsable(false), System.ComponentModel.DesignerSerializationVisibility(System.ComponentModel.DesignerSerializationVisibility.Hidden)]
        public System.Windows.Forms.TableLayoutPanel composerLayout;
        [System.ComponentModel.Browsable(false), System.ComponentModel.DesignerSerializationVisibility(System.ComponentModel.DesignerSerializationVisibility.Hidden)]
        public Blackcancer.ChatUI.Views.ChatInputView inputTextBox;
        [System.ComponentModel.Browsable(false), System.ComponentModel.DesignerSerializationVisibility(System.ComponentModel.DesignerSerializationVisibility.Hidden)]
        public System.Windows.Forms.TableLayoutPanel actionsLayout;
        [System.ComponentModel.Browsable(false), System.ComponentModel.DesignerSerializationVisibility(System.ComponentModel.DesignerSerializationVisibility.Hidden)]
        public System.Windows.Forms.Panel sendActionsPanel;
        [System.ComponentModel.Browsable(false), System.ComponentModel.DesignerSerializationVisibility(System.ComponentModel.DesignerSerializationVisibility.Hidden)]
        public System.Windows.Forms.ComboBox modePicker;
        [System.ComponentModel.Browsable(false), System.ComponentModel.DesignerSerializationVisibility(System.ComponentModel.DesignerSerializationVisibility.Hidden)]
        public Blackcancer.ChatUI.Views.ChatActionButton sendButton;
        [System.ComponentModel.Browsable(false), System.ComponentModel.DesignerSerializationVisibility(System.ComponentModel.DesignerSerializationVisibility.Hidden)]
        public Blackcancer.ChatUI.Views.ChatActionButton stopButton;
        [System.ComponentModel.Browsable(false), System.ComponentModel.DesignerSerializationVisibility(System.ComponentModel.DesignerSerializationVisibility.Hidden)]
        public System.Windows.Forms.TableLayoutPanel providerLayout;
        [System.ComponentModel.Browsable(false), System.ComponentModel.DesignerSerializationVisibility(System.ComponentModel.DesignerSerializationVisibility.Hidden)]
        public System.Windows.Forms.ComboBox providerPicker;
        [System.ComponentModel.Browsable(false), System.ComponentModel.DesignerSerializationVisibility(System.ComponentModel.DesignerSerializationVisibility.Hidden)]
        public System.Windows.Forms.ComboBox modelPicker;
        [System.ComponentModel.Browsable(false), System.ComponentModel.DesignerSerializationVisibility(System.ComponentModel.DesignerSerializationVisibility.Hidden)]
        public System.Windows.Forms.TableLayoutPanel footerLayout;
        [System.ComponentModel.Browsable(false), System.ComponentModel.DesignerSerializationVisibility(System.ComponentModel.DesignerSerializationVisibility.Hidden)]
        public Blackcancer.ChatUI.Views.ChatActionButton latestButton;
        [System.ComponentModel.Browsable(false), System.ComponentModel.DesignerSerializationVisibility(System.ComponentModel.DesignerSerializationVisibility.Hidden)]
        public System.Windows.Forms.Label statusLabel;
        [System.ComponentModel.Browsable(false), System.ComponentModel.DesignerSerializationVisibility(System.ComponentModel.DesignerSerializationVisibility.Hidden)]
        public System.Windows.Forms.ProgressBar activityBar;
        [System.ComponentModel.Browsable(false), System.ComponentModel.DesignerSerializationVisibility(System.ComponentModel.DesignerSerializationVisibility.Hidden)]
        public System.Windows.Forms.TextBox historySearchTextBox;
        [System.ComponentModel.Browsable(false), System.ComponentModel.DesignerSerializationVisibility(System.ComponentModel.DesignerSerializationVisibility.Hidden)]
        public System.Windows.Forms.CheckBox showArchivedCheckBox;
        [System.ComponentModel.Browsable(false), System.ComponentModel.DesignerSerializationVisibility(System.ComponentModel.DesignerSerializationVisibility.Hidden)]
        public System.Windows.Forms.TextBox chatTitleEditor;
        [System.ComponentModel.Browsable(false), System.ComponentModel.DesignerSerializationVisibility(System.ComponentModel.DesignerSerializationVisibility.Hidden)]
        public Blackcancer.ChatUI.Views.ChatActionButton renameChatButton;
        [System.ComponentModel.Browsable(false), System.ComponentModel.DesignerSerializationVisibility(System.ComponentModel.DesignerSerializationVisibility.Hidden)]
        public Blackcancer.ChatUI.Views.ChatActionButton archiveChatButton;
        [System.ComponentModel.Browsable(false), System.ComponentModel.DesignerSerializationVisibility(System.ComponentModel.DesignerSerializationVisibility.Hidden)]
        public Blackcancer.ChatUI.Views.ChatActionButton pinChatButton;
        [System.ComponentModel.Browsable(false), System.ComponentModel.DesignerSerializationVisibility(System.ComponentModel.DesignerSerializationVisibility.Hidden)]
        public Blackcancer.ChatUI.Views.ChatActionButton exportChatButton;
        [System.ComponentModel.Browsable(false), System.ComponentModel.DesignerSerializationVisibility(System.ComponentModel.DesignerSerializationVisibility.Hidden)]
        public Blackcancer.ChatUI.Views.ChatActionButton deleteChatButton;
        [System.ComponentModel.Browsable(false), System.ComponentModel.DesignerSerializationVisibility(System.ComponentModel.DesignerSerializationVisibility.Hidden)]
        public System.Windows.Forms.FlowLayoutPanel historyActions;
        [System.ComponentModel.Browsable(false), System.ComponentModel.DesignerSerializationVisibility(System.ComponentModel.DesignerSerializationVisibility.Hidden)]
        public Blackcancer.ChatUI.Views.ChatActionButton memoryToggleButton;
        [System.ComponentModel.Browsable(false), System.ComponentModel.DesignerSerializationVisibility(System.ComponentModel.DesignerSerializationVisibility.Hidden)]
        public System.Windows.Forms.GroupBox memoryPanel;
        [System.ComponentModel.Browsable(false), System.ComponentModel.DesignerSerializationVisibility(System.ComponentModel.DesignerSerializationVisibility.Hidden)]
        public System.Windows.Forms.TableLayoutPanel memoryLayout;
        [System.ComponentModel.Browsable(false), System.ComponentModel.DesignerSerializationVisibility(System.ComponentModel.DesignerSerializationVisibility.Hidden)]
        public System.Windows.Forms.TextBox memoryEditor;
        [System.ComponentModel.Browsable(false), System.ComponentModel.DesignerSerializationVisibility(System.ComponentModel.DesignerSerializationVisibility.Hidden)]
        public Blackcancer.ChatUI.Views.ChatActionButton saveMemoryButton;
        [System.ComponentModel.Browsable(false), System.ComponentModel.DesignerSerializationVisibility(System.ComponentModel.DesignerSerializationVisibility.Hidden)]
        public System.Windows.Forms.CheckBox attachMemoryCheckBox;
        [System.ComponentModel.Browsable(false), System.ComponentModel.DesignerSerializationVisibility(System.ComponentModel.DesignerSerializationVisibility.Hidden)]
        public System.Windows.Forms.FlowLayoutPanel contextChips;
        [System.ComponentModel.Browsable(false), System.ComponentModel.DesignerSerializationVisibility(System.ComponentModel.DesignerSerializationVisibility.Hidden)]
        public Blackcancer.ChatUI.Views.ChatActionButton attachFileButton;
        [System.ComponentModel.Browsable(false), System.ComponentModel.DesignerSerializationVisibility(System.ComponentModel.DesignerSerializationVisibility.Hidden)]
        public Blackcancer.ChatUI.Views.ChatActionButton attachContextButton;
        [System.ComponentModel.Browsable(false), System.ComponentModel.DesignerSerializationVisibility(System.ComponentModel.DesignerSerializationVisibility.Hidden)]
        public Blackcancer.ChatUI.Views.ChatActionButton attachSelectionButton;
        [System.ComponentModel.Browsable(false), System.ComponentModel.DesignerSerializationVisibility(System.ComponentModel.DesignerSerializationVisibility.Hidden)]
        public System.Windows.Forms.ContextMenuStrip cadContextMenu;
        [System.ComponentModel.Browsable(false), System.ComponentModel.DesignerSerializationVisibility(System.ComponentModel.DesignerSerializationVisibility.Hidden)]
        public System.Windows.Forms.ToolStripMenuItem attachDocumentMenuItem;
        [System.ComponentModel.Browsable(false), System.ComponentModel.DesignerSerializationVisibility(System.ComponentModel.DesignerSerializationVisibility.Hidden)]
        public System.Windows.Forms.ToolStripMenuItem attachSelectionMenuItem;
        [System.ComponentModel.Browsable(false), System.ComponentModel.DesignerSerializationVisibility(System.ComponentModel.DesignerSerializationVisibility.Hidden)]
        public System.Windows.Forms.OpenFileDialog fileAttachmentDialog;
        public ChatTranscriptView transcriptPanel;
    }
}



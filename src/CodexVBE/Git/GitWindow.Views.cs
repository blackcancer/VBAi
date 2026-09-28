namespace CodexVBE
{
    /// <summary>Control references from child Git views used by the parent window's shared actions.</summary>
internal sealed partial class GitWindow
    {
        /// <summary>Remote URL label in the connection view.</summary>
private System.Windows.Forms.Label remoteLabel;
        /// <summary>Active branch label in the connection view.</summary>
private System.Windows.Forms.Label branchLabel;
        /// <summary>Connection help text.</summary>
private System.Windows.Forms.Label help;
        /// <summary>Remote repository URL input.</summary>
private System.Windows.Forms.TextBox remote;
        /// <summary>Current branch input.</summary>
private System.Windows.Forms.TextBox branch;
        /// <summary>Connects to the configured repository.</summary>
private System.Windows.Forms.Button connect;
        /// <summary>Commit message caption.</summary>
private System.Windows.Forms.Label messageLabel;
        /// <summary>Commit message input.</summary>
private System.Windows.Forms.TextBox commitMessage;
        /// <summary>Commits selected changes.</summary>
private System.Windows.Forms.Button commit;
        /// <summary>Split layout for the changes list and code comparison.</summary>
private System.Windows.Forms.SplitContainer changeSplit;
        /// <summary>Code comparison for selected changes.</summary>
private CodexVBE.CodeDiffView diff;
        /// <summary>Selectable changed files and modules.</summary>
private System.Windows.Forms.CheckedListBox changes;
        /// <summary>Review actions for the selected change.</summary>
private System.Windows.Forms.FlowLayoutPanel reviewActions;
        /// <summary>Opens the selected VBA module.</summary>
private System.Windows.Forms.Button openModule;
        /// <summary>Restores the selected module version.</summary>
private System.Windows.Forms.Button restoreModule;
        /// <summary>Commit history entries.</summary>
private System.Windows.Forms.ListBox history;
        /// <summary>Details for the selected history entry.</summary>
private System.Windows.Forms.TextBox historyDetails;
        /// <summary>Compares the selected history entry with the current project.</summary>
private System.Windows.Forms.Button historyCompare;
        /// <summary>Actions for creating and switching branches.</summary>
private System.Windows.Forms.FlowLayoutPanel branchActions;
        /// <summary>Local and remote branch entries.</summary>
private System.Windows.Forms.ListBox branchList;
        /// <summary>Branch name input.</summary>
private System.Windows.Forms.ComboBox branchName;
        /// <summary>Lists remote branches.</summary>
private System.Windows.Forms.Button branchRemote;
        /// <summary>Creates a branch.</summary>
private System.Windows.Forms.Button branchCreate;
        /// <summary>Tracks a selected remote branch.</summary>
private System.Windows.Forms.Button branchTrack;
        /// <summary>Switches to the selected branch.</summary>
private System.Windows.Forms.Button branchSwitch;
        /// <summary>Begins a merge for the selected branch.</summary>
private System.Windows.Forms.Button mergeBegin;
        /// <summary>Actions for creating and restoring local checkpoints.</summary>
private System.Windows.Forms.FlowLayoutPanel checkpointActions;
        /// <summary>Saved project checkpoints.</summary>
private System.Windows.Forms.ListBox checkpointList;
        /// <summary>Checkpoint name input.</summary>
private System.Windows.Forms.TextBox checkpointName;
        /// <summary>Creates a checkpoint.</summary>
private System.Windows.Forms.Button checkpointCreate;
        /// <summary>Restores the selected checkpoint.</summary>
private System.Windows.Forms.Button checkpointRestore;
        /// <summary>Actions for reviewing and resolving merge conflicts.</summary>
private System.Windows.Forms.FlowLayoutPanel conflictActions;
        /// <summary>Files with unresolved merge conflicts.</summary>
private System.Windows.Forms.ListBox conflictList;
        /// <summary>Manual merge resolution text.</summary>
private System.Windows.Forms.TextBox resolutionText;
        /// <summary>Chooses the current branch's version for the selected conflict.</summary>
private System.Windows.Forms.Button mergeOurs;
        /// <summary>Chooses the incoming branch's version for the selected conflict.</summary>
private System.Windows.Forms.Button mergeTheirs;
        /// <summary>Applies the text entered as the conflict resolution.</summary>
private System.Windows.Forms.Button mergeText;
        /// <summary>Completes the active merge after all conflicts are resolved.</summary>
private System.Windows.Forms.Button mergeComplete;
        /// <summary>Aborts the active merge.</summary>
private System.Windows.Forms.Button mergeAbort;
        /// <summary>Side-by-side conflict comparison grid.</summary>
private System.Windows.Forms.DataGridView conflictDiff;
        /// <summary>Current branch side of the conflict comparison.</summary>
private System.Windows.Forms.DataGridViewTextBoxColumn conflictOurs;
        /// <summary>Incoming branch side of the conflict comparison.</summary>
private System.Windows.Forms.DataGridViewTextBoxColumn conflictTheirs;
        /// <summary>Common ancestor version of the conflict text.</summary>
private System.Windows.Forms.TextBox baseContent;
        /// <summary>Layout for conflict comparison labels and text.</summary>
private System.Windows.Forms.TableLayoutPanel conflictLayout;
        /// <summary>Common ancestor column caption.</summary>
private System.Windows.Forms.Label ancestorLabel;
        /// <summary>Resolved result column caption.</summary>
private System.Windows.Forms.Label resultLabel;
        /// <summary>Summary of a repository import preview.</summary>
private System.Windows.Forms.TextBox importSummary;
        /// <summary>Binds controls in the active child views to fields used by GitWindow handlers.</summary>
private void BindViews()
        {
            remoteLabel = gitConnectionView.remoteLabel;
            branchLabel = gitConnectionView.branchLabel;
            help = gitConnectionView.help;
            remote = gitConnectionView.remote;
            branch = gitConnectionView.branch;
            connect = gitConnectionView.connect;
            messageLabel = gitChangesView.messageLabel;
            commitMessage = gitChangesView.commitMessage;
            commit = gitChangesView.commit;
            changeSplit = gitChangesView.changeSplit;
            diff = gitChangesView.diff;
            changes = gitChangesView.changes;
            reviewActions = gitChangesView.reviewActions;
            openModule = gitChangesView.openModule;
            restoreModule = gitChangesView.restoreModule;
            history = gitHistoryView.history;
            historyDetails = gitHistoryView.historyDetails;
            historyCompare = gitHistoryView.historyCompare;
            branchActions = gitBranchesView.branchActions;
            branchList = gitBranchesView.branchList;
            branchName = gitBranchesView.branchName;
            branchRemote = gitBranchesView.branchRemote;
            branchCreate = gitBranchesView.branchCreate;
            branchTrack = gitBranchesView.branchTrack;
            branchSwitch = gitBranchesView.branchSwitch;
            mergeBegin = gitBranchesView.mergeBegin;
            checkpointActions = gitCheckpointsView.checkpointActions;
            checkpointList = gitCheckpointsView.checkpointList;
            checkpointName = gitCheckpointsView.checkpointName;
            checkpointCreate = gitCheckpointsView.checkpointCreate;
            checkpointRestore = gitCheckpointsView.checkpointRestore;
            conflictActions = gitConflictsView.conflictActions;
            conflictList = gitConflictsView.conflictList;
            resolutionText = gitConflictsView.resolutionText;
            mergeOurs = gitConflictsView.mergeOurs;
            mergeTheirs = gitConflictsView.mergeTheirs;
            mergeText = gitConflictsView.mergeText;
            mergeComplete = gitConflictsView.mergeComplete;
            mergeAbort = gitConflictsView.mergeAbort;
            conflictDiff = gitConflictsView.conflictDiff;
            conflictOurs = gitConflictsView.conflictOurs;
            conflictTheirs = gitConflictsView.conflictTheirs;
            baseContent = gitConflictsView.baseContent;
            conflictLayout = gitConflictsView.conflictLayout;
            ancestorLabel = gitConflictsView.ancestorLabel;
            resultLabel = gitConflictsView.resultLabel;
            importSummary = gitImportView.importSummary;
            this.connect.Click += new System.EventHandler(this.Connect_Click);
            this.commit.Click += new System.EventHandler(this.Commit_Click);
            this.changes.SelectedIndexChanged += new System.EventHandler(this.Changes_SelectedIndexChanged);
            this.branchCreate.Click += new System.EventHandler(this.BranchCreate_Click);
            this.branchRemote.Click += new System.EventHandler(this.RemoteBranches_Click);
            this.branchTrack.Click += new System.EventHandler(this.BranchTrack_Click);
            this.branchSwitch.Click += new System.EventHandler(this.BranchSwitch_Click);
            this.mergeBegin.Click += new System.EventHandler(this.MergeBegin_Click);
            this.checkpointCreate.Click += new System.EventHandler(this.CheckpointCreate_Click);
            this.checkpointRestore.Click += new System.EventHandler(this.CheckpointRestore_Click);
            this.conflictList.SelectedIndexChanged += new System.EventHandler(this.ConflictList_SelectedIndexChanged);
            this.mergeOurs.Click += new System.EventHandler(this.MergeOurs_Click);
            this.mergeTheirs.Click += new System.EventHandler(this.MergeTheirs_Click);
            this.mergeText.Click += new System.EventHandler(this.MergeText_Click);
            this.mergeComplete.Click += new System.EventHandler(this.MergeComplete_Click);
            this.mergeAbort.Click += new System.EventHandler(this.MergeAbort_Click);
            this.openModule.Click += new System.EventHandler(this.OpenModule_Click);
            this.restoreModule.Click += new System.EventHandler(this.RestoreModule_Click);
            this.history.SelectedIndexChanged += new System.EventHandler(this.HistoryChanged);
            this.checkpointList.SelectedIndexChanged += new System.EventHandler(this.CheckpointChanged);
            this.historyCompare.Click += new System.EventHandler(this.HistoryChanged);
        }
    }
}

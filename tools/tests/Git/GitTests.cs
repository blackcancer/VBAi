using System;
using System.Collections;
using System.Collections.Generic;
using System.ComponentModel.Design;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using CodexVBE;

internal static partial class GitTests
{
    private static int checks;
    private static string root;
    [STAThread]
    private static int Main()
    {
        System.Windows.Forms.Application.EnableVisualStyles();
        // Regression: .NET Framework otherwise prepends this BOM to redirected binary stdin.
        Console.InputEncoding = new UTF8Encoding(true);
        root = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "runs", Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(root);
        try { Snapshots(); Repositories(); ProjectImport(); Workflow(); Advanced(); Designer(); Console.WriteLine("PASS " + checks + " Git checks (local Git + simulated VBE). Real hosts: NOT_RUN."); return 0; }
        catch (Exception ex) { Console.Error.WriteLine(ex); return 1; }
    }
    private static void Assert(bool ok, string message) { checks++; if (!ok) throw new Exception(message); }
    private static void Reject(Action action, string message)
    {
        bool rejected = false; try { action(); } catch { rejected = true; } Assert(rejected, message);
    }
    private static VbaGitSnapshot Snapshot(string value = "1")
    {
        return new VbaGitSnapshot(new VbaGitManifest { References = "test-reference:1:0", Components = new[] {
            new VbaGitComponent { Name = "Module1", Type = 1 }
        } }, new Dictionary<string, byte[]> { { "Module1.bas", Encoding.UTF8.GetBytes("Attribute VB_Name = \"Module1\"\nOption Explicit\nPublic Const Value = " + value + "\n") } });
    }
    private static void Snapshots()
    {
        var one = Snapshot(); Assert(one.SameAs(VbaGitSnapshot.Read(one.Serialize())), "Snapshot round trip");
        Assert(!one.SameAs(Snapshot("2")), "Changed source detected");
        Assert(Snapshot("2").Changes(one).SequenceEqual(new[] { "~ Module1.bas" }), "Changed file list");
        var bad = one.Serialize(); bad["unexpected.txt"] = new byte[0]; Reject(() => VbaGitSnapshot.Read(bad), "Reject unknown files");
        bad = one.Serialize(); bad["Module1.bas"] = Encoding.UTF8.GetBytes("Attribute VB_Name = \"Other\"\n"); Reject(() => VbaGitSnapshot.Read(bad), "Reject renamed native identity");
        bad = one.Serialize(); bad["Module1.bas"] = Encoding.UTF8.GetBytes("Attribute VB_Name = \"Module1\"\n<<<<<<< HEAD\n"); Reject(() => VbaGitSnapshot.Read(bad), "Reject conflict markers");
        Reject(() => new VbaGitSnapshot(new VbaGitManifest { References = "", Components = new[] { new VbaGitComponent { Name = "../escape", Type = 1 } } }, new Dictionary<string, byte[]>()), "Reject traversal");
        Reject(() => MacroGitRepository.ValidateRemote("https://token@github.com/me/test"), "Reject credentials in remote");
        Reject(() => new MacroGitRepository(root, "../../HEAD"), "Reject invalid branch");
        Reject(() => VbaGitSnapshot.ValidateName("CON"), "Reject Windows device name before export");
        Assert(MacroGitRepository.ValidateRemote("https://github.com/me/test.git") == "https://github.com/me/test.git", "GitHub HTTPS remote");
    }
    private static void Repositories()
    {
        string remote = Path.Combine(root, "origin.git"), seed = Path.Combine(root, "seed");
        Git(root, "init --bare \"" + remote + "\"");
        Directory.CreateDirectory(seed); Git(seed, "init -b main"); Identity(seed, false);
        File.WriteAllText(Path.Combine(seed, "README.md"), "Keep unrelated root files.\n");
        Git(seed, "add README.md"); Git(seed, "commit -m Seed"); Git(seed, "remote add origin \"" + remote + "\""); Git(seed, "push origin main");
        var a = Repo("a", remote); string initial = a.Fetch();
        Assert(initial != null && a.Read(initial) == null, "Existing remote without VBA can be linked");
        string first = a.Commit(Snapshot(), initial, "First VBA commit with accents é and multiline\nDetails");
        a.SetRef(a.Head, first); a.SetRef(MacroGitRepository.Baseline, first);
        Assert(a.Read(first).SameAs(Snapshot()), "Git object snapshot round trip");
        a.PrepareRecovery(Snapshot()); a.RecordImportedState(Snapshot()); a.CompleteRecovery();
        a.Push(first);
        Assert(Git(root, "--git-dir=\"" + remote + "\" show main:README.md").Contains("Keep unrelated"), "Unrelated root files preserved");
        Assert(!Git(root, "--git-dir=\"" + remote + "\" for-each-ref").Contains("refs/codex"), "Private backups never pushed");
        Assert(!Directory.Exists(Path.Combine(root, "a", "vba")), "No working source directory");
        var b = Repo("b", remote); string incoming = b.Fetch(); b.SetRef(b.Head, incoming); b.SetRef(MacroGitRepository.Baseline, incoming);
        string second = b.Commit(Snapshot("2"), incoming, "Second"); b.SetRef(b.Head, second); b.Push(second);
        Assert(a.Fetch() == second, "Fetch remote update"); a.RequireFastForward(first, second); checks++;
        string diverged = a.Commit(Snapshot("3"), first, "Diverged");
        Reject(() => a.RequireFastForward(diverged, second), "Reject divergent pull");
        Reject(() => a.Push(diverged), "Reject non fast forward push");
        a.PrepareRecovery(Snapshot()); Assert(a.RecoveryPending, "Durable recovery marker before mutation");
        a.RecordImportedState(Snapshot("2")); Assert(a.Read(a.Resolve(MacroGitRepository.Backup)).SameAs(Snapshot()), "Durable backup survives reopening");
        a.CompleteRecovery(); Assert(!a.RecoveryPending, "Clear recovery only after success");
        Assert(a.History().Length >= 2, "Git history includes original repository");
        Assert(a.SynchronizationStatus().Contains("↓ 1"), "Incoming count");
    }
    private static MacroGitRepository Repo(string name, string remote)
    {
        string path = Path.Combine(root, name); var repo = new MacroGitRepository(path, "main"); repo.Initialize(remote); Identity(path, true); return repo;
    }
    private static void Identity(string path, bool bare)
    {
        string prefix = bare ? "--git-dir=\"" + path + "\" " : "";
        Git(path, prefix + "config user.name Test"); Git(path, prefix + "config user.email test@example.invalid");
    }
    private static string Git(string cwd, string args)
    {
        using (var p = Process.Start(new ProcessStartInfo("git.exe", args) { WorkingDirectory = cwd, UseShellExecute = false,
            CreateNoWindow = true, RedirectStandardOutput = true, RedirectStandardError = true }))
        {
            var stdout = p.StandardOutput.ReadToEndAsync(); var stderr = p.StandardError.ReadToEndAsync();
            if (!p.WaitForExit(30000)) throw new Exception("Git fixture timeout");
            if (p.ExitCode != 0) throw new Exception(stderr.Result); return stdout.Result;
        }
    }
    private static void ProjectImport()
    {
        var host = new FakeProject { FileName = Path.Combine(root, "macro.xlsm") };
        host.VBComponents.Add(new FakeComponent("Module1", 1, "Attribute VB_Name = \"Module1\"\nOption Explicit\nPublic Const Value = 1\n"));
        var document = new FakeComponent("ThisWorkbook", 100, "Option Explicit\nPrivate Sub Workbook_Open()\nEnd Sub"); host.VBComponents.Add(document);
        host.VBComponents.Add(new FakeComponent("Form1", 3, "VERSION 5.00\nBegin VB.UserForm Form1\n   OleObjectBlob = \"Form1.frx\":0000\nEnd\nAttribute VB_Name = \"Form1\"\n") { Resource = new byte[] { 0, 1, 2, 255 } });
        var project = new VbaGitProject(() => host, host.FileName);
        var before = project.Capture(); Assert(before.Files.ContainsKey("Form1.frx"), "Export FRX companion");
        var files = before.Serialize(); files["Module1.bas"] = Encoding.UTF8.GetBytes("Attribute VB_Name = \"Module1\"\nOption Explicit\nPublic Const Value = 2\n");
        files["ThisWorkbook.vba"] = Encoding.UTF8.GetBytes("Option Explicit\n' Modifié é\nPrivate Sub Workbook_Open()\nEnd Sub");
        var target = VbaGitSnapshot.Read(files);
        project.Apply(target, before); Assert(project.Capture().SameAs(target), "Import and readback");
        Assert(object.ReferenceEquals(document, host.VBComponents.Item("ThisWorkbook")), "Host document module updated in place");
        project.Apply(before, target); Assert(project.Capture().SameAs(before), "Rollback source and form resources");
        bool started = false;
        Reject(() => project.Apply(target, target, () => started = true), "Reject stale source before edits");
        Assert(!started, "A stale preflight does not identify unrelated edits as imported state");
        var wrongRefs = new VbaGitSnapshot(new VbaGitManifest { References = "different", Components = before.Manifest.Components }, before.Files);
        Reject(() => project.Apply(wrongRefs, before), "Reject mismatched references");
        Assert(project.Capture().SameAs(before), "Preflight preserves live code");
        host.VBComponents.ThrowAfterImport = true;
        Reject(() => project.Apply(target, before), "Do not retry COM import after applied exception");
        Assert(host.VBComponents.ImportAttempts == 3, "One import per attempt; two successes and one failed call");
        Assert(host.VBComponents.Count(x => x.Name == "Module1") == 1, "No duplicate component after uncertain import");
        host.Mode = 1; Reject(() => project.Capture(), "Reject running project"); host.Mode = 2;
        host.FileName += ".other"; Reject(() => project.Capture(), "Reject changed document identity");
    }
    private static void Designer()
    {
        using (var surface = new DesignSurface(typeof(GitWindow)))
        {
            Assert(surface.IsLoaded && surface.LoadErrors.Count == 0, "GitWindow loads on DesignSurface without runtime services");
            var host = (IDesignerHost)surface.GetService(typeof(IDesignerHost));
            var form = (System.Windows.Forms.Form)host.RootComponent;
            Assert(form.Controls.Find("tabs", true).Length == 1, "Designer tabs");
            Assert(form.Controls.Find("diff", true).Length == 1, "Designer diff grid");
            Assert(form.Controls.Find("fetch", true).Length == 1 && form.Controls.Find("commit", true).Length == 1, "Separate Git actions");
            Assert(form.Controls.Find("branchSwitch", true).Length == 1 && form.Controls.Find("checkpointRestore", true).Length == 1 && form.Controls.Find("conflictDiff", true).Length == 1, "Advanced Git controls available in designer");
        }
        using (var form = new GitWindow())
        {
            var flags = System.Reflection.BindingFlags.Instance | System.Reflection.BindingFlags.NonPublic;
            typeof(GitWindow).GetField("displayedBaseline", flags).SetValue(form, Snapshot());
            typeof(GitWindow).GetField("displayedLive", flags).SetValue(form, Snapshot("2"));
            var files = (System.Windows.Forms.ListBox)form.Controls.Find("changes", true)[0];
            files.Items.Add("~ Module1.bas"); files.SelectedIndex = 0;
            form.Show(); form.Refresh();
            using (var bitmap = new System.Drawing.Bitmap(form.Width, form.Height))
            { form.DrawToBitmap(bitmap, new System.Drawing.Rectangle(0, 0, form.Width, form.Height)); bitmap.Save(Path.Combine(root, "git-window.png")); }
            var split = (System.Windows.Forms.SplitContainer)form.Controls.Find("changeSplit", true)[0];
            Assert(split.Panel2.Width > 550, "Diff has usable width");
            var tabs = (System.Windows.Forms.TabControl)form.Controls.Find("tabs", true)[0];
            foreach (string tab in new[] { "branchesTab", "checkpointsTab", "conflictsTab" })
            {
                var page = (System.Windows.Forms.TabPage)form.Controls.Find(tab, true)[0];
                page.Enabled = true; tabs.SelectedTab = page; form.Refresh();
                using (var bitmap = new System.Drawing.Bitmap(form.Width, form.Height))
                { form.DrawToBitmap(bitmap, new System.Drawing.Rectangle(0, 0, form.Width, form.Height)); bitmap.Save(Path.Combine(root, tab + ".png")); }
            }
            form.Close();
        }
    }

    private static void Workflow()
    {
        string remote = Path.Combine(root, "workflow-origin.git"); Git(root, "init --bare \"" + remote + "\"");
        var a = Repo("workflow-a", remote); var b = Repo("workflow-b", remote);
        var host = new FakeProject { FileName = Path.Combine(root, "workflow.xlsm") };
        host.VBComponents.Add(new FakeComponent("Module1", 1, "Attribute VB_Name = \"Module1\"\nOption Explicit\nPublic Const Value = 1\n"));
        var project = new VbaGitProject(() => host, host.FileName);
        var initial = project.Capture();
        using (var form = new GitWindow())
        {
            Set(form, "project", project); Set(form, "repository", a);
            ((System.Windows.Forms.TextBox)form.Controls.Find("commitMessage", true)[0]).Text = "Initial local commit";
            form.Show();
            Action<string> run = method => {
                typeof(GitWindow).GetMethod(method, System.Reflection.BindingFlags.Instance | System.Reflection.BindingFlags.NonPublic).Invoke(form, new object[] { null, EventArgs.Empty });
                var clock = Stopwatch.StartNew();
                while ((bool)Get(form, "running")) { System.Windows.Forms.Application.DoEvents(); System.Threading.Thread.Sleep(5); if (clock.ElapsedMilliseconds > 30000) throw new Exception("Workflow timeout: " + method); }
            };
            Func<string> status = () => ((System.Windows.Forms.Label)form.Controls.Find("status", true)[0]).Text;
            run("Commit_Click"); Assert(status().StartsWith("Commit local créé"), "UI creates local commit: " + status());
            Assert(a.Fetch() == null, "Commit does not publish");
            run("Push_Click"); Assert(status().StartsWith("Push terminé"), "Separate UI push: " + status());
            string first = b.Fetch();
            var files = initial.Serialize(); files["Module1.bas"] = Encoding.UTF8.GetBytes("Attribute VB_Name = \"Module1\"\nOption Explicit\nPublic Const Value = 2\n");
            var target = VbaGitSnapshot.Read(files);
            string second = b.Commit(target, first, "Remote change"); b.Push(second);
            host.VBComponents.Item("Module1").CodeModule.Text += "' Local uncommitted\n";
            run("Pull_Click"); Assert(status().Contains("modifications locales"), "UI rejects dirty VBA before pull");
            host.VBComponents.Item("Module1").CodeModule.Text = Encoding.UTF8.GetString(initial.Files["Module1.bas"]);
            run("Fetch_Click"); Assert(project.Capture().SameAs(initial), "UI Fetch never imports");
            run("Pull_Click"); Assert(status().StartsWith("Pull et import terminés") && project.Capture().SameAs(target), "UI pull and import: " + status());
            string backupBefore = a.Resolve(MacroGitRepository.Backup);
            run("Pull_Click"); Assert(a.Resolve(MacroGitRepository.Backup) == backupBefore, "No-op pull preserves rollback");
            run("Restore_Click"); Assert(project.Capture().SameAs(initial), "UI rollback restores initial VBA: " + status());
            var list = (System.Windows.Forms.ListBox)form.Controls.Find("changes", true)[0];
            Assert(list.Items.Count > 0, "Restoration is shown as an uncommitted change");
            run("Pull_Click"); Assert(status().Contains("modifications locales"), "Pull does not overwrite local restoration");
            run("Commit_Click"); run("Push_Click"); Assert(a.Read(a.Fetch()).SameAs(initial), "Restoration published as new commit");
            string third = b.Commit(target, b.Fetch(), "Another remote edit"); b.Push(third);
            host.VBComponents.ThrowAfterImport = true;
            run("Pull_Click"); Assert(a.RecoveryPending, "Partial import leaves durable recovery marker");
            run("Commit_Click"); Assert(status().Contains("interrompu"), "Partial import blocks committing");
            host.VBComponents.ThrowAfterImport = false;
            run("Restore_Click"); Assert(!a.RecoveryPending && project.Capture().SameAs(initial), "Restore recovers partial import: " + status());
            form.Close();
        }
    }
    private static object Get(object target, string name) { return target.GetType().GetField(name, System.Reflection.BindingFlags.Instance | System.Reflection.BindingFlags.NonPublic).GetValue(target); }
    private static void Set(object target, string name, object value) { target.GetType().GetField(name, System.Reflection.BindingFlags.Instance | System.Reflection.BindingFlags.NonPublic).SetValue(target, value); }
}

public sealed class FakeProject
{
    public string FileName { get; set; }
    public int Mode { get; set; } = 2;
    public int Protection { get; set; }
    public FakeComponents VBComponents { get; } = new FakeComponents();
    public object[] References { get; } = new object[0];
}
public sealed class FakeComponents : IEnumerable<FakeComponent>
{
    private readonly List<FakeComponent> items = new List<FakeComponent>();
    public bool ThrowAfterImport;
    public int ImportAttempts;
    public void Add(FakeComponent item) { items.Add(item); }
    public FakeComponent Item(string name) { return items.Single(x => x.Name == name); }
    public void Remove(FakeComponent item) { if (item.Type == 100) throw new Exception("Cannot remove host module"); items.Remove(item); }
    public FakeComponent Import(string path)
    {
        ImportAttempts++;
        string text = File.ReadAllText(path, Encoding.GetEncoding(System.Globalization.CultureInfo.CurrentCulture.TextInfo.ANSICodePage));
        string name = Regex.Match(text, "Attribute VB_Name = \"([^\"]+)\"").Groups[1].Value;
        var component = new FakeComponent(name, path.EndsWith(".bas") ? 1 : path.EndsWith(".cls") ? 2 : 3, text);
        string frx = Path.ChangeExtension(path, ".frx"); if (File.Exists(frx)) component.Resource = File.ReadAllBytes(frx);
        Add(component); if (ThrowAfterImport) throw new Exception("Simulated failure after applied import"); return component;
    }
    public IEnumerator<FakeComponent> GetEnumerator() { return items.GetEnumerator(); }
    IEnumerator IEnumerable.GetEnumerator() { return GetEnumerator(); }
}
public sealed class FakeComponent
{
    public string Name { get; set; }
    public int Type { get; set; }
    public FakeCodeModule CodeModule { get; }
    public byte[] Resource;
    public FakeComponent(string name, int type, string text) { Name = name; Type = type; CodeModule = new FakeCodeModule(text); }
    public void Export(string path)
    {
        File.WriteAllText(path, CodeModule.Text, Encoding.GetEncoding(System.Globalization.CultureInfo.CurrentCulture.TextInfo.ANSICodePage));
        if (Resource != null) File.WriteAllBytes(Path.ChangeExtension(path, ".frx"), Resource);
    }
}
public sealed class FakeCodeModule
{
    public string Text;
    public FakeCodeModule(string text) { Text = text.Replace("\r\n", "\n"); }
    public int CountOfLines { get { return Text.Length == 0 ? 0 : Text.Split('\n').Length; } }
    public FakeCodeModule Lines { get { return this; } }
    public string this[int start, int count] { get { return string.Join("\n", Text.Split('\n').Skip(start - 1).Take(count)); } }
    public void DeleteLines(int start, int count) { var lines = Text.Split('\n').ToList(); lines.RemoveRange(start - 1, count); Text = string.Join("\n", lines); }
    public void InsertLines(int start, string text) { var lines = Text.Length == 0 ? new List<string>() : Text.Split('\n').ToList(); lines.InsertRange(start - 1, text.Replace("\r\n", "\n").Split('\n')); Text = string.Join("\n", lines); }
}

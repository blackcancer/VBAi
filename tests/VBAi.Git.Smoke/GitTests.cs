using System;
using System.Collections;
using System.Collections.Generic;
using System.ComponentModel.Design;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using VBAi;

/// <summary>Exécute les scénarios smoke Git avec dépôt local et projet VBE simulé.</summary>
internal static partial class GitTests
{
    /// <summary>Nombre d’assertions validées par la suite.</summary>
    private static int checks;
    /// <summary>Répertoire temporaire des dépôts et documents créés pendant l’exécution.</summary>
    private static string root;
    /// <summary>Point d’entrée du programme smoke et traduit une exception en code de sortie non nul.</summary>
    /// <returns>Zéro si les vérifications réussissent, sinon un.</returns>
    [STAThread]
    private static int Main()
    {
        try { RunSuite(); Console.WriteLine("PASS " + checks + " Git checks (local Git + simulated VBE). Real hosts: NOT_RUN."); return 0; }
        catch (Exception ex) { Console.Error.WriteLine(ex); return 1; }
    }

    /// <summary>Configure WinForms et exécute les scénarios snapshots, dépôts, import, workflow et Designer.</summary>
    internal static void RunSuite()
    {
        checks = 0;
        System.Windows.Forms.Application.EnableVisualStyles();
        var previousContext = System.Threading.SynchronizationContext.Current;
        System.Threading.SynchronizationContext.SetSynchronizationContext(
            new System.Windows.Forms.WindowsFormsSynchronizationContext());
        Exception uiError = null;
        System.Threading.ThreadExceptionEventHandler onUiError = (sender, args) => uiError = args.Exception;
        System.Windows.Forms.Application.ThreadException += onUiError;
        // Regression: .NET Framework otherwise prepends this BOM to redirected binary stdin.
        Console.InputEncoding = new UTF8Encoding(true);
        root = GitScratchDirectory.Create().Root;
        try
        {
            Snapshots(); Repositories(); ProjectImport(); Workflow(); Advanced(); Designer();
            if (uiError != null) throw new Exception("WinForms drawing failed during Git tests.", uiError);
        }
        finally
        {
            System.Windows.Forms.Application.ThreadException -= onUiError;
            System.Threading.SynchronizationContext.SetSynchronizationContext(previousContext);
        }
    }
    /// <summary>Incrémente le compteur puis échoue avec le message si la condition est fausse.</summary>
    /// <param name="ok">Condition attendue.</param>
    /// <param name="message">Message d’échec associé.</param>
    private static void Assert(bool ok, string message) { checks++; if (!ok) throw new Exception(message); }
    /// <summary>Vérifie qu’une action lève une exception.</summary>
    /// <param name="action">Action exécutée.</param>
    /// <param name="message">Message d’échec si aucune exception n’est levée.</param>
    private static void Reject(Action action, string message)
    {
        bool rejected = false; try { action(); } catch { rejected = true; }
        Assert(rejected, message);
    }
    /// <summary>Crée un snapshot de test à un module dont la constante peut être changée.</summary>
    /// <param name="value">Valeur écrite dans la constante VBA.</param>
    /// <returns>Snapshot minimal du module Module1.</returns>
    private static VbaGitSnapshot Snapshot(string value = "1")
    {
        return new VbaGitSnapshot(new VbaGitManifest
        {
            References = "test-reference:1:0",
            Components = new[] {
            new VbaGitComponent { Name = "Module1", Type = 1 }
        }
        }, new Dictionary<string, byte[]> { { "Module1.bas", Encoding.UTF8.GetBytes("Attribute VB_Name = \"Module1\"\nOption Explicit\nPublic Const Value = " + value + "\n") } });
    }
    /// <summary>Vérifie la sérialisation, les différences et le rejet des snapshots mal formés.</summary>
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
    /// <summary>Teste les commits locaux, fetch, push, fast-forward et marqueurs de récupération durables.</summary>
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
    /// <summary>Initialise un dépôt de test et configure son identité Git.</summary>
    /// <param name="name">Nom du répertoire local sous la racine temporaire.</param>
    /// <param name="remote">Chemin du dépôt distant bare.</param>
    /// <returns>Dépôt initialisé sur la branche main.</returns>
    private static MacroGitRepository Repo(string name, string remote)
    {
        string path = Path.Combine(root, name); var repo = new MacroGitRepository(path, "main"); repo.Initialize(remote); Identity(path, true); return repo;
    }
    /// <summary>Configure le nom et l’adresse e-mail Git pour un dépôt bare ou de travail.</summary>
    /// <param name="path">Chemin du dépôt.</param>
    /// <param name="bare"><see langword="true"/> si le dépôt est bare.</param>
    private static void Identity(string path, bool bare)
    {
        string prefix = bare ? "--git-dir=\"" + path + "\" " : "";
        Git(path, prefix + "config user.name Test"); Git(path, prefix + "config user.email test@example.invalid");
    }
    /// <summary>Exécute git.exe dans le répertoire demandé avec une limite de trente secondes.</summary>
    /// <param name="cwd">Répertoire de travail du processus.</param>
    /// <param name="args">Arguments Git.</param>
    /// <returns>Sortie standard.</returns>
    /// <exception cref="Exception">Git échoue ou dépasse le délai d’attente.</exception>
    private static string Git(string cwd, string args)
    {
        using (var p = Process.Start(new ProcessStartInfo("git.exe", args)
        {
            WorkingDirectory = cwd,
            UseShellExecute = false,
            CreateNoWindow = true,
            RedirectStandardOutput = true,
            RedirectStandardError = true
        }))
        {
            var stdout = p.StandardOutput.ReadToEndAsync(); var stderr = p.StandardError.ReadToEndAsync();
            if (!p.WaitForExit(30000)) throw new Exception("Git fixture timeout");
            if (p.ExitCode != 0) throw new Exception(stderr.Result); return stdout.Result;
        }
    }
    /// <summary>Vérifie capture, import, restauration et refus des changements de projet incompatibles.</summary>
    private static void ProjectImport()
    {
        var host = new FakeProject { FileName = Path.Combine(root, "macro.xlsm") };
        host.VBComponents.Add(new FakeComponent("Module1", 1, "Attribute VB_Name = \"Module1\"\nOption Explicit\nPublic Const Value = 1\n"));
        var document = new FakeComponent("ThisWorkbook", 100, "Option Explicit\nPrivate Sub Workbook_Open()\nEnd Sub"); host.VBComponents.Add(document);
        host.VBComponents.Add(new FakeComponent("Form1", 3, "VERSION 5.00\nBegin VB.UserForm Form1\n   OleObjectBlob = \"Form1.frx\":0000\nEnd\nAttribute VB_Name = \"Form1\"\n") { Resource = SyntheticFormResource() });
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

    /// <summary>Builds a bounded LB/08 and MS-CFB v3 resource containing an empty synthetic UserForm.</summary>
    private static byte[] SyntheticFormResource()
    {
        const uint end = 0xfffffffe, free = 0xffffffff;
        var resource = new byte[24 + 5 * 512];
        Action<int, uint> write = (offset, value) => Buffer.BlockCopy(BitConverter.GetBytes(value), 0, resource, offset, 4);
        resource[0] = 0x4c; resource[1] = 0x42; resource[2] = 8;
        write(4, 5 * 512);
        byte[] signature = { 0xd0, 0xcf, 0x11, 0xe0, 0xa1, 0xb1, 0x1a, 0xe1 };
        Buffer.BlockCopy(signature, 0, resource, 24, signature.Length);
        resource[48] = 0x3e; resource[50] = 3; resource[52] = 0xfe; resource[53] = 0xff;
        resource[54] = 9; resource[56] = 6;
        write(24 + 44, 1); write(24 + 48, 1); write(24 + 56, 4096);
        write(24 + 60, 2); write(24 + 64, 1); write(24 + 68, end);
        for (int i = 0; i < 109; i++) write(24 + 76 + i * 4, i == 0 ? 0u : free);
        for (int i = 0; i < 128; i++)
        {
            write(24 + 512 + i * 4, i == 0 ? 0xfffffffd : i < 4 ? end : free);
            write(24 + 1536 + i * 4, i == 0 ? end : free);
        }
        Action<int, string, byte, uint, uint, uint, uint> entry = (index, name, kind, child, right, start, size) =>
        {
            int position = 24 + 1024 + index * 128;
            byte[] text = Encoding.Unicode.GetBytes(name + "\0");
            Buffer.BlockCopy(text, 0, resource, position, text.Length);
            resource[position + 64] = (byte)text.Length; resource[position + 66] = kind; resource[position + 67] = 1;
            write(position + 68, free); write(position + 72, right); write(position + 76, child);
            write(position + 116, start); write(position + 120, size);
        };
        entry(0, "Root Entry", 5, 1, free, 3, 64);
        byte[] formClass = new Guid("C62A69F0-16DC-11CE-9E98-00AA00574A4F").ToByteArray();
        Buffer.BlockCopy(formClass, 0, resource, 24 + 1024 + 80, formClass.Length);
        entry(1, "f", 2, free, 2, 0, 22);
        entry(2, "o", 2, free, free, end, 0);
        resource[24 + 1024 + 2 * 128 + 67] = 0; // Red right child of the black f tree root.
        // FormControl v4, required DrawBuffer, empty class table and zero sites.
        byte[] form = { 0, 4, 8, 0, 0, 0, 0, 8, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0 };
        Buffer.BlockCopy(form, 0, resource, 24 + 2048, form.Length);
        return resource;
    }
    /// <summary>Charge GitWindow dans le Designer et vérifie le rendu et les commandes visuelles.</summary>
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
            typeof(GitWindow).GetMethod("PopulateChanges", flags).Invoke(form, new object[] { Snapshot("2"), Snapshot() });
            var files = (System.Windows.Forms.ListBox)form.Controls.Find("changes", true)[0];
            Assert(files.Items.Count == 1, "Module changes displayed as a selectable commit unit"); files.SelectedIndex = 0;
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

    /// <summary>Parcourt les actions de l’interface, du commit local à la restauration après import partiel.</summary>
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
            Action<string> run = method =>
            {
                Console.WriteLine("Git UI: " + method);
                typeof(GitWindow).GetMethod(method, System.Reflection.BindingFlags.Instance | System.Reflection.BindingFlags.NonPublic).Invoke(form, new object[] { null, EventArgs.Empty });
                var clock = Stopwatch.StartNew();
                while ((bool)Get(form, "running"))
                {
                    if (clock.ElapsedMilliseconds > 30000) throw new Exception("Workflow timeout: " + method);
                    System.Windows.Forms.Application.DoEvents(); System.Threading.Thread.Sleep(5);
                }
            };
            Func<string> status = () => ((System.Windows.Forms.Label)Get(form, "status")).Text;
            run("Compare_Click");
            run("Commit_Click"); Assert(status() == UiText.Get("Local commit created. Use Push to publish it."), "UI creates local commit: " + status());
            Assert(a.Fetch() == null, "Commit does not publish");
            run("Push_Click"); Assert(status() == UiText.Get("Push complete."), "Separate UI push: " + status());
            string first = b.Fetch();
            var files = initial.Serialize(); files["Module1.bas"] = Encoding.UTF8.GetBytes("Attribute VB_Name = \"Module1\"\nOption Explicit\nPublic Const Value = 2\n");
            var target = VbaGitSnapshot.Read(files);
            string second = b.Commit(target, first, "Remote change"); b.Push(second);
            host.VBComponents.Item("Module1").CodeModule.Text += "' Local uncommitted\n";
            run("Pull_Click"); Assert(status().StartsWith(UiText.Get("VBA contains uncommitted local changes."), StringComparison.Ordinal), "UI rejects dirty VBA before pull");
            host.VBComponents.Item("Module1").CodeModule.Text = Encoding.UTF8.GetString(initial.Files["Module1.bas"]);
            run("Fetch_Click"); Assert(project.Capture().SameAs(initial), "UI Fetch never imports");
            run("Pull_Click"); Assert(status() == UiText.Get("Pull and import complete. Check and save the document.") && project.Capture().SameAs(target), "UI pull and import: " + status());
            string backupBefore = a.Resolve(MacroGitRepository.Backup);
            run("Pull_Click"); Assert(a.Resolve(MacroGitRepository.Backup) == backupBefore, "No-op pull preserves rollback");
            run("Restore_Click"); Assert(project.Capture().SameAs(initial), "UI rollback restores initial VBA: " + status());
            var list = (System.Windows.Forms.ListBox)form.Controls.Find("changes", true)[0];
            Assert(list.Items.Count > 0, "Restoration is shown as an uncommitted change");
            run("Pull_Click"); Assert(status().StartsWith(UiText.Get("VBA contains uncommitted local changes."), StringComparison.Ordinal), "Pull does not overwrite local restoration");
            run("Commit_Click"); run("Push_Click"); Assert(a.Read(a.Fetch()).SameAs(initial), "Restoration published as new commit");
            string third = b.Commit(target, b.Fetch(), "Another remote edit"); b.Push(third);
            host.VBComponents.ThrowAfterImport = true;
            run("Pull_Click"); Assert(a.RecoveryPending, "Partial import leaves durable recovery marker");
            run("Commit_Click"); Assert(status().StartsWith(UiText.Get("Restore the interrupted import before continuing."), StringComparison.Ordinal), "Partial import blocks committing");
            host.VBComponents.ThrowAfterImport = false;
            run("Restore_Click"); Assert(!a.RecoveryPending && project.Capture().SameAs(initial), "Restore recovers partial import: " + status());
            form.Close();
        }
    }
    /// <summary>Lit un champ d’instance non public par réflexion dans le fixture.</summary>
    /// <param name="target">Objet contenant le champ.</param>
    /// <param name="name">Nom du champ.</param>
    /// <returns>Valeur du champ.</returns>
    private static object Get(object target, string name) { return target.GetType().GetField(name, System.Reflection.BindingFlags.Instance | System.Reflection.BindingFlags.NonPublic).GetValue(target); }
    /// <summary>Écrit un champ d’instance non public par réflexion dans le fixture.</summary>
    /// <param name="target">Objet contenant le champ.</param>
    /// <param name="name">Nom du champ.</param>
    /// <param name="value">Nouvelle valeur du champ.</param>
    private static void Set(object target, string name, object value) { target.GetType().GetField(name, System.Reflection.BindingFlags.Instance | System.Reflection.BindingFlags.NonPublic).SetValue(target, value); }
}

/// <summary>Projet VBE factice exposant les propriétés lues et modifiées par l’adaptateur.</summary>
public sealed class FakeProject
{
    /// <summary>Chemin du document hôte simulé.</summary>
    /// <value>Chemin du document hôte simulé.</value>
    public string FileName { get; set; }
    /// <summary>Mode d’exécution du projet, avec deux comme mode création par défaut.</summary>
    /// <value>Mode d’exécution courant.</value>
    public int Mode { get; set; } = 2;
    /// <summary>État de protection du projet.</summary>
    /// <value>Valeur de protection configurée.</value>
    public int Protection { get; set; }
    /// <summary>Collection ordonnée des composants VBA simulés.</summary>
    /// <value>Collection des composants du projet.</value>
    public FakeComponents VBComponents { get; } = new FakeComponents();
    /// <summary>Références du projet simulé.</summary>
    /// <value>Références exposées par le projet factice.</value>
    public object[] References { get; } = new object[0];
}
/// <summary>Collection mutable de composants factices compatible avec l’énumération VBE.</summary>
public sealed class FakeComponents : IEnumerable<FakeComponent>
{
    /// <summary>Composants actuellement présents dans la collection.</summary>
    private readonly List<FakeComponent> items = new List<FakeComponent>();
    /// <summary>Provoque une exception après l’ajout d’un composant importé.</summary>
    public bool ThrowAfterImport;
    /// <summary>Nombre de tentatives d’import effectuées.</summary>
    public int ImportAttempts;
    /// <summary>Ajoute un composant à la collection.</summary>
    /// <param name="item">Composant à ajouter.</param>
    public void Add(FakeComponent item) { items.Add(item); }
    /// <summary>Retourne le composant portant le nom demandé.</summary>
    /// <param name="name">Nom du composant.</param>
    /// <returns>Composant correspondant.</returns>
    public FakeComponent Item(string name) { return items.Single(x => x.Name == name); }
    /// <summary>Retire le composant, sauf le module document hôte de type 100.</summary>
    /// <param name="item">Composant à retirer.</param>
    public void Remove(FakeComponent item) { if (item.Type == 100) throw new Exception("Cannot remove host module"); items.Remove(item); }
    /// <summary>Importe un fichier de composant et son éventuel fichier de ressources FRX.</summary>
    /// <param name="path">Chemin du fichier exporté.</param>
    /// <returns>Composant construit à partir du fichier.</returns>
    public FakeComponent Import(string path)
    {
        ImportAttempts++;
        string text = File.ReadAllText(path, Encoding.GetEncoding(System.Globalization.CultureInfo.CurrentCulture.TextInfo.ANSICodePage));
        string name = Regex.Match(text, "Attribute VB_Name = \"([^\"]+)\"").Groups[1].Value;
        var component = new FakeComponent(name, path.EndsWith(".bas") ? 1 : path.EndsWith(".cls") ? 2 : 3, text);
        string frx = Path.ChangeExtension(path, ".frx"); if (File.Exists(frx)) component.Resource = File.ReadAllBytes(frx);
        Add(component); if (ThrowAfterImport) throw new Exception("Simulated failure after applied import"); return component;
    }
    /// <summary>Crée l’énumérateur générique de la collection.</summary>
    /// <returns>Énumérateur des composants.</returns>
    public IEnumerator<FakeComponent> GetEnumerator() { return items.GetEnumerator(); }
    /// <summary>Crée l’énumérateur non générique de la collection.</summary>
    /// <returns>Énumérateur des composants.</returns>
    IEnumerator IEnumerable.GetEnumerator() { return GetEnumerator(); }
}
/// <summary>Composant VBA simulé avec module de code et ressources facultatives.</summary>
public sealed class FakeComponent
{
    /// <summary>Nom du composant.</summary>
    /// <value>Nom utilisé pour l’identité du composant.</value>
    public string Name { get; set; }
    /// <summary>Type selon les constantes VBE.</summary>
    /// <value>Type VBE du composant.</value>
    public int Type { get; set; }
    /// <summary>Module de code associé au composant.</summary>
    /// <value>Module de code factice.</value>
    public FakeCodeModule CodeModule { get; }
    /// <summary>Octets du fichier FRX facultatif.</summary>
    public byte[] Resource;
    /// <summary>Crée un composant simulé avec son nom, son type et son code.</summary>
    /// <param name="name">Nom du composant.</param>
    /// <param name="type">Type VBE du composant.</param>
    /// <param name="text">Texte initial du module.</param>
    public FakeComponent(string name, int type, string text) { Name = name; Type = type; CodeModule = new FakeCodeModule(text); }
    /// <summary>Écrit le code et, s’il existe, le fichier de ressources FRX compagnon.</summary>
    /// <param name="path">Chemin d’export du composant.</param>
    public void Export(string path)
    {
        File.WriteAllText(path, CodeModule.Text, Encoding.GetEncoding(System.Globalization.CultureInfo.CurrentCulture.TextInfo.ANSICodePage));
        if (Resource != null) File.WriteAllBytes(Path.ChangeExtension(path, ".frx"), Resource);
    }
}
/// <summary>Module de code simulé avec opérations de lignes utilisées lors des imports.</summary>
public sealed class FakeCodeModule
{
    /// <summary>Texte complet du module avec des séparateurs LF.</summary>
    public string Text;
    /// <summary>Crée le module en normalisant ses fins de ligne en LF.</summary>
    /// <param name="text">Texte initial du module.</param>
    public FakeCodeModule(string text) { Text = text.Replace("\r\n", "\n"); }
    /// <summary>Nombre de lignes du module, ou zéro si son texte est vide.</summary>
    /// <value>Nombre de lignes séparées par LF.</value>
    public int CountOfLines { get { return Text.Length == 0 ? 0 : Text.Split('\n').Length; } }
    /// <summary>Retourne le module lui-même pour simuler l’objet <c>Lines</c> du VBE.</summary>
    /// <value>Instance courante du module.</value>
    public FakeCodeModule Lines { get { return this; } }
    /// <summary>Lit un segment de lignes, avec indexation à partir de un.</summary>
    /// <param name="start">Première ligne à lire.</param>
    /// <param name="count">Nombre maximal de lignes.</param>
    /// <value>Texte des lignes demandées, joint par LF.</value>
    public string this[int start, int count] { get { return string.Join("\n", Text.Split('\n').Skip(start - 1).Take(count)); } }
    /// <summary>Supprime un segment de lignes, avec indexation à partir de un.</summary>
    /// <param name="start">Première ligne à supprimer.</param>
    /// <param name="count">Nombre de lignes à supprimer.</param>
    public void DeleteLines(int start, int count) { var lines = Text.Split('\n').ToList(); lines.RemoveRange(start - 1, count); Text = string.Join("\n", lines); }
    /// <summary>Insère les lignes fournies avant la position spécifiée.</summary>
    /// <param name="start">Position d’insertion indexée à partir de un.</param>
    /// <param name="text">Lignes à insérer, séparées par CRLF ou LF.</param>
    public void InsertLines(int start, string text) { var lines = Text.Length == 0 ? new List<string>() : Text.Split('\n').ToList(); lines.InsertRange(start - 1, text.Replace("\r\n", "\n").Split('\n')); Text = string.Join("\n", lines); }
}

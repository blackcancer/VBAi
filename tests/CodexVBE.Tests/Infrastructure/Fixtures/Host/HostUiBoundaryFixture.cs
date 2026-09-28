using System;
using System.Collections;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;
using CodexVBE;
using CodexVBE.Tests.Unit;
using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace CodexVBE.Tests.Infrastructure
{
    internal sealed class HostUiScope : IDisposable
    {
        private readonly LlmBoundaryScope boundaries = new LlmBoundaryScope();
        private readonly Dictionary<FieldInfo, object> defaults = new Dictionary<FieldInfo, object>();
        private readonly SynchronizationContext context = SynchronizationContext.Current;
        private readonly string originalStorage = LlmSettings.StoragePathOverride;
        internal readonly List<string> Logs = new List<string>();
        internal readonly List<string> Notices = new List<string>();
        internal readonly List<Type> Dialogs = new List<Type>();
        internal readonly NativeHost Host;
        internal readonly LlmSettings Settings = new LlmSettings { ProviderName = "Ollama" };
        internal readonly string Cache;
        internal HostUiScope()
        {
            foreach (var type in new[] { typeof(AddIn), typeof(ChatWindow) })
                foreach (var field in type.GetFields(BindingFlags.Static | BindingFlags.NonPublic).Where(x => typeof(Delegate).IsAssignableFrom(x.FieldType))) defaults[field] = field.GetValue(null);
            SynchronizationContext.SetSynchronizationContext(new System.Windows.Threading.DispatcherSynchronizationContext());
            Host = new NativeHost(Path.Combine(boundaries.Root, "project.xlsm"));
            Cache = MacroGitRepository.ScopeDirectory(Host.Project.FileName); Assert.IsFalse(Directory.Exists(Cache), "This fixture must own a fresh cache scope.");
            LlmSettings.StoragePathOverride = Path.Combine(boundaries.Root, "settings.json");
            AddIn.WriteLog = Logs.Add; AddIn.StartBridge = bridge => { };
            AddIn.ReadSettings = () => Settings;
            AddIn.StartUpdateCheck = () => { };
            AddIn.StopUpdateCheck = () => { };
            AddIn.CreateCrashReporter = show => new CrashReporter(show, directory: Path.Combine(boundaries.Root, "CrashReports"));
            AddIn.ShowNotice = (text, title, buttons, icon) => { Notices.Add(text); return DialogResult.OK; };
            AddIn.ShowModal = (form, owner) => { Dialogs.Add(form.GetType()); Assert.AreEqual(Host.MainWindow.HWnd, owner.Handle.ToInt64()); return DialogResult.Cancel; };
            ChatWindow.ReadSettings = () => Settings; ChatWindow.WriteSettings = value => { };
            ChatWindow.HistoryPath = () => Path.Combine(boundaries.Root, "chat.db");
            ChatWindow.ReadModelCatalogue = (provider, settings) => Task.FromResult(new LlmModelOption[0]);
            ChatWindow.ReadHost = (session, request) => Response.Success(request.Command == "list_projects" ? (object)new object[0] : new { });
            ChatWindow.ShowModal = (form, owner) => { Dialogs.Add(form.GetType()); return DialogResult.Cancel; };
            ChatWindow.ShowNotice = (owner, text, title, buttons, icon) => DialogResult.OK;
        }
        internal AddIn Connected(object addIn = null)
        {
            var instance = new AddIn(); object[] custom = null; instance.OnConnection(Host, 0, addIn ?? Host.AddIns.AddIn, ref custom); return instance;
        }
        internal void Close(AddIn instance) { object[] custom = null; instance.OnDisconnection(0, ref custom); }
        public void Dispose()
        {
            Host.Dispose();
            foreach (var pair in defaults) pair.Key.SetValue(null, pair.Value);
            LlmSettings.StoragePathOverride = originalStorage; SynchronizationContext.SetSynchronizationContext(context);
            if (Directory.Exists(Cache)) Directory.Delete(Cache, true);
            boundaries.Dispose();
        }
    }
    public sealed class NativeHost : IDisposable
    {
        public NativeMainWindow MainWindow { get; }
        public NativeAddIns AddIns { get; } = new NativeAddIns();
        public object ActiveWindow { get; set; }
        public object ActiveCodePane { get; set; }
        public NativeWindows Windows { get; }
        public List<VbeSessionTests.FakeProject> VBProjects { get; } = new List<VbeSessionTests.FakeProject>();
        public VbeSessionTests.FakeProject ActiveVBProject { get; set; }
        public VbeSessionTests.FakeProject Project { get; }
        public VbeMenuLifecycleTests.FakeBar[] CommandBars { get; }
        private readonly Form owner;
        internal Form Owner => owner;
        internal void Workspace(bool available = true) { owner.IsMdiContainer = available; owner.ClientSize = new System.Drawing.Size(1000, 720); MainWindow.HWnd = owner.Handle.ToInt64(); owner.Show(); }
        public NativeHost(string path)
        {
            owner = new Form { Left = -10000, Top = -10000, ShowInTaskbar = false };
            MainWindow = new NativeMainWindow { HWnd = owner.Handle.ToInt64() };
            Windows = new NativeWindows();
            Project = new VbeSessionTests.FakeProject { Name = "P", FileName = path, Mode = 2 }; VBProjects.Add(Project); ActiveVBProject = Project;
            var controls = new VbeMenuLifecycleTests.FakeControls(); controls.Items.Add(new VbeMenuLifecycleTests.FakeButton { Caption = "View" }); controls.Items.Add(new VbeMenuLifecycleTests.FakeButton { Caption = "Tools" });
            CommandBars = new[] { new VbeMenuLifecycleTests.FakeBar { Type = 1, Controls = controls } };
        }
        public void Dispose() { Windows.Dispose(); owner.Dispose(); }
    }
    public sealed class NativeMainWindow
    {
        private long handle;
        public bool RejectHandle;
        public long HWnd { get { if (RejectHandle) throw new IOException("owner unavailable"); return handle; } set { handle = value; } }
        public NativeLinkedWindows LinkedWindows { get; } = new NativeLinkedWindows();
    }
    public sealed class NativeLinkedWindows
    {
        public bool Reject;
        public int Adds;
        public int Removes;
        public void Add(object window) { if (Reject) throw new IOException("position rejected"); Adds++; }
        public void Remove(object window) { if (Reject) throw new IOException("position rejected"); Removes++; }
    }
    public sealed class NativeAddIn : IVbeAddIn { public string ProgId { get; set; } = "CodexVBE.AddIn"; }
    public sealed class NativeAddIns
    {
        public NativeAddIn AddIn { get; } = new NativeAddIn();
        public bool Reject;
        public object Item(string id) { if (Reject) throw new IOException("lookup rejected"); Assert.AreEqual("CodexVBE.AddIn", id); return AddIn; }
    }
    public sealed class NativeWindow : IDisposable
    {
        internal readonly Form Form = new Form { Left = -10000, Top = -10000, ShowInTaskbar = false, ClientSize = new System.Drawing.Size(520, 760) };
        public Action Focusing;
        public bool RejectClose;
        public int FocusCount, CloseCount;
        public object LinkedWindowFrame { get; set; }
        public int Width { get { return Form.Width; } set { Form.Width = value; } }
        public int Height { get { return Form.Height; } set { Form.Height = value; } }
        public int Left { get { return Form.Left; } set { Form.Left = value; } }
        public int Top { get { return Form.Top; } set { Form.Top = value; } }
        public bool Visible { get { return Form.Visible; } set { if (value) Form.Show(); else Form.Hide(); } }
        public void SetFocus() { FocusCount++; Focusing?.Invoke(); }
        public void Close() { CloseCount++; if (RejectClose) throw new IOException("close rejected"); Form.Close(); }
        public void Dispose() { Form.Dispose(); }
    }
    public sealed class NativeWindows : IVbeWindows, IDisposable
    {
        internal NativeWindow Window;
        internal ChatToolWindow Control;
        internal NativeWindow EditorWindow;
        internal ChatToolWindow EditorControl;
        internal bool MissingControl;
        internal bool RejectCreation;
        internal Action AfterCreation;
        public object VBE => null;
        public object Parent => null;
        public int Count => 0;
        public object Item(object index) => null;
        public IEnumerator GetEnumerator() { return new object[0].GetEnumerator(); }
        object IVbeWindows.CreateToolWindow(IVbeAddIn addIn, string progId, string caption, string position, ref object document)
        {
            Assert.IsNotNull(addIn); Assert.AreEqual("CodexVBE.ChatToolWindow", progId);
            if (caption == UiText.Get("VBAi editor"))
            {
                StringAssert.Contains(position, "CC57B0DE");
                if (RejectCreation) throw new IOException("creation rejected");
                EditorWindow?.Dispose(); EditorWindow = new NativeWindow();
                EditorControl = MissingControl ? null : new ChatToolWindow();
                if (EditorControl != null) { EditorWindow.Form.Controls.Add(EditorControl); var handle = EditorControl.Handle; }
                document = EditorControl; AfterCreation?.Invoke(); return EditorWindow;
            }
            Assert.AreEqual("VBAi", caption); StringAssert.Contains(position, "B5C96ED5");
            if (RejectCreation) throw new IOException("creation rejected");
            Window?.Dispose(); Window = new NativeWindow();
            Control = MissingControl ? null : new ChatToolWindow();
            if (Control != null) { Window.Form.Controls.Add(Control); var handle = Control.Handle; }
            document = Control; AfterCreation?.Invoke(); return Window;
        }
        public void Dispose() { Window?.Dispose(); Control?.Dispose(); EditorWindow?.Dispose(); EditorControl?.Dispose(); }
    }
}

namespace CodexVBE.Tests.Unit
{
    public sealed partial class AddInCoverageTests
    {
        private static void Call(AddIn instance, string method) { CodexVBE.Tests.Infrastructure.LlmBoundaryScope.Call(instance, method); }
        private static ChatWindow Chat(AddIn instance) { return CodexVBE.Tests.Infrastructure.LlmBoundaryScope.Get<ChatWindow>(instance, "chat"); }
    }
}

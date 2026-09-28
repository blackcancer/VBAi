namespace CodexVBE.Tests.Unit
{
    using System;
    using System.Collections.Generic;
    using System.Linq;
    using System.Runtime.InteropServices;
    using System.Security.Cryptography;
    using System.Text;
    using CodexVBE;

    // VBIDE contract objects are constructed normally; actions model native lifecycle changes.
    public sealed class EditorDebugFixture
    {
        internal readonly Host Vbe = new Host();
        internal readonly DebugProject Project;
        internal readonly DebugModule Module;
        internal readonly DebugPane Pane;
        internal readonly VbeDebug Service;
        internal EditorDebugFixture()
        {
            Project = new DebugProject(); Vbe.VBProjects.Add(Project); Vbe.ActiveVBProject = Project;
            Module = AddModule("Module1"); Pane = Module.CodePane;
            Vbe.ActiveCodePane = Pane; Vbe.ActiveWindow = Pane.Window;
            Service = new VbeDebug(Vbe);
        }
        internal DebugModule AddModule(string name, string code = "Debug.Print 1\r\nDebug.Print 2")
        {
            var component = new Component { Name = name, Collection = Project.VBComponents };
            var module = new DebugModule { Parent = component, Code = code };
            component.CodeModule = module;
            module.CodePane = new DebugPane { CodeModule = module, Window = new Window { Caption = name + " (Code)" } };
            module.CodePane.OnShow = () => Vbe.ActiveCodePane = module.CodePane;
            Project.VBComponents.Add(component); Vbe.Panes.Add(module.CodePane); Vbe.Windows.Add(module.CodePane.Window);
            return module;
        }
        internal Request Location(string action = null)
        {
            return new Request { Project = Project.Name, Module = Module.Parent.Name, StartLine = 1,
                ExpectedSha256 = Hash(Module.Code), ExpectedMode = Project.Mode, Action = action, ControlCaption = action };
        }
        internal Request Inspect(string action = null)
        {
            var request = Location(action);
            dynamic layout = Service.CodePaneLayout(request);
            dynamic row = ((System.Collections.IEnumerable)layout.Panes).Cast<object>().First();
            request.Pane = row.Pane; request.ExpectedWindowVersion = row.WindowVersion;
            return request;
        }
        internal Control Command(int id, string caption, Action execute = null)
        {
            var control = new Control { Id = id, Caption = caption, OnExecute = execute };
            Vbe.CommandBars[0].Controls.Add(control); return control;
        }
        internal static string Hash(string code)
        {
            using (var sha = SHA256.Create()) return BitConverter.ToString(sha.ComputeHash(Encoding.UTF8.GetBytes(code))).Replace("-", "").ToLowerInvariant();
        }
        public sealed class Host
        {
            public List<DebugProject> VBProjects { get; } = new List<DebugProject>();
            public List<object> Panes { get; } = new List<object>();
            public Action OnReadPanes { get; set; }
            public List<object> CodePanes { get { OnReadPanes?.Invoke(); return Panes; } }
            public List<Window> Windows { get; } = new List<Window>();
            public Bars CommandBars { get; } = new Bars { new Bar() };
            public DebugProject ActiveVBProject { get; set; }
            public object ActiveCodePane { get; set; }
            public Window ActiveWindow { get; set; }
            public Window MainWindow { get; } = new Window { HWnd = 12345 };
        }
        public sealed class DebugProject
        {
            public string Name { get; set; } = "CoverageProject";
            public bool FailPath { get; set; }
            private string path = @"C:\Coverage\Debug.xlsm";
            public string FileName { get { if (FailPath) throw new COMException("unsaved project"); return path; } set { path = value; } }
            public int Mode { get; set; } = 2;
            public Components VBComponents { get; }
            public DebugProject() { VBComponents = new Components { Parent = this }; }
        }
        public sealed class Components : List<Component> { public DebugProject Parent { get; set; } }
        public sealed class Component
        {
            public string Name { get; set; }
            public int Type { get; set; }
            public Components Collection { get; set; }
            public DebugModule CodeModule { get; set; }
            public bool HasOpenDesigner { get; set; }
            public Window Designer { get; set; }
            public Window DesignerWindow() { return Designer; }
        }
        public sealed class DebugModule
        {
            public Component Parent { get; set; }
            public DebugPane CodePane { get; set; }
            public string Code { get; set; }
            public int CountOfLines => string.IsNullOrEmpty(Code) ? 0 : Code.Split(new[] { "\r\n" }, StringSplitOptions.None).Length;
            public Lines Lines { get { return new Lines(this); } }
        }
        public sealed class Lines
        {
            private readonly DebugModule module;
            public Lines(DebugModule value) { module = value; }
            public string this[int start, int count] => string.Join("\r\n", module.Code.Split(new[] { "\r\n" }, StringSplitOptions.None).Skip(start - 1).Take(count));
        }
        public sealed class DebugPane
        {
            public int ModuleError { get; set; }
            private DebugModule module;
            public DebugModule CodeModule { get { if (ModuleError != 0) throw new COMException("expired pane", ModuleError); return module; } set { module = value; } }
            public Window Window { get; set; }
            public int Start = 1, Column = 1, End = 1, EndColumn = 1;
            public int CountOfVisibleLines { get; set; } = 20;
            public int CodePaneView { get; set; } = 1;
            public Action OnShow { get; set; }
            public Action OnReadSelection { get; set; }
            public Action<int> OnSetTop { get; set; }
            private int top = 1;
            public int TopLine { get { return top; } set { OnSetTop?.Invoke(value); top = value; } }
            public void Show() { OnShow?.Invoke(); }
            public void SetSelection(int start, int column, int end, int endColumn) { Start = start; Column = column; End = end; EndColumn = endColumn; }
            public void GetSelection(ref int start, ref int column, ref int end, ref int endColumn)
            { OnReadSelection?.Invoke(); start = Start; column = Column; end = End; endColumn = EndColumn; }
        }
        public sealed class Window
        {
            public int Type { get; set; }
            public string Caption { get; set; }
            public bool Visible { get; set; } = true;
            public object LinkedWindowFrame { get; set; }
            public int WindowState { get; set; }
            public int Left { get; set; }
            public int Top { get; set; }
            public int Width { get; set; } = 300;
            public int Height { get; set; } = 200;
            public int HWnd { get; set; }
            public Action OnFocus { get; set; }
            public void SetFocus() { OnFocus?.Invoke(); }
        }
        public sealed class Bars : List<Bar>
        {
            public Action OnFind { get; set; }
            public object FindControl(int type, int id) { OnFind?.Invoke(); return this.SelectMany(x => x.Controls).FirstOrDefault(x => x.Id == id); }
        }
        public sealed class Bar
        {
            public string Name { get; set; } = "Commands";
            public List<Control> Controls { get; } = new List<Control>();
        }
        public sealed class Control
        {
            public int Id { get; set; }
            public string Caption { get; set; }
            public bool Enabled { get; set; } = true;
            public List<Control> Controls { get; } = new List<Control>();
            public Action OnExecute { get; set; }
            public int Executions { get; private set; }
            public void Execute() { Executions++; OnExecute?.Invoke(); }
        }
    }
}

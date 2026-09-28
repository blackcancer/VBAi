using System;
using System.Collections;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using CodexVBE;

namespace CodexVBE.Tests.Infrastructure
{
    /// <summary>Modélise les contrats VBIDE avec fichiers réels et frontières de panne explicites ; aucun hôte Office.</summary>
    public sealed class EditorVbeContract
    {
        public readonly Host Vbe = new Host();
        public readonly ContractProject Project = new ContractProject();
        public readonly Component Original;
        internal readonly EditorVbeModule Adapter;
        public EditorVbeContract(int type = 1, string code = "Option Explicit\nPublic Sub Hello()\n    Debug.Print 1\nEnd Sub")
        {
            Vbe.VBProjects.Add(Project); Project.VBComponents.Parent = Project;
            Original = new Component { Name = "Module1", Type = type, Collection = Project.VBComponents };
            Original.CodeModule.Raw = code; Project.VBComponents.Items.Add(Original);
            Vbe.ActiveVBProject = Project; Vbe.ActiveCodePane = Original.CodeModule.CodePane;
            Original.CodeModule.CodePane.OnShow = () => Vbe.ActiveCodePane = Original.CodeModule.CodePane;
            Adapter = new EditorVbeModule(Vbe, Project, Original);
        }
        public sealed class Host
        {
            public List<object> VBProjects { get; } = new List<object>();
            public Window MainWindow { get; } = new Window();
            public object ActiveVBProject { get; set; }
            public object ActiveCodePane { get; set; }
            public List<CodexVBE.Tests.Unit.VbeDebugTests.FakeBar> CommandBars { get; } = new List<CodexVBE.Tests.Unit.VbeDebugTests.FakeBar>();
        }
        public sealed class ContractProject
        {
            public string Name { get; set; } = "Project1";
            public string Path;
            public bool FailFileName;
            public string FileName => FailFileName ? throw new COMException("FileName unavailable") : Path;
            public int Mode { get; set; } = 2;
            public int Protection { get; set; }
            public Components VBComponents { get; } = new Components();
        }
        public sealed class Components : IEnumerable
        {
            public readonly List<Component> Items = new List<Component>();
            public ContractProject Parent { get; set; }
            public Action<Component, string> AfterImport;
            public Action<Component> BeforeRemove, AfterRemove;
            public bool ImportThrowsBefore, ImportThrowsAfter;
            public Component Item(string name) => Items.Single(item => item.Name == name);
            public Component Import(string path)
            {
                if (ImportThrowsBefore) throw new IOException("Import before allocation");
                var text = File.ReadAllText(path, Encoding.Default);
                var match = System.Text.RegularExpressions.Regex.Match(text, "Attribute VB_Name\\s*=\\s*\"([^\"]+)\"");
                int type = System.IO.Path.GetExtension(path) == ".frm" ? 3 : System.IO.Path.GetExtension(path) == ".cls" ? 2 : 1;
                var component = new Component { Name = match.Groups[1].Value, Type = type, Collection = this };
                component.CodeModule.Raw = EditorAttributeRewrite.CodeSection(text);
                if (type == 3) component.CodeModule.Raw = "\n" + component.CodeModule.Raw;
                Items.Add(component); AfterImport?.Invoke(component, path);
                if (ImportThrowsAfter) throw new IOException("Import after allocation");
                return component;
            }
            public void Remove(object item)
            { var component = (Component)item; BeforeRemove?.Invoke(component); Items.Remove(component); AfterRemove?.Invoke(component); }
            public IEnumerator GetEnumerator() => Items.GetEnumerator();
        }
        public sealed class Component
        {
            public Component() { CodeModule.Parent = this; }
            private string name;
            public Action<string> BeforeName, AfterName;
            public string Name { get => name; set { BeforeName?.Invoke(value); name = value; AfterName?.Invoke(value); } }
            public int Type { get; set; }
            public Components Collection { get; set; }
            public Module CodeModule { get; } = new Module();
            public Designer Designer { get; } = new Designer();
            public object[] Properties => new object[] { new Property("Name", () => Name), new Property("Caption", () => Designer.Caption) };
            public Action<string> BeforeExport, AfterExport;
            public void Export(string path)
            {
                BeforeExport?.Invoke(path);
                string header = Type == 2 ? "VERSION 1.0 CLASS\nBEGIN\n  MultiUse = -1\nEND\n" :
                    Type == 3 ? "VERSION 5.00\nBegin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} " + Name + "\nEnd\n" : "";
                string code = string.Join("\n", CodeModule.Raw.Split('\n').Where(line => !line.StartsWith("Attribute VB_Name", StringComparison.OrdinalIgnoreCase)));
                File.WriteAllText(path, (header + "Attribute VB_Name = \"" + Name + "\"\n" + code).Replace("\n", "\r\n"), Encoding.Default);
                if (Type == 3) File.WriteAllBytes(System.IO.Path.ChangeExtension(path, ".frx"), new byte[] { 1, 2, 3 });
                AfterExport?.Invoke(path);
            }
        }
        public sealed class Designer
        {
            public string Name { get; set; } = "Form1";
            public string Caption { get; set; } = "Caption";
            public List<object> Controls { get; } = new List<object>();
        }
        public sealed class Property
        {
            private readonly Func<object> value;
            public Property(string name, Func<object> value) { Name = name; this.value = value; }
            public string Name { get; }
            public int NumIndices => 0;
            public object Value => value();
        }
        public sealed class Module
        {
            public Component Parent { get; set; }
            public string Raw = "";
            public readonly Pane CodePane = new Pane();
            public Module() { CodePane.CodeModule = this; }
            public Action<string> BeforeOperation, AfterOperation;
            public Action BeforeRead;
            public bool FailAdd;
            public bool LoseAttributesOnAdd;
            public int CountOfLines => Visible().Count;
            public Range Lines => new Range(this);
            private List<string> Visible() => Raw.Length == 0 ? new List<string>() : Raw.Split('\n').Where(line => !IsAttribute(line)).ToList();
            private static bool IsAttribute(string line) => line.TrimStart().StartsWith("Attribute ", StringComparison.OrdinalIgnoreCase);
            private List<int> VisibleIndices(List<string> raw) => Enumerable.Range(0, raw.Count).Where(i => !IsAttribute(raw[i])).ToList();
            public void ReplaceLine(int line, string text)
            {
                BeforeOperation?.Invoke("replace"); var raw = Raw.Split('\n').ToList(); raw[VisibleIndices(raw)[line - 1]] = text;
                Raw = string.Join("\n", raw); AfterOperation?.Invoke("replace");
            }
            public void DeleteLines(int line, int count)
            {
                BeforeOperation?.Invoke("delete"); var raw = Raw.Split('\n').ToList(); var indices = VisibleIndices(raw);
                int first = indices[line - 1], last = indices[line + count - 2] + 1;
                while (last < raw.Count && IsAttribute(raw[last])) last++;
                raw.RemoveRange(first, last - first); Raw = string.Join("\n", raw); AfterOperation?.Invoke("delete");
            }
            public void InsertLines(int line, string text)
            {
                BeforeOperation?.Invoke("insert"); var raw = Raw.Length == 0 ? new List<string>() : Raw.Split('\n').ToList();
                var indices = VisibleIndices(raw); int position = line > indices.Count ? raw.Count : indices[line - 1];
                raw.InsertRange(position, EditorDocument.Normalize(text).Split('\n')); Raw = string.Join("\n", raw); AfterOperation?.Invoke("insert");
            }
            public void AddFromFile(string path)
            {
                BeforeOperation?.Invoke("add"); if (FailAdd) throw new IOException("AddFromFile");
                string text = EditorAttributeRewrite.CodeSection(File.ReadAllText(path, Encoding.Default));
                if (LoseAttributesOnAdd) text = string.Join("\n", text.Split('\n').Where(line => !IsAttribute(line)));
                Raw = CountOfLines == 0 ? text : Raw + "\n" + text; AfterOperation?.Invoke("add");
            }
            public sealed class Range
            {
                private readonly Module owner;
                public Range(Module owner) { this.owner = owner; }
                public string this[int first, int count] { get { owner.BeforeRead?.Invoke(); return string.Join("\r\n", owner.Visible().Skip(first - 1).Take(count)); } }
            }
        }
        public sealed class Pane
        {
            public Window Window { get; } = new Window();
            public Module CodeModule { get; set; }
            public int Shows, Line, Column;
            public Action OnShow;
            public void Show() { Shows++; OnShow?.Invoke(); }
            public void SetSelection(int first, int column, int last, int endColumn) { Line = first; Column = column; }
            public void GetSelection(ref int first, ref int column, ref int last, ref int endColumn) { first = Line; column = Column; last = Line; endColumn = Column; }
        }
        public sealed class Window
        {
            private bool visible;
            public Action OnVisible;
            public int HWnd { get; set; }
            public bool Visible { get => visible; set { OnVisible?.Invoke(); visible = value; } }
            public bool FailClose;
            public int Closes;
            public void Close() { Closes++; if (FailClose) throw new COMException("Pane already gone"); }
        }
    }
}

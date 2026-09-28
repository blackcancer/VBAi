namespace CodexVBE.Tests.Unit
{
    using System;
    using System.Collections;
    using System.Collections.Generic;
    using System.IO;
    using System.Linq;
    using System.Text;
    using System.Text.RegularExpressions;
    using CodexVBE;

    public sealed partial class VbaGitProjectTests
    {
        public sealed class ProjectFixture
        {
            public string FileName { get; set; } = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "disposable-transfer.xlsm");
            public int Mode { get; set; } = 2;
            public int Protection { get; set; }
            public ComponentsFixture VBComponents { get; } = new ComponentsFixture();
            public List<ReferenceFixture> References { get; } = new List<ReferenceFixture>();
        }
        public sealed class ReferenceFixture
        {
            public bool IsBroken { get; set; }
            public string GUID { get; set; } = "{00020430-0000-0000-c000-000000000046}";
            public int Major { get; set; } = 2;
            public int Minor { get; set; }
        }
        public sealed class ComponentsFixture : IEnumerable<ComponentFixture>
        {
            internal readonly List<ComponentFixture> Items = new List<ComponentFixture>();
            internal bool WrongName, WrongType, LoseText;
            public ComponentFixture Item(string name) => Items.Single(component => component.Name == name);
            public void Remove(ComponentFixture component) => Items.Remove(component);
            public ComponentFixture Import(string path)
            {
                string text = File.ReadAllText(path, Encoding.Default);
                string name = Regex.Match(text, "Attribute VB_Name = \"([^\"]+)\"").Groups[1].Value;
                var component = new ComponentFixture(name, path.EndsWith(".bas") ? 1 : path.EndsWith(".cls") ? 2 : 3, text);
                if (WrongName) component.Name = "Unexpected";
                if (WrongType) component.Type = 100;
                if (LoseText) component.CodeModule.Text += "' native host changed the imported source\n";
                string resource = Path.ChangeExtension(path, ".frx");
                if (File.Exists(resource)) component.Resource = File.ReadAllBytes(resource);
                Items.Add(component); return component;
            }
            public IEnumerator<ComponentFixture> GetEnumerator() => Items.GetEnumerator();
            IEnumerator IEnumerable.GetEnumerator() => GetEnumerator();
        }
        public sealed class ComponentFixture
        {
            public string Name { get; set; }
            public int Type { get; set; }
            public ModuleFixture CodeModule { get; }
            public byte[] Resource { get; set; }
            public ComponentFixture(string name, int type, string text) { Name = name; Type = type; CodeModule = new ModuleFixture(text); }
            public void Export(string path)
            {
                File.WriteAllText(path, CodeModule.Text, Encoding.Default);
                if (Resource != null) File.WriteAllBytes(Path.ChangeExtension(path, ".frx"), Resource);
            }
        }
        public sealed class ModuleFixture
        {
            public string Text { get; set; }
            public int CountOfLines => Text.Length == 0 ? 0 : Text.Split('\n').Length;
            public ModuleFixture Lines => this;
            public PaneFixture CodePane { get; } = new PaneFixture();
            public ModuleFixture(string text) { Text = text.Replace("\r\n", "\n"); }
            public string this[int start, int count] => string.Join("\n", Text.Split('\n').Skip(start - 1).Take(count));
            public void DeleteLines(int start, int count) { var lines = Text.Split('\n').ToList(); lines.RemoveRange(start - 1, count); Text = string.Join("\n", lines); }
            public void InsertLines(int start, string text) { var lines = Text.Length == 0 ? new List<string>() : Text.Split('\n').ToList(); lines.InsertRange(start - 1, text.Replace("\r\n", "\n").Split('\n')); Text = string.Join("\n", lines); }
        }
        public sealed class PaneFixture
        {
            public bool Shown { get; private set; }
            public int Line { get; private set; }
            public void Show() { Shown = true; }
            public void SetSelection(int startLine, int startColumn, int endLine, int endColumn) { Line = startLine; }
        }
        internal static VbaGitProject Adapter(ProjectFixture host) => new VbaGitProject(() => host, host.FileName);
        internal static string Code(string name, string extra = "") => "Attribute VB_Name = \"" + name + "\"\nOption Explicit\n" + extra;
    }
}

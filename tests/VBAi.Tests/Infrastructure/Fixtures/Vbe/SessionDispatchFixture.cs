using System;
using System.Collections;
using System.Collections.Generic;
using VBAi;

namespace VBAi.Tests.Unit
{
    public sealed class SessionDispatchFixture
    {
        public sealed class Host
        {
            public List<object> VBProjects { get; } = new List<object>();
            public List<Window> Windows { get; } = new List<Window>();
            public List<object> CodePanes { get; } = new List<object>();
            public List<object> CommandBars { get; } = new List<object>();
            public List<VbeAddInReadbackTests.Plugin> AddIns { get; } = new List<VbeAddInReadbackTests.Plugin>();
            public object ActiveVBProject { get; set; }
            public Window ActiveWindow { get; set; }
            public object ActiveCodePane { get; set; }
            public Events Events { get; } = new Events();
        }
        public sealed class Events { public ReferenceEvents ReferencesEvents { get; } = new ReferenceEvents(); }
        public sealed class ReferenceEvents { public object this[object project] => project; }
        public sealed class Window
        {
            public string Caption { get; set; } = "Pane";
            public int Type { get; set; } = 4;
            public bool Visible { get; set; } = true;
            public int WindowState { get; set; }
            public int Left { get; set; }
            public int Top { get; set; }
            public int Width { get; set; } = 300;
            public int Height { get; set; } = 200;
            public Window LinkedWindowFrame { get; set; }
            public Members LinkedWindows { get; }
            public Action Focus { get; set; }
            public Window() { LinkedWindows = new Members(this); }
            public void SetFocus() { Focus?.Invoke(); }
        }
        public sealed class Members : IEnumerable<Window>
        {
            private readonly Window frame;
            private readonly List<Window> entries = new List<Window>();
            public Members(Window frame) { this.frame = frame; }
            public void Add(Window window) { entries.Add(window); window.LinkedWindowFrame = frame; }
            public void Remove(Window window) { entries.Remove(window); window.LinkedWindowFrame = null; }
            public IEnumerator<Window> GetEnumerator() { return entries.GetEnumerator(); }
            IEnumerator IEnumerable.GetEnumerator() { return GetEnumerator(); }
        }
        internal sealed class Clipboard : ICodeClipboard
        {
            internal CodeClipboardSnapshot State = new CodeClipboardSnapshot { HasText = true, Text = "initial", Version = "1" };
            private int revision = 1;
            public CodeClipboardSnapshot Read() { return State; }
            public CodeClipboardSnapshot Write(string text)
            { return State = new CodeClipboardSnapshot { HasText = true, Text = text, Version = (++revision).ToString() }; }
        }
        public sealed class CorruptProject
        {
            public string Name { get; set; } = "P";
            public string FileName { get; set; } = "";
            public int Mode { get; set; } = 2;
            public List<CorruptComponent> VBComponents { get; } = new List<CorruptComponent>();
        }
        public sealed class CorruptComponent { public string Name { get; set; } = "M"; public CorruptModule CodeModule { get; set; } = new CorruptModule(); }
        public sealed class CorruptModule
        {
            public string Code = "original";
            private int inserts;
            public string PartialAfterFailure = "";
            public bool RestoreExactly;
            public int CountOfLines => Code.Length == 0 ? 0 : 1;
            public CorruptLines Lines { get; }
            public CorruptModule() { Lines = new CorruptLines(this); }
            public void DeleteLines(int start, int count) { Code = ""; }
            public void InsertLines(int start, string text)
            { if (++inserts == 1) { Code = PartialAfterFailure; throw new InvalidOperationException("initial insert failed"); } Code = RestoreExactly ? text : "unexpected normalization"; }
        }
        public sealed class CorruptLines
        {
            private readonly CorruptModule module;
            public CorruptLines(CorruptModule module) { this.module = module; }
            public string this[int start, int count] => module.Code;
        }
    }
}

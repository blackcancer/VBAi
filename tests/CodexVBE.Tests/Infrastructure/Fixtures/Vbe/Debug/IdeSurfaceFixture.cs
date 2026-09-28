using System;
using System.Collections.Generic;
using System.Linq;
using CodexVBE;

namespace CodexVBE.Tests.Infrastructure
{
    /// <summary>Doubles bornés des nouvelles surfaces IDE, sans exécution de macro.</summary>
    public static class IdeSurfaceFixture
    {
        internal sealed class NavigationProbe : VbeDebugWindows.INavigationSurfaceProbe
        {
            internal VbeDebugWindows.NavigationSurface State = new VbeDebugWindows.NavigationSurface {
                Available = true, Identity = "tree-1", Caption = "Project", Nodes = new[] {
                    new VbeDebugWindows.NavigationNode { Token = "1", Name = "VBAProject", Enabled = true, Expansion = "Collapsed", Selected = false },
                    new VbeDebugWindows.NavigationNode { Token = "2", ParentToken = "1", Name = "Module1", Enabled = true, Selected = false } } };
            internal Action BeforeRead, AfterAction;
            internal int Reads, Actions;
            public VbeDebugWindows.NavigationSurface Read(string pane) { Reads++; BeforeRead?.Invoke(); return State; }
            public void Act(string pane, string token, string action)
            {
                Actions++;
                var node = State.Nodes.Single(n => n.Token == token);
                if (action == "select") node.Selected = true; else node.Expansion = action == "expand" ? "Expanded" : "Collapsed";
                AfterAction?.Invoke();
            }
            internal Request Change(string action = "expand", string token = "1")
            {
                dynamic snapshot = VbeDebugWindows.ReadNavigationSurface(new Request { Pane = "project" }, this);
                return new Request { Pane = "project", Action = action, Control = token, ExpectedWindowVersion = snapshot.WindowVersion };
            }
        }

        public sealed class Vbe { public List<Project> VBProjects { get; } = new List<Project>(); }
        public sealed class Project
        {
            public string Name { get; set; } = "P";
            public string FileName { get; set; } = "";
            public int Mode { get; set; } = 2;
            public int Protection { get; set; }
            public bool Saved { get; set; } = true;
            public string HelpFile { get; set; }
            public object HelpContextID { get; set; } = 0;
            public List<Component> VBComponents { get; } = new List<Component>();
            public List<object> References { get; } = new List<object>();
        }
        public sealed class Component
        {
            public string Name { get; set; } = "M";
            public int Type { get; set; } = 1;
            public Module CodeModule { get; set; } = new Module();
        }
        public sealed class Module
        {
            public string Source { get; set; } = "";
            public int CountOfLines => Source.Length == 0 ? 0 : Source.Replace("\r", "").Split('\n').Length;
            public SourceLines Lines => new SourceLines(this);
        }
        public sealed class SourceLines
        {
            private readonly Module module;
            public SourceLines(Module module) { this.module = module; }
            public string this[int start, int count] => string.Join("\r\n", module.Source.Replace("\r", "").Split('\n').Skip(start - 1).Take(count));
        }
    }
}

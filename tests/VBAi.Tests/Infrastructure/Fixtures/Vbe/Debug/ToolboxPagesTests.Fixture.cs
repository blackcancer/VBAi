using System;
using System.Collections.Generic;

namespace VBAi.Tests.Infrastructure
{
    /// <summary>Arbre MSAA déterministe sans opérations de sélection ni objets COM.</summary>
    internal sealed class ToolboxPagesFixture : VbeDebugWindows.IToolboxAccessibleNode
    {
        internal int RoleValue, StateValue, CountOverride = -1;
        internal string NameValue;
        internal bool FailCount, FailRole, FailState, FailName, FailChild;
        internal readonly List<ToolboxPagesFixture> Children = new List<ToolboxPagesFixture>();
        internal ToolboxPagesFixture(int role = 20, string name = null, int flags = 0)
        { RoleValue = role; NameValue = name; StateValue = flags; }
        private ToolboxPagesFixture Resolve(int child) => child == 0 ? this : Children[child - 1];
        public int Role(int child) { if (FailRole) throw new InvalidOperationException(); return Resolve(child).RoleValue; }
        public int State(int child) { if (FailState) throw new InvalidOperationException(); return Resolve(child).StateValue; }
        public string Name(int child) { if (FailName) throw new InvalidOperationException(); return Resolve(child).NameValue; }
        public int Count { get { if (FailCount) throw new InvalidOperationException(); return CountOverride == -1 ? Children.Count : CountOverride; } }
        public VbeDebugWindows.IToolboxAccessibleNode Child(int child)
        { if (FailChild) throw new InvalidOperationException(); return Children[child - 1]; }
        public void Dispose() { }
        internal static ToolboxPagesFixture Create()
        {
            var root = new ToolboxPagesFixture(); var tabs = new ToolboxPagesFixture(60);
            tabs.Children.Add(new ToolboxPagesFixture(37, "Contrôles", 2097154));
            tabs.Children.Add(new ToolboxPagesFixture(37, "Autres", 2097152));
            root.Children.Add(tabs); return root;
        }
        internal sealed class Probe : VbeDebugWindows.INavigationSurfaceProbe
        {
            internal VbeDebugWindows.NavigationSurface State;
            internal int Actions;
            public VbeDebugWindows.NavigationSurface Read(string pane) => State;
            public void Act(string pane, string token, string action) { Actions++; }
        }
    }
}
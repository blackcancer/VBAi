using System;
using System.Collections.Generic;
using System.IO;
using VBAi;
using VBAi.Tests.Infrastructure;

namespace VBAi.Tests.Unit
{
    public sealed class AddInEditorActiveWindow { public int Type { get; set; } }
    internal static class OwnedMdiWorkspace { [System.Runtime.InteropServices.DllImport("user32.dll")] internal static extern IntPtr GetParent(IntPtr child); }
    public sealed class AddInEditorCollection { public object Parent { get; set; } }
    public sealed class AddInEditorComponent { public string Name { get; set; } = "Module1"; public AddInEditorCollection Collection { get; set; } }
    public sealed class AddInEditorCode { public object Parent { get; set; } }
    public sealed class AddInEditorPane { public AddInEditorCode CodeModule { get; set; } }
    internal sealed class AddInModernEditorFixture : IDisposable
    {
        internal readonly HostUiScope Scope = new HostUiScope();
        internal readonly List<ModernEditorWindow> Editors = new List<ModernEditorWindow>();
        internal readonly AddIn Instance;
        internal AddInModernEditorFixture(bool connect = false)
        {
            Scope.Host.Workspace();
            var create = AddIn.CreateModernEditor;
            AddIn.CreateModernEditor = () =>
            {
                var editor = create();
                // Hold the existing initialization-in-progress guard while the Host
                // orchestration is tested; no WebView environment/profile is created.
                LlmBoundaryScope.Set(editor, "initializing", true);
                editor.Drafts = new EditorDraftStore(Path.Combine(Path.GetDirectoryName(Scope.Host.Project.FileName), "editor-drafts"));
                Editors.Add(editor); return editor;
            };
            Instance = connect ? Scope.Connected() : new AddIn();
            if (!connect) { LlmBoundaryScope.Set(Instance, "vbe", Scope.Host); LlmBoundaryScope.Set(Instance, "addIn", Scope.Host.AddIns.AddIn); }
        }
        internal ModernEditorWindow Get(bool show = true) => (ModernEditorWindow)LlmBoundaryScope.Call(Instance, "GetModernEditor", show);
        internal void Call(string name) => LlmBoundaryScope.Call(Instance, name);
        internal AddInEditorComponent Active(int mode = 2)
        {
            Scope.Host.ActiveWindow = new AddInEditorActiveWindow();
            Scope.Host.Project.Mode = mode;
            var component = new AddInEditorComponent { Collection = new AddInEditorCollection { Parent = Scope.Host.Project } };
            Scope.Host.ActiveCodePane = new AddInEditorPane { CodeModule = new AddInEditorCode { Parent = component } }; return component;
        }
        public void Dispose()
        { Scope.Close(Instance); foreach (var editor in Editors) editor.Dispose(); Scope.Dispose(); }
    }
}

namespace VBAi.Tests.Unit
{
    using System;
    using System.Collections.Generic;
    using System.Reflection;
    using System.Windows.Automation;
    using VBAi;

    public sealed partial class VbeDebugWindowsSystemTests
    {
        // The real UIA transport targets only our disposable provider windows.
        private sealed class ObjectBrowserScene : IDisposable
        {
            internal readonly NativeDebugScene Native = new NativeDebugScene();
            internal readonly NativeDebugScene.Window Root, Pane;
            private readonly List<AutomationHost> hosts = new List<AutomationHost>();
            private readonly Dictionary<System.Reflection.FieldInfo, object> saved = new Dictionary<System.Reflection.FieldInfo, object>();
            internal readonly List<Tuple<IntPtr, int, IntPtr, IntPtr>> Notifications = new List<Tuple<IntPtr, int, IntPtr, IntPtr>>();
            internal bool PostSucceeded = true;
            internal Action<int> OnPause;
            internal int Pauses;
            internal readonly HashSet<IntPtr> Unreadable = new HashSet<IntPtr>();
            internal ObjectBrowserScene()
            {
                foreach (string name in new[] { "ObjectBrowserElement", "PostMessage", "PauseNative", "GetDlgCtrlID" })
                {
                    var field = typeof(VbeDebugWindows).GetField(name, BindingFlags.NonPublic | BindingFlags.Static); saved.Add(field, field.GetValue(null));
                }
                var acquire = VbeDebugWindows.ObjectBrowserElement;
                VbeDebugWindows.ObjectBrowserElement = handle =>
                {
                    if (Unreadable.Contains(handle)) throw new ElementNotAvailableException();
                    return acquire(handle);
                };
                VbeDebugWindows.PostMessage = (handle, message, w, l) => { Notifications.Add(Tuple.Create(handle, message, w, l)); return PostSucceeded; };
                VbeDebugWindows.GetDlgCtrlID = handle => 42;
                VbeDebugWindows.PauseNative = milliseconds => { Pauses++; OnPause?.Invoke(Pauses); };
                Root = Native.Add("wndclass_desked_gsk", "VBE");
                Pane = Native.Add("VbaWindow", "Object Browser", Root);
            }
            internal NativeDebugScene.Window Add(AutomationNode node, string kind = "ListBox")
            {
                var host = new AutomationHost(node); hosts.Add(host);
                var window = Native.Add(kind, node.Name, Pane); window.Handle = host.Handle; return window;
            }
            public void Dispose()
            {
                try { foreach (var pair in saved) pair.Key.SetValue(null, pair.Value); }
                finally { try { foreach (var host in hosts) host.Dispose(); } finally { Native.Dispose(); } }
            }
        }
        private static AutomationNode BrowserList(params string[] labels)
        {
            var list = new AutomationNode { Name = "List", Kind = ControlType.List }.With(SelectionPattern.Pattern);
            foreach (string label in labels) list.Add(new AutomationNode { Name = label }.With(SelectionItemPattern.Pattern));
            return list;
        }
        private static AutomationNode BrowserLibraries(string name = "Libraries", params string[] labels)
        {
            var combo = new AutomationNode { Name = name, Kind = ControlType.ComboBox }.With(SelectionPattern.Pattern, ValuePattern.Pattern);
            foreach (string label in labels) combo.Add(new AutomationNode { Name = label }.With(SelectionItemPattern.Pattern));
            return combo;
        }
    }
}

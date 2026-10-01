namespace VBAi.Tests.Unit
{
    using System;
    using System.Collections.Generic;
    using System.Diagnostics;
    using System.Linq;
    using System.Reflection;
    using System.Runtime.InteropServices;
    using System.Windows.Forms;
    using VBAi;

    public sealed partial class VbeDebugWindowsSystemTests
    {
        private sealed class SystemWindow
        {
            public IntPtr Handle;
            public IntPtr Parent;
            public string Class = "#32770";
            public string Text;
            public bool Visible = true;
            public bool Enabled = true;
            public uint ProcessId = (uint)Process.GetCurrentProcess().Id;
            public uint ThreadId = 1;
            public int ControlId;
            public AccessibleObject Accessible;
        }

        private sealed class SystemScene : IDisposable
        {
            private readonly Dictionary<FieldInfo, object> saved = new Dictionary<FieldInfo, object>();
            public readonly List<SystemWindow> Windows = new List<SystemWindow>();
            public readonly List<Tuple<IntPtr, int, IntPtr, IntPtr>> Messages = new List<Tuple<IntPtr, int, IntPtr, IntPtr>>();
            public readonly List<Tuple<IntPtr, string>> Replacements = new List<Tuple<IntPtr, string>>();
            public Action<int> OnPause;
            public Action<SystemWindow, int> OnMessage;
            public int Pauses;
            public int AccessibilityHResult;
            public object AccessibilityOverride;
            public bool OverrideAccessibility;
            public bool PostSucceeds = true;
            public IntPtr IntegerMessageResult;

            public SystemScene()
            {
                foreach (string name in new[] { "EnumWindows", "EnumChildWindows", "GetWindowThreadProcessId", "GetClassName", "GetWindowText", "GetDlgCtrlID", "GetDlgItem", "IsWindowVisible", "OptionsWindowEnabled", "PostMessage", "SendMessageText", "SendMessageInt", "AccessibleObjectFromWindow", "PauseNative" })
                {
                    var field = typeof(VbeDebugWindows).GetField(name, BindingFlags.Static | BindingFlags.NonPublic);
                    saved.Add(field, field.GetValue(null));
                }
                VbeDebugWindows.EnumWindows = (callback, parameter) => {
                    foreach (var window in Windows.Where(w => w.Parent == IntPtr.Zero).ToArray())
                        if (!callback(window.Handle, parameter)) break;
                    return true;
                };
                VbeDebugWindows.EnumChildWindows = (parent, callback, parameter) => {
                    foreach (var window in Windows.Where(w => IsBelow(w, parent)).ToArray())
                        if (!callback(window.Handle, parameter)) break;
                    return true;
                };
                VbeDebugWindows.GetWindowThreadProcessId = (IntPtr handle, out uint processId) => {
                    var window = Find(handle);
                    processId = window?.ProcessId ?? 0;
                    return window?.ThreadId ?? 0;
                };
                VbeDebugWindows.GetClassName = (handle, text, capacity) => { text.Append(Find(handle)?.Class); return text.Length; };
                VbeDebugWindows.GetWindowText = (handle, text, capacity) => { text.Append(Find(handle)?.Text); return text.Length; };
                VbeDebugWindows.GetDlgCtrlID = handle => Find(handle)?.ControlId ?? 0;
                VbeDebugWindows.GetDlgItem = (dialog, id) => Windows.FirstOrDefault(w => w.Parent == dialog && w.ControlId == id)?.Handle ?? IntPtr.Zero;
                VbeDebugWindows.IsWindowVisible = handle => Find(handle)?.Visible ?? false;
                VbeDebugWindows.OptionsWindowEnabled = handle => Find(handle)?.Enabled ?? false;
                VbeDebugWindows.PostMessage = (handle, message, wParam, lParam) => {
                    Messages.Add(Tuple.Create(handle, message, wParam, lParam));
                    OnMessage?.Invoke(Find(handle), message); return PostSucceeds;
                };
                VbeDebugWindows.SendMessageText = (handle, message, parameter, text) => {
                    Replacements.Add(Tuple.Create(handle, text));
                    var window = Find(handle); if (window != null) window.Text = text;
                    return IntPtr.Zero;
                };
                VbeDebugWindows.SendMessageInt = (handle, message, wParam, lParam) => IntegerMessageResult;
                VbeDebugWindows.AccessibleObjectFromWindow = (IntPtr handle, uint objectId, ref Guid iid, out object accessible) => {
                    accessible = OverrideAccessibility ? AccessibilityOverride : Find(handle)?.Accessible;
                    return AccessibilityHResult;
                };
                VbeDebugWindows.PauseNative = milliseconds => { Pauses++; OnPause?.Invoke(Pauses); };
            }

            private bool IsBelow(SystemWindow window, IntPtr parent)
            {
                for (var current = window; current != null && current.Parent != IntPtr.Zero; current = Find(current.Parent))
                    if (current.Parent == parent) return true;
                return false;
            }

            public SystemWindow Add(string text, string kind = "#32770", SystemWindow parent = null, int id = 0)
            {
                var window = new SystemWindow { Handle = new IntPtr(100 + Windows.Count), Text = text,
                    Class = kind, Parent = parent?.Handle ?? IntPtr.Zero, ControlId = id,
                    ProcessId = parent?.ProcessId ?? (uint)Process.GetCurrentProcess().Id, ThreadId = parent?.ThreadId ?? 1 };
                Windows.Add(window); return window;
            }

            public SystemWindow Find(IntPtr handle) { return Windows.FirstOrDefault(w => w.Handle == handle); }
            public void Dispose() { foreach (var pair in saved) pair.Key.SetValue(null, pair.Value); }
        }

        private sealed class AccessibleNode : AccessibleObject
        {
            public readonly List<AccessibleNode> Children = new List<AccessibleNode>();
            public string Label;
            public AccessibleRole NativeRole;
            public Action OnInvoke;
            public bool FailName;
            public bool FailRole;
            public AccessibleStates StateValue;
            public Action OnName;
            public override string Name { get { OnName?.Invoke(); if (FailName) throw new COMException("name unavailable"); return Label; } set { Label = value; } }
            public override AccessibleStates State => StateValue;
            public override AccessibleRole Role { get { if (FailRole) throw new COMException("role unavailable"); return NativeRole; } }
            public override int GetChildCount() { return Children.Count; }
            public override AccessibleObject GetChild(int index) { return Children[index]; }
            public override void DoDefaultAction() { OnInvoke?.Invoke(); }
        }

        private sealed class ChildValueAccessible : AccessibleObject, Accessibility.IAccessible
        {
            public object ChildValue;
            object Accessibility.IAccessible.get_accChild(object child) => ChildValue;
        }

        private static AccessibleNode Label(string text) { return new AccessibleNode { Label = text, NativeRole = AccessibleRole.StaticText }; }
        private static AccessibleNode Button(string text, Action click = null) { return new AccessibleNode { Label = text, NativeRole = AccessibleRole.PushButton, OnInvoke = click }; }

        private static AccessibleNode Signature(SystemWindow dialog, string current, string signAs)
        {
            var root = new AccessibleNode();
            root.Children.AddRange(new[] { Label(current), Label("Certificate name:"), Label("The VBA project is currently signed as"),
                Label(signAs), Label("Certificate name:"), Label("Sign as"), Button("Choose..."), Button("OK", () => dialog.Visible = false),
                Button("Cancel", () => dialog.Visible = false) });
            dialog.Accessible = root; return root;
        }

        private static T Native<T>(string name) { return (T)Activator.CreateInstance(typeof(VbeDebugWindows).GetNestedType(name, BindingFlags.NonPublic), true); }
        private static object Call(string name, params object[] arguments)
        {
            return typeof(VbeDebugWindows).GetMethod(name, BindingFlags.Static | BindingFlags.NonPublic).Invoke(null, arguments);
        }
    }
}

namespace VBAi.Tests.Unit
{
    using System;
    using System.Collections.Generic;
    using System.Linq;
    using System.Reflection;
    using VBAi;

    // A deterministic user32 boundary: no handles belong to an actual host.
    internal sealed class NativeDebugScene : IDisposable
    {
        internal sealed class Window
        {
            internal IntPtr Handle, Parent, Owner;
            internal string Kind, Caption;
            internal uint Process = (uint)System.Diagnostics.Process.GetCurrentProcess().Id;
            internal bool Visible = true, Enabled = true, BoundsAvailable = true, FailBounds;
            internal VbeDebugWindows.ViewRect Bounds;
        }
        private readonly Dictionary<FieldInfo, object> saved = new Dictionary<FieldInfo, object>();
        internal readonly List<Window> Windows = new List<Window>();
        internal readonly List<Tuple<IntPtr, int, IntPtr, IntPtr>> Messages = new List<Tuple<IntPtr, int, IntPtr, IntPtr>>();
        internal bool EnumerationSucceeded = true;
        internal Action OnClick;
        internal NativeDebugScene()
        {
            foreach (string name in new[] { "EnumWindows", "EnumChildWindows", "GetWindowThreadProcessId", "GetClassName", "GetWindowText", "ReadObserverTextMessage", "IsWindowVisible", "SendMessageInt", "ViewBounds", "RuntimeWindowOwner", "ObjectBrowserParent", "ObjectBrowserEnabled" })
            {
                FieldInfo field = typeof(VbeDebugWindows).GetField(name, BindingFlags.NonPublic | BindingFlags.Static);
                saved.Add(field, field.GetValue(null));
            }
            VbeDebugWindows.EnumWindows = (callback, parameter) =>
            {
                foreach (Window window in Windows.Where(w => w.Parent == IntPtr.Zero).ToArray())
                    if (!callback(window.Handle, parameter)) break;
                return EnumerationSucceeded;
            };
            VbeDebugWindows.EnumChildWindows = (parent, callback, parameter) =>
            {
                foreach (Window window in Windows.Where(w => Below(w, parent)).ToArray())
                    if (!callback(window.Handle, parameter)) break;
                return true;
            };
            VbeDebugWindows.GetWindowThreadProcessId = (IntPtr handle, out uint pid) => { pid = Find(handle)?.Process ?? 0; return 1; };
            VbeDebugWindows.GetClassName = (handle, value, capacity) => { value.Append(Find(handle)?.Kind); return value.Length; };
            VbeDebugWindows.GetWindowText = (handle, value, capacity) => { value.Append(Find(handle)?.Caption); return value.Length; };
            VbeDebugWindows.ReadObserverTextMessage = (IntPtr handle, uint message, IntPtr capacity, System.Text.StringBuilder value,
                uint flags, uint milliseconds, out UIntPtr result) =>
            {
                value.Append(Find(handle)?.Caption);
                result = new UIntPtr((uint)value.Length);
                return new IntPtr(1);
            };
            VbeDebugWindows.IsWindowVisible = handle => Find(handle)?.Visible ?? false;
            VbeDebugWindows.ObjectBrowserEnabled = handle => Find(handle)?.Enabled ?? false;
            VbeDebugWindows.ObjectBrowserParent = handle => Find(handle)?.Parent ?? IntPtr.Zero;
            VbeDebugWindows.RuntimeWindowOwner = (handle, command) => Find(handle)?.Owner ?? IntPtr.Zero;
            VbeDebugWindows.ViewBounds = (IntPtr handle, out VbeDebugWindows.ViewRect bounds) =>
            {
                Window window = Find(handle);
                if (window.FailBounds) throw new InvalidOperationException("bounds unavailable");
                bounds = window.Bounds; return window.BoundsAvailable;
            };
            VbeDebugWindows.SendMessageInt = (handle, message, w, l) =>
            {
                Messages.Add(Tuple.Create(handle, message, w, l));
                if (message == 0x202) OnClick?.Invoke();
                return IntPtr.Zero;
            };
        }
        private bool Below(Window window, IntPtr parent)
        {
            for (Window current = window; current != null && current.Parent != IntPtr.Zero; current = Find(current.Parent))
                if (current.Parent == parent) return true;
            return false;
        }
        internal Window Find(IntPtr handle) { return Windows.FirstOrDefault(w => w.Handle == handle); }
        internal Window Add(string kind, string caption = "", Window parent = null, int left = 0, int top = 0, int width = 40, int height = 20)
        {
            var window = new Window
            {
                Handle = new IntPtr(5000 + Windows.Count),
                Kind = kind,
                Caption = caption,
                Parent = parent?.Handle ?? IntPtr.Zero,
                Bounds = new VbeDebugWindows.ViewRect { Left = left, Top = top, Right = left + width, Bottom = top + height }
            };
            Windows.Add(window); return window;
        }
        internal Window CodeWindow()
        {
            Window root = Add("wndclass_desked_gsk", "VBE");
            Window code = Add("VbaWindow", "Module1 (Code)", root);
            Add("ScrollBar", parent: code, left: 40, width: 200);
            return Add("ObtbarWndClass", parent: code);
        }
        public void Dispose() { foreach (var entry in saved) entry.Key.SetValue(null, entry.Value); }
    }
}

using System;
using System.Collections;
using System.Collections.Generic;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Text;
using VBAi;
using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace VBAi.Tests.Unit
{
    internal sealed class NativeThemeNativeFixture : IDisposable
    {
        internal readonly NativeThemeFixture State = new NativeThemeFixture();
        private readonly VbeNativeTheme.ReadWindowThread nativeThread = VbeNativeTheme.WindowThread;
        private readonly Func<IntPtr, bool> deleteBrush = VbeNativeTheme.ReleaseBrush;
        private readonly Func<int, IntPtr> createBrush = VbeNativeTheme.NewBrush;
        private readonly Func<IntPtr, uint, IntPtr, IntPtr, int> send = VbeNativeTheme.NativeSend;
        private readonly string local = Environment.GetEnvironmentVariable(VbeNativeTheme.LocalRefreshExperimentVariable);
        private readonly List<IntPtr> windows = new List<IntPtr>(), brushes = new List<IntPtr>();
        private readonly Dictionary<int, Delegate> exports = new Dictionary<int, Delegate>();
        internal readonly Dictionary<IntPtr, string> Classes = new Dictionary<IntPtr, string>(), Captions = new Dictionary<IntPtr, string>();
        internal readonly Dictionary<IntPtr, int> Styles = new Dictionary<IntPtr, int>();
        internal readonly Dictionary<IntPtr, IntPtr> Parents = new Dictionary<IntPtr, IntPtr>(), Owners = new Dictionary<IntPtr, IntPtr>();
        internal readonly HashSet<IntPtr> ForeignThread = new HashSet<IntPtr>(), ForeignProcess = new HashSet<IntPtr>(), CodeChildren = new HashSet<IntPtr>();
        internal readonly HashSet<int> MissingExports = new HashSet<int>();
        internal readonly List<Tuple<IntPtr, bool>> DarkCalls = new List<Tuple<IntPtr, bool>>();
        internal readonly List<int> Modes = new List<int>();
        internal readonly List<IntPtr> Installations = new List<IntPtr>(), Removals = new List<IntPtr>(), Redraws = new List<IntPtr>(), Borders = new List<IntPtr>(), Combos = new List<IntPtr>();
        internal readonly List<Tuple<IntPtr, uint>> Posts = new List<Tuple<IntPtr, uint>>(), Sends = new List<Tuple<IntPtr, uint>>();
        internal readonly List<Tuple<IntPtr, string, string>> Themes = new List<Tuple<IntPtr, string, string>>();
        internal readonly List<Tuple<IntPtr, int, int>> Attributes = new List<Tuple<IntPtr, int, int>>();
        internal readonly List<Tuple<IntPtr, bool, IntPtr, bool, bool, bool>> Paints = new List<Tuple<IntPtr, bool, IntPtr, bool, bool, bool>>();
        internal readonly List<Tuple<IntPtr, VbeNativeTheme.NativeRect>> Rows = new List<Tuple<IntPtr, VbeNativeTheme.NativeRect>>();
        internal int VersionResult, Build = 22000, Major = 10, Dwm20Result, Flushes, RefreshPolicies, HookCalls, Unhooks, NativeCalls;
        internal bool ModuleAvailable = true, SubclassSucceeds = true, PostSucceeds = true, ClientSucceeds = true;
        internal IntPtr NativeResult = new IntPtr(77);
        internal Action<IntPtr, uint, IntPtr, IntPtr> DuringNative { get; set; }
        internal IntPtr Main => State.Handle;

        [DllImport("user32.dll", CharSet = CharSet.Unicode)] private static extern IntPtr CreateWindowEx(int exStyle, string className,
            string title, int style, int x, int y, int width, int height, IntPtr parent, IntPtr id, IntPtr instance, IntPtr parameter);
        [DllImport("user32.dll")] private static extern bool DestroyWindow(IntPtr window);
        [DllImport("gdi32.dll")] private static extern uint GetObjectType(IntPtr value);

        internal NativeThemeNativeFixture()
        {
            Classes[Main] = "wndclass_desked_gsk"; State.Set("editorWindow", Main);
            Environment.SetEnvironmentVariable(VbeNativeTheme.LocalRefreshExperimentVariable, null);
            AddExport(135, "setPreferredMode", "Preferred"); AddExport(133, "allowDarkMode", "Allow");
            AddExport(104, "flushMenuThemes", "RefreshPolicy"); AddExport(136, "flushMenuThemes", "Flush");
            VbeNativeTheme.ReadClassName = (window, text, capacity) => { string value; if (!Classes.TryGetValue(window, out value)) value = ""; text.Append(value); return value.Length; };
            VbeNativeTheme.ReadWindowText = (window, text, capacity) => { string value; if (!Captions.TryGetValue(window, out value)) value = ""; text.Append(value); return value.Length; };
            var nativeStyle = VbeNativeTheme.ReadStyle;
            VbeNativeTheme.ReadStyle = (window, index) => { int style; return Styles.TryGetValue(window, out style) ? style : nativeStyle(window, index); };
            VbeNativeTheme.WindowThread = Thread;
            VbeNativeTheme.Ancestor = Ancestor;
            VbeNativeTheme.ChildRelation = IsChild;
            VbeNativeTheme.WindowRelation = (window, command) => { IntPtr owner; Assert.AreEqual((uint)4, command); return Owners.TryGetValue(window, out owner) ? owner : IntPtr.Zero; };
            VbeNativeTheme.EnumerateChildren = (parent, callback, parameter) => { foreach (var window in new List<IntPtr>(Classes.Keys)) if (IsChild(parent, window)) callback(window, parameter); return true; };
            VbeNativeTheme.ThemeModule = name => { Assert.AreEqual("uxtheme.dll", name); return ModuleAvailable ? new IntPtr(42) : IntPtr.Zero; };
            VbeNativeTheme.NativeEntryPoint = (module, ordinal) => { Assert.AreEqual(new IntPtr(42), module); int id = ordinal.ToInt32(); return MissingExports.Contains(id) ? IntPtr.Zero : Marshal.GetFunctionPointerForDelegate(exports[id]); };
            VbeNativeTheme.ReadVersion = (ref VbeNativeTheme.NativeOsVersion version) => { Assert.IsTrue(version.Size > 0); version.Major = (uint)Major; version.Minor = 0; version.Build = (uint)Build; return VersionResult; };
            VbeNativeTheme.InstallSubclass = (window, callback, id, data) => { Assert.IsNotNull(callback); Installations.Add(window); return SubclassSucceeds; };
            VbeNativeTheme.RemoveSubclass = (window, callback, id) => { Removals.Add(window); return true; };
            VbeNativeTheme.SetNativeTheme = (window, app, ids) => { Themes.Add(Tuple.Create(window, app, ids)); return 0; };
            VbeNativeTheme.SetAttribute = (IntPtr window, int attribute, ref int value, int size) => { Assert.AreEqual(4, size); Attributes.Add(Tuple.Create(window, attribute, value)); return attribute == 20 ? Dwm20Result : 0; };
            VbeNativeTheme.InstallWindowHook = (first, last, module, callback, process, thread, flags) =>
            { Assert.AreEqual((uint)0x8000, first); Assert.AreEqual((uint)0x8002, last); Assert.IsNotNull(callback); HookCalls++; return new IntPtr(9); };
            VbeNativeTheme.RemoveWindowHook = hook => { Assert.AreEqual(new IntPtr(9), hook); Unhooks++; return true; };
            VbeNativeTheme.Redraw = (window, update, region, flags) => { Redraws.Add(window); return true; };
            VbeNativeTheme.NativePost = (window, message, first, second) => { Posts.Add(Tuple.Create(window, message)); return PostSucceeds; };
            VbeNativeTheme.NativeSend = (window, message, first, second) => { Sends.Add(Tuple.Create(window, message)); return message == 0x31a ? 0 : send(window, message, first, second); };
            VbeNativeTheme.NativeProcedure = (window, message, first, second) => { NativeCalls++; DuringNative?.Invoke(window, message, first, second); return NativeResult; };
            VbeNativeTheme.FindChild = (window, after, className, caption) => { Assert.AreEqual("ObtbarWndClass", className); return CodeChildren.Contains(window) ? new IntPtr(10) : IntPtr.Zero; };
            VbeNativeTheme.ReadClientBounds = (IntPtr window, out VbeNativeTheme.NativeRect rectangle) => { rectangle = new VbeNativeTheme.NativeRect { Right = 8, Bottom = 8 }; return ClientSucceeds; };
            VbeNativeTheme.NewBrush = color => { IntPtr value = createBrush(color); brushes.Add(value); return value; };
            VbeNativeTheme.ReleaseBrush = value => deleteBrush(value);
            VbeNativeTheme.DrawChrome = (window, client, dc, hosted, preserve, code) => Paints.Add(Tuple.Create(window, client, dc, hosted, preserve, code));
            VbeNativeTheme.DrawBorder = window => Borders.Add(window); VbeNativeTheme.DrawCombo = window => Combos.Add(window);
            VbeNativeTheme.DrawPropertyRow = (dc, bounds) => Rows.Add(Tuple.Create(dc, bounds));
            State.Set("allowDarkMode", exports[133]); State.Set("setPreferredMode", exports[135]); State.Set("flushMenuThemes", exports[136]);
        }
        private void AddExport(int ordinal, string field, string method)
        {
            var type = typeof(VbeNativeTheme).GetField(field, BindingFlags.Static | BindingFlags.NonPublic).FieldType;
            exports[ordinal] = Delegate.CreateDelegate(type, this, GetType().GetMethod(method, BindingFlags.Instance | BindingFlags.NonPublic));
        }
        private int Preferred(int mode) { Modes.Add(mode); return 3; }
        private bool Allow(IntPtr window, bool allow) { DarkCalls.Add(Tuple.Create(window, allow)); return true; }
        private void RefreshPolicy() { RefreshPolicies++; }
        private void Flush() { Flushes++; }
        private uint Thread(IntPtr window, out uint process)
        {
            uint current = nativeThread(window, out process);
            if (ForeignThread.Contains(window)) current++;
            if (ForeignProcess.Contains(window)) process++;
            return current;
        }
        private IntPtr Ancestor(IntPtr window, uint kind)
        {
            IntPtr parent;
            if (kind == 1) return Parents.TryGetValue(window, out parent) ? parent : IntPtr.Zero;
            Assert.AreEqual((uint)2, kind);
            if (!Classes.ContainsKey(window)) return IntPtr.Zero;
            while (Parents.TryGetValue(window, out parent) && parent != IntPtr.Zero) window = parent;
            return window;
        }
        private bool IsChild(IntPtr parent, IntPtr window)
        {
            IntPtr next;
            while (Parents.TryGetValue(window, out next) && next != IntPtr.Zero) { if (next == parent) return true; window = next; }
            return false;
        }
        internal IntPtr Add(string className, IntPtr? parent = null, IntPtr? owner = null)
        {
            IntPtr nativeParent = parent ?? owner ?? Main;
            IntPtr window = CreateWindowEx(0, "STATIC", "Synthetic Theme node", parent.HasValue ? 0x40000000 : 0,
                0, 0, 8, 8, nativeParent, IntPtr.Zero, IntPtr.Zero, IntPtr.Zero);
            Assert.AreNotEqual(IntPtr.Zero, window); windows.Add(window); Classes[window] = className;
            Parents[window] = parent ?? IntPtr.Zero; Owners[window] = owner ?? Main; return window;
        }
        internal static T Invoke<T>(string method, params object[] args)
        {
            try { return (T)typeof(VbeNativeTheme).GetMethod(method, BindingFlags.Static | BindingFlags.NonPublic).Invoke(null, args); }
            catch (TargetInvocationException error) { System.Runtime.ExceptionServices.ExceptionDispatchInfo.Capture(error.InnerException).Throw(); throw; }
        }
        internal void ApplyWindow(IntPtr window) => Invoke<object>("ApplyWindow", window);
        internal IntPtr Message(IntPtr window, uint message, IntPtr first = default(IntPtr), IntPtr second = default(IntPtr)) =>
            ((VbeNativeTheme.SubclassCallback)State.Read("Subclass"))(window, message, first, second, (UIntPtr)State.Read("SubclassId"), IntPtr.Zero);
        internal void Event(uint kind, IntPtr window, int objectId = 0, int childId = 0) =>
            ((VbeNativeTheme.WinEventCallback)State.Read("WindowEvent"))(IntPtr.Zero, kind, window, objectId, childId, 0, 0);
        internal bool Belongs(IntPtr window) => Invoke<bool>("BelongsToEditor", window);
        internal bool IsCode(IntPtr window, string className) => Invoke<bool>("IsCodeSurface", window, className);
        internal void Paint(IntPtr window, string className, IntPtr dc) => Invoke<object>("PaintChrome", window, className, dc);
        internal void Queue(IntPtr container) => Invoke<object>("QueueContainerChromeRefresh", container);
        internal VbeNativePropertyTabs AddTabs(IntPtr parent)
        {
            var tabs = State.AddPropertyTabs();
            IntPtr handle = (IntPtr)typeof(VbeNativePropertyTabs).GetField("window", BindingFlags.Instance | BindingFlags.NonPublic).GetValue(tabs);
            Classes[handle] = "SysTabControl32"; Parents[handle] = parent; return tabs;
        }
        internal IntPtr TabsWindow(VbeNativePropertyTabs tabs) => (IntPtr)typeof(VbeNativePropertyTabs)
            .GetField("window", BindingFlags.Instance | BindingFlags.NonPublic).GetValue(tabs);
        public void Dispose()
        {
            foreach (IDisposable renderer in ((IDictionary)State.Read("propertyTabs")).Values) renderer.Dispose();
            foreach (IntPtr window in windows) DestroyWindow(window);
            foreach (IntPtr brush in brushes) if (GetObjectType(brush) != 0) deleteBrush(brush);
            State.Dispose(); Environment.SetEnvironmentVariable(VbeNativeTheme.LocalRefreshExperimentVariable, local);
            foreach (var value in exports.Values) GC.KeepAlive(value);
        }
    }
}

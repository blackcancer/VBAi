using System;
using System.Collections;
using System.Collections.Generic;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Threading;
using System.Windows.Forms;
using VBAi;
using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace VBAi.Tests.Unit
{
    public sealed class NativeThemePane
    {
        public int Type { get; set; }
        public long HWnd { get; set; }
        public string Caption { get; set; }
    }
    public sealed class NativeThemeMainWindow { public long HWnd { get; set; } }
    public sealed class NativeThemeVbe
    {
        public string Version { get; set; } = "7.1";
        public object[] Windows { get; set; } = new object[0];
        public NativeThemeMainWindow MainWindow { get; set; } = new NativeThemeMainWindow();
    }
    public sealed class NativeThemeBrokenVbe
    {
        public string Version => "7.1";
        public object[] Windows => throw new InvalidOperationException("synthetic pane enumeration failure");
    }

    internal sealed class NativeThemeFixture : IDisposable
    {
        private static readonly BindingFlags Flags = BindingFlags.Static | BindingFlags.NonPublic;
        private readonly Dictionary<FieldInfo, object> fields = new Dictionary<FieldInfo, object>();
        private readonly Dictionary<FieldInfo, object[]> collections = new Dictionary<FieldInfo, object[]>();
        private readonly string experiment = Environment.GetEnvironmentVariable(VbeNativeTheme.ExperimentVariable);
        private readonly Func<IntPtr, int> previousApply = VbeNativeTheme.ApplyNativeTheme;
        private readonly Func<object, IntPtr, VbeNativePalette> previousFactory = VbeNativeTheme.CreatePalette;
        private readonly Action<string, string> previousLog = LoadLog.AppendText;
        private readonly List<VbeNativePalette> palettes = new List<VbeNativePalette>();
        internal readonly NativePaletteFixture Recovery = new NativePaletteFixture();
        internal readonly List<string> Messages = new List<string>();
        internal readonly List<IntPtr> RestoredWindows = new List<IntPtr>();
        internal readonly List<int> PreferredModes = new List<int>();
        internal int Applies, Creates, Flushes;
        internal Exception ApplyFailure;
        internal Form Owner = new Form { ShowInTaskbar = false };
        private readonly List<IntPtr> nativeWindows = new List<IntPtr>();
        private readonly List<IntPtr> ownedBrushes = new List<IntPtr>();
        private readonly List<IntPtr> ownedHooks = new List<IntPtr>();
        private static readonly VbeNativeTheme.WinEventCallback NoEvents = (a, b, c, d, e, f, g) => { };

        [DllImport("gdi32.dll")] private static extern IntPtr CreateSolidBrush(int color);
        [DllImport("gdi32.dll")] private static extern bool DeleteObject(IntPtr brush);
        [DllImport("gdi32.dll")] internal static extern uint GetObjectType(IntPtr value);
        [DllImport("user32.dll", CharSet = CharSet.Unicode)] private static extern IntPtr CreateWindowEx(int exStyle, string className,
            string title, int style, int x, int y, int width, int height, IntPtr parent, IntPtr id, IntPtr instance, IntPtr parameter);
        [DllImport("user32.dll")] private static extern bool DestroyWindow(IntPtr window);
        [DllImport("user32.dll")] private static extern IntPtr SetWinEventHook(uint first, uint last, IntPtr module,
            VbeNativeTheme.WinEventCallback callback, uint process, uint thread, uint flags);
        [DllImport("user32.dll")] private static extern bool UnhookWinEvent(IntPtr hook);
        [DllImport("comctl32.dll")] private static extern bool SetWindowSubclass(IntPtr window,
            VbeNativeTheme.SubclassCallback callback, UIntPtr id, IntPtr reference);
        [DllImport("comctl32.dll", EntryPoint = "#411")] private static extern bool GetWindowSubclass(IntPtr window,
            VbeNativeTheme.SubclassCallback callback, UIntPtr id, out IntPtr reference);

        internal NativeThemeFixture()
        {
            foreach (var field in typeof(VbeNativeTheme).GetFields(Flags))
            {
                if (field.IsLiteral) continue;
                object value = field.GetValue(null);
                if (!field.IsInitOnly) { fields[field] = value; continue; }
                if (value is IList || value is IDictionary || (field.FieldType.IsGenericType && field.FieldType.GetGenericTypeDefinition() == typeof(HashSet<>)))
                {
                    var items = new List<object>();
                    if (value is IDictionary) foreach (DictionaryEntry item in (IDictionary)value) items.Add(item);
                    else foreach (object item in (IEnumerable)value) items.Add(item);
                    collections[field] = items.ToArray(); field.FieldType.GetMethod("Clear").Invoke(value, null);
                }
            }
            var stateNames = new HashSet<string> { "toolbarPaintCount", "toolbarDeferredPaintCount", "localChromeRefresh",
                "allowDarkMode", "setPreferredMode", "flushMenuThemes", "backgroundBrush", "editorWindow", "previousPreferredMode",
                "windowEventHook", "preferredModeChanged", "nativePalette", "immediateWindow", "immediateCaption" };
            foreach (var field in fields.Keys) if (stateNames.Contains(field.Name))
                field.SetValue(null, field.FieldType.IsValueType ? Activator.CreateInstance(field.FieldType) : null);
            Environment.SetEnvironmentVariable(VbeNativeTheme.ExperimentVariable, null);
            LoadLog.AppendText = (path, message) => Messages.Add(message);
            VbeNativeTheme.ApplyNativeTheme = window =>
            { Applies++; Add("themedWindows", window); if (ApplyFailure != null) throw ApplyFailure; return 1; };
            VbeNativeTheme.CreatePalette = (vbe, window) =>
            { Creates++; var palette = new VbeNativePalette(vbe, window, Recovery.PathName); palettes.Add(palette); return palette; };
        }

        internal static void OnSta(Action action)
        {
            Exception failure = null;
            var thread = new Thread(() => { try { action(); } catch (Exception error) { failure = error; } });
            thread.SetApartmentState(ApartmentState.STA); thread.Start();
            Assert.IsTrue(thread.Join(15000), "Theme orchestration fixture timed out.");
            if (failure != null) throw failure;
        }
        internal IntPtr Handle => Owner.Handle;
        internal object Read(string name) => typeof(VbeNativeTheme).GetField(name, Flags).GetValue(null);
        internal void Set(string name, object value) => typeof(VbeNativeTheme).GetField(name, Flags).SetValue(null, value);
        internal void Add(string name, IntPtr window)
        { var value = Read(name); value.GetType().GetMethod("Add", new[] { typeof(IntPtr) }).Invoke(value, new object[] { window }); }
        internal int Count(string name) => (int)Read(name).GetType().GetProperty("Count").GetValue(Read(name), null);
        internal VbeNativePalette Palette => (VbeNativePalette)Read("nativePalette");
        internal static object PaletteField(VbeNativePalette palette, string name) => typeof(VbeNativePalette)
            .GetField(name, BindingFlags.Instance | BindingFlags.NonPublic).GetValue(palette);
        internal void ConfigureCallbacks(bool allow = true, bool preferred = true, bool flush = true)
        {
            Configure("allowDarkMode", allow ? "Allow" : null); Configure("setPreferredMode", preferred ? "Preferred" : null);
            Configure("flushMenuThemes", flush ? "Flush" : null);
        }
        private void Configure(string name, string method)
        {
            var field = typeof(VbeNativeTheme).GetField(name, Flags);
            field.SetValue(null, method == null ? null : Delegate.CreateDelegate(field.FieldType, this,
                GetType().GetMethod(method, BindingFlags.Instance | BindingFlags.NonPublic)));
        }
        private bool Allow(IntPtr window, bool enabled) { Assert.IsFalse(enabled); RestoredWindows.Add(window); return true; }
        private int Preferred(int mode) { PreferredModes.Add(mode); return 0; }
        private void Flush() { Flushes++; }
        internal IntPtr AddDialog()
        {
            IntPtr window = CreateWindowEx(0, "#32770", "Synthetic theme fixture", 0x40000000, 0, 0, 8, 8,
                Handle, IntPtr.Zero, IntPtr.Zero, IntPtr.Zero);
            Assert.AreNotEqual(IntPtr.Zero, window); nativeWindows.Add(window); return window;
        }
        [StructLayout(LayoutKind.Sequential, CharSet = CharSet.Unicode)]
        private struct TabItem
        {
            internal uint Mask, State, StateMask;
            [MarshalAs(UnmanagedType.LPWStr)] internal string Text;
            internal int Capacity, Image;
            internal IntPtr Data;
        }
        [DllImport("user32.dll", EntryPoint = "SendMessageW", CharSet = CharSet.Unicode)]
        private static extern IntPtr InsertTab(IntPtr window, uint message, IntPtr index, ref TabItem item);
        internal VbeNativePropertyTabs AddPropertyTabs()
        {
            using (var initialize = new TabControl()) { var unused = initialize.Handle; }
            IntPtr window = CreateWindowEx(0, "SysTabControl32", "Synthetic properties tabs", 0x40000000,
                0, 0, 240, 60, Handle, IntPtr.Zero, IntPtr.Zero, IntPtr.Zero);
            Assert.AreNotEqual(IntPtr.Zero, window); nativeWindows.Add(window);
            var item = new TabItem { Mask = 1, Text = "Alphabetic" };
            Assert.AreEqual(0, InsertTab(window, 0x133e, IntPtr.Zero, ref item).ToInt32());
            item.Text = "Categorized";
            Assert.AreEqual(1, InsertTab(window, 0x133e, new IntPtr(1), ref item).ToInt32());
            VbeNativePropertyTabs tabs; Assert.IsTrue(VbeNativePropertyTabs.TryCreate(window, out tabs));
            ((IDictionary)Read("propertyTabs")).Add(window, tabs); return tabs;
        }

        internal IntPtr AddBrush()
        {
            IntPtr brush = CreateSolidBrush(0x00332211); Assert.AreNotEqual(IntPtr.Zero, brush);
            ownedBrushes.Add(brush); Set("backgroundBrush", brush); return brush;
        }
        internal IntPtr AddEventHook()
        {
            IntPtr hook = SetWinEventHook(0x8000, 0x8002, IntPtr.Zero, NoEvents,
                (uint)System.Diagnostics.Process.GetCurrentProcess().Id, 0, 0);
            Assert.AreNotEqual(IntPtr.Zero, hook); ownedHooks.Add(hook); Set("windowEventHook", hook); return hook;
        }
        internal void AddSubclass(IntPtr window)
        {
            Assert.IsTrue(SetWindowSubclass(window, (VbeNativeTheme.SubclassCallback)Read("Subclass"), (UIntPtr)Read("SubclassId"), IntPtr.Zero));
            Add("subclassedWindows", window);
        }
        internal bool IsSubclassed(IntPtr window)
        { IntPtr unused; return GetWindowSubclass(window, (VbeNativeTheme.SubclassCallback)Read("Subclass"), (UIntPtr)Read("SubclassId"), out unused); }
        public void Dispose()
        {
            foreach (var palette in palettes) palette.Dispose();
            foreach (var hook in ownedHooks) UnhookWinEvent(hook);
            foreach (var brush in ownedBrushes) if (GetObjectType(brush) != 0) DeleteObject(brush);
            foreach (var window in nativeWindows) DestroyWindow(window);
            Owner.Dispose(); Recovery.Dispose();
            foreach (var pair in fields) pair.Key.SetValue(null, pair.Value);
            foreach (var pair in collections)
            {
                object value = pair.Key.GetValue(null); var type = value.GetType(); type.GetMethod("Clear").Invoke(value, null);
                foreach (var item in pair.Value)
                {
                    if (value is IDictionary) { var entry = (DictionaryEntry)item; ((IDictionary)value).Add(entry.Key, entry.Value); }
                    else type.GetMethod("Add").Invoke(value, new[] { item });
                }
            }
            VbeNativeTheme.ApplyNativeTheme = previousApply; VbeNativeTheme.CreatePalette = previousFactory;
            Environment.SetEnvironmentVariable(VbeNativeTheme.ExperimentVariable, experiment); LoadLog.AppendText = previousLog;
        }
    }
}

using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Reflection;
using System.Runtime.InteropServices;
using CodexVBE;
using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace CodexVBE.Tests.Infrastructure
{
    public sealed class NativeOtherHostsFixture : IDisposable
    {
        internal readonly VbeProjectComponents.NativeOtherHostProbe Native = new VbeProjectComponents.NativeOtherHostProbe();
        internal readonly ApplicationContract Application = new ApplicationContract();
        internal readonly Dictionary<long, uint> Owners = new Dictionary<long, uint>();
        internal readonly Dictionary<long, string> Classes = new Dictionary<long, string>();
        internal readonly Dictionary<long, object> Accessible = new Dictionary<long, object>();
        internal readonly List<long> Roots = new List<long> { 10, 11 };
        internal readonly List<long> Children = new List<long> { 20, 21, 22, 23 };
        internal readonly Dictionary<FieldInfo, object> Saved = new Dictionary<FieldInfo, object>();
        internal string Kind = "Word";
        internal int AccessibilityError, AccessibleCalls;
        internal NativeOtherHostsFixture()
        {
            uint pid = (uint)Process.GetCurrentProcess().Id;
            Owners[1] = Owners[10] = Owners[20] = Owners[21] = Owners[23] = pid; Owners[11] = Owners[22] = 999999; Owners[2] = 0;
            foreach (string name in new[] { "EnumWindows", "EnumChildWindows", "GetWindowThreadProcessId", "GetClassName", "AccessibleObjectFromWindow" })
            {
                var field = typeof(VbeDebugWindows).GetField(name, BindingFlags.Static | BindingFlags.NonPublic); Saved.Add(field, field.GetValue(null));
            }
            Classes[20] = Classes[22] = Classes[23] = "_WwG"; Classes[21] = "Other";
            Application.HWND = 1; Application.Windows.Add(new WindowContract { Hwnd = 1 });
            Accessible[20] = Accessible[23] = new AutomationContract { Application = Application };
            Native.ReadHostKind = () => Kind;
            Native.ReadOwner = window => Owners.TryGetValue(window.ToInt64(), out uint owner) ? owner : 0;
            Native.ReadActiveApplication = name => { Assert.AreEqual(Kind == "Word" ? "Word.Application" : "PowerPoint.Application", name); return Application; };
            VbeDebugWindows.EnumWindows = (callback, parameter) => { foreach (long root in Roots) callback(new IntPtr(root), parameter); return true; };
            VbeDebugWindows.EnumChildWindows = (root, callback, parameter) => { Assert.AreEqual(new IntPtr(10), root); foreach (long child in Children) callback(new IntPtr(child), parameter); return true; };
            VbeDebugWindows.GetWindowThreadProcessId = (IntPtr window, out uint owner) => { owner = Owners.TryGetValue(window.ToInt64(), out uint value) ? value : 0; return 1; };
            VbeDebugWindows.GetClassName = (window, text, capacity) => { string name = Classes[window.ToInt64()]; text.Append(name); return name.Length; };
            VbeDebugWindows.AccessibleObjectFromWindow = (IntPtr window, uint id, ref Guid iid, out object value) =>
            {
                Assert.AreEqual(0xFFFFFFF0u, id); Assert.AreEqual(new Guid("00020400-0000-0000-C000-000000000046"), iid); AccessibleCalls++;
                value = Accessible.TryGetValue(window.ToInt64(), out object item) ? item : null;
                if (value is Exception error) throw error;
                return window == new IntPtr(20) ? AccessibilityError : 0;
            };
        }
        public void Dispose() { foreach (var pair in Saved) pair.Key.SetValue(null, pair.Value); }
        public sealed class WindowContract { public long Hwnd { get; set; } }
        public sealed class AutomationContract { public object Application { get; set; } }
        public sealed class ApplicationContract
        {
            public long HWND { get; set; }
            public List<WindowContract> Windows { get; } = new List<WindowContract>();
            public List<object> Documents { get; } = new List<object>();
            public List<object> Presentations { get; } = new List<object>();
        }
        public sealed class DocumentContract
        {
            public object VBProject { get; set; }
            public string Path { get; set; } = @"C:\Temp";
            public string FullName { get; set; } = @"C:\Temp\Document.docm";
            public object ReadOnly { get; set; } = false;
            public object Saved { get; set; } = true;
            public int SaveFormat { get; set; } = 13;
            public int SaveCalls, WordSaveAsCalls, PowerPointSaveAsCalls, Format;
            public string Destination;
            public void Save() { SaveCalls++; }
            public void SaveAs2(string destination, int format) { WordSaveAsCalls++; Destination = destination; Format = format; }
            public void SaveAs(string destination, int format) { PowerPointSaveAsCalls++; Destination = destination; Format = format; }
        }
    }
}

namespace VBAi.Tests.Unit
{
    using System;
    using System.Collections.Generic;
    using System.Diagnostics;
    using System.Runtime.InteropServices;
    using VBAi;

    public sealed partial class VbeProjectProtectionTests
    {
        private sealed class Win32ProtectionScope : IDisposable
        {
            private readonly Func<VbeDebugWindows.EnumWindowCallback, IntPtr, bool> enumeration = VbeDebugWindows.EnumWindows;
            private readonly Func<IntPtr, VbeDebugWindows.EnumWindowCallback, IntPtr, bool> children = VbeDebugWindows.EnumChildWindows;
            private readonly VbeDebugWindows.WindowProcessReader owner = VbeDebugWindows.GetWindowThreadProcessId;
            private readonly Func<IntPtr, System.Text.StringBuilder, int, int> classes = VbeDebugWindows.GetClassName;
            private readonly Func<IntPtr, System.Text.StringBuilder, int, int> titles = VbeDebugWindows.GetWindowText;
            private readonly Func<IntPtr, bool> visible = VbeDebugWindows.IsWindowVisible, enabled = VbeDebugWindows.OptionsWindowEnabled;
            private readonly Func<IntPtr, int> ids = VbeDebugWindows.GetDlgCtrlID;
            private readonly Func<IntPtr, int, IntPtr> items = VbeDebugWindows.GetDlgItem;
            private readonly Func<IntPtr, int, IntPtr, IntPtr, IntPtr> messages = VbeDebugWindows.SendMessageInt;
            private readonly Func<IntPtr, int, IntPtr, string, IntPtr> text = VbeDebugWindows.SendMessageText;
            private readonly Func<IntPtr, int, IntPtr, IntPtr, bool> post = VbeDebugWindows.PostMessage;
            private readonly Func<IntPtr, long> styles = VbeDebugWindows.ProtectionWindowStyle;
            private readonly Func<IntPtr, IntPtr> parents = VbeDebugWindows.ProtectionWindowParent;
            private readonly Func<IntPtr, int, string> tabText = VbeDebugWindows.ProtectionTabText;
            private readonly Func<IntPtr, IntPtr, int, bool> selectTab = VbeDebugWindows.ProtectionSelectTab;
            private readonly Func<int, IntPtr> allocate = VbeDebugWindows.AllocateProtectionBuffer;
            private readonly Action<int> pause = VbeDebugWindows.PauseNative;
            internal readonly VbeDebugWindows.IProjectProtectionProbe Native;
            internal int Selected, Locked, Writes, Posts, TabReads;
            internal bool DuplicateTab, FailTabRead, FailSelect, WrongCheckbox, MissingPassword, WrongOwner, WrongStyle, IgnoreClick, IgnorePassword;
            internal int Length1, Length2;
            private string Kind(IntPtr h)
            {
                int id = h.ToInt32();
                return id == 73 ? "SysTabControl32" : id == 74 || id == 77 || id == 78 ? "Button" :
                    id == 75 || id == 76 ? "Edit" : "#32770";
            }
            internal Win32ProtectionScope()
            {
                var type = typeof(VbeDebugWindows).GetNestedType("NativeProjectProtectionProbe", System.Reflection.BindingFlags.NonPublic);
                Native = (VbeDebugWindows.IProjectProtectionProbe)Activator.CreateInstance(type, true);
                VbeDebugWindows.EnumWindows = (callback, data) => { callback(new IntPtr(71), data); return true; };
                VbeDebugWindows.EnumChildWindows = (root, callback, data) =>
                { foreach (int id in new[] { 72, 73, 74, 75, 76, 77, 78 }) if (!MissingPassword || id != 76) callback(new IntPtr(id), data); return true; };
                VbeDebugWindows.GetWindowThreadProcessId = (IntPtr h, out uint pid) =>
                { pid = (uint)Process.GetCurrentProcess().Id + (WrongOwner && h.ToInt32() == 73 ? 1u : 0u); return 1; };
                VbeDebugWindows.GetClassName = (h, buffer, length) => { buffer.Append(Kind(h)); return buffer.Length; };
                VbeDebugWindows.GetWindowText = (h, buffer, length) =>
                { buffer.Append(h.ToInt32() == 71 ? "ExactProject - Propriétés du projet" : h.ToInt32() == 74 ? "&Verrouiller le projet pour l'affichage" : ""); return buffer.Length; };
                VbeDebugWindows.IsWindowVisible = h => true;
                VbeDebugWindows.OptionsWindowEnabled = h => true;
                VbeDebugWindows.GetDlgCtrlID = h => h.ToInt32() == 74 ? WrongCheckbox ? 999 : 5463 :
                    h.ToInt32() == 75 ? 5461 : h.ToInt32() == 76 ? 5462 : h.ToInt32() == 77 ? 1 : h.ToInt32() == 78 ? 2 : 12320;
                VbeDebugWindows.GetDlgItem = (h, id) => new IntPtr(id == 1 ? 77 : 78);
                VbeDebugWindows.ProtectionWindowParent = h => new IntPtr(h.ToInt32() == 74 || h.ToInt32() == 75 || h.ToInt32() == 76 ? 72 : 71);
                VbeDebugWindows.ProtectionWindowStyle = h => h.ToInt32() == 74 ? 3 : h.ToInt32() == 75 || h.ToInt32() == 76 ? WrongStyle ? 0 : 0x20 : 0;
                VbeDebugWindows.SendMessageInt = Message;
                VbeDebugWindows.SendMessageText = (h, m, w, value) =>
                { Writes++; if (!IgnorePassword) { if (h.ToInt32() == 75) Length1 = value.Length; else Length2 = value.Length; } return IntPtr.Zero; };
                VbeDebugWindows.PostMessage = (h, m, w, l) => { Posts++; return true; };
            }
            private IntPtr Message(IntPtr h, int message, IntPtr w, IntPtr l)
            {
                if (message == 0x1304) return new IntPtr(2);
                if (message == 0x130b) return new IntPtr(Selected);
                if (message == 0x0465) { if (FailSelect) return IntPtr.Zero; Selected = w.ToInt32(); return new IntPtr(1); }
                if (message == 0x133c)
                {
                    TabReads++;
                    if (FailTabRead) return IntPtr.Zero;
                    var type = typeof(VbeDebugWindows).GetNestedType("ProtectionTabItem", System.Reflection.BindingFlags.NonPublic);
                    object item = Marshal.PtrToStructure(l, type);
                    IntPtr buffer = (IntPtr)type.GetField("Text").GetValue(item);
                    string title = w.ToInt32() == 1 || DuplicateTab ? "Protection" : "Général";
                    for (int i = 0; i < title.Length; i++) Marshal.WriteInt16(buffer, i * 2, (short)title[i]);
                    Marshal.WriteInt16(buffer, title.Length * 2, 0);
                    return new IntPtr(1);
                }
                if (message == 0xf0) return new IntPtr(Locked);
                if (message == 0xf5) { if (!IgnoreClick) Locked = Locked == 0 ? 1 : 0; return IntPtr.Zero; }
                if (message == 0xe) return new IntPtr(h.ToInt32() == 75 ? Length1 : Length2);
                return IntPtr.Zero;
            }
            public void Dispose()
            {
                VbeDebugWindows.EnumWindows = enumeration; VbeDebugWindows.EnumChildWindows = children;
                VbeDebugWindows.GetWindowThreadProcessId = owner; VbeDebugWindows.GetClassName = classes; VbeDebugWindows.GetWindowText = titles;
                VbeDebugWindows.IsWindowVisible = visible; VbeDebugWindows.OptionsWindowEnabled = enabled;
                VbeDebugWindows.GetDlgCtrlID = ids; VbeDebugWindows.GetDlgItem = items;
                VbeDebugWindows.SendMessageInt = messages; VbeDebugWindows.SendMessageText = text; VbeDebugWindows.PostMessage = post;
                VbeDebugWindows.ProtectionWindowStyle = styles; VbeDebugWindows.ProtectionWindowParent = parents;
                VbeDebugWindows.ProtectionTabText = tabText; VbeDebugWindows.ProtectionSelectTab = selectTab;
                VbeDebugWindows.AllocateProtectionBuffer = allocate; VbeDebugWindows.PauseNative = pause;
            }
        }
    }
}

using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Windows.Forms;

namespace VBAi.Tests.Infrastructure
{
    /// <summary>Fenêtres natives possédées par le test ; les messages synthétiques ne pilotent aucun VBE.</summary>
    public sealed class EditorNavigationFixture : IDisposable
    {
        internal readonly Form Dispatcher = new Form();
        internal readonly System.Windows.Forms.NativeWindow Root = new System.Windows.Forms.NativeWindow(), Pane = new System.Windows.Forms.NativeWindow(), Nested = new System.Windows.Forms.NativeWindow();
        internal readonly OwnedTree Tree = new OwnedTree();
        internal readonly TreeReplies Replies;
        internal readonly Host Vbe = new Host();
        internal const string Caption = "VBAi owned project fixture";
        private static readonly WindowProcedure procedure = DefWindowProc;
        private static readonly object registration = new object();
        private static bool registered;

        internal EditorNavigationFixture()
        {
            lock (registration)
            {
                if (!registered)
                {
                    var value = new WindowClass { ClassName = "wndclass_desked_gsk", Procedure = Marshal.GetFunctionPointerForDelegate(procedure), Instance = GetModuleHandle(null) };
                    if (RegisterClass(ref value) == 0 && Marshal.GetLastWin32Error() != 1410) throw new Win32Exception();
                    registered = true;
                }
            }
            var dispatch = Dispatcher.Handle;
            IntPtr root = CreateWindowEx(0, "wndclass_desked_gsk", "Owned test root", 0, 0, 0, 40, 40, IntPtr.Zero, IntPtr.Zero, GetModuleHandle(null), IntPtr.Zero);
            if (root == IntPtr.Zero) throw new Win32Exception();
            Root.AssignHandle(root);
            Pane.CreateHandle(new CreateParams { Caption = Caption, Parent = Root.Handle, Style = 0x40000000 });
            Nested.CreateHandle(new CreateParams { Caption = Caption, Parent = Pane.Handle, Style = 0x40000000 });
            var controls = new CommonControls { Size = 8, Classes = 2 }; InitCommonControlsEx(ref controls);
            IntPtr tree = CreateWindowEx(0, "SysTreeView32", "Owned tree", 0x40000000, 0, 0, 40, 40, Nested.Handle, IntPtr.Zero, GetModuleHandle(null), IntPtr.Zero);
            if (tree == IntPtr.Zero) throw new Win32Exception();
            Tree.AssignHandle(tree);
            Replies = new TreeReplies(Tree.Handle);
            Vbe.Windows.Add(new ProjectWindow { Type = 6, Caption = Caption });
        }

        internal static object Get(object value, string name) => value.GetType().GetField(name, BindingFlags.Instance | BindingFlags.NonPublic).GetValue(value);
        internal static object Invoke(object value, string name) => value.GetType().GetMethod(name, BindingFlags.Instance | BindingFlags.NonPublic).Invoke(value, null);
        internal void DoubleClick() => SendMessage(Tree.Handle, 0x0203, IntPtr.Zero, IntPtr.Zero);
        public void Dispose()
        {
            Replies.ReleaseHandle(); Tree.Dispose(); Nested.DestroyHandle(); Pane.DestroyHandle(); Root.DestroyHandle(); Dispatcher.Dispose();
        }

        public sealed class Host
        {
            public readonly List<ProjectWindow> Windows = new List<ProjectWindow>();
            public readonly List<object> VBProjects = new List<object>();
            public EditorVbeContract.Window MainWindow { get; } = new EditorVbeContract.Window();
            public bool FailSelection;
            public object Selection;
            public object SelectedVBComponent => FailSelection ? throw new COMException("Selection unavailable") : Selection;
        }
        public sealed class ProjectWindow { public int Type { get; set; } public string Caption { get; set; } }
        internal sealed class OwnedTree : System.Windows.Forms.NativeWindow, IDisposable { public void Dispose() { if (Handle != IntPtr.Zero) DestroyHandle(); } }
        internal sealed class TreeReplies : System.Windows.Forms.NativeWindow
        {
            internal uint Flags = 0x4;
            internal IntPtr Item = new IntPtr(1), Selected = new IntPtr(1);
            internal int Hits;
            internal TreeReplies(IntPtr handle) { AssignHandle(handle); }
            protected override void WndProc(ref Message message)
            {
                if (message.Msg == 0x1111)
                {
                    var value = (Hit)Marshal.PtrToStructure(message.LParam, typeof(Hit));
                    value.Flags = Flags; value.Item = Item; Marshal.StructureToPtr(value, message.LParam, false);
                    Hits++; message.Result = Item; return;
                }
                if (message.Msg == 0x110A && message.WParam == new IntPtr(9)) { message.Result = Selected; return; }
                base.WndProc(ref message);
            }
        }
        [StructLayout(LayoutKind.Sequential)] private struct Hit { internal int X, Y; internal uint Flags; internal IntPtr Item; }
        [StructLayout(LayoutKind.Sequential)] private struct CommonControls { internal int Size, Classes; }
        [StructLayout(LayoutKind.Sequential, CharSet = CharSet.Unicode)] private struct WindowClass
        {
            internal uint Style; internal IntPtr Procedure; internal int ClassExtra, WindowExtra;
            internal IntPtr Instance, Icon, Cursor, Background; internal string MenuName, ClassName;
        }
        private delegate IntPtr WindowProcedure(IntPtr window, uint message, IntPtr first, IntPtr second);
        [DllImport("user32.dll", CharSet = CharSet.Unicode, SetLastError = true)] private static extern ushort RegisterClass(ref WindowClass value);
        [DllImport("kernel32.dll", CharSet = CharSet.Unicode)] private static extern IntPtr GetModuleHandle(string name);
        [DllImport("user32.dll", CharSet = CharSet.Unicode)] private static extern IntPtr DefWindowProc(IntPtr window, uint message, IntPtr first, IntPtr second);
        [DllImport("comctl32.dll")] private static extern bool InitCommonControlsEx(ref CommonControls value);
        [DllImport("user32.dll", CharSet = CharSet.Unicode, SetLastError = true)] private static extern IntPtr CreateWindowEx(int extendedStyle, string className, string title, int style, int x, int y, int width, int height, IntPtr parent, IntPtr menu, IntPtr instance, IntPtr data);
        [DllImport("user32.dll")] private static extern IntPtr SendMessage(IntPtr window, int message, IntPtr first, IntPtr second);
    }
}

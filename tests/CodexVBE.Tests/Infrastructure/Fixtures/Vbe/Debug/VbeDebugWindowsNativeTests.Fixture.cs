namespace CodexVBE.Tests.Unit
{
    using System;
    using System.Collections.Generic;
    using System.Linq;
    using System.Threading;
    using CodexVBE;
    using Microsoft.VisualStudio.TestTools.UnitTesting;

    public sealed partial class VbeDebugWindowsNativeTests
    {
        private static FakeNative DialogWith(string message, string button)
        {
            var native = new FakeNative
            {
                DialogHandle = new IntPtr(2),
                AccessibleMessage = message
            };
            native.Controls.Add(Control(3, "Static", message, true));
            native.Controls.Add(Control(4, "Button", button, true));
            return native;
        }

        private static VbeDebugWindows.NativeControl Control(int handle, string kind, string text, bool visible)
        {
            return new VbeDebugWindows.NativeControl
            {
                Handle = new IntPtr(handle),
                Kind = kind,
                Text = text,
                Visible = visible
            };
        }

        private sealed class FakeNative : VbeDebugWindows.INativeProbe
        {
            public IntPtr Root;
            public IntPtr DialogHandle;
            public readonly Dictionary<string, IntPtr> Panes = new Dictionary<string, IntPtr>();
            public readonly List<VbeDebugWindows.NativeControl> Controls = new List<VbeDebugWindows.NativeControl>();
            public readonly List<IntPtr> ListReads = new List<IntPtr>();
            public IntPtr ImmediateHandle;
            public IntPtr CallStackHandle;
            public int CallStackReads;
            public int Clicks;
            public bool ClickSucceeds = true;
            public bool CloseAfterClick;
            public int PausesSinceReset;
            public string AccessibleMessage;
            public Action<int> OnPause;
            public IntPtr VbeRoot()
            {
                return Root;
            }

            public List<IntPtr> Children(IntPtr root)
            {
                return Panes.Values.ToList();
            }

            public IntPtr Pane(IEnumerable<IntPtr> panes, params string[] names)
            {
                foreach (string name in names)
                    if (Panes.TryGetValue(name, out IntPtr handle))
                        return handle;
                return IntPtr.Zero;
            }

            public object List(IntPtr handle)
            {
                ListReads.Add(handle);
                return new
                {
                    Available = handle != IntPtr.Zero
                };
            }

            public object Immediate(IntPtr handle)
            {
                ImmediateHandle = handle;
                return new
                {
                    Available = handle != IntPtr.Zero
                };
            }

            public object CallStack(IntPtr locals)
            {
                CallStackHandle = locals;
                CallStackReads++;
                return new
                {
                    Available = locals != IntPtr.Zero
                };
            }

            public IntPtr Dialog(params string[] titles)
            {
                return DialogHandle;
            }

            public List<VbeDebugWindows.NativeControl> DialogControls(IntPtr dialog)
            {
                return Controls;
            }

            public string DialogMessage(IntPtr dialog)
            {
                return AccessibleMessage;
            }

            public bool Click(IntPtr handle)
            {
                Clicks++;
                return ClickSucceeds;
            }

            public bool Visible(IntPtr handle)
            {
                return !(CloseAfterClick && Clicks > 0);
            }

            public void Pause(int milliseconds)
            {
                PausesSinceReset++;
                OnPause?.Invoke(PausesSinceReset);
            }
        }
    }
}

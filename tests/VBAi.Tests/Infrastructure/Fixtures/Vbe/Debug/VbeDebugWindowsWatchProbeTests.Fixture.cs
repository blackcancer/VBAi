namespace VBAi.Tests.Unit
{
    using System;
    using System.Collections.Generic;
    using VBAi;

    public sealed partial class VbeDebugWindowsWatchProbeTests
    {
        private static Request Add()
        {
            return new Request
            {
                Project = "Book",
                Module = "Module1",
                Procedure = "Run",
                Expression = "x",
                WatchType = "break_when_true"
            };
        }

        private static Request Edit()
        {
            return new Request
            {
                Project = "Book",
                Module = "Module1",
                Expression = "x",
                NewExpression = "y",
                Context = "Module1.Run",
                WatchType = "break_when_changed"
            };
        }

        private static WatchFake QuickFake()
        {
            var fake = new WatchFake();
            fake.Texts[4751] = "x";
            fake.Texts[4752] = "42";
            fake.Texts[4753] = "Book.Module1.Run";
            return fake;
        }

        private sealed class WatchFake : VbeDebugWindows.IWatchProbe
        {
            public readonly Dictionary<int, string> Texts = new Dictionary<int, string>
            {
                {
                    4853,
                    "x"
                },
                {
                    4858,
                    "Book"
                },
                {
                    4857,
                    "Module1"
                },
                {
                    4856,
                    "Run"
                }
            };
            public bool Open = true, CloseAfterOk = true, Check = true, ReplaceEcho = true, ErrorOpen;
            public int Missing, RefuseClick, Closes, Pauses, NewMatches, OldMatches, WatchReadAttempts, SelectAttempts, DisappearAfter;
            public bool SelectSucceeds = true;
            public IntPtr Root, PaneHandle, ListHandle;
            private bool clickedOk;
            public IntPtr Dialog(params string[] titles)
            {
                if (titles[0].StartsWith("Microsoft"))
                    return ErrorOpen && clickedOk ? new IntPtr(9) : IntPtr.Zero;
                return Open && !(clickedOk && CloseAfterOk) ? new IntPtr(2) : IntPtr.Zero;
            }

            public IntPtr Item(IntPtr dialog, int id)
            {
                return id == Missing ? IntPtr.Zero : new IntPtr(id);
            }

            public string Text(IntPtr handle)
            {
                return Texts.TryGetValue(handle.ToInt32(), out string value) ? value : "";
            }

            public bool Click(IntPtr handle)
            {
                if (handle.ToInt32() == RefuseClick)
                    return false;
                if (handle.ToInt32() == 1)
                    clickedOk = true;
                return true;
            }

            public bool Checked(IntPtr handle)
            {
                return Check;
            }

            public void Replace(IntPtr handle, string value)
            {
                if (ReplaceEcho)
                    Texts[handle.ToInt32()] = value;
            }

            public void Pause(int milliseconds)
            {
                Pauses++;
            }

            public string Message(IntPtr dialog)
            {
                return "invalid expression";
            }

            public void Close(IntPtr dialog)
            {
                Closes++;
                if (dialog.ToInt32() == 2)
                    Open = false;
            }

            public IntPtr VbeRoot()
            {
                return Root;
            }

            public List<IntPtr> Children(IntPtr root)
            {
                return new List<IntPtr>
                {
                    PaneHandle
                };
            }

            public IntPtr Pane(IEnumerable<IntPtr> panes, params string[] names)
            {
                return PaneHandle;
            }

            public object List(IntPtr handle)
            {
                ListHandle = handle;
                return new
                {
                    Available = handle != IntPtr.Zero
                };
            }

            public int WatchMatches(IntPtr pane, string expression, string context)
            {
                WatchReadAttempts++;
                if (DisappearAfter > 0 && WatchReadAttempts >= DisappearAfter)
                    return 0;
                return expression == "y" ? NewMatches : OldMatches;
            }

            public bool SelectWatchRow(IntPtr pane, string expression, string context)
            {
                SelectAttempts++;
                return SelectSucceeds;
            }
        }
    }
}

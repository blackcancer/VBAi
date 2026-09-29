namespace VBAi.Tests.Unit
{
    using System;
    using System.Collections.Generic;
    using VBAi;
    using Microsoft.VisualStudio.TestTools.UnitTesting;

    public sealed partial class VbeDebugWindowsImmediateProbeTests
    {
        private static ImmediateFake Ready(string command)
        {
            var fake = new ImmediateFake
            {
                Root = new IntPtr(1),
                PaneHandle = new IntPtr(2)
            };
            fake.Readbacks.Add("> " + command);
            return fake;
        }

        private sealed class ImmediateFake : VbeDebugWindows.IImmediateProbe
        {
            public IntPtr Root, PaneHandle;
            public int RootReads, Prepares, Enters, Pauses;
            public char RejectChar;
            public bool EnterSucceeds = true;
            public readonly List<string> Readbacks = new List<string>();
            private int readIndex;
            private string last = "> ";
            public IntPtr VbeRoot()
            {
                RootReads++;
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

            public string Prepare(IntPtr pane)
            {
                Prepares++;
                return "> ";
            }

            public string Text(IntPtr pane)
            {
                if (readIndex < Readbacks.Count)
                    last = Readbacks[readIndex++];
                return last;
            }

            public bool PostChar(IntPtr pane, char character)
            {
                return character != RejectChar;
            }

            public bool PostEnter(IntPtr pane)
            {
                Enters++;
                return EnterSucceeds;
            }

            public void Pause(int milliseconds)
            {
                Pauses++;
            }
        }
    }
}

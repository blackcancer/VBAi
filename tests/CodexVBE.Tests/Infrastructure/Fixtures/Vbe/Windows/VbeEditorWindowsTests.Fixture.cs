namespace CodexVBE.Tests.Unit
{
    using System;
    using System.Collections.Generic;
    using CodexVBE;
    using Microsoft.VisualStudio.TestTools.UnitTesting;

    public sealed partial class VbeEditorWindowsTests
    {
        public sealed class FakeVbe
        {
            public string Version { get; set; } = "7.1";
            public List<FakeWindow> Windows { get; } = new List<FakeWindow>();
            public List<object> CodePanes { get; } = new List<object>();
            public List<FakeAddIn> AddIns { get; } = new List<FakeAddIn>();
            public List<FakeProject> VBProjects { get; } = new List<FakeProject>();
            public FakeProject ActiveVBProject { get; set; }
            public FakeWindow ActiveWindow { get; set; }
            public object ActiveCodePane { get; set; }
        }

        public sealed class FakeProject
        {
            public string Name { get; set; }
        }

        public sealed class FakePaneProject
        {
            public string Name { get; set; }
            public string FileName { get; set; }
        }

        public sealed class FakePaneCollection
        {
            public FakePaneProject Parent { get; set; }
        }

        public sealed class FakePaneComponent
        {
            public string Name { get; set; }
            public FakePaneCollection Collection { get; set; }
        }

        public sealed class FakePaneModule
        {
            public FakePaneComponent Parent { get; set; }
        }

        public sealed class FakePaneWindow
        {
            public string Caption { get; set; }
        }

        public sealed class FakePane
        {
            public FakePaneModule CodeModule { get; set; }
            public int CodePaneView { get; set; } = 1;
            public int TopLine { get; set; }
            public int CountOfVisibleLines { get; set; }
            public FakePaneWindow Window { get; set; }
            public bool SelectionThrows { get; set; }
            public int SelectionReads { get; private set; }

            public void GetSelection(ref int startLine, ref int startColumn, ref int endLine, ref int endColumn)
            {
                SelectionReads++;
                if (SelectionThrows)
                    throw new InvalidOperationException("Selection unavailable");
                startLine = 2;
                startColumn = 1;
                endLine = 4;
                endColumn = 8;
            }
        }

        public sealed class FakeAddIn
        {
            public string ProgId { get; set; }
            public string Guid { get; set; }
            public string Description { get; set; }
            public bool Connect { get; set; }
        }

        public sealed class FakeFrame
        {
            public string Caption { get; set; }
            public List<FakeWindow> LinkedWindows { get; set; }
        }

        public sealed class FakeWindow
        {
            private readonly FakeVbe host;
            public FakeWindow(FakeVbe host)
            {
                this.host = host;
            }

            public string Caption { get; set; }
            public int Type { get; set; }
            public bool Visible { get; set; }
            public int WindowState { get; set; }
            public int Left { get; set; }
            public int Top { get; set; }
            public int Width { get; set; } = 500;
            public int Height { get; set; } = 300;
            public FakeFrame LinkedWindowFrame { get; set; }
            public int FocusCount { get; private set; }

            public void SetFocus()
            {
                FocusCount++;
                host.ActiveWindow = this;
            }

            public void Close()
            {
                Visible = false;
            }
        }
    }
}

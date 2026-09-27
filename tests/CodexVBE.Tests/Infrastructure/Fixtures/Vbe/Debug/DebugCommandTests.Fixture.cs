namespace CodexVBE.Tests.Unit
{
    using System;
    using System.Collections.Generic;
    using System.Linq;
    using CodexVBE;
    using Microsoft.VisualStudio.TestTools.UnitTesting;

    public sealed partial class DebugCommandTests
    {
        private static FakeVbe Host()
        {
            var debug = new FakeControl
            {
                Caption = "&Debug",
                Id = 1,
                Enabled = true
            };
            debug.Controls.Add(new FakeControl { Caption = "&Locals Window", Id = 2555, Enabled = true });
            var watch = new FakeControl
            {
                Caption = "&Watch Window",
                Id = 2556,
                Enabled = true
            };
            var disabled = new FakeControl
            {
                Caption = "Immediate Window",
                Id = 2554,
                Enabled = false
            };
            var other = new FakeControl
            {
                Caption = "Run",
                Id = 186,
                Enabled = true
            };
            return new FakeVbe
            {
                CommandBars = new List<FakeBar>
                {
                    new FakeBar
                    {
                        Name = "Debug",
                        Controls = new List<FakeControl>
                        {
                            debug,
                            watch,
                            disabled,
                            other
                        }
                    }
                },
                VBProjects = new List<FakeProject>
                {
                    new FakeProject
                    {
                        Name = "Projet",
                        FileName = @"C:\Temp\Projet.xlsm",
                        Mode = 2
                    }
                }
            };
        }

        public sealed class FakeVbe
        {
            public List<FakeBar> CommandBars { get; set; }
            public List<FakeProject> VBProjects { get; set; }
            public FakeProject ActiveVBProject { get; set; }
            public FakeCodePane ActiveCodePane { get; set; }
        }

        public sealed class FakeProject
        {
            public string Name { get; set; }
            public string FileName { get; set; }
            public int Mode { get; set; }
        }

        public sealed class FakeBar
        {
            public string Name { get; set; }
            public List<FakeControl> Controls { get; set; }
        }

        public sealed class FakeControl
        {
            public string Caption { get; set; }
            public int Id { get; set; }
            public bool Enabled { get; set; }
            public int Executions { get; private set; }
            public List<FakeControl> Controls { get; } = new List<FakeControl>();

            public void Execute()
            {
                Executions++;
            }
        }

        public sealed class FakeCodePane
        {
            public FakeCodeModule CodeModule { get; } = new FakeCodeModule();

            public void GetSelection(ref int startLine, ref int startColumn, ref int endLine, ref int endColumn)
            {
                startLine = 4;
                startColumn = 2;
                endLine = 4;
                endColumn = 9;
            }
        }

        public sealed class FakeCodeModule
        {
            public FakeComponent Parent { get; } = new FakeComponent();
        }

        public sealed class FakeComponent
        {
            public string Name { get; } = "Module1";
        }
    }
}

using System;
using System.Collections.Generic;
using System.Linq;
using CodexVBE;
using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace CodexVBE.Tests.Unit
{
    [TestClass]
    [TestCategory("Unit")]
    public sealed class DebugCommandTests
    {
        [TestMethod]
        public void CommandCatalogSearchesNestedMenusAndPagesResults()
        {
            var host = Host();
            var commands = new VbeDebug(host);
            dynamic all = commands.ListCommands(null, 0, 0);
            Assert.AreEqual(5, ((object[])all).Length);
            dynamic match = commands.ListCommands("WATCh", 0, 10);
            Assert.AreEqual(1, ((object[])match).Length);
            Assert.AreEqual("Debug > &Watch Window", (string)match[0].Path);
            Assert.AreEqual(2556, (int)match[0].Id);
            Assert.AreEqual(2, ((object[])commands.ListCommands(null, 1, 2)).Length);
            Assert.AreEqual(0, ((object[])commands.ListCommands("missing", 0, 10)).Length);
            Assert.ThrowsException<ArgumentOutOfRangeException>(() => commands.ListCommands(null, -1, 1));
            Assert.ThrowsException<ArgumentOutOfRangeException>(() => commands.ListCommands(null, 0, -1));
        }

        [TestMethod]
        public void OpeningDebugPanesRequiresAnEnabledMatchingNativeCommand()
        {
            var host = Host();
            var commands = new VbeDebug(host);
            Assert.ThrowsException<ArgumentException>(() => commands.OpenDebugPane("output", null));
            Assert.ThrowsException<InvalidOperationException>(() => commands.OpenDebugPane("immediate", null));
            dynamic locals = commands.OpenDebugPane("LOCALS", null);
            Assert.AreEqual(2555, (int)locals.ControlId);
            Assert.IsTrue((bool)locals.VerificationPending);
            Assert.AreEqual(1, host.CommandBars[0].Controls[0].Controls[0].Executions);
            dynamic watches = commands.OpenDebugPane("watches", null);
            Assert.AreEqual(2556, (int)watches.ControlId);
            Assert.AreEqual(1, host.CommandBars[0].Controls[1].Executions);
        }

        [TestMethod]
        public void DebugStateReportsTheActiveSelectionWithoutMovingIt()
        {
            var host = Host();
            var commands = new VbeDebug(host);
            dynamic idle = commands.State("Projet");
            Assert.AreEqual(2, (int)idle.Mode);
            Assert.IsNull((object)idle.Selection);
            host.ActiveVBProject = host.VBProjects.Single();
            host.ActiveCodePane = new FakeCodePane();
            dynamic selected = commands.State("Projet");
            Assert.AreEqual("Projet", (string)selected.SelectedProject);
            Assert.AreEqual("Module1", (string)selected.ActiveModule);
            Assert.AreEqual(4, (int)selected.Selection.StartLine);
            Assert.AreEqual(9, (int)selected.Selection.EndColumn);
        }

        private static FakeVbe Host()
        {
            var debug = new FakeControl { Caption = "&Debug", Id = 1, Enabled = true };
            debug.Controls.Add(new FakeControl { Caption = "&Locals Window", Id = 2555, Enabled = true });
            var watch = new FakeControl { Caption = "&Watch Window", Id = 2556, Enabled = true };
            var disabled = new FakeControl { Caption = "Immediate Window", Id = 2554, Enabled = false };
            var other = new FakeControl { Caption = "Run", Id = 186, Enabled = true };
            return new FakeVbe {
                CommandBars = new List<FakeBar> { new FakeBar { Name = "Debug", Controls = new List<FakeControl> { debug, watch, disabled, other } } },
                VBProjects = new List<FakeProject> { new FakeProject { Name = "Projet", FileName = @"C:\Temp\Projet.xlsm", Mode = 2 } }
            };
        }

        public sealed class FakeVbe
        {
            public List<FakeBar> CommandBars { get; set; }
            public List<FakeProject> VBProjects { get; set; }
            public FakeProject ActiveVBProject { get; set; }
            public FakeCodePane ActiveCodePane { get; set; }
        }
        public sealed class FakeProject { public string Name { get; set; } public string FileName { get; set; } public int Mode { get; set; } }
        public sealed class FakeBar { public string Name { get; set; } public List<FakeControl> Controls { get; set; } }
        public sealed class FakeControl
        {
            public string Caption { get; set; }
            public int Id { get; set; }
            public bool Enabled { get; set; }
            public int Executions { get; private set; }
            public List<FakeControl> Controls { get; } = new List<FakeControl>();
            public void Execute() { Executions++; }
        }
        public sealed class FakeCodePane
        {
            public FakeCodeModule CodeModule { get; } = new FakeCodeModule();
            public void GetSelection(ref int startLine, ref int startColumn, ref int endLine, ref int endColumn)
            { startLine = 4; startColumn = 2; endLine = 4; endColumn = 9; }
        }
        public sealed class FakeCodeModule { public FakeComponent Parent { get; } = new FakeComponent(); }
        public sealed class FakeComponent { public string Name { get; } = "Module1"; }
    }
}

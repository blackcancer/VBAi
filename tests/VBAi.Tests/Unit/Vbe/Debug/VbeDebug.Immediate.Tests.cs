using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using VBAi;

namespace VBAi.Tests.Unit
{
    [TestClass]
    [TestCategory("Unit")]
    public sealed class VbeDebugImmediateTests
    {
        [WinFormsTestMethod]
        public void ReadImmediateUsesExactCommandsAndRestoresFocus()
        {
            var fixture = Create();
            dynamic result = fixture.Service.ReadImmediateAsync(new Request { Project = "P", ExpectedMode = 1 }).GetAwaiter().GetResult();
            Assert.AreEqual("auditValue = 41", (string)result.Text);
            Assert.AreEqual("NativeCopy", (string)result.Method);
            Assert.IsTrue((bool)result.ClipboardRestored);
            Assert.IsTrue((bool)result.FocusRestored);
            Assert.IsFalse((bool)result.SelectionRestored);
            Assert.AreEqual(1, fixture.Select.Executions);
            Assert.AreEqual(1, fixture.Copy.Executions);
            Assert.AreSame(fixture.Previous, fixture.Vbe.ActiveWindow);
            Assert.AreEqual("original", fixture.Data.GetData(DataFormats.UnicodeText, false));
        }

        [STATestMethod]
        public void ReadImmediateKeepsOwningStaAcrossRealNativeYields()
        {
            var fixture = Create();
            var previousContext = SynchronizationContext.Current;
            var ambient = new SynchronizationContext();
            int ownerThread = Thread.CurrentThread.ManagedThreadId;
            int yields = 0;
            Action verifyThread = () => {
                Assert.AreEqual(ownerThread, Thread.CurrentThread.ManagedThreadId);
                Assert.AreEqual(ApartmentState.STA, Thread.CurrentThread.GetApartmentState());
            };
            try
            {
                SynchronizationContext.SetSynchronizationContext(ambient);
                fixture.Select.OnExecute = verifyThread;
                fixture.Copy.OnExecute = () => {
                    verifyThread(); fixture.Data = Text("auditValue = 41"); fixture.Sequence++;
                };
                fixture.Service.ImmediateClipboard = new VbeDebugClipboard {
                    Sequence = () => { verifyThread(); return fixture.Sequence; },
                    YieldNative = async () => { await Task.Delay(30); verifyThread(); yields++; },
                    ReadData = () => { verifyThread(); return fixture.Data; },
                    IsHostOwner = () => { verifyThread(); return true; },
                    WriteData = data => { verifyThread(); fixture.Data = data; fixture.Sequence++; }
                };
                Task<object> pending = fixture.Service.ReadImmediateAsync(new Request { Project = "P", ExpectedMode = 1 });
                Assert.AreSame(ambient, SynchronizationContext.Current);
                var elapsed = Stopwatch.StartNew();
                while (!pending.IsCompleted && elapsed.ElapsedMilliseconds < 3000)
                {
                    Application.DoEvents();
                    Thread.Sleep(1);
                }
                Assert.IsTrue(pending.IsCompleted, "Native Immediate capture did not complete while pumping STA messages.");
                dynamic result = pending.GetAwaiter().GetResult();
                Assert.AreEqual("auditValue = 41", (string)result.Text);
                Assert.IsTrue((bool)result.ClipboardRestored);
                Assert.IsTrue((bool)result.FocusRestored);
                Assert.AreSame(fixture.Previous, fixture.Vbe.ActiveWindow);
                Assert.AreEqual("original", fixture.Data.GetData(DataFormats.UnicodeText, false));
                Assert.IsTrue(yields >= 1, "The capture must cross a real asynchronous native yield.");
            }
            finally { SynchronizationContext.SetSynchronizationContext(previousContext); }
        }

        [WinFormsTestMethod]
        public void ProjectModeAndUniqueVisibleImmediateAreRequiredBeforeCommands()
        {
            var fixture = Create();
            Assert.ThrowsException<ArgumentException>(() => fixture.Service.ReadImmediateAsync(new Request { Project = "", ExpectedMode = 1 }).GetAwaiter().GetResult());
            Assert.ThrowsException<InvalidOperationException>(() => fixture.Service.ReadImmediateAsync(new Request { Project = "P", ExpectedMode = 2 }).GetAwaiter().GetResult());
            fixture.Vbe.ActiveVBProject = new FakeProject { Name = "Other", Mode = 1 };
            Assert.ThrowsException<InvalidOperationException>(() => fixture.Service.ReadImmediateAsync(new Request { Project = "P", ExpectedMode = 1 }).GetAwaiter().GetResult());
            fixture.Vbe.ActiveVBProject = fixture.Project;
            fixture.Immediate.Visible = false;
            Assert.ThrowsException<InvalidOperationException>(() => fixture.Service.ReadImmediateAsync(new Request { Project = "P", ExpectedMode = 1 }).GetAwaiter().GetResult());
            fixture.Immediate.Visible = true;
            fixture.Vbe.Windows.Add(new FakeWindow(fixture.Vbe) { Type = 5, Visible = true });
            Assert.ThrowsException<InvalidOperationException>(() => fixture.Service.ReadImmediateAsync(new Request { Project = "P", ExpectedMode = 1 }).GetAwaiter().GetResult());
            Assert.AreEqual(0, fixture.Select.Executions);
            Assert.AreEqual(0, fixture.Copy.Executions);
        }

        [WinFormsTestMethod]
        public void WrongCaptionOrDisabledCommandNeverCopies()
        {
            foreach (Action<Fixture> invalidate in new Action<Fixture>[] {
                f => f.Select.Caption = "Select Row",
                f => f.Select.Enabled = false,
                f => f.Copy.Caption = "Cut",
                f => f.Copy.Enabled = false
            })
            {
                var fixture = Create();
                invalidate(fixture);
                Assert.ThrowsException<InvalidOperationException>(() => fixture.Service.ReadImmediateAsync(
                    new Request { Project = "P", ExpectedMode = 1 }).GetAwaiter().GetResult());
                Assert.AreEqual(0, fixture.Copy.Executions);
                Assert.AreEqual("original", fixture.Data.GetData(DataFormats.UnicodeText, false));
                Assert.AreSame(fixture.Previous, fixture.Vbe.ActiveWindow);
            }
        }

        [WinFormsTestMethod]
        public void CopyRefusesWhenSelectionChangesActiveWindow()
        {
            var fixture = Create();
            fixture.Select.OnExecute = () => fixture.Previous.SetFocus();
            Assert.ThrowsException<InvalidOperationException>(() => fixture.Service.ReadImmediateAsync(
                new Request { Project = "P", ExpectedMode = 1 }).GetAwaiter().GetResult());
            Assert.AreEqual(1, fixture.Select.Executions);
            Assert.AreEqual(0, fixture.Copy.Executions);
            Assert.AreSame(fixture.Previous, fixture.Vbe.ActiveWindow);
        }

        [TestMethod]
        public void ImmediateCopyCaptionsAcceptOnlyKnownFrenchAndEnglishCommands()
        {
            Assert.IsTrue(VbeDebug.IsImmediateCopyCaption("&Select All", true));
            Assert.IsTrue(VbeDebug.IsImmediateCopyCaption("&Sélectionner tout", true));
            Assert.IsTrue(VbeDebug.IsImmediateCopyCaption("&Copy", false));
            Assert.IsTrue(VbeDebug.IsImmediateCopyCaption("&Copier", false));
            Assert.IsFalse(VbeDebug.IsImmediateCopyCaption("Cut", false));
            Assert.IsFalse(VbeDebug.IsImmediateCopyCaption("Select Row", true));
        }

        private static Fixture Create()
        {
            var fixture = new Fixture();
            fixture.Project = new FakeProject { Name = "P", FileName = @"C:\Temp\P.xlsm", Mode = 1 };
            fixture.Vbe = new FakeVbe { ActiveVBProject = fixture.Project };
            fixture.Vbe.VBProjects.Add(fixture.Project);
            fixture.Previous = new FakeWindow(fixture.Vbe) { Type = 0, Visible = true };
            fixture.Immediate = new FakeWindow(fixture.Vbe) { Type = 5, Visible = true };
            fixture.Vbe.Windows.Add(fixture.Previous);
            fixture.Vbe.Windows.Add(fixture.Immediate);
            fixture.Vbe.ActiveWindow = fixture.Previous;
            fixture.Select = new FakeControl { Id = 756, Caption = "&Select All", Enabled = true };
            fixture.Copy = new FakeControl { Id = 19, Caption = "&Copy", Enabled = true };
            fixture.Vbe.CommandBars.Add(new FakeBar { Name = "Edit", Controls = { fixture.Select, fixture.Copy } });
            fixture.Data = Text("original");
            fixture.Service = new VbeDebug(fixture.Vbe);
            fixture.Select.OnExecute = () => { };
            fixture.Copy.OnExecute = () => { fixture.Data = Text("auditValue = 41"); fixture.Sequence++; };
            fixture.Service.ImmediateClipboard = new VbeDebugClipboard {
                Sequence = () => fixture.Sequence,
                YieldNative = () => Task.CompletedTask,
                ReadData = () => fixture.Data,
                IsHostOwner = () => true,
                WriteData = data => { fixture.Data = data; fixture.Sequence++; }
            };
            return fixture;
        }

        private static DataObject Text(string text)
        {
            var data = new DataObject();
            data.SetData(DataFormats.UnicodeText, false, text);
            return data;
        }

        private sealed class Fixture
        {
            public VbeDebug Service;
            public FakeVbe Vbe;
            public FakeProject Project;
            public FakeWindow Previous, Immediate;
            public FakeControl Select, Copy;
            public IDataObject Data;
            public uint Sequence = 1;
        }

        public sealed class FakeVbe
        {
            public readonly List<FakeProject> VBProjects = new List<FakeProject>();
            public readonly List<FakeWindow> Windows = new List<FakeWindow>();
            public readonly List<FakeBar> CommandBars = new List<FakeBar>();
            public FakeProject ActiveVBProject { get; set; }
            public FakeWindow ActiveWindow { get; set; }
        }

        public sealed class FakeProject
        {
            public string Name { get; set; }
            public string FileName { get; set; }
            public int Mode { get; set; }
        }

        public sealed class FakeWindow
        {
            private readonly FakeVbe owner;
            public FakeWindow(FakeVbe owner) { this.owner = owner; }
            public int Type { get; set; }
            public bool Visible { get; set; }
            public void SetFocus() { owner.ActiveWindow = this; }
        }

        public sealed class FakeBar
        {
            public string Name { get; set; }
            public List<FakeControl> Controls { get; } = new List<FakeControl>();
        }

        public sealed class FakeControl
        {
            public int Id { get; set; }
            public string Caption { get; set; }
            public bool Enabled { get; set; }
            public int Executions { get; private set; }
            public Action OnExecute { get; set; }
            public List<FakeControl> Controls { get; } = new List<FakeControl>();
            public void Execute() { Executions++; OnExecute?.Invoke(); }
        }
    }
}

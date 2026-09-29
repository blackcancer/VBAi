using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading;
using System.Threading.Tasks;
using VBAi;
using VBAi.Tests.Infrastructure;
using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace VBAi.Tests.Unit
{
    [TestClass, TestCategory("Unit")]
    public sealed class ModernEditorReconciliationTests
    {
        private sealed class CountingModule : IEditorModule
        {
            private readonly string key = Guid.NewGuid().ToString("N");
            internal string DisplayName = "Module";
            internal int NameReads;
            public string Name { get { NameReads++; return DisplayName; } }
            public string Key => key;
            public bool CanWrite => true;
            internal int Reads;
            internal string Code = "Sub Main()\nDebug.Print 1\nEnd Sub";
            public string Read() { Reads++; return Code; }
            public string Write(string expected, string text) { Assert.AreEqual(expected, Code); return Code = text; }
            public void ShowNative(int line, int column) { }
        }

        [STATestMethod]
        public void BackgroundTicksBoundReadsAndEventuallyObserveEveryInactiveDocument()
        {
            using (var fixture = new ModernEditorDebugFixture())
            {
                var modules = Enumerable.Range(0, 29).Select(i => new CountingModule()).ToArray();
                foreach (var module in modules) ModernEditorDebugFixture.Wait(fixture.Window.OpenModule(module));
                fixture.Ready(true);
                foreach (var module in modules) { module.Reads = 0; module.NameReads = 0; }
                for (int tick = 0; tick < 15; tick++)
                {
                    int before = modules.Sum(m => m.Reads);
                    int namesBefore = modules.Sum(m => m.NameReads);
                    ModernEditorDebugFixture.Wait((Task)UiInvoke.Call(typeof(ModernEditorWindow), "ProcessBackgroundDocuments", fixture.Window));
                    Assert.IsTrue(modules.Sum(m => m.Reads) - before <= 3);
                    Assert.IsTrue(modules.Sum(m => m.NameReads) - namesBefore <= 3, "Status layout must not read every native caption.");
                    Assert.AreEqual(tick + 1, modules.Last().Reads, "The selected document is checked on every tick.");
                }
                Assert.IsTrue(modules.All(m => m.Reads > 0));
                var semaphore = fixture.Get<SemaphoreSlim>("debugCommands");
                semaphore.Wait();
                try
                {
                    int before = modules.Sum(m => m.Reads);
                    ModernEditorDebugFixture.Wait((Task)UiInvoke.Call(typeof(ModernEditorWindow), "ProcessBackgroundDocuments", fixture.Window));
                    Assert.AreEqual(before, modules.Sum(m => m.Reads), "Native commands take priority over background reconciliation.");
                }
                finally { semaphore.Release(); }
            }
        }

        [STATestMethod]
        public void StatusUsesCachedCaptionsAndReconciliationObservesRename()
        {
            using (var fixture = new ModernEditorDebugFixture())
            {
                var module = new CountingModule { DisplayName = "Original" };
                var document = ModernEditorDebugFixture.Wait(fixture.Window.OpenModule(module));
                fixture.Ready(true);
                module.NameReads = 0;
                module.DisplayName = "Renamed";
                for (int i = 0; i < 20; i++) UiInvoke.Call(typeof(ModernEditorWindow), "UpdateStatus", fixture.Window);
                Assert.AreEqual(0, module.NameReads);
                var tabs = UiInvoke.Field<System.Windows.Forms.TabControl>(fixture.Window, "tabs");
                var page = tabs.TabPages.Cast<System.Windows.Forms.TabPage>().Single(tab => (string)tab.Tag == document.Id);
                Assert.AreEqual("Original", page.Text);
                ModernEditorDebugFixture.Wait((Task)UiInvoke.Call(typeof(ModernEditorWindow), "ProcessBackgroundDocuments", fixture.Window));
                UiInvoke.Call(typeof(ModernEditorWindow), "UpdateStatus", fixture.Window);
                Assert.AreEqual(1, module.NameReads);
                Assert.AreEqual("Renamed", page.Text);
            }
        }

        [STATestMethod]
        public void CoalescedStatusDoesNotOverwriteANewerCommandResult()
        {
            using (var fixture = new ModernEditorDebugFixture())
            {
                for (int i = 0; i < 100; i++) UiInvoke.Call(typeof(ModernEditorWindow), "SetStatus", fixture.Window);
                UiInvoke.Call(typeof(ModernEditorWindow), "SetResultStatus", fixture.Window, "command finished");
                System.Windows.Forms.Application.DoEvents();
                Assert.AreEqual("command finished", UiInvoke.Field<System.Windows.Forms.Label>(fixture.Window, "status").Text);
                Assert.IsFalse(fixture.Get<bool>("statusUpdatePending"));
            }
        }
    }
}

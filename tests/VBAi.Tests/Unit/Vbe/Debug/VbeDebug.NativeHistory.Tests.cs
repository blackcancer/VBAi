namespace VBAi.Tests.Unit
{
    using Microsoft.VisualStudio.TestTools.UnitTesting;
    using System;
    using System.Collections;
    using System.Linq;
    using VBAi;

    [TestClass, TestCategory("Unit"), DoNotParallelize]
    public sealed class VbeNativeHistoryTests
    {
        private Func<IntPtr, bool> saved;
        [TestInitialize] public void IsolateNativeWindowCheck() { saved = VbeDebug.HistoryWindowEnabled; VbeDebug.HistoryWindowEnabled = handle => true; }
        [TestCleanup] public void RestoreNativeWindowCheck() { VbeDebug.HistoryWindowEnabled = saved; }
        private static Request Inspected(EditorDebugFixture f, string action = "undo")
        {
            var request = f.Location(action);
            dynamic state = f.Service.NativeCodeHistoryState(request);
            request.ExpectedProjectVersion = state.HistoryVersion;
            return request;
        }
        [TestMethod]
        public void HistoryStateRejectsSharedStacksRuntimeModeAndDesignerProjects()
        {
            var f = new EditorDebugFixture();
            f.Vbe.VBProjects.Add(new EditorDebugFixture.DebugProject { Name = "other" });
            Assert.ThrowsException<InvalidOperationException>(() => f.Service.NativeCodeHistoryState(f.Location()));
            f.Vbe.VBProjects.RemoveAt(1); f.Project.Mode = 1;
            Assert.ThrowsException<InvalidOperationException>(() => f.Service.NativeCodeHistoryState(f.Location()));
            f.Project.Mode = 2; f.Module.Parent.Type = 3;
            Assert.ThrowsException<InvalidOperationException>(() => f.Service.NativeCodeHistoryState(f.Location()));
            f.Module.Parent.Type = 1; f.AddModule("Empty", "");
            f.Command(128, "undo"); f.Command(129, "redo"); f.Command(10, "other");
            dynamic state = f.Service.NativeCodeHistoryState(f.Location());
            Assert.AreEqual(2, ((object[])state.Modules).Length); Assert.AreEqual(2, ((object[])state.Commands).Length);
            Assert.AreEqual("Empty", (string)((dynamic)((object[])state.Modules)[0]).Module);
            Assert.IsFalse(string.IsNullOrWhiteSpace((string)state.HistoryVersion));
        }
        [TestMethod]
        public void HistoryRequiresActionDesignModeVersionCaptionAndEnabledExactCommand()
        {
            var f = new EditorDebugFixture(); var request = f.Location("invalid");
            Assert.ThrowsException<ArgumentException>(() => f.Service.NativeCodeHistory(request));
            request.Action = "undo"; request.ExpectedMode = 1;
            Assert.ThrowsException<ArgumentException>(() => f.Service.NativeCodeHistory(request));
            request.ExpectedMode = 2; Assert.ThrowsException<ArgumentException>(() => f.Service.NativeCodeHistory(request));
            request = Inspected(f); request.ControlCaption = " ";
            Assert.ThrowsException<ArgumentException>(() => f.Service.NativeCodeHistory(request));
            request = Inspected(f); request.ExpectedProjectVersion = "stale";
            Assert.ThrowsException<InvalidOperationException>(() => f.Service.NativeCodeHistory(request));
            request = Inspected(f); VbeDebug.HistoryWindowEnabled = handle => false;
            Assert.ThrowsException<InvalidOperationException>(() => f.Service.NativeCodeHistory(request));
            VbeDebug.HistoryWindowEnabled = handle => true;
            Assert.ThrowsException<InvalidOperationException>(() => f.Service.NativeCodeHistory(request));
            var command = f.Command(128, "wrong"); request = Inspected(f);
            Assert.ThrowsException<InvalidOperationException>(() => f.Service.NativeCodeHistory(request));
            command.Caption = "undo"; command.Enabled = false; request = Inspected(f);
            Assert.ThrowsException<InvalidOperationException>(() => f.Service.NativeCodeHistory(request));
            Assert.AreEqual(0, command.Executions);
        }
        [TestMethod]
        public void HistoryReportsActualChangesTopologyAndUnchangedNativeResults()
        {
            foreach (int scenario in new[] { 0, 1, 2, 3, 4 })
            {
                var f = new EditorDebugFixture(); f.AddModule("Unchanged", "");
                string action = scenario == 4 ? "redo" : "undo";
                f.Command(action == "undo" ? 128 : 129, action, () =>
                {
                    if (scenario == 0 || scenario == 2 || scenario == 4) f.Module.Code = "restored source";
                    if (scenario == 2 || scenario == 3) f.Project.VBComponents.Remove(f.Project.VBComponents.Single(x => x.Name == "Unchanged"));
                    if (scenario == 3) f.AddModule("New");
                });
                dynamic result = f.Service.NativeCodeHistory(Inspected(f, action));
                int count = ((IEnumerable)result.Changes).Cast<object>().Count();
                Assert.AreEqual(scenario == 0 || scenario == 2 || scenario == 4 ? 1 : 0, count);
                Assert.AreEqual(scenario == 2 || scenario == 3, (bool)result.TopologyChanged);
                Assert.AreEqual(scenario == 0 || scenario == 4, (bool)result.Verified);
                Assert.AreEqual(scenario == 1, (bool)result.VerificationPending); Assert.IsTrue((bool)result.Executed);
                Assert.IsNull(result.NativeError);
            }
        }
        [TestMethod]
        public void HistoryPreservesExecutionErrorWhenReadbackAlsoFails()
        {
            foreach (int scenario in new[] { 0, 1, 2, 3 })
            {
                var f = new EditorDebugFixture();
                f.Command(128, "undo", () =>
                {
                    if (scenario == 1 || scenario == 2) f.Project.Mode = 0;
                    if (scenario == 3) f.Module.Code = "changed before failure";
                    if (scenario != 1) throw new InvalidOperationException("execute failed");
                });
                dynamic result = f.Service.NativeCodeHistory(Inspected(f));
                Assert.IsFalse((bool)result.Executed); Assert.IsFalse((bool)result.Verified);
                StringAssert.Contains((string)result.NativeError, scenario == 1 ? "design mode" : "execute failed");
                Assert.AreEqual(scenario != 3, (bool)result.VerificationPending);
            }
        }
    }
}

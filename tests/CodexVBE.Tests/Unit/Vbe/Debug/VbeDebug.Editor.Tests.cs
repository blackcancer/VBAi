namespace CodexVBE.Tests.Unit
{
    using System;
    using CodexVBE;
    using Microsoft.VisualStudio.TestTools.UnitTesting;

    [TestClass, TestCategory("Unit")]
    public sealed class VbeEditorSplitTests
    {
        [TestMethod]
        public void SplitRejectsInvalidActionCaptionModeAndTopology()
        {
            var f = new EditorDebugFixture(); var request = f.Location("invalid");
            Assert.ThrowsException<ArgumentException>(() => f.Service.SetCodeSplit(request));
            request.Action = "split"; request.ControlCaption = " ";
            Assert.ThrowsException<ArgumentException>(() => f.Service.SetCodeSplit(request));
            request.ControlCaption = "split";
            foreach (int mode in new[] { 0, 3 }) { f.Project.Mode = mode; request.ExpectedMode = mode; Assert.ThrowsException<InvalidOperationException>(() => f.Service.SetCodeSplit(request)); }
            f.Project.Mode = 1; request.ExpectedMode = 2;
            Assert.ThrowsException<InvalidOperationException>(() => f.Service.SetCodeSplit(request));
            request.ExpectedMode = 1; f.Pane.OnShow = () => f.Vbe.ActiveCodePane = null;
            Assert.ThrowsException<InvalidOperationException>(() => f.Service.SetCodeSplit(request));
            var other = f.AddModule("other"); f.Pane.OnShow = () => f.Vbe.ActiveCodePane = other.CodePane;
            Assert.ThrowsException<InvalidOperationException>(() => f.Service.SetCodeSplit(request));
            f.Pane.OnShow = () => f.Vbe.ActiveCodePane = f.Pane; f.Vbe.Panes.Clear();
            Assert.ThrowsException<InvalidOperationException>(() => f.Service.SetCodeSplit(request));
            f.Vbe.Panes.Add(f.Pane); f.Vbe.Panes.Add(new EditorDebugFixture.DebugPane { CodeModule = f.Module }); f.Vbe.Panes.Add(new EditorDebugFixture.DebugPane { CodeModule = f.Module });
            Assert.ThrowsException<InvalidOperationException>(() => f.Service.SetCodeSplit(request));
        }
        [TestMethod]
        public void SplitUsesExactEnabledCommandAndVerifiesNativePaneCounts()
        {
            var f = new EditorDebugFixture(); var request = f.Location("split");
            var command = f.Command(301, "split");
            Assert.ThrowsException<InvalidOperationException>(() => f.Service.SetCodeSplit(request));
            command.Id = 302; command.Enabled = false;
            Assert.ThrowsException<InvalidOperationException>(() => f.Service.SetCodeSplit(request));
            command.Enabled = true; command.Caption = "wrong";
            Assert.ThrowsException<InvalidOperationException>(() => f.Service.SetCodeSplit(request));
            command.Caption = "split"; command.OnExecute = () => f.Vbe.Panes.Add(new EditorDebugFixture.DebugPane { CodeModule = f.Module });
            f.Vbe.Panes.Add(new EditorDebugFixture.DebugPane { ModuleError = unchecked((int)0x80020010) });
            f.AddModule("unrelated");
            dynamic result = f.Service.SetCodeSplit(request);
            Assert.IsTrue((bool)result.Verified); Assert.AreEqual(2, (int)result.PaneCount); Assert.IsNull(result.NextRead);
            dynamic unchanged = f.Service.SetCodeSplit(request);
            Assert.IsFalse((bool)unchanged.Applied); Assert.AreEqual(1, command.Executions);
            request.Action = "unsplit"; request.ControlCaption = "unsplit"; command.Caption = "unsplit";
            command.OnExecute = () => f.Vbe.Panes.RemoveAt(f.Vbe.Panes.Count - 1);
            Assert.IsTrue((bool)((dynamic)f.Service.SetCodeSplit(request)).Verified);
        }
        [TestMethod]
        public void SplitReportsExecutionAndReadbackErrorsWithoutClaimingVerification()
        {
            foreach (int failure in new[] { 0, 1, 2, 3 })
            {
                var f = new EditorDebugFixture(); var request = f.Location("split");
                f.Command(302, "split", () => {
                    if (failure == 1 || failure == 2) f.Vbe.OnReadPanes = () => { throw new InvalidOperationException("readback failed"); };
                    if (failure == 0 || failure == 2) throw new InvalidOperationException("execute failed");
                });
                dynamic result = f.Service.SetCodeSplit(request);
                Assert.IsFalse((bool)result.Verified); Assert.IsTrue((bool)result.VerificationPending);
                Assert.AreEqual("code_panes", (string)result.NextRead);
                if (failure == 3) { Assert.IsTrue((bool)result.Applied); Assert.IsNull(result.NativeError); }
                else { Assert.IsNull(result.Applied); Assert.AreEqual(failure == 1 ? "readback failed" : "execute failed", (string)result.NativeError); }
            }
        }
    }
}

namespace VBAi.Tests.Unit
{
    using System;
    using VBAi;
    using Microsoft.VisualStudio.TestTools.UnitTesting;

    [TestClass, TestCategory("Unit"), DoNotParallelize]
    public sealed class VbeCodeViewTests
    {
        [TestMethod]
        public void NativeViewClicksOnlyTheVerifiedLocalButtonCenters()
        {
            using (var scene = new NativeDebugScene())
            {
                var toolbar = scene.CodeWindow();
                var code = scene.Find(toolbar.Parent);
                scene.Add("other", parent: code);
                scene.Add("ScrollBar", parent: code, width: 10, height: 60);
                scene.Add("ObtbarWndClass", parent: code, top: 25);
                scene.Add("ObtbarWndClass", parent: code, height: 19);
                scene.Add("ObtbarWndClass", parent: code, left: 90);
                scene.Add("other", parent: code).Visible = false;
                scene.Add("other", parent: code).BoundsAvailable = false;
                dynamic procedure = VbeDebugWindows.ChangeCodeView("Module1 (Code)", true);
                dynamic module = VbeDebugWindows.ChangeCodeView("Module1 (Code)", false);
                Assert.AreEqual(10, (int)procedure.X); Assert.AreEqual(30, (int)module.X);
                Assert.AreEqual(20, (int)module.Height); Assert.AreEqual(40, (int)module.Width);
                Assert.AreEqual(4, scene.Messages.Count);
                Assert.AreEqual(toolbar.Handle, scene.Messages[0].Item1);
                Assert.AreEqual(0x201, scene.Messages[0].Item2);
                Assert.AreEqual(new IntPtr((10 << 16) | 10), scene.Messages[0].Item4);
                Assert.AreEqual(0x202, scene.Messages[1].Item2);
                Assert.AreEqual(IntPtr.Zero, scene.Messages[1].Item3);
            }
        }
        [TestMethod]
        public void NativeViewRejectsMissingAmbiguousAndModalCodeWindows()
        {
            using (var scene = new NativeDebugScene())
            {
                Assert.ThrowsException<InvalidOperationException>(() => VbeDebugWindows.ChangeCodeView("missing", true));
                var toolbar = scene.CodeWindow(); var code = scene.Find(toolbar.Parent); var root = scene.Find(code.Parent);
                Assert.ThrowsException<InvalidOperationException>(() => VbeDebugWindows.ChangeCodeView("wrong", true));
                var duplicate = scene.Add("VbaWindow", code.Caption, root);
                Assert.ThrowsException<InvalidOperationException>(() => VbeDebugWindows.ChangeCodeView(code.Caption, true));
                scene.Windows.Remove(duplicate); root.Enabled = false;
                Assert.ThrowsException<InvalidOperationException>(() => VbeDebugWindows.ChangeCodeView(code.Caption, true));
                Assert.AreEqual(0, scene.Messages.Count);
            }
        }
        [TestMethod]
        public void NativeViewRejectsScrollbarToolbarAndGeometryAmbiguityBeforeSending()
        {
            using (var scene = new NativeDebugScene())
            {
                var toolbar = scene.CodeWindow(); var code = scene.Find(toolbar.Parent);
                var scrollbar = scene.Windows[2]; scrollbar.Visible = false;
                Assert.ThrowsException<InvalidOperationException>(() => VbeDebugWindows.ChangeCodeView(code.Caption, true));
                scrollbar.Visible = true;
                var second = scene.Add("ScrollBar", parent: code, width: 200);
                Assert.ThrowsException<InvalidOperationException>(() => VbeDebugWindows.ChangeCodeView(code.Caption, true));
                scene.Windows.Remove(second); toolbar.Visible = false;
                Assert.ThrowsException<InvalidOperationException>(() => VbeDebugWindows.ChangeCodeView(code.Caption, true));
                toolbar.Visible = true;
                second = scene.Add("ObtbarWndClass", parent: code);
                Assert.ThrowsException<InvalidOperationException>(() => VbeDebugWindows.ChangeCodeView(code.Caption, true));
                scene.Windows.Remove(second);
                foreach (var dimensions in new[] { Tuple.Create(40,9), Tuple.Create(140,65), Tuple.Create(20,20), Tuple.Create(80,20) })
                {
                    toolbar.Bounds.Right = dimensions.Item1; toolbar.Bounds.Bottom = dimensions.Item2;
                    scrollbar.Bounds.Top = toolbar.Bounds.Top; scrollbar.Bounds.Bottom = toolbar.Bounds.Bottom;
                    scrollbar.Bounds.Left = toolbar.Bounds.Right;
                    Assert.ThrowsException<InvalidOperationException>(() => VbeDebugWindows.ChangeCodeView(code.Caption, true));
                }
                toolbar.Bounds.Right = 40; toolbar.Bounds.Bottom = 20; scrollbar.Bounds.Left = 40; scrollbar.Bounds.Bottom = 20;
                toolbar.Enabled = false;
                Assert.ThrowsException<InvalidOperationException>(() => VbeDebugWindows.ChangeCodeView(code.Caption, true));
                Assert.AreEqual(0, scene.Messages.Count);
            }
        }
    }
}

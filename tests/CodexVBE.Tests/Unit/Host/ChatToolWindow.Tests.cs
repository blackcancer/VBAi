using System;
using System.Drawing;
using System.Windows.Forms;
using CodexVBE;
using CodexVBE.Tests.Infrastructure;
using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace CodexVBE.Tests.Unit
{
        /// <summary>Vérifie le dimensionnement natif et le cycle d’attachement de ChatToolWindow.</summary>
[TestClass, TestCategory("Unit")]
    public sealed class ChatToolWindowCoverageTests
    {
                /// <summary>Utilise des handles WinForms réels pour vérifier le redimensionnement, l’attachement et la destruction.</summary>
[STATestMethod]
        public void NativeSiteResizingAttachDetachTimerAndDisposeUseActualWinFormsHandles()
        {
            using (var scope = new HostUiScope())
            using (var owner = new Form { Left = -10000, Top = -10000, ClientSize = new Size(600, 800), ShowInTaskbar = false })
            using (var tool = new ChatToolWindow())
            using (var chat = new ChatWindow(new VbeSession(scope.Host)))
            {
                Assert.AreEqual(DockStyle.None, tool.Dock); Assert.AreEqual(new Size(520, 760), tool.Size);
                LlmBoundaryScope.Call(tool, "FitNativeSite"); var handle = tool.Handle; LlmBoundaryScope.Call(tool, "FitNativeSite");
                owner.Controls.Add(tool); owner.Show(); tool.Attach(chat); Application.DoEvents(); Assert.AreEqual(DockStyle.Fill, tool.Dock); Assert.AreSame(tool, chat.Parent); Assert.IsFalse(chat.TopLevel);
                var timer = LlmBoundaryScope.Get<Timer>(tool, "siteResizeTimer"); Assert.IsTrue(timer.Enabled); LlmBoundaryScope.Call(timer, "OnTick", EventArgs.Empty);
                LlmBoundaryScope.Call(tool, "FitNativeSite"); Assert.AreEqual(owner.ClientSize, tool.Size);
                tool.Dock = DockStyle.None; tool.Location = new Point(11, 13); tool.Size = new Size(100, 100);
                LlmBoundaryScope.Call(tool, "FitNativeSite"); Assert.AreEqual(new Size(owner.ClientSize.Width - 11, owner.ClientSize.Height - 13), tool.Size);
                tool.Detach(chat); Assert.IsTrue(chat.TopLevel); Assert.AreEqual(FormBorderStyle.Sizable, chat.FormBorderStyle); Assert.IsFalse(timer.Enabled);
                tool.Dispose(); LlmBoundaryScope.Call(tool, "FitNativeSite");
            }
            using (var tool = new ChatToolWindow()) LlmBoundaryScope.Call(tool, "Dispose", false);
        }
                /// <summary>Vérifie les garde-fous d’interop et les dimensions minimales en cas de géométrie invalide.</summary>
[STATestMethod]
        public void InteropFailureGuardsAndClampedDimensionsPreserveTheNativeSite()
        {
            using (var tool = new ChatToolWindow())
            {
                var handle = tool.Handle; tool.ParentReader = h => new IntPtr(123);
                tool.ClientReader = (IntPtr h, out ChatToolWindow.NativeRect r) => { r = new ChatToolWindow.NativeRect(); return false; }; LlmBoundaryScope.Call(tool, "FitNativeSite"); Assert.AreEqual(520, tool.Width);
                tool.ClientReader = (IntPtr h, out ChatToolWindow.NativeRect r) => { r = new ChatToolWindow.NativeRect { Right = 300, Bottom = 400 }; return true; };
                tool.WindowReader = (IntPtr h, out ChatToolWindow.NativeRect r) => { r = new ChatToolWindow.NativeRect(); return false; }; LlmBoundaryScope.Call(tool, "FitNativeSite"); Assert.AreEqual(760, tool.Height);
                tool.WindowReader = (IntPtr h, out ChatToolWindow.NativeRect r) => { r = new ChatToolWindow.NativeRect { Left = 10, Top = 20 }; return true; };
                tool.CoordinateConverter = (IntPtr h, ref ChatToolWindow.NativePoint p) => false; LlmBoundaryScope.Call(tool, "FitNativeSite"); Assert.AreEqual(520, tool.Width);
                tool.CoordinateConverter = (IntPtr h, ref ChatToolWindow.NativePoint p) => true; int moves = 0; tool.PositionWindow = (h, after, x, y, w, height, flags) => { moves++; Assert.AreEqual(0x14u, flags); return true; };
                tool.Size = new Size(290, 380); LlmBoundaryScope.Call(tool, "FitNativeSite"); Assert.AreEqual(0, moves);
                tool.Width = 291; LlmBoundaryScope.Call(tool, "FitNativeSite"); Assert.AreEqual(1, moves); Assert.AreEqual(new Size(290, 380), tool.Size);
                tool.Height = 381; LlmBoundaryScope.Call(tool, "FitNativeSite"); Assert.AreEqual(2, moves);
                tool.ClientReader = (IntPtr h, out ChatToolWindow.NativeRect r) => { r = new ChatToolWindow.NativeRect(); return true; }; LlmBoundaryScope.Call(tool, "FitNativeSite"); Assert.AreEqual(new Size(1, 1), tool.Size);
            }
        }
    }
}

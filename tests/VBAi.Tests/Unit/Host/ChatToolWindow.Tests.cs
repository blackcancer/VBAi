using Microsoft.VisualStudio.TestTools.UnitTesting;
using System;
using System.Drawing;
using System.Runtime.InteropServices;
using System.Windows.Forms;
using VBAi.Tests.Infrastructure;

namespace VBAi.Tests.Unit
{
    /// <summary>Vérifie le dimensionnement natif et le cycle d’attachement de ChatToolWindow.</summary>
    [TestClass, TestCategory("Unit")]
    public sealed class ChatToolWindowCoverageTests
    {
        private sealed class UnavailableBackgroundParent : Control
        {
            internal bool Unavailable;
            internal int Reads;
            public override Color BackColor
            {
                get
                {
                    if (Unavailable) { Reads++; throw new InvalidComObjectException("Ancestor background has detached."); }
                    return base.BackColor;
                }
                set { base.BackColor = value; }
            }
        }

        /// <summary>Exercises the actual framework paint path with an unavailable inherited background, without a native host or message-loop exception.</summary>
        [STATestMethod]
        [DataRow(false, false), DataRow(true, false), DataRow(false, true), DataRow(true, true)]
        public void OwnedBackgroundAvoidsUnavailableAncestorDuringActualFrameworkPainting(bool handle, bool dark)
        {
            using (var parent = new UnavailableBackgroundParent())
            using (var tool = new ChatToolWindow())
            using (var bitmap = new Bitmap(24, 24))
            using (var graphics = Graphics.FromImage(bitmap))
            {
                parent.Controls.Add(tool);
                if (handle) { var ignored = tool.Handle; }
                Color expected = dark ? Color.FromArgb(30, 30, 30) : SystemColors.Control;
                if (dark) tool.BackColor = expected;
                parent.Unavailable = true;
                try
                {
                    Assert.AreEqual(expected, tool.BackColor);
                    var paint = typeof(Control).GetMethod("OnPaintBackground", System.Reflection.BindingFlags.Instance | System.Reflection.BindingFlags.NonPublic);
                    paint.Invoke(tool, new object[] { new PaintEventArgs(graphics, new Rectangle(0, 0, 24, 24)) });
                    // OnPaintBackground asks GetClientRect(InternalHandle). Without
                    // an HWND the rectangle is empty; it must not create a window
                    // merely to draw this synthetic paint notification.
                    Assert.AreEqual(handle ? expected.ToArgb() : 0, bitmap.GetPixel(12, 12).ToArgb());
                    Assert.AreEqual(0, parent.Reads); Assert.AreEqual(handle, tool.IsHandleCreated);
                }
                finally { parent.Unavailable = false; }
            }
        }

        /// <summary>The native container copies its form's local background when attaching a chat or test explorer.</summary>
        [STATestMethod, DataRow(false), DataRow(true)]
        public void AttachedSurfaceOwnsItsFormsBackgroundInsteadOfTheActiveXAmbientColor(bool dark)
        {
            using (var tool = new ChatToolWindow())
            using (var child = new Form { BackColor = dark ? Color.FromArgb(30, 30, 30) : Color.White })
            {
                var ignored = tool.Handle;
                tool.Attach(child);
                Assert.AreEqual(child.BackColor, tool.BackColor);
                Assert.AreSame(tool, child.Parent);
                tool.StopSiteTrackingForShutdown(); child.Dispose();
                Assert.IsFalse(tool.IsDisposed);
            }
        }

        /// <summary>Shutdown leaves the child attached for disposal; ordinary detach still restores its floating layout.</summary>
        [STATestMethod]
        [DataRow(false, false)]
        [DataRow(false, true)]
        [DataRow(true, false)]
        [DataRow(true, true)]
        public void DetachmentSeparatesShutdownLifetimeFromFloatingPlacement(bool shutdown, bool handleCreated)
        {
            using (var tool = new ChatToolWindow())
            using (var child = new Form { TopLevel = false, FormBorderStyle = FormBorderStyle.None, Dock = DockStyle.Fill })
            {
                tool.Controls.Add(child);
                if (handleCreated) { var handle = child.Handle; child.Show(); }
                int creations = 0, destruction = 0;
                bool visible = child.Visible;
                child.HandleCreated += (sender, args) => creations++;
                child.HandleDestroyed += (sender, args) => destruction++;
                var timer = LlmBoundaryScope.Get<Timer>(tool, "siteResizeTimer"); timer.Start();
                if (shutdown) tool.StopSiteTrackingForShutdown(); else tool.Detach(child);
                Assert.IsFalse(timer.Enabled);
                if (shutdown) Assert.AreSame(tool, child.Parent); else Assert.IsNull(child.Parent);
                Assert.AreEqual(shutdown && visible, child.Visible);
                Assert.AreEqual(!shutdown, child.TopLevel);
                Assert.AreEqual(shutdown ? FormBorderStyle.None : FormBorderStyle.Sizable, child.FormBorderStyle);
                Assert.AreEqual(shutdown ? DockStyle.Fill : DockStyle.None, child.Dock);
                if (shutdown) { Assert.AreEqual(0, creations); Assert.AreEqual(0, destruction); }
            }
        }

        /// <summary>Only the owned form is disposed; its native site and unrelated children retain their ownership.</summary>
        [STATestMethod]
        [DataRow(false)]
        [DataRow(true)]
        public void ShutdownTrackingLeavesChildRemovalToItsDisposalAndDoesNotDisposeTheNativeSite(bool shown)
        {
            using (var tool = new ChatToolWindow())
            using (var child = new Form { TopLevel = false, FormBorderStyle = FormBorderStyle.None })
            using (var unrelated = new Panel())
            {
                tool.Controls.Add(child); tool.Controls.Add(unrelated);
                if (shown) child.Show();
                int parentChanges = 0, visibilityChanges = 0;
                child.ParentChanged += (sender, args) =>
                {
                    parentChanges++;
                    Assert.IsTrue(child.Disposing || child.IsDisposed, "Parent removal belongs to disposal, after its lifetime guard is set.");
                };
                child.VisibleChanged += (sender, args) => visibilityChanges++;
                var timer = LlmBoundaryScope.Get<Timer>(tool, "siteResizeTimer"); timer.Start();
                tool.StopSiteTrackingForShutdown();
                Assert.AreEqual(0, parentChanges); Assert.AreEqual(0, visibilityChanges);
                Assert.AreSame(tool, child.Parent); Assert.IsFalse(timer.Enabled);
                child.Dispose();
                Assert.IsTrue(child.IsDisposed); Assert.IsNull(child.Parent);
                Assert.IsFalse(tool.Controls.Contains(child)); Assert.IsFalse(tool.IsDisposed);
                Assert.AreSame(tool, unrelated.Parent); Assert.IsFalse(unrelated.IsDisposed);
            }
        }

        /// <summary>Preserves focus events while tolerating a detached native ActiveX site.</summary>
        [STATestMethod]
        public void LostFocusToleratesDetachedSiteWithoutRepeatingTheNotification()
        {
            using (var tool = new ChatToolWindow())
            {
                int attempts = 0;
                tool.LostFocus += (sender, args) =>
                {
                    attempts++;
                    throw new InvalidComObjectException("Detached native site");
                };
                LlmBoundaryScope.Call(tool, "OnLostFocus", EventArgs.Empty);
                Assert.AreEqual(1, attempts);
                Assert.IsFalse(tool.IsDisposed);
            }
        }

        /// <summary>Normal focus notifications still reach their subscribers.</summary>
        [STATestMethod]
        public void LostFocusStillForwardsTheOriginalEvent()
        {
            using (var tool = new ChatToolWindow())
            {
                int notifications = 0;
                tool.LostFocus += (sender, args) =>
                {
                    notifications++;
                    Assert.AreSame(tool, sender);
                    Assert.AreSame(EventArgs.Empty, args);
                };
                LlmBoundaryScope.Call(tool, "OnLostFocus", EventArgs.Empty);
                Assert.AreEqual(1, notifications);
            }
        }

        /// <summary>Unrelated focus errors remain observable instead of being hidden.</summary>
        [STATestMethod]
        public void LostFocusPropagatesUnrelatedErrorsWithoutRepeatingTheNotification()
        {
            using (var tool = new ChatToolWindow())
            {
                int attempts = 0;
                var failure = new InvalidOperationException("Unrelated focus failure");
                tool.LostFocus += (sender, args) => { attempts++; throw failure; };
                var wrapper = Assert.ThrowsException<System.Reflection.TargetInvocationException>(
                    () => LlmBoundaryScope.Call(tool, "OnLostFocus", EventArgs.Empty));
                Assert.AreSame(failure, wrapper.InnerException);
                Assert.AreEqual(1, attempts);
            }
        }

        /// <summary>Les coordonnées inversées ou dégénérées ne produisent jamais une taille négative.</summary>
        [STATestMethod]
        public void NativeSiteSizeClampsBothDimensionsWithoutUsingTheControlBounds()
        {
            using (var tool = new ChatToolWindow())
            {
                var handle = tool.Handle; tool.ParentReader = h => new IntPtr(123);
                foreach (var rect in new[] { new ChatToolWindow.NativeRect { Left = 10, Top = 20, Right = 210, Bottom = 320 }, new ChatToolWindow.NativeRect { Left = 10, Top = 20, Right = 5, Bottom = 3 }, new ChatToolWindow.NativeRect() })
                {
                    tool.ClientReader = (IntPtr h, out ChatToolWindow.NativeRect result) => { Assert.AreEqual(new IntPtr(123), h); result = rect; return true; };
                    Size size; Assert.IsTrue(tool.TryGetNativeSiteSize(out size));
                    Assert.AreEqual(new Size(Math.Max(0, rect.Right - rect.Left), Math.Max(0, rect.Bottom - rect.Top)), size);
                }
            }
        }
        [STATestMethod]
        public void NativeSiteSizeUsesParentClientAreaAndRefusesUnavailableHandles()
        {
            using (var tool = new ChatToolWindow())
            using (var owner = new Form { Left = -10000, Top = -10000, ShowInTaskbar = false, ClientSize = new Size(600, 800) })
            {
                Size size;
                Assert.IsFalse(tool.TryGetNativeSiteSize(out size));
                var handle = tool.Handle;
                var readParent = tool.ParentReader;
                tool.ParentReader = h => IntPtr.Zero;
                Assert.IsFalse(tool.TryGetNativeSiteSize(out size));
                tool.ParentReader = readParent;
                owner.Controls.Add(tool);
                owner.Show();
                Assert.IsTrue(tool.TryGetNativeSiteSize(out size));
                Assert.AreEqual(owner.ClientSize, size);
                tool.ClientReader = (IntPtr h, out ChatToolWindow.NativeRect r) => { r = default(ChatToolWindow.NativeRect); return false; };
                Assert.IsFalse(tool.TryGetNativeSiteSize(out size));
                tool.Dispose();
                Assert.IsFalse(tool.TryGetNativeSiteSize(out size));
            }
        }
        /// <summary>Utilise des handles WinForms réels pour vérifier le redimensionnement, l’attachement et la destruction.</summary>
        [STATestMethod]
        public void NativeSiteResizingAttachDetachTimerAndDisposeUseActualWinFormsHandles()
        {
            using (var scope = new HostUiScope())
            using (var owner = new Form { Left = -10000, Top = -10000, ClientSize = new Size(600, 800), ShowInTaskbar = false })
            using (var tool = new ChatToolWindow())
            using (var chat = new ChatWindow(new VbeSession(scope.Host)))
            {
                LlmBoundaryScope.Pump(LlmBoundaryScope.Get<System.Threading.Tasks.Task>(chat, "scopeLoad"));
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

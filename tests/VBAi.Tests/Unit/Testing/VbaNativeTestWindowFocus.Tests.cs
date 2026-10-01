using System;
using System.Collections.Generic;
using System.Linq;
using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace VBAi.Tests.Unit
{
    [TestClass, TestCategory("Unit")]
    public sealed class VbaNativeTestWindowFocusTests
    {
        internal sealed class Windows : VbaNativeTestWindowFocus.IWindows
        {
            internal readonly Dictionary<IntPtr, VbaNativeTestWindowFocus.Window> Items = new Dictionary<IntPtr, VbaNativeTestWindowFocus.Window>();
            internal readonly Dictionary<IntPtr, string> Captions = new Dictionary<IntPtr, string>();
            internal readonly List<IntPtr> CaptionReads = new List<IntPtr>();
            internal Action Activated, FocusAction;
            internal Func<IntPtr, VbaNativeTestWindowFocus.Window> ReadOverride;
            internal Func<IntPtr, IntPtr[]> ChildrenOverride;
            internal int Activations, Focuses;
            public uint CurrentProcess => 7;
            public uint CurrentThread => 8;
            public IntPtr Focused { get; set; }
            internal Windows()
            {
                Add(1, 0, "VbeMain"); Add(2, 1, "MDIClient"); Add(3, 2, "VbaWindow");
                Add(4, 3, "RichEdit20A"); Captions[new IntPtr(3)] = "Owned support (Code)";
            }
            internal void Add(int handle, int parent, string kind)
            {
                Items[new IntPtr(handle)] = new VbaNativeTestWindowFocus.Window { Handle = new IntPtr(handle), Parent = new IntPtr(parent),
                    Class = kind, Process = 7, Thread = 8, Exists = true, Enabled = true, Visible = true };
            }
            public VbaNativeTestWindowFocus.Window Read(IntPtr handle)
            {
                if (ReadOverride != null) return ReadOverride(handle);
                VbaNativeTestWindowFocus.Window item;
                if (!Items.TryGetValue(handle, out item)) return new VbaNativeTestWindowFocus.Window { Handle = handle };
                return new VbaNativeTestWindowFocus.Window { Handle = item.Handle, Parent = item.Parent, Class = item.Class,
                    Process = item.Process, Thread = item.Thread, Exists = item.Exists, Visible = item.Visible, Enabled = item.Enabled };
            }
            public IntPtr[] Children(IntPtr parent) => ChildrenOverride == null ? Items.Values.Where(item => item.Parent == parent).Select(item => item.Handle).ToArray() : ChildrenOverride(parent);
            public string Caption(IntPtr handle) { CaptionReads.Add(handle); string value; return Captions.TryGetValue(handle, out value) ? value : null; }
            public void Activate(IntPtr mdi, IntPtr child) { Assert.AreEqual(new IntPtr(2), mdi); Assert.AreEqual(new IntPtr(3), child); Activations++; Activated?.Invoke(); }
            public void Focus(IntPtr child) { Assert.AreEqual(new IntPtr(3), child); Focuses++; Focused = child; FocusAction?.Invoke(); }
        }

        [DataTestMethod]
        [DataRow(false)]
        [DataRow(true)]
        public void ExactCodeWindowAndOwnedDescendantFocusAreAcceptedOnce(bool descendant)
        {
            var windows = new Windows();
            if (descendant) windows.FocusAction = () => windows.Focused = new IntPtr(4);
            var target = VbaNativeTestWindowFocus.Focus(windows, new IntPtr(1), "Owned support (Code)");
            VbaNativeTestWindowFocus.VerifyFocus(windows, target);
            Assert.AreEqual(new IntPtr(3), target.Code.Handle);
            Assert.AreEqual(1, windows.Activations); Assert.AreEqual(1, windows.Focuses);
        }

        [DataTestMethod]
        [DataRow("mainProcess")]
        [DataRow("mainThread")]
        [DataRow("mainAbsent")]
        [DataRow("mdiAmbiguous")]
        [DataRow("codeAmbiguous")]
        [DataRow("codeProcess")]
        [DataRow("codeThread")]
        [DataRow("codeParent")]
        [DataRow("codeClass")]
        [DataRow("codeHidden")]
        [DataRow("codeDisabled")]
        [DataRow("caption")]
        [DataRow("emptyCaption")]
        public void InvalidNativeMappingCannotActivateOrFocus(string change)
        {
            var windows = new Windows();
            var root = windows.Items[new IntPtr(1)]; var code = windows.Items[new IntPtr(3)];
            switch (change)
            {
                case "mainProcess": root.Process = 9; break;
                case "mainThread": root.Thread = 9; break;
                case "mainAbsent": root.Exists = false; break;
                case "mdiAmbiguous": windows.Add(5, 1, "MDIClient"); break;
                case "codeAmbiguous": windows.Add(5, 2, "VbaWindow"); windows.Captions[new IntPtr(5)] = windows.Captions[new IntPtr(3)]; break;
                case "codeProcess": code.Process = 9; break;
                case "codeThread": code.Thread = 9; break;
                case "codeParent": code.Parent = new IntPtr(1); break;
                case "codeClass": code.Class = "ToolWindow"; break;
                case "codeHidden": code.Visible = false; break;
                case "codeDisabled": code.Enabled = false; break;
                case "caption": windows.Captions[new IntPtr(3)] = "Foreign caption"; break;
            }
            Assert.ThrowsException<InvalidOperationException>(() => VbaNativeTestWindowFocus.Focus(windows, new IntPtr(1),
                change == "emptyCaption" ? " " : "Owned support (Code)"));
            Assert.AreEqual(0, windows.Activations); Assert.AreEqual(0, windows.Focuses);
            if (change == "codeProcess" || change == "codeThread" || change == "codeParent" || change == "codeClass")
                Assert.IsFalse(windows.CaptionReads.Contains(new IntPtr(3)), "An ineligible window caption is private.");
        }

        [DataTestMethod]
        [DataRow("process")]
        [DataRow("thread")]
        [DataRow("parent")]
        [DataRow("class")]
        [DataRow("removed")]
        [DataRow("caption")]
        [DataRow("duplicate")]
        public void ActivationThatChangesTheNativeMappingCannotProceedToFocus(string change)
        {
            var windows = new Windows();
            windows.Activated = () =>
            {
                var code = windows.Items[new IntPtr(3)];
                switch (change)
                {
                    case "process": code.Process = 9; break;
                    case "thread": code.Thread = 9; break;
                    case "parent": code.Parent = new IntPtr(1); break;
                    case "class": code.Class = "Replacement"; break;
                    case "removed": code.Exists = false; break;
                    case "caption": windows.Captions[new IntPtr(3)] = "Replacement"; break;
                    case "duplicate": windows.Add(5, 2, "VbaWindow"); windows.Captions[new IntPtr(5)] = windows.Captions[new IntPtr(3)]; break;
                }
            };
            Assert.ThrowsException<InvalidOperationException>(() => VbaNativeTestWindowFocus.Focus(windows, new IntPtr(1), "Owned support (Code)"));
            Assert.AreEqual(1, windows.Activations); Assert.AreEqual(0, windows.Focuses);
        }

        [DataTestMethod]
        [DataRow("none")]
        [DataRow("root")]
        [DataRow("foreignProcess")]
        [DataRow("foreignThread")]
        [DataRow("cycle")]
        [DataRow("changedCode")]
        public void UnconfirmedOrChangedFocusIsRefusedWithoutASecondAttempt(string change)
        {
            var windows = new Windows();
            windows.FocusAction = () =>
            {
                windows.Focused = new IntPtr(4);
                switch (change)
                {
                    case "none": windows.Focused = IntPtr.Zero; break;
                    case "root": windows.Focused = new IntPtr(1); break;
                    case "foreignProcess": windows.Items[new IntPtr(4)].Process = 9; break;
                    case "foreignThread": windows.Items[new IntPtr(4)].Thread = 9; break;
                    case "cycle": windows.Items[new IntPtr(4)].Parent = new IntPtr(4); break;
                    case "changedCode": windows.Items[new IntPtr(3)].Class = "Replacement"; break;
                }
            };
            Assert.ThrowsException<InvalidOperationException>(() => VbaNativeTestWindowFocus.Focus(windows, new IntPtr(1), "Owned support (Code)"));
            Assert.AreEqual(1, windows.Activations); Assert.AreEqual(1, windows.Focuses);
        }
        private static void OnSta(Action action)
        {
            Exception error = null;
            var thread = new System.Threading.Thread(() => { try { action(); } catch (Exception caught) { error = caught; } });
            thread.SetApartmentState(System.Threading.ApartmentState.STA); thread.Start(); thread.Join();
            if (error != null) System.Runtime.ExceptionServices.ExceptionDispatchInfo.Capture(error).Throw();
        }

        [TestMethod]
        public void NativeBoundaryReadsOnlyDisposableWindowsAndBoundsEnumerationAndCaptions()
        {
            OnSta(() =>
            {
                var windows = new VbaNativeTestWindowFocus.NativeWindows();
                using (var root = new System.Windows.Forms.Form { Text = "Owned test window" })
                using (var child = new System.Windows.Forms.TextBox { Text = "Child" })
                {
                    root.Controls.Add(child); root.Show();
                    var rootHandle = root.Handle; var childHandle = child.Handle;
                    var observed = windows.Read(childHandle);
                    Assert.IsTrue(observed.Exists); Assert.IsTrue(observed.Visible); Assert.IsTrue(observed.Enabled);
                    Assert.AreEqual(windows.CurrentProcess, observed.Process); Assert.AreEqual(windows.CurrentThread, observed.Thread);
                    Assert.AreEqual(rootHandle, observed.Parent);
                    CollectionAssert.Contains(windows.Children(rootHandle), childHandle);
                    Assert.AreEqual("Owned test window", windows.Caption(rootHandle));
                    root.Text = ""; Assert.IsNull(windows.Caption(rootHandle));
                    root.Text = new string('X', 1025); Assert.IsNull(windows.Caption(rootHandle));
                    Assert.IsNull(windows.Caption(IntPtr.Zero));
                    windows.Activate(rootHandle, childHandle); windows.Focus(childHandle);
                    Assert.AreEqual(childHandle, windows.Focused);
                    Assert.IsFalse(windows.Read(IntPtr.Zero).Exists);
                    for (int i = 0; i < 512; i++) { var c = new System.Windows.Forms.Control(); root.Controls.Add(c); var handle = c.Handle; }
                    Assert.ThrowsException<InvalidOperationException>(() => windows.Children(rootHandle));
                }
            });
        }

        [TestMethod]
        public void NullZeroAndDiagnosticControlCharactersAreHandledWithoutCaptionInspection()
        {
            Assert.ThrowsException<ArgumentNullException>(() => VbaNativeTestWindowFocus.Focus(null, IntPtr.Zero, "Code"));
            var windows = new Windows();
            Assert.ThrowsException<InvalidOperationException>(() => VbaNativeTestWindowFocus.Focus(windows, new IntPtr(1), new string('x', 1025)));
            foreach (string failure in new[] { "zero", "zeroPid", "zeroThread", "nullClass", "controlClass", "longClass", "hiddenMain", "disabledMain" })
            {
                windows = new Windows(); var root = windows.Items[new IntPtr(1)];
                switch (failure)
                {
                    case "zero": root.Handle = IntPtr.Zero; break;
                    case "zeroPid": root.Process = 0; break;
                    case "zeroThread": root.Thread = 0; break;
                    case "nullClass": root.Class = null; root.Exists = false; break;
                    case "controlClass": root.Class = "Bad\r\nClass"; root.Exists = false; break;
                    case "longClass": root.Class = new string('x', 97); root.Exists = false; break;
                    case "hiddenMain": root.Visible = false; break;
                    case "disabledMain": root.Enabled = false; break;
                }
                var error = Assert.ThrowsException<InvalidOperationException>(() => VbaNativeTestWindowFocus.Focus(windows, new IntPtr(1), "Owned support (Code)"));
                Assert.IsFalse(error.Message.Contains("Bad\r\nClass"));
                Assert.AreEqual(0, windows.CaptionReads.Count);
            }
        }

        [TestMethod]
        public void CapturedTargetIntegrityAndDepthBoundAreRevalidated()
        {
            foreach (string failure in new[] { "handle", "process", "thread", "mdiParent", "codeParent", "mdiClass", "codeClass", "depth" })
            {
                var windows = new Windows();
                var target = VbaNativeTestWindowFocus.Focus(windows, new IntPtr(1), "Owned support (Code)");
                switch (failure)
                {
                    case "handle": target.Main.Handle = new IntPtr(4); break;
                    case "process": target.Main.Process = 6; break;
                    case "thread": target.Main.Thread = 6; break;
                    case "mdiParent": target.Mdi.Parent = new IntPtr(4); windows.Items[new IntPtr(2)].Parent = new IntPtr(4); break;
                    case "codeParent": target.Code.Parent = new IntPtr(4); windows.Items[new IntPtr(3)].Parent = new IntPtr(4); break;
                    case "mdiClass": target.Mdi.Class = "Other"; windows.Items[new IntPtr(2)].Class = "Other"; break;
                    case "codeClass": target.Code.Class = "Other"; windows.Items[new IntPtr(3)].Class = "Other"; break;
                    case "depth": for (int i = 5; i < 40; i++) windows.Add(i, i - 1, "Child"); windows.Focused = new IntPtr(39); break;
                }
                Assert.ThrowsException<InvalidOperationException>(() => VbaNativeTestWindowFocus.VerifyFocus(windows, target));
                Assert.AreEqual(1, windows.Activations); Assert.AreEqual(1, windows.Focuses);
            }
        }

        private sealed class CaptionRaceWindow : System.Windows.Forms.NativeWindow, IDisposable
        {
            internal CaptionRaceWindow() { CreateHandle(new System.Windows.Forms.CreateParams { Caption = "Short" }); }
            protected override void WndProc(ref System.Windows.Forms.Message message)
            {
                if (message.Msg == 0x000E) { message.Result = new IntPtr(10); return; }
                base.WndProc(ref message);
            }
            public void Dispose() { DestroyHandle(); }
        }

        [TestMethod]
        public void ChangedCaptionLengthCannotProduceAnUnverifiedCaption()
        {
            OnSta(() => {
                using (var window = new CaptionRaceWindow())
                    Assert.IsNull(new VbaNativeTestWindowFocus.NativeWindows().Caption(window.Handle));
            });
        }

        [TestMethod]
        public void NullSnapshotsAndIndirectChildrenCannotBypassNativeWindowOwnership()
        {
            var empty = new Windows { ReadOverride = _ => null };
            StringAssert.Contains(Assert.ThrowsException<InvalidOperationException>(() => VbaNativeTestWindowFocus.Focus(empty, new IntPtr(1), "Code")).Message, "failed=Null");
            var f = new Windows();
            f.Add(5, 4, "VbaWindow");
            f.Add(6, 1, "Other");
            f.Add(7, 4, "MDIClient");
            f.Captions[new IntPtr(5)] = "Owned support (Code)";
            f.ChildrenOverride = parent => parent == new IntPtr(1) ? new[] { new IntPtr(2), new IntPtr(6), new IntPtr(7) } : new[] { new IntPtr(3), new IntPtr(4), new IntPtr(5) };
            var target = VbaNativeTestWindowFocus.Focus(f, new IntPtr(1), "Owned support (Code)");
            Assert.IsFalse(f.CaptionReads.Contains(new IntPtr(5)));
            f.Add(8, 1, "MDIClient");
            f.ChildrenOverride = parent => parent == new IntPtr(1) ? new[] { new IntPtr(2), new IntPtr(8) } : new[] { new IntPtr(3) };
            Assert.ThrowsException<InvalidOperationException>(() => VbaNativeTestWindowFocus.VerifyFocus(f, target));
        }

    }
}

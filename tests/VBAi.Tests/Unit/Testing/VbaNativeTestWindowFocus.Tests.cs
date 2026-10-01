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
                VbaNativeTestWindowFocus.Window item;
                if (!Items.TryGetValue(handle, out item)) return new VbaNativeTestWindowFocus.Window { Handle = handle };
                return new VbaNativeTestWindowFocus.Window { Handle = item.Handle, Parent = item.Parent, Class = item.Class,
                    Process = item.Process, Thread = item.Thread, Exists = item.Exists, Visible = item.Visible, Enabled = item.Enabled };
            }
            public IntPtr[] Children(IntPtr parent) => Items.Values.Where(item => item.Parent == parent).Select(item => item.Handle).ToArray();
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
    }
}

namespace VBAi.Tests.Unit
{
    using Microsoft.VisualStudio.TestTools.UnitTesting;
    using System;
    using System.Collections;
    using System.Collections.Generic;
    using VBAi;

    [TestClass]
    [TestCategory("Unit")]
    public sealed class VbeWindowLayoutTests
    {
        public sealed class Host { public List<Window> Windows { get; } = new List<Window>(); }
        public sealed class Window
        {
            public Window(string name, int type) { Caption = name; Type = type; LinkedWindows = new Members(this); }
            public string Caption { get; set; }
            public int Type { get; set; }
            public bool Visible { get; set; } = true;
            private int state;
            public bool FailAfterStateChange { get; set; }
            public bool IgnoreStateChange { get; set; }
            public string CaptionAfterStateChange { get; set; }
            public int WindowState
            {
                get => state;
                set
                {
                    if (IgnoreStateChange) return;
                    state = value;
                    if (CaptionAfterStateChange != null) Caption = CaptionAfterStateChange;
                    if (FailAfterStateChange) throw new InvalidOperationException("Partial state failure");
                }
            }
            public int Left { get; set; }
            public int Top { get; set; }
            private int width = 300;
            public bool ConstrainNextWidth { get; set; }
            public int Width { get => width; set { width = ConstrainNextWidth ? 250 : value; ConstrainNextWidth = false; } }
            public int Height { get; set; } = 200;
            public Window LinkedWindowFrame { get; set; }
            public Members LinkedWindows { get; }
        }
        public sealed class Members : IEnumerable<Window>
        {
            private readonly Window owner;
            private readonly List<Window> windows = new List<Window>();
            public bool FailAfterAdd { get; set; }
            public int AddCalls { get; private set; }
            public Members(Window owner) { this.owner = owner; }
            public void Add(Window window)
            {
                AddCalls++;
                window.LinkedWindowFrame?.LinkedWindows.windows.Remove(window);
                windows.Add(window); window.LinkedWindowFrame = owner;
                if (FailAfterAdd) throw new InvalidOperationException("Partial native failure");
            }
            public void Remove(Window window)
            {
                windows.Remove(window);
                var floating = new Window("Floating", 11);
                floating.LinkedWindows.Add(window);
            }
            public IEnumerator<Window> GetEnumerator() => windows.GetEnumerator();
            IEnumerator IEnumerable.GetEnumerator() => GetEnumerator();
        }
        private static string Version(VbeEditorWindows editor, Window window)
        { return (string)((dynamic)editor.WindowLayout(window.Caption, window.Type)).WindowVersion; }
        [TestMethod]
        public void WindowStateRequiresCurrentRevisionAndStandaloneTarget()
        {
            var host = new Host(); var window = new Window("Code", 0); host.Windows.Add(window);
            var editor = new VbeEditorWindows(host);
            var request = new Request { WindowCaption = "Code", WindowType = 0, Action = "maximize", ExpectedWindowVersion = Version(editor, window) };
            window.Top++;
            Assert.ThrowsException<InvalidOperationException>(() => editor.SetWindowState(request));
            Assert.AreEqual(0, window.WindowState);
            request.ExpectedWindowVersion = Version(editor, window);
            dynamic result = editor.SetWindowState(request);
            Assert.IsTrue((bool)result.Verified); Assert.AreEqual(2, window.WindowState);
            request.ExpectedWindowVersion = Version(editor, window);
            result = editor.SetWindowState(request);
            Assert.IsFalse((bool)result.Applied); Assert.IsTrue((bool)result.Verified);
            window.LinkedWindowFrame = new Window("Frame", 11);
            request.ExpectedWindowVersion = Version(editor, window);
            request.Action = "restore";
            Assert.ThrowsException<InvalidOperationException>(() => editor.SetWindowState(request));
            Assert.AreEqual(2, window.WindowState);
        }

        [TestMethod]
        public void StateReadbackFollowsChangedCaptionAndReportsPartialFailures()
        {
            var host = new Host(); var window = new Window("Code", 0) { CaptionAfterStateChange = "Book - Code" }; host.Windows.Add(window);
            var editor = new VbeEditorWindows(host);
            var request = new Request { WindowCaption = window.Caption, WindowType = 0, Action = "maximize", ExpectedWindowVersion = Version(editor, window) };
            dynamic result = editor.SetWindowState(request);
            Assert.IsTrue((bool)result.Verified);
            Assert.AreEqual("Book - Code", (string)result.After.Window.Properties["Caption"]);
            request.WindowCaption = window.Caption; request.Action = "restore"; request.ExpectedWindowVersion = Version(editor, window);
            window.IgnoreStateChange = true;
            result = editor.SetWindowState(request);
            Assert.IsFalse((bool)result.Verified); Assert.IsTrue((bool)result.VerificationPending);
            window.IgnoreStateChange = false; window.FailAfterStateChange = true;
            result = editor.SetWindowState(request);
            Assert.IsFalse((bool)result.Verified); Assert.IsNull((bool?)result.Applied);
            Assert.AreEqual(0, (int)result.ActualState);
            StringAssert.Contains((string)result.NativeError, "Partial state failure");
        }

        [TestMethod]
        public void LinkThenUnlinkVerifiesMembershipIncludingFloatingFrames()
        {
            var host = new Host(); var pane = new Window("Locals", 4); var frame = new Window("Target", 11);
            host.Windows.Add(pane); host.Windows.Add(frame);
            var editor = new VbeEditorWindows(host);
            var request = new Request
            {
                WindowCaption = pane.Caption,
                WindowType = pane.Type,
                ExpectedWindowVersion = Version(editor, pane),
                Action = "link",
                TargetWindowCaption = frame.Caption,
                TargetWindowType = frame.Type,
                ExpectedTargetWindowVersion = Version(editor, frame)
            };
            dynamic linked = editor.LinkWindow(request);
            Assert.IsTrue((bool)linked.Verified); Assert.AreSame(frame, pane.LinkedWindowFrame);
            request.Action = "unlink"; request.ExpectedWindowVersion = Version(editor, pane);
            dynamic unlinked = editor.LinkWindow(request);
            Assert.IsTrue((bool)unlinked.Verified);
            Assert.AreNotSame(frame, pane.LinkedWindowFrame);
            Assert.IsFalse((bool)unlinked.PersistenceVerified);
        }
        [TestMethod]
        public void StaleDestinationAndAmbiguousIdentityRefuseMutation()
        {
            var host = new Host(); var pane = new Window("Locals", 4); var frame = new Window("Target", 11);
            host.Windows.Add(pane); host.Windows.Add(frame);
            var editor = new VbeEditorWindows(host);
            var request = new Request
            {
                WindowCaption = pane.Caption,
                WindowType = pane.Type,
                ExpectedWindowVersion = Version(editor, pane),
                Action = "link",
                TargetWindowCaption = frame.Caption,
                TargetWindowType = frame.Type,
                ExpectedTargetWindowVersion = Version(editor, frame)
            };
            frame.Width++;
            Assert.ThrowsException<InvalidOperationException>(() => editor.LinkWindow(request));
            Assert.AreEqual(0, frame.LinkedWindows.AddCalls);
            request.ExpectedTargetWindowVersion = Version(editor, frame);
            frame.LinkedWindows.Add(new Window("Watch", 3));
            Assert.ThrowsException<InvalidOperationException>(() => editor.LinkWindow(request));
            Assert.AreEqual(1, frame.LinkedWindows.AddCalls);
            host.Windows.Add(new Window("Locals", 4));
            Assert.ThrowsException<InvalidOperationException>(() => editor.WindowLayout("Locals", 4));
        }
        [TestMethod]
        public void PartialNativeFailureNeverClaimsVerifiedOrRolledBack()
        {
            var host = new Host(); var pane = new Window("Locals", 4); var frame = new Window("Target", 11);
            host.Windows.Add(pane); host.Windows.Add(frame); frame.LinkedWindows.FailAfterAdd = true;
            var editor = new VbeEditorWindows(host);
            dynamic result = editor.LinkWindow(new Request
            {
                WindowCaption = pane.Caption,
                WindowType = pane.Type,
                ExpectedWindowVersion = Version(editor, pane),
                Action = "link",
                TargetWindowCaption = frame.Caption,
                TargetWindowType = frame.Type,
                ExpectedTargetWindowVersion = Version(editor, frame)
            });
            Assert.IsNull((bool?)result.Applied); Assert.IsFalse((bool)result.Verified);
            Assert.IsTrue((bool)result.VerificationPending); Assert.AreSame(frame, pane.LinkedWindowFrame);
            StringAssert.Contains((string)result.NativeError, "Partial native failure");
        }
        [TestMethod]
        public void ConstrainedBoundsAreRestoredAndStaleRevisionIsRefused()
        {
            var host = new Host(); var window = new Window("Floating", 11); host.Windows.Add(window);
            var editor = new VbeEditorWindows(host);
            var request = new Request { WindowCaption = window.Caption, WindowType = window.Type, ExpectedWindowVersion = Version(editor, window), Left = 10, Top = 20, Width = 400, Height = 300 };
            window.ConstrainNextWidth = true;
            var error = Assert.ThrowsException<InvalidOperationException>(() => editor.SetWindowBounds(request));
            StringAssert.Contains(error.Message, "Original bounds restored"); Assert.AreEqual(300, window.Width); Assert.AreEqual(0, window.Left);
            dynamic result = editor.SetWindowBounds(request); Assert.IsTrue((bool)result.Verified);
            Assert.ThrowsException<InvalidOperationException>(() => editor.SetWindowBounds(request));
        }
    }
}

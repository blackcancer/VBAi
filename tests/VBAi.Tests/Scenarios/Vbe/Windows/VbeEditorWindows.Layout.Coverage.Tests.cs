using System;
using System.Collections;
using System.Collections.Generic;
using System.Dynamic;
using VBAi;
using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace VBAi.Tests.Unit
{
    [TestClass]
    [TestCategory("Unit")]
    public sealed class VbeWindowLayoutCoverageTests
    {
        public sealed class Host { public List<object> Windows { get; } = new List<object>(); }
        public sealed class NativeWindow : DynamicObject
        {
            public readonly Dictionary<string, object> Values = new Dictionary<string, object> {
                { "Caption", "Window" }, { "Type", 0 }, { "Visible", true }, { "WindowState", 0 },
                { "Left", 0 }, { "Top", 0 }, { "Width", 300 }, { "Height", 200 }, { "LinkedWindowFrame", null } };
            public readonly HashSet<string> FailReads = new HashSet<string>();
            public Action<string, object> BeforeWrite;
            public string IgnoreWrite, FailAfterWrite;
            public int Writes;
            public NativeWindow(string caption, int type)
            { Values["Caption"] = caption; Values["Type"] = type; Values["LinkedWindows"] = new Members(this); }
            public override bool TryGetMember(GetMemberBinder binder, out object value)
            {
                if (FailReads.Contains(binder.Name) || Writes > 0 && binder.Name == FailAfterWrite) throw new InvalidOperationException("read " + binder.Name);
                return Values.TryGetValue(binder.Name, out value);
            }
            public override bool TrySetMember(SetMemberBinder binder, object value)
            { Writes++; BeforeWrite?.Invoke(binder.Name, value); if (IgnoreWrite != binder.Name) Values[binder.Name] = value; return true; }
        }
        public sealed class Members : IEnumerable
        {
            public readonly List<NativeWindow> Items = new List<NativeWindow>();
            public readonly NativeWindow Owner;
            public bool FailEnumeration;
            public int Enumerations, ReturnNullOnEnumeration;
            public Action<NativeWindow> AddNative, RemoveNative;
            public Members(NativeWindow owner) { Owner = owner; }
            public void Add(NativeWindow window)
            { if (AddNative != null) { AddNative(window); return; } Items.Add(window); window.Values["LinkedWindowFrame"] = Owner; }
            public void Remove(NativeWindow window)
            { if (RemoveNative != null) { RemoveNative(window); return; } Items.Remove(window); window.Values["LinkedWindowFrame"] = null; }
            public IEnumerator GetEnumerator()
            { if (++Enumerations == ReturnNullOnEnumeration) return null; if (FailEnumeration) throw new InvalidOperationException("members unavailable"); return Items.ToArray().GetEnumerator(); }
        }
        private static string Version(VbeEditorWindows editor, NativeWindow window)
        { return (string)((dynamic)editor.WindowLayout((string)window.Values["Caption"], (int)window.Values["Type"])).WindowVersion; }
        private static Request Request(VbeEditorWindows editor, NativeWindow window, string action = "restore")
        { return new Request { WindowCaption = (string)window.Values["Caption"], WindowType = (int)window.Values["Type"], Action = action,
            ExpectedWindowVersion = Version(editor, window), Left = 10, Top = 20, Width = 400, Height = 300 }; }

        [TestMethod]
        public void NativeWindowLayoutsExposeUnreadableFramesAndRequireCompleteRevisions()
        {
            var host = new Host(); var window = new NativeWindow("Frame", 12); host.Windows.Add(window);
            var editor = new VbeEditorWindows(host); var members = (Members)window.Values["LinkedWindows"];
            members.Items.Add(new NativeWindow("Pane", 4)); editor.WindowLayout("Frame", 12);
            members.FailEnumeration = true;
            Assert.IsNull(Version(editor, window));
            Assert.ThrowsException<InvalidOperationException>(() => editor.SetWindowState(Request(editor, window)));
            members.FailEnumeration = false;
            window.FailReads.Add("LinkedWindowFrame"); Assert.IsNull(Version(editor, window)); window.FailReads.Clear();
            foreach (string expected in new[] { null, "wrong" })
            {
                var request = Request(editor, window); request.ExpectedWindowVersion = expected;
                Assert.ThrowsException<InvalidOperationException>(() => editor.SetWindowState(request));
            }
            Assert.ThrowsException<ArgumentException>(() => editor.SetWindowState(Request(editor, window, "invalid")));
            foreach (string action in new[] { "minimize", "maximize", "restore" })
                Assert.IsTrue((bool)((dynamic)editor.SetWindowState(Request(editor, window, action))).Verified);
            window.FailAfterWrite = "WindowState"; window.Writes = 0;
            dynamic result = editor.SetWindowState(Request(editor, window, "maximize"));
            Assert.IsFalse((bool)result.Verified); Assert.IsNotNull((string)result.ReadbackError);
        }

        [TestMethod]
        public void NativeBoundsValidateEveryCoordinateAndVerifyEachWriteAndRollback()
        {
            var host = new Host(); var window = new NativeWindow("Window", 0); host.Windows.Add(window);
            var editor = new VbeEditorWindows(host);
            foreach (double value in new[] { double.NaN, double.PositiveInfinity, 1.5, -32769, 32768 })
            {
                var request = Request(editor, window); request.Left = value;
                Assert.ThrowsException<ArgumentException>(() => editor.SetWindowBounds(request));
            }
            foreach (var size in new[] { new[] { 79, 300 }, new[] { 400, 59 } })
            {
                var request = Request(editor, window); request.Width = size[0]; request.Height = size[1];
                Assert.ThrowsException<ArgumentException>(() => editor.SetWindowBounds(request));
            }
            window.Values["WindowState"] = 2;
            Assert.ThrowsException<InvalidOperationException>(() => editor.SetWindowBounds(Request(editor, window)));
            window.Values["WindowState"] = 0; window.Values["LinkedWindowFrame"] = new NativeWindow("Frame", 11);
            Assert.ThrowsException<InvalidOperationException>(() => editor.SetWindowBounds(Request(editor, window)));
            window.Values["LinkedWindowFrame"] = null;
            foreach (string property in new[] { "Left", "Top", "Width", "Height" })
            {
                window.IgnoreWrite = property;
                StringAssert.Contains(Assert.ThrowsException<InvalidOperationException>(() => editor.SetWindowBounds(Request(editor, window))).Message, "Original bounds restored");
            }
            window.IgnoreWrite = "Width";
            window.BeforeWrite = (name, value) => { if (name == "Width") window.Values[name] = (int)value + 1; };
            StringAssert.Contains(Assert.ThrowsException<InvalidOperationException>(() => editor.SetWindowBounds(Request(editor, window))).Message, "rollback failed");
            window.BeforeWrite = null; window.IgnoreWrite = null;
            Assert.IsTrue((bool)((dynamic)editor.SetWindowBounds(Request(editor, window))).Verified);
        }

        [TestMethod]
        public void NativeLinkingValidatesTypesAnchorsAndReportsMembershipReadbacks()
        {
            var sameNameHost = new Host();
            var sameNamePane = new NativeWindow("Shared", 4); var sameNameFrame = new NativeWindow("Shared", 11);
            sameNameHost.Windows.Add(sameNamePane); sameNameHost.Windows.Add(sameNameFrame);
            ((Members)sameNameFrame.Values["LinkedWindows"]).Items.Add(new NativeWindow("Existing", 3));
            var sameNameEditor = new VbeEditorWindows(sameNameHost);
            var sameNameRequest = Request(sameNameEditor, sameNamePane, "link");
            sameNameRequest.TargetWindowCaption = "Shared"; sameNameRequest.TargetWindowType = 11;
            sameNameRequest.ExpectedTargetWindowVersion = Version(sameNameEditor, sameNameFrame);
            Assert.IsTrue((bool)((dynamic)sameNameEditor.LinkWindow(sameNameRequest)).Verified);
            var nullHost = new Host(); var nullPane = new NativeWindow("Pane", 4); var nullFrame = new NativeWindow("Frame", 11);
            nullHost.Windows.Add(nullPane); nullHost.Windows.Add(nullFrame); var nullEditor = new VbeEditorWindows(nullHost);
            var nullRequest = Request(nullEditor, nullPane, "link"); nullRequest.TargetWindowCaption = "Frame"; nullRequest.TargetWindowType = 11;
            nullRequest.ExpectedTargetWindowVersion = Version(nullEditor, nullFrame);
            var nullMembers = (Members)nullFrame.Values["LinkedWindows"]; nullMembers.Enumerations = 0; nullMembers.ReturnNullOnEnumeration = 2;
            Assert.ThrowsException<NullReferenceException>(() => nullEditor.LinkWindow(nullRequest));
            foreach (int type in new[] { 1, 8, 9, 11, 12, 14, 16 })
                Assert.ThrowsException<ArgumentException>(() => new VbeEditorWindows(new Host()).LinkWindow(new Request { Action = "link", WindowType = type }));
            Assert.ThrowsException<ArgumentException>(() => new VbeEditorWindows(new Host()).LinkWindow(new Request { Action = "invalid" }));
            foreach (int paneType in new[] { 2, 7, 10, 15 })
            foreach (int frameType in new[] { 11, 12, 4, 10 })
            {
                var host = new Host(); var pane = new NativeWindow("Pane", paneType); var frame = new NativeWindow("Frame", frameType);
                var actualFrame = frameType == 4 || frameType == 10 ? new NativeWindow("Container", 11) : frame;
                if (frameType == 4 || frameType == 10) frame.Values["LinkedWindowFrame"] = actualFrame;
                host.Windows.Add(pane); host.Windows.Add(frame); var editor = new VbeEditorWindows(host);
                var r = Request(editor, pane, "link"); r.TargetWindowCaption = "Pane"; r.TargetWindowType = paneType;
                Assert.ThrowsException<ArgumentException>(() => editor.LinkWindow(r));
                r.TargetWindowCaption = "Frame"; r.TargetWindowType = frameType; r.ExpectedTargetWindowVersion = Version(editor, frame);
                if (frameType == 4)
                {
                    frame.Values["LinkedWindowFrame"] = null; r.ExpectedTargetWindowVersion = Version(editor, frame);
                    Assert.ThrowsException<InvalidOperationException>(() => editor.LinkWindow(r));
                    frame.Values["LinkedWindowFrame"] = actualFrame; r.ExpectedTargetWindowVersion = Version(editor, frame);
                }
                Assert.IsTrue((bool)((dynamic)editor.LinkWindow(r)).Verified);
                r = Request(editor, pane, "unlink"); Assert.IsTrue((bool)((dynamic)editor.LinkWindow(r)).Verified);
                r = Request(editor, pane, "unlink"); Assert.IsFalse((bool)((dynamic)editor.LinkWindow(r)).Applied);
            }
        }

        [TestMethod]
        public void NativeLinkFailuresNeverClaimVerifiedRecoveryOrAssumeOldFramesSurvive()
        {
            foreach (int outcome in new[] { 0, 1, 2, 3, 4, 5, 6 })
            {
                var host = new Host(); var pane = new NativeWindow("Pane", 4); var frame = new NativeWindow("Frame", 11);
                host.Windows.Add(pane); host.Windows.Add(frame); var editor = new VbeEditorWindows(host);
                var members = (Members)frame.Values["LinkedWindows"];
                var r = Request(editor, pane, "link"); r.TargetWindowCaption = "Frame"; r.TargetWindowType = 11; r.ExpectedTargetWindowVersion = Version(editor, frame);
                members.AddNative = item => {
                    if (outcome == 0) return;
                    item.Values["LinkedWindowFrame"] = outcome == 1 ? new NativeWindow("Other", 11) : frame;
                    members.Items.Add(new NativeWindow("OtherCaption", 4)); members.Items.Add(new NativeWindow("Pane", 3));
                    if (outcome >= 3) members.Items.Add(item);
                    if (outcome == 4) members.Items.Add(item);
                    if (outcome == 5 || outcome == 6)
                    { if (outcome == 6) item.FailReads.Add("Caption"); throw new InvalidOperationException("native link failed"); }
                };
                dynamic result = editor.LinkWindow(r);
                Assert.AreEqual(outcome == 3, (bool)result.Verified);
                if (outcome == 6) Assert.IsNotNull((string)result.ReadbackError);
            }
            foreach (int outcome in new[] { 0, 1, 2, 3 })
            {
                var host = new Host(); var pane = new NativeWindow("Pane", 4); var frame = new NativeWindow("Frame", 11);
                host.Windows.Add(pane); var members = (Members)frame.Values["LinkedWindows"]; members.Items.Add(pane); pane.Values["LinkedWindowFrame"] = frame;
                var editor = new VbeEditorWindows(host); var r = Request(editor, pane, "unlink");
                members.RemoveNative = item => {
                    if (outcome == 0) return;
                    var floating = new NativeWindow("Floating", 11); item.Values["LinkedWindowFrame"] = floating;
                    var items = (Members)floating.Values["LinkedWindows"];
                    items.Items.Add(outcome == 1 ? new NativeWindow("Wrong", 4) : item);
                    if (outcome == 3) items.Items.Add(new NativeWindow("Other", 3));
                };
                Assert.AreEqual(outcome == 2, (bool)((dynamic)editor.LinkWindow(r)).Verified);
            }
        }
    }
}

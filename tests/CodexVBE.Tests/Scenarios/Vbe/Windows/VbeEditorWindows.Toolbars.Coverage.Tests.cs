using System;
using System.Collections.Generic;
using System.Dynamic;
using System.Reflection;
using CodexVBE;
using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace CodexVBE.Tests.Unit
{
    [TestClass]
    [TestCategory("Unit")]
    public sealed class VbeToolbarCoverageTests
    {
        public sealed class Host { public List<object> CommandBars { get; } = new List<object>(); }
        public sealed class BrokenMessageException : Exception
        { public override string Message { get { throw new InvalidOperationException("error details unavailable"); } } }
        public sealed class NativeBar : DynamicObject
        {
            public readonly Dictionary<string, object> Values = new Dictionary<string, object> {
                { "Name", "Standard" }, { "Type", 0 }, { "Visible", true }, { "Enabled", true }, { "BuiltIn", true },
                { "Protection", 0 }, { "Position", 4 }, { "RowIndex", 1 }, { "Left", 0 }, { "Top", 0 }, { "Width", 200 }, { "Height", 24 } };
            public readonly HashSet<string> FailReads = new HashSet<string>();
            public string FailAfterWrite, IgnoreWrite, AlterAfterWrite;
            public bool FailWrite, InvalidErrorDetails;
            public int Writes;
            public override bool TryGetMember(GetMemberBinder binder, out object result)
            {
                if (FailReads.Contains(binder.Name) || (Writes > 0 && binder.Name == FailAfterWrite))
                { if (InvalidErrorDetails) throw new BrokenMessageException(); throw new InvalidOperationException("read " + binder.Name); }
                return Values.TryGetValue(binder.Name, out result);
            }
            public override bool TrySetMember(SetMemberBinder binder, object value)
            {
                Writes++;
                if (FailWrite) throw new InvalidOperationException("write failed");
                if (binder.Name != IgnoreWrite) Values[binder.Name] = value;
                if (AlterAfterWrite != null) Values[AlterAfterWrite] = AlterAfterWrite == "Visible" ? (object)false : -1;
                return true;
            }
        }
        private static Request Command(VbeEditorWindows editor, string action)
        {
            var method = typeof(VbeEditorWindows).GetMethod("ToolbarSnapshot", BindingFlags.Static | BindingFlags.NonPublic);
            var bars = (Host)typeof(VbeEditorWindows).GetField("vbe", BindingFlags.Instance | BindingFlags.NonPublic).GetValue(editor);
            dynamic state = method.Invoke(null, new[] { bars.CommandBars[0] });
            return new Request { ObjectName = "Standard", Action = action, ExpectedWindowVersion = state.WindowVersion,
                ExpectedToolbarLayoutVersion = state.ToolbarLayoutVersion, ToolbarLeft = 20, ToolbarTop = 30, RowIndex = 2 };
        }
        private static VbeEditorWindows Editor(NativeBar bar)
        { var host = new Host(); host.CommandBars.Add(bar); return new VbeEditorWindows(host); }

        [TestMethod]
        public void ToolbarCatalogueAndMutationsRejectIncompleteIdentityAndGeometry()
        {
            var bar = new NativeBar(); var editor = Editor(bar);
            var valid = Command(editor, "hide");
            foreach (string action in new[] { null, "invalid" })
            {
                Assert.ThrowsException<ArgumentException>(() => editor.SetToolbarVisibility(new Request { Action = action }));
                Assert.ThrowsException<ArgumentException>(() => editor.SetToolbarPlacement(new Request { Action = action }));
                Assert.ThrowsException<ArgumentException>(() => editor.SetToolbarPosition(new Request { Action = action }));
            }
            foreach (string name in new[] { null, " " })
            {
                Assert.ThrowsException<ArgumentException>(() => editor.SetToolbarVisibility(new Request { Action = "hide", ObjectName = name, ExpectedWindowVersion = valid.ExpectedWindowVersion }));
                Assert.ThrowsException<ArgumentException>(() => editor.SetToolbarPosition(new Request { Action = "top", ObjectName = name }));
            }
            Assert.ThrowsException<ArgumentException>(() => editor.SetToolbarVisibility(new Request { Action = "hide", ObjectName = "Standard" }));
            Assert.ThrowsException<InvalidOperationException>(() => editor.SetToolbarPosition(new Request { Action = "top", ObjectName = "Missing" }));
            bar.FailReads.Add("Type"); editor.Toolbars(); bar.FailReads.Clear();
            bar.Values["Type"] = 1; editor.Toolbars(); bar.Values["Type"] = 0;
            foreach (string property in new[] { "Name", "Width" })
            {
                bar.FailReads.Add(property); editor.Toolbars();
                if (property == "Width")
                {
                    valid.Action = "float";
                    Assert.ThrowsException<InvalidOperationException>(() => editor.SetToolbarPlacement(valid));
                    valid.Action = "hide";
                }
                else
                    Assert.ThrowsException<InvalidOperationException>(() => editor.SetToolbarVisibility(valid));
                bar.FailReads.Clear();
            }
            foreach (string revision in new[] { null, "wrong" })
            {
                var request = Command(editor, "float"); request.ExpectedToolbarLayoutVersion = revision;
                Assert.ThrowsException<InvalidOperationException>(() => editor.SetToolbarPlacement(request));
                request.Action = "top"; Assert.ThrowsException<InvalidOperationException>(() => editor.SetToolbarPosition(request));
            }
            bar.FailReads.Add("Enabled");
            Assert.ThrowsException<InvalidOperationException>(() => editor.SetToolbarVisibility(valid));
            bar.FailReads.Clear(); Assert.AreEqual(0, bar.Writes);
        }

        [TestMethod]
        public void ToolbarPlacementValidatesEachBoundDockingStateAndProtectionBeforeWrites()
        {
            var bar = new NativeBar(); var editor = Editor(bar);
            foreach (var coordinates in new int?[][] { new int?[] { null, 0 }, new int?[] { 0, null },
                new int?[] { -32769, 0 }, new int?[] { 32768, 0 }, new int?[] { 0, -32769 }, new int?[] { 0, 32768 } })
            {
                var r = Command(editor, "float"); r.ToolbarLeft = coordinates[0]; r.ToolbarTop = coordinates[1];
                Assert.ThrowsException<ArgumentException>(() => editor.SetToolbarPlacement(r));
            }
            foreach (int? row in new int?[] { null, 0 })
            {
                var r = Command(editor, "row"); r.RowIndex = row;
                Assert.ThrowsException<ArgumentException>(() => editor.SetToolbarPlacement(r));
            }
            foreach (int position in new[] { -1, 4, 5 })
            {
                bar.Values["Position"] = position;
                Assert.ThrowsException<InvalidOperationException>(() => editor.SetToolbarPlacement(Command(editor, "row")));
            }
            bar.Values["Position"] = 0;
            Assert.ThrowsException<InvalidOperationException>(() => editor.SetToolbarPlacement(Command(editor, "float")));
            bar.Values["Enabled"] = false;
            Assert.ThrowsException<InvalidOperationException>(() => editor.SetToolbarPlacement(Command(editor, "row")));
            Assert.ThrowsException<InvalidOperationException>(() => editor.SetToolbarPosition(Command(editor, "float")));
            bar.Values["Enabled"] = true;
            foreach (int protection in new[] { 4, 16 })
            {
                bar.Values["Protection"] = protection;
                Assert.ThrowsException<InvalidOperationException>(() => editor.SetToolbarPlacement(Command(editor, "row")));
            }
            bar.Values["Protection"] = 0;
            foreach (string action in new[] { "left", "top", "right", "bottom", "float", "float" })
                editor.SetToolbarPosition(Command(editor, action));
        }

        [TestMethod]
        public void ToolbarMutationsReportPartialFailuresAndUnverifiableReadbacks()
        {
            foreach (string action in new[] { "visibility", "position", "float", "row" })
            foreach (int fault in new[] { 0, 1, 2, 3, 4, 5, 6, 7, 8 })
            {
                var bar = new NativeBar();
                bar.Values["Position"] = action == "row" ? 1 : 4;
                var editor = Editor(bar); var r = Command(editor, action == "visibility" ? "hide" : action == "position" ? "top" : action);
                string property = action == "visibility" ? "Visible" : action == "position" ? "Position" : action == "float" ? "Left" : "RowIndex";
                bar.FailWrite = fault == 1 || fault == 6;
                bar.IgnoreWrite = fault == 2 ? property : fault == 7 && action == "float" ? "Top" : null;
                bar.FailAfterWrite = fault == 4 || fault == 5 || fault == 6 ? property : null;
                bar.InvalidErrorDetails = fault == 5 || fault == 6;
                bar.AlterAfterWrite = fault == 3 ? (action == "visibility" ? "Enabled" : "Visible") : fault == 8 && (action == "float" || action == "row") ? "Position" : null;
                dynamic result = action == "visibility" ? editor.SetToolbarVisibility(r) : action == "position" ? editor.SetToolbarPosition(r) : editor.SetToolbarPlacement(r);
                if (fault == 0) Assert.IsTrue((bool)result.Verified);
                if (fault == 1 || fault >= 4 && fault <= 6) Assert.IsFalse((bool)result.Verified);
                if (fault == 1) StringAssert.Contains((string)result.NativeError, "write failed");
                if (fault == 5) StringAssert.Contains((string)result.NativeError, "error details unavailable");
                if (fault == 6) StringAssert.Contains((string)result.NativeError, "write failed");
            }
        }
    }
}

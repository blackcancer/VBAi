namespace VBAi.Tests.Unit
{
    using System;
    using System.Security.Cryptography;
    using System.Text;
    using System.Web.Script.Serialization;
    using VBAi;
    using Microsoft.VisualStudio.TestTools.UnitTesting;

    [TestClass]
    [TestCategory("Unit")]
    public sealed class VbeEditorCompletionTests
    {
        public sealed class CodeSnapshot { public string Code { get; set; } public string Sha256 { get; set; } }
        private static string Hash(string text)
        { using (var sha = SHA256.Create()) return BitConverter.ToString(sha.ComputeHash(Encoding.UTF8.GetBytes(text))).Replace("-", "").ToLowerInvariant(); }
        [TestMethod]
        public void LiteralReplacementHonorsRangeAndDoesNotExpandDollarTokens()
        {
            Assert.AreEqual("value\r\n$1 value2\r\nvalue", VbaTextEdits.Transform("value\r\nvalue value2\r\nvalue",
                new Request { StartLine = 2, Count = 1, Action = "replace", Query = "value", Text = "$1", WholeWord = true }));
        }
        [TestMethod]
        public void IdentifierReplacementSkipsStringsCommentsAndLongerNames()
        {
            string source = "value = \"value\" ' value\r\nIf True Then Rem value\r\nvalue2 = VALUE\r\n[ value ] = #1/1/2020#";
            string expected = "renamed = \"value\" ' value\r\nIf True Then Rem value\r\nvalue2 = renamed\r\n[ value ] = #1/1/2020#";
            Assert.AreEqual(expected, VbaTextEdits.ReplaceIdentifier(source, "value", "renamed"));
        }
        [TestMethod]
        public void InvalidRangesAndIndentSizesAreRejected()
        {
            Assert.ThrowsException<ArgumentException>(() => VbaTextEdits.Transform("one", new Request { StartLine = 1, Count = 2, Action = "comment" }));
            Assert.ThrowsException<ArgumentException>(() => VbaTextEdits.Transform("one", new Request { StartLine = 1, Count = 1, Action = "indent", InsertIndex = 0 }));
        }
        [TestMethod]
        public void CommentsAndIndentationLeaveUnselectedLinesUntouched()
        {
            var request = new Request { StartLine = 2, Count = 1, Action = "comment" };
            Assert.AreEqual("a\r\n'  b\r\nc", VbaTextEdits.Transform("a\r\n  b\r\nc", request));
            request.Action = "unindent"; request.InsertIndex = 4;
            Assert.AreEqual("a\r\nb\r\nc", VbaTextEdits.Transform("a\r\n  b\r\nc", request));
        }
        [TestMethod]
        public void UndoRedoRejectsExternalChangesAndNewEditsClearRedo()
        {
            string code = "one"; int writes = 0; VbeCodeEdits edits = null;
            edits = new VbeCodeEdits(r => {
                if (r.Command == "read_module") return Response.Success(new CodeSnapshot { Code = code });
                Assert.AreEqual(Hash(code), r.ExpectedSha256);
                string before = code; code = r.Text; writes++;
                edits.Record(r.Project, r.Module, before, code);
                return Response.Success(new { Sha256 = Hash(code) });
            });
            var request = new Request { Project = "p", Module = "m", ExpectedSha256 = Hash(code), StartLine = 1, Count = 1, Action = "replace", Query = "one", Text = "two" };
            edits.Edit(request, true); Assert.AreEqual(0, writes);
            edits.Edit(request, false); Assert.AreEqual("two", code);
            code = "external"; request.ExpectedSha256 = Hash(code);
            Assert.ThrowsException<InvalidOperationException>(() => edits.Replay(request, false));
            Assert.AreEqual(1, writes);
            code = "two"; request.ExpectedSha256 = Hash(code); edits.Replay(request, false); Assert.AreEqual("one", code);
            request.ExpectedSha256 = Hash(code); edits.Replay(request, true); Assert.AreEqual("two", code);
            request.ExpectedSha256 = Hash(code); edits.Replay(request, false);
            request.ExpectedSha256 = Hash(code); request.Text = "three"; edits.Edit(request, false);
            request.ExpectedSha256 = Hash(code);
            Assert.ThrowsException<InvalidOperationException>(() => edits.Replay(request, true));
        }
        private static FormLayoutBox Box(string name, double left, double width = 10)
        { return new FormLayoutBox { Path = name, Left = left, Top = 5, Width = width, Height = 10 }; }
        [TestMethod]
        public void LayoutAnchorsFirstSelectedControlWithoutMutatingInput()
        {
            var input = new[] { Box("b", 40), Box("a", 10, 20) };
            var result = FormLayoutPlan.Create(input, "align_right", 0, 100, 100);
            Assert.AreEqual(30d, result[1].Left); Assert.AreEqual(10d, input[1].Left);
        }
        [TestMethod]
        public void CenterAlignmentPreservesSizeAndRejectsNegativeCoordinates()
        {
            var input = new[] { Box("anchor", 40, 11), Box("other", 10, 20) };
            input[0].Top = 30; input[0].Height = 11; input[1].Height = 20;
            var centers = FormLayoutPlan.Create(input, "align_centers", 0, 100, 100);
            Assert.AreEqual(35.5, centers[1].Left); Assert.AreEqual(20d, centers[1].Width);
            Assert.AreEqual(input[1].Top, centers[1].Top);
            var middles = FormLayoutPlan.Create(input, "align_middles", 0, 100, 100);
            Assert.AreEqual(25.5, middles[1].Top); Assert.AreEqual(20d, middles[1].Height);
            Assert.AreEqual(input[1].Left, middles[1].Left);
            input[0].Left = 0; input[0].Top = 0;
            Assert.ThrowsException<InvalidOperationException>(() => FormLayoutPlan.Create(input, "align_centers", 0, 100, 100));
            Assert.ThrowsException<InvalidOperationException>(() => FormLayoutPlan.Create(input, "align_middles", 0, 100, 100));
        }

        [TestMethod]
        public void SpacingKeepsVisualOrderAndChangesEachOriginalGap()
        {
            var input = new[] { Box("last", 80, 15), Box("first", 10, 20), Box("middle", 40, 10) };
            var fixedGap = FormLayoutPlan.Create(input, "space_horizontal", 0, 200, 100);
            Assert.AreEqual(40d, fixedGap[0].Left); Assert.AreEqual(10d, fixedGap[1].Left); Assert.AreEqual(30d, fixedGap[2].Left);
            var increased = FormLayoutPlan.Create(input, "increase_horizontal_spacing", 5, 200, 100);
            Assert.AreEqual(90d, increased[0].Left); Assert.AreEqual(45d, increased[2].Left);
            var decreased = FormLayoutPlan.Create(input, "decrease_horizontal_spacing", 5, 200, 100);
            Assert.AreEqual(70d, decreased[0].Left); Assert.AreEqual(35d, decreased[2].Left);
            Assert.AreEqual(80d, input[0].Left);
            Assert.ThrowsException<InvalidOperationException>(() => FormLayoutPlan.Create(input, "decrease_horizontal_spacing", 11, 200, 100));
            Assert.ThrowsException<InvalidOperationException>(() => FormLayoutPlan.Create(input, "space_horizontal", 100, 200, 100));
            Assert.ThrowsException<ArgumentException>(() => FormLayoutPlan.Create(input, "space_horizontal", double.NaN, 200, 100));
        }

        [TestMethod]
        public void DistributionUsesEqualGapsAndRejectsImpossibleBounds()
        {
            var result = FormLayoutPlan.Create(new[] { Box("a", 0), Box("b", 12), Box("c", 50) }, "distribute_horizontal", 0, 100, 100);
            Assert.AreEqual(25d, result[1].Left);
            Assert.ThrowsException<InvalidOperationException>(() => FormLayoutPlan.Create(new[] { Box("a", 0, 30), Box("b", 12, 30), Box("c", 50, 30) }, "distribute_horizontal", 0, 100, 100));
        }
        [TestMethod]
        public void LayoutRejectsInvalidDimensionsAndOutOfBoundsResults()
        {
            var input = new[] { Box("a", 0), Box("b", 90, 20) };
            Assert.ThrowsException<ArgumentException>(() => FormLayoutPlan.Create(input, "align_top", 0, double.NaN, 100));
            Assert.ThrowsException<InvalidOperationException>(() => FormLayoutPlan.Create(input, "align_top", 0, 100, 100));
        }
        [TestMethod]
        public void ContextMonitorDetectsSameLineCountEditsWithoutKeyCollisions()
        {
            string hash = "one"; int changed = 0;
            var monitor = new VbeContextMonitor(r => {
                if (r.Command == "list_projects") return Response.Success(new[] { new { Name = "p" } });
                if (r.Command == "list_modules") return Response.Success(new[] { new { Name = "projects", Lines = 1 } });
                if (r.Command == "read_module") return Response.Success(new { Sha256 = hash });
                return Response.Success(new object[0]);
            });
            monitor.Changed += () => changed++;
            for (int i = 0; i < 6; i++) monitor.Step("p");
            Assert.AreEqual(0, changed);
            hash = "two";
            for (int i = 0; i < 3; i++) monitor.Step("p");
            Assert.AreEqual(1, changed);
            monitor.Step("another"); Assert.AreEqual(1, changed);
        }
        [TestMethod]
        public void BookmarksReturnNamesAndRejectStaleTargets()
        {
            string code = "abc"; Request selection = null;
            var navigation = new VbeNavigationHistory(null, r => {
                if (r.Command == "read_module") return Response.Success(new CodeSnapshot { Code = code, Sha256 = Hash(code) });
                selection = r; return Response.Success(new { Selected = true });
            });
            var request = new Request { Project = "p", Module = "m", Query = "entry", Action = "add", StartLine = 1, StartColumn = 2, ExpectedSha256 = Hash(code) };
            navigation.Bookmark(request);
            request.Action = "list";
            string json = new JavaScriptSerializer().Serialize(navigation.Bookmark(request));
            StringAssert.Contains(json, "\"Name\":\"entry\""); Assert.IsFalse(json.Contains("\\u0000"));
            request.Action = "go"; navigation.Bookmark(request);
            Assert.AreEqual("select_code_range", selection.Command); Assert.AreEqual(2, selection.EndColumn);
            code = "changed";
            Assert.ThrowsException<InvalidOperationException>(() => navigation.Bookmark(request));
        }
    }
}

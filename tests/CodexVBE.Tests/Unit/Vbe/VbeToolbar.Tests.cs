using System;
using System.Collections.Generic;
using System.Web.Script.Serialization;
using CodexVBE;
using Microsoft.VisualStudio.TestTools.UnitTesting;
namespace CodexVBE.Tests.Unit
{
    [TestClass, TestCategory("Unit")]
    public sealed class VbeToolbarTests
    {
        public sealed class Bar
        {
            public string Name = "Standard"; public int Type; public bool Enabled = true, BuiltIn = true;
            public int Protection, Position, RowIndex, Left, Top, Width = 200, Height = 24, Writes;
            private bool visible = true;
            public bool Visible { get { return visible; } set { visible = value; Writes++; } }
        }
        public sealed class Host { public List<Bar> CommandBars = new List<Bar>(); }
        private static Request Request(VbeEditorWindows windows, string action)
        {
            var json = new JavaScriptSerializer();
            var data = (Dictionary<string, object>)json.DeserializeObject(json.Serialize(windows.Toolbars()));
            var first = (Dictionary<string, object>)((object[])data["Toolbars"])[0];
            return new Request { ObjectName = "Standard", Action = action, ExpectedWindowVersion = (string)first["WindowVersion"], ExpectedToolbarLayoutVersion = (string)first["ToolbarLayoutVersion"] };
        }
        [TestMethod]
        public void StaleVisibilityRevisionDoesNotWriteAndNoopDoesNotWrite()
        {
            var bar = new Bar(); var host = new Host(); host.CommandBars.Add(bar);
            var windows = new VbeEditorWindows(host); var stale = Request(windows, "show");
            windows.SetToolbarVisibility(Request(windows, "hide")); Assert.IsFalse(bar.Visible);
            Assert.ThrowsException<InvalidOperationException>(() => windows.SetToolbarVisibility(stale));
            windows.SetToolbarVisibility(Request(windows, "hide")); Assert.AreEqual(1, bar.Writes);
            windows.SetToolbarVisibility(Request(windows, "show")); Assert.IsTrue(bar.Visible);
        }
        [TestMethod]
        public void PlacementRequiresCoordinatesAndFreshGeometryAndCorrectDockingMode()
        {
            var bar = new Bar { Position = 4 }; var host = new Host(); host.CommandBars.Add(bar);
            var windows = new VbeEditorWindows(host); var request = Request(windows, "float");
            Assert.ThrowsException<ArgumentException>(() => windows.SetToolbarPlacement(request));
            request.ToolbarLeft = 200; request.ToolbarTop = 150;
            windows.SetToolbarPlacement(request); Assert.AreEqual(200, bar.Left); Assert.AreEqual(150, bar.Top);
            Assert.ThrowsException<InvalidOperationException>(() => windows.SetToolbarPlacement(request));
            request = Request(windows, "row"); request.RowIndex = 2;
            Assert.ThrowsException<InvalidOperationException>(() => windows.SetToolbarPlacement(request));
            bar.Position = 1; request = Request(windows, "row"); request.RowIndex = 2;
            windows.SetToolbarPlacement(request); Assert.AreEqual(2, bar.RowIndex);
            bar.Protection = 4; request = Request(windows, "row"); request.RowIndex = 3;
            Assert.ThrowsException<InvalidOperationException>(() => windows.SetToolbarPlacement(request));
            Assert.AreEqual(2, bar.RowIndex);
        }
        [DataTestMethod]
        [DataRow(4, "top")]
        [DataRow(16, "top")]
        [DataRow(32, "left")]
        [DataRow(32, "right")]
        [DataRow(64, "top")]
        [DataRow(64, "bottom")]
        public void NativeDockingProtectionIsNotOverridden(int protection, string action)
        {
            var bar = new Bar { Position = 4, Protection = protection }; var host = new Host(); host.CommandBars.Add(bar);
            var windows = new VbeEditorWindows(host);
            var error = Assert.ThrowsException<InvalidOperationException>(() => windows.SetToolbarPosition(Request(windows, action)));
            StringAssert.Contains(error.Message, "protection"); Assert.AreEqual(4, bar.Position); Assert.AreEqual(protection, bar.Protection);
        }
        [TestMethod]
        public void ChangedGeometryIsRejectedAndHideProtectionIsHonored()
        {
            var bar = new Bar { Position = 4 }; var host = new Host(); host.CommandBars.Add(bar);
            var windows = new VbeEditorWindows(host); var request = Request(windows, "top");
            bar.Left++;
            Assert.ThrowsException<InvalidOperationException>(() => windows.SetToolbarPosition(request));
            Assert.AreEqual(4, bar.Position);
            bar.Protection = 8;
            Assert.ThrowsException<InvalidOperationException>(() => windows.SetToolbarVisibility(Request(windows, "hide")));
            Assert.IsTrue(bar.Visible); Assert.AreEqual(0, bar.Writes);
        }
        [TestMethod]
        public void MenusAmbiguityAndDisabledBarsCannotBeChanged()
        {
            var bar = new Bar(); var host = new Host(); host.CommandBars.Add(bar);
            var windows = new VbeEditorWindows(host); var request = Request(windows, "hide");
            bar.Type = 1;
            Assert.ThrowsException<InvalidOperationException>(() => windows.SetToolbarVisibility(request));
            bar.Type = 0; host.CommandBars.Add(new Bar());
            Assert.ThrowsException<InvalidOperationException>(() => windows.SetToolbarVisibility(request));
            host.CommandBars.RemoveAt(1); bar.Visible = false; bar.Enabled = false; bar.Writes = 0;
            Assert.ThrowsException<InvalidOperationException>(() => windows.SetToolbarVisibility(Request(windows, "show")));
            Assert.AreEqual(0, bar.Writes); Assert.IsFalse(bar.Enabled);
        }
    }
}

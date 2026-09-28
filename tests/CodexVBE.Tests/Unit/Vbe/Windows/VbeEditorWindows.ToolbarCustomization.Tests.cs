namespace CodexVBE.Tests.Unit
{
    using System;
    using System.Collections.Generic;
    using System.Linq;
    using System.Web.Script.Serialization;
    using CodexVBE;
    using Microsoft.VisualStudio.TestTools.UnitTesting;

    [TestClass, TestCategory("Unit")]
    public sealed class ToolbarCustomizationTests
    {
        private static readonly JavaScriptSerializer Json = new JavaScriptSerializer();
        private static Dictionary<string, object> Data(object value) => Json.Deserialize<Dictionary<string, object>>(Json.Serialize(value));
        [TestMethod]
        public void NativeCommandCustomizationRoundTripProtectsOriginalButtonsAndRevisions()
        {
            var host = new Host(); var bar = host.CommandBars.Add("Standard", 1, false, true); bar.BuiltIn = true;
            var original = bar.Controls.Add(1, 42, Type.Missing, 1, false);
            original.Caption = "Native command"; original.BuiltIn = true;
            var service = new VbeEditorWindows(host);
            string revision = (string)Data(service.ToolbarControls(new Request { ObjectName = "Standard" }))["ToolbarControlsVersion"];
            var request = new Request { ObjectName = "Standard", ExpectedToolbarControlsVersion = revision, ControlId = 42, ControlCaption = "Native command", InsertIndex = 1 };
            var added = Data(service.AddToolbarCommand(request));
            Assert.IsTrue((bool)added["Verified"]); Assert.AreEqual(2, bar.Controls.Count);
            Assert.IsTrue(bar.Controls[1].Temporary);
            Assert.ThrowsException<InvalidOperationException>(() => service.AddToolbarCommand(request));
            request.ExpectedToolbarControlsVersion = (string)Data(service.ToolbarControls(request))["ToolbarControlsVersion"];
            request.InsertIndex = 2;
            Assert.ThrowsException<InvalidOperationException>(() => service.RemoveToolbarCommand(request));
            request.InsertIndex = 1;
            Assert.IsTrue((bool)Data(service.RemoveToolbarCommand(request))["Verified"]);
            Assert.AreSame(original, bar.Controls[1]);
            request.ExpectedToolbarControlsVersion = (string)Data(service.ToolbarControls(request))["ToolbarControlsVersion"];
            request.ControlCaption = "Different command";
            Assert.ThrowsException<InvalidOperationException>(() => service.AddToolbarCommand(request));
            Assert.AreEqual(1, bar.Controls.Count);
            request.ControlCaption = "Native command"; request.Temporary = false;
            var persistent = Data(service.AddToolbarCommand(request));
            Assert.IsTrue((bool)persistent["Verified"]); Assert.IsFalse((bool)persistent["Temporary"]);
            Assert.IsFalse(bar.Controls[1].Temporary); Assert.IsFalse((bool)persistent["PersistenceVerified"]);
        }
        [TestMethod]
        public void EmptyCustomToolbarCreationDeletionAndProtectedBarsAreVerified()
        {
            var host = new Host(); var standard = host.CommandBars.Add("Standard", 1, false, true); standard.BuiltIn = true;
            var service = new VbeEditorWindows(host);
            string version = (string)Data(service.Toolbars())["ToolbarCollectionVersion"];
            var request = new Request { ObjectName = "Coverage", ExpectedToolbarCollectionVersion = version };
            var created = Data(service.CreateToolbar(request)); Assert.IsTrue((bool)created["Verified"]);
            Assert.AreEqual("VBAi - Coverage", (string)created["ObjectName"]);
            Assert.ThrowsException<InvalidOperationException>(() => service.CreateToolbar(request));
            request.ObjectName = "VBAi - Coverage"; request.ExpectedToolbarCollectionVersion = (string)created["ToolbarCollectionVersion"];
            request.ExpectedToolbarControlsVersion = (string)Data(service.ToolbarControls(request))["ToolbarControlsVersion"];
            Assert.IsTrue((bool)Data(service.RemoveToolbar(request))["Verified"]);
            request.ObjectName = "Standard";
            request.ExpectedToolbarCollectionVersion = (string)Data(service.Toolbars())["ToolbarCollectionVersion"];
            request.ExpectedToolbarControlsVersion = (string)Data(service.ToolbarControls(request))["ToolbarControlsVersion"];
            Assert.ThrowsException<InvalidOperationException>(() => service.RemoveToolbar(request));
            standard.Protection = 1; request.ExpectedToolbarControlsVersion = (string)Data(service.ToolbarControls(request))["ToolbarControlsVersion"];
            Assert.ThrowsException<InvalidOperationException>(() => service.AddToolbarCommand(request));
        }
        public sealed class Host { public Bars CommandBars { get; } = new Bars(); }
        public sealed class Bars : List<Bar>
        {
            public Bar Add(string name, int position, bool menuBar, bool temporary)
            { var bar = new Bar { Owner = this, Name = name, Position = position }; Add(bar); return bar; }
            public Button FindControl(int type, int id) => this.SelectMany(x => x.Controls).FirstOrDefault(x => x.Type == type && x.Id == id);
        }
        public sealed class Bar
        {
            internal Bars Owner;
            public string Name { get; set; }
            public int Type => 0;
            public bool BuiltIn { get; set; }
            public bool Visible { get; set; } = true;
            public bool Enabled { get; set; } = true;
            public int Protection { get; set; }
            public int Position { get; set; }
            public int Left => 0;
            public int Top => 0;
            public int Width => 200;
            public int Height => 20;
            public int RowIndex => 1;
            public Buttons Controls { get; } = new Buttons();
            public void Delete() => Owner.Remove(this);
        }
        public sealed class Buttons : List<Button>
        {
            public new Button this[int oneBased] => base[oneBased - 1];
            public Button Add(int type, int id, object parameter, int before, bool temporary)
            {
                var button = new Button { Owner = this, Id = id, Type = type, Temporary = temporary, Caption = "Native command", BuiltIn = true };
                Insert(before - 1, button); return button;
            }
        }
        public sealed class Button
        {
            internal Buttons Owner;
            public int Index => Owner.IndexOf(this) + 1;
            public int Id { get; set; }
            public int Type { get; set; }
            public string Caption { get; set; }
            public string Tag { get; set; } = "";
            public bool BuiltIn { get; set; }
            public bool Visible => true;
            public bool Temporary { get; set; }
            public void Delete() => Owner.Remove(this);
        }
    }
}

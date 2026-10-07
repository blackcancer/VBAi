using Microsoft.VisualStudio.TestTools.UnitTesting;
using System;
using System.Collections.Generic;

namespace VBAi.Tests.Unit
{
    [TestClass]
    public sealed class VbeDebugGeneralCommandTests
    {
        public sealed class Control
        {
            public int Id { get; set; } = 2578; public int Type { get; set; } = 1; public string Caption { get; set; } = "Project Properties...";
            public bool Enabled { get; set; } = true; public List<Control> Controls { get; } = new List<Control>();
            public int Executes { get; private set; }
            public void Execute() { Executes++; }
        }
        public sealed class Bar { public string Name { get; set; } public List<Control> Controls { get; } = new List<Control>(); }
        public sealed class Bars : List<Bar> { public Bar ActiveMenuBar => this[0]; }
        public sealed class Host { public Bars CommandBars { get; } = new Bars(); }
        private static Host Create(Control command)
        {
            var host = new Host(); var menu = new Bar { Name = "Menu" }; menu.Controls.Add(command); host.CommandBars.Add(menu);
            var other = new Bar { Name = "Project Context" }; other.Controls.Add(new Control()); host.CommandBars.Add(other); return host;
        }
        [TestMethod]
        public void GeneralCaptionDiscoveryUsesExactLocalizedActiveMenuWithoutExecution()
        {
            var control = new Control { Caption = "Propriétés de Source..." }; var host = Create(control);
            Assert.AreEqual(control.Caption, new VbeDebug(host).ReadGeneralCommandCaption());
            Assert.AreEqual(0, control.Executes);
        }
        [DataTestMethod]
        [DataRow("disabled")]
        [DataRow("duplicate")]
        [DataRow("type")]
        [DataRow("id")]
        [DataRow("empty-menu")]
        [DataRow("empty-caption")]
        public void GeneralCaptionDiscoveryRefusesUnknownOrAmbiguousActiveMenu(string fault)
        {
            var control = new Control(); var host = Create(control);
            if (fault == "disabled") control.Enabled = false;
            if (fault == "duplicate") host.CommandBars[0].Controls.Add(new Control());
            if (fault == "type") control.Type = 10;
            if (fault == "id") control.Id = 1;
            if (fault == "empty-menu") host.CommandBars[0].Name = "";
            if (fault == "empty-caption") control.Caption = "";
            Assert.ThrowsException<InvalidOperationException>(() => new VbeDebug(host).ReadGeneralCommandCaption());
            Assert.AreEqual(0, control.Executes);
        }
        [TestMethod]
        public void ExactMenuRouteIgnoresOtherContextsAndExecutesOnlyAfterEntryGuardOnce()
        {
            var control = new Control(); var host = Create(control); var service = new VbeDebug(host); int live = 0, entry = 0;
            var execute = service.CaptureGeneralCommand(new Request { ControlCaption = control.Caption }, () => live++);
            execute(() => { entry++; Assert.AreEqual(0, control.Executes); Assert.AreEqual(2, live); });
            Assert.AreEqual(1, control.Executes); Assert.AreEqual(1, entry);
            Assert.ThrowsException<InvalidOperationException>(() => execute(() => entry++)); Assert.AreEqual(1, control.Executes);
        }
        [DataTestMethod]
        [DataRow("caption")]
        [DataRow("id")]
        [DataRow("path")]
        [DataRow("disabled")]
        [DataRow("duplicate")]
        [DataRow("type")]
        [DataRow("authorization")]
        [DataRow("entry")]
        public void ChangedResolvedCommandOrFinalAuthorityNeverExecutesAndCannotRetry(string fault)
        {
            var control = new Control(); var host = Create(control); bool valid = true; var service = new VbeDebug(host);
            var execute = service.CaptureGeneralCommand(new Request { ControlCaption = control.Caption }, () => { if (!valid) throw new InvalidOperationException("revoked"); });
            if (fault == "caption") control.Caption = "other"; if (fault == "id") control.Id = 999; if (fault == "path") host.CommandBars[0].Name = "Other";
            if (fault == "disabled") control.Enabled = false; if (fault == "duplicate") host.CommandBars[0].Controls.Add(new Control()); if (fault == "type") control.Type = 10; if (fault == "authorization") valid = false;
            Assert.ThrowsException<InvalidOperationException>(() => execute(() => { if (fault == "entry") throw new InvalidOperationException("entry proof changed"); }));
            Assert.AreEqual(0, control.Executes); Assert.ThrowsException<InvalidOperationException>(() => execute(() => { })); Assert.AreEqual(0, control.Executes);
        }
    }
}

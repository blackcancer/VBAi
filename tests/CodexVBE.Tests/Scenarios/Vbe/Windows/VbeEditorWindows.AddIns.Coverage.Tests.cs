using System;
using System.Collections.Generic;
using CodexVBE;
using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace CodexVBE.Tests.Unit
{
    [TestClass]
    [TestCategory("Unit")]
    public sealed class VbeAddInReadbackTests
    {
        public sealed class Host { public List<Plugin> AddIns { get; } = new List<Plugin>(); }
        public sealed class Plugin
        {
            public string ProgId { get; set; } = "Probe.AddIn";
            public string Guid { get; set; } = "invalid-guid";
            public bool FailDescription;
            public string Description { get { if (FailDescription) throw new InvalidOperationException("description unavailable"); return "probe"; } }
            public Action AfterWrite;
            private bool connected;
            public bool Connect { get { return connected; } set { connected = value; AfterWrite?.Invoke(); } }
        }
        private static Request Request(VbeEditorWindows editor, Plugin plugin)
        { dynamic data = editor.AddIns(); return new Request { Action = "connect", ProgId = plugin.ProgId, ExpectedAddInVersion = data.AddIns[0].AddInVersion }; }

        [TestMethod]
        public void AddInConnectionRejectsMalformedIdentityAndReportsDisappearanceOrIdentityChanges()
        {
            var host = new Host(); var plugin = new Plugin(); host.AddIns.Add(plugin); var editor = new VbeEditorWindows(host);
            Assert.ThrowsException<ArgumentException>(() => editor.SetAddInConnection(new Request()));
            foreach (string identity in new[] { null, " ", new string('x', 256) })
                Assert.ThrowsException<ArgumentException>(() => editor.SetAddInConnection(new Request { Action = "connect", ProgId = identity }));
            var request = Request(editor, plugin); request.ExpectedAddInVersion = null;
            Assert.ThrowsException<InvalidOperationException>(() => editor.SetAddInConnection(request));
            request = Request(editor, plugin); plugin.FailDescription = true;
            Assert.ThrowsException<InvalidOperationException>(() => editor.SetAddInConnection(request));
            plugin.FailDescription = false;
            plugin.AfterWrite = () => host.AddIns.Clear();
            dynamic result = editor.SetAddInConnection(Request(editor, plugin));
            Assert.IsFalse((bool)result.Verified); Assert.IsNotNull((string)result.ReadbackError);
            host.AddIns.Add(plugin); plugin.AfterWrite = null; plugin.Connect = false;
            plugin.AfterWrite = () => plugin.Guid = System.Guid.NewGuid().ToString();
            result = editor.SetAddInConnection(Request(editor, plugin)); Assert.IsFalse((bool)result.Verified);
            plugin.AfterWrite = null; plugin.Connect = false;
            plugin.AfterWrite = () => plugin.FailDescription = true;
            result = editor.SetAddInConnection(Request(editor, plugin)); Assert.IsFalse((bool)result.Verified);
        }
    }
}

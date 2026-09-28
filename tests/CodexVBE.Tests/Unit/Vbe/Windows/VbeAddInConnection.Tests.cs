namespace CodexVBE.Tests.Unit
{
    using System;
    using System.Collections.Generic;
    using CodexVBE;
    using Microsoft.VisualStudio.TestTools.UnitTesting;
    [TestClass]
    [TestCategory("Unit")]
    public sealed class VbeAddInConnectionTests
    {
        public sealed class Host { public List<Plugin> AddIns { get; } = new List<Plugin>(); }
        public sealed class Plugin
        {
            public string ProgId { get; set; } = "Probe.AddIn";
            public string Guid { get; set; } = "{DB940A6C-2AB8-43AF-9B02-B68D77DF068A}";
            public string Description { get; set; } = "Disposable test double";
            private bool connected;
            public int Writes { get; private set; }
            public bool ThrowAfterWrite { get; set; }
            public bool IgnoreWrite { get; set; }
            public bool Connect { get => connected; set { Writes++; if (!IgnoreWrite) connected = value; if (ThrowAfterWrite) throw new InvalidOperationException("Native connection failure"); } }
        }
        private static Request Request(VbeEditorWindows service, string action = "connect")
        {
            dynamic state = service.AddIns();
            return new Request { ProgId = state.AddIns[0].Properties["ProgId"], ExpectedAddInVersion = state.AddIns[0].AddInVersion, Action = action };
        }
        [TestMethod]
        public void ConnectionRoundtripUsesFreshRevisionAndNoOpDoesNotWrite()
        {
            var host = new Host(); var plugin = new Plugin(); host.AddIns.Add(plugin); var service = new VbeEditorWindows(host);
            var request = Request(service);
            Assert.IsTrue((bool)((dynamic)service.SetAddInConnection(request)).Verified);
            Assert.ThrowsException<InvalidOperationException>(() => service.SetAddInConnection(request));
            Assert.IsFalse((bool)((dynamic)service.SetAddInConnection(Request(service))).Applied);
            Assert.AreEqual(1, plugin.Writes);
            Assert.IsTrue((bool)((dynamic)service.SetAddInConnection(Request(service, "disconnect"))).Verified);
            Assert.IsFalse(plugin.Connect);
        }
        [TestMethod]
        public void SelfProtectionUsesBothProgIdAndGuid()
        {
            var host = new Host(); var plugin = new Plugin { ProgId = "CodexVBE.AddIn" }; host.AddIns.Add(plugin); var service = new VbeEditorWindows(host);
            Assert.ThrowsException<InvalidOperationException>(() => service.SetAddInConnection(Request(service)));
            plugin.ProgId = "Alias.AddIn"; plugin.Guid = typeof(AddIn).GUID.ToString("B");
            Assert.ThrowsException<InvalidOperationException>(() => service.SetAddInConnection(Request(service)));
            Assert.AreEqual(0, plugin.Writes);
        }
        [TestMethod]
        public void DuplicateOrMissingIdentityIsRefusedBeforeWriting()
        {
            var host = new Host(); var plugin = new Plugin(); host.AddIns.Add(plugin); var service = new VbeEditorWindows(host);
            var request = Request(service); host.AddIns.Add(new Plugin());
            Assert.ThrowsException<InvalidOperationException>(() => service.SetAddInConnection(request));
            host.AddIns.Clear();
            Assert.ThrowsException<InvalidOperationException>(() => service.SetAddInConnection(request));
            Assert.AreEqual(0, plugin.Writes);
        }
        [TestMethod]
        public void NativeErrorsAndIgnoredWritesStayUnverified()
        {
            var host = new Host(); var plugin = new Plugin { IgnoreWrite = true }; host.AddIns.Add(plugin); var service = new VbeEditorWindows(host);
            dynamic ignored = service.SetAddInConnection(Request(service));
            Assert.IsFalse((bool)ignored.Verified); Assert.IsTrue((bool)ignored.VerificationPending);
            plugin.IgnoreWrite = false; plugin.ThrowAfterWrite = true;
            dynamic partial = service.SetAddInConnection(Request(service));
            Assert.IsFalse((bool)partial.Verified); Assert.IsNull((bool?)partial.Applied);
            Assert.IsTrue((bool)partial.After.Properties["Connect"]);
            Assert.AreEqual(2, plugin.Writes);
        }
    }
}

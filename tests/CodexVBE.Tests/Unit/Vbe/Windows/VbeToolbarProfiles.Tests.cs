namespace CodexVBE.Tests.Unit
{
    using System;
    using System.IO;
    using System.Collections.Generic;
    using System.Web.Script.Serialization;
    using CodexVBE;
    using Microsoft.VisualStudio.TestTools.UnitTesting;
    [TestClass, TestCategory("Unit")]
    public sealed class ToolbarProfilesTests
    {
        private static Dictionary<string, object> Data(object value)
        { var json = new JavaScriptSerializer(); return json.Deserialize<Dictionary<string, object>>(json.Serialize(value)); }
        [TestMethod]
        public void ProfilesMergeAcrossInstancesRejectCorruptionAndDeleteExactBars()
        {
            string root = Path.Combine(Path.GetTempPath(), "VBAi-toolbar-" + Guid.NewGuid().ToString("N"));
            try
            {
                var first = new VbeToolbarProfiles(Path.Combine(root, "profile.sqlite"));
                first.Update("VBAi - First", new VbeToolbarProfiles.Bar { Name = "VBAi - First", Visible = true, Position = 1, Commands = new VbeToolbarProfiles.Command[0] });
                var second = new VbeToolbarProfiles(Path.Combine(root, "profile.sqlite"));
                second.Update("VBAi - Second", new VbeToolbarProfiles.Bar { Name = "VBAi - Second", Position = 4, Commands = new VbeToolbarProfiles.Command[0] });
                Assert.AreEqual(2, first.Read().Length);
                second.Update("VBAi - First", null); Assert.AreEqual("VBAi - Second", first.Read()[0].Name);
                Assert.ThrowsException<InvalidOperationException>(() => second.Update("ThirdParty", new VbeToolbarProfiles.Bar { Name = "ThirdParty", Position = 1, Commands = new VbeToolbarProfiles.Command[0] }));
                Assert.AreEqual(1, first.Read().Length);
            }
            finally { if (Directory.Exists(root)) Directory.Delete(root, true); }
        }
        [TestMethod]
        public void RestartRestoresOnlyPersistentOwnedButtonsWithoutDuplicatingNativeControls()
        {
            string root = Path.Combine(Path.GetTempPath(), "VBAi-toolbar-" + Guid.NewGuid().ToString("N"));
            try
            {
                var store = new VbeToolbarProfiles(Path.Combine(root, "profile.sqlite"));
                var host = NativeHost(); var service = new VbeEditorWindows(host) { ToolbarProfiles = store };
                var created = Data(service.CreateToolbar(new Request { ObjectName = "Persistent", Temporary = false,
                    ExpectedToolbarCollectionVersion = (string)Data(service.Toolbars())["ToolbarCollectionVersion"] }));
                var request = new Request { ObjectName = (string)created["ObjectName"], ControlId = 42, ControlCaption = "Native command", Temporary = false };
                request.ExpectedToolbarControlsVersion = (string)Data(service.ToolbarControls(request))["ToolbarControlsVersion"];
                Assert.IsTrue((bool)Data(service.AddToolbarCommand(request))["Verified"]);
                request.Temporary = true; request.ExpectedToolbarControlsVersion = (string)Data(service.ToolbarControls(request))["ToolbarControlsVersion"];
                service.AddToolbarCommand(request); Assert.AreEqual(1, store.Read()[0].Commands.Length);
                service.RemoveTemporaryToolbarCommands(); Assert.AreEqual(1, host.CommandBars[1].Controls.Count); Assert.AreEqual(1, host.CommandBars[0].Controls.Count);
                var restarted = NativeHost(); restarted.CommandBars[0].Controls[1].Caption = "Runtime-context caption"; var after = new VbeEditorWindows(restarted) { ToolbarProfiles = new VbeToolbarProfiles(Path.Combine(root, "profile.sqlite")) };
                after.RestoreToolbarProfiles(); after.RestoreToolbarProfiles();
                Assert.AreEqual(0, after.ToolbarProfileErrors.Count, string.Join("; ", after.ToolbarProfileErrors)); Assert.AreEqual(2, restarted.CommandBars.Count);
                Assert.AreEqual(1, restarted.CommandBars[1].Controls.Count); Assert.AreEqual(1, restarted.CommandBars[0].Controls.Count);
                request.ExpectedToolbarControlsVersion = (string)Data(after.ToolbarControls(request))["ToolbarControlsVersion"]; request.InsertIndex = 1; request.ControlCaption = "Runtime-context caption";
                after.RemoveToolbarCommand(request); Assert.AreEqual(0, store.Read()[0].Commands.Length);
                request.ExpectedToolbarControlsVersion = (string)Data(after.ToolbarControls(request))["ToolbarControlsVersion"];
                request.ExpectedToolbarCollectionVersion = (string)Data(after.Toolbars())["ToolbarCollectionVersion"];
                after.RemoveToolbar(request); Assert.AreEqual(0, store.Read().Length);
            }
            finally { if (Directory.Exists(root)) Directory.Delete(root, true); }
        }
        private static ToolbarCustomizationTests.Host NativeHost()
        {
            var host = new ToolbarCustomizationTests.Host(); var bar = host.CommandBars.Add("Standard", 1, false, true); bar.BuiltIn = true;
            var button = bar.Controls.Add(1, 42, Type.Missing, 1, false); button.BuiltIn = true; return host;
        }
    }
}

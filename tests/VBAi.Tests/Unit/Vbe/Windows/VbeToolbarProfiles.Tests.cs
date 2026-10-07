namespace VBAi.Tests.Unit
{
    using Microsoft.VisualStudio.TestTools.UnitTesting;
    using System;
    using System.Collections.Generic;
    using System.IO;
    using System.Web.Script.Serialization;
    using VBAi;
    [TestClass, TestCategory("Unit")]
    public sealed partial class ToolbarProfilesTests
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
                var created = Data(service.CreateToolbar(new Request
                {
                    ObjectName = "Persistent",
                    Temporary = false,
                    ExpectedToolbarCollectionVersion = (string)Data(service.Toolbars())["ToolbarCollectionVersion"]
                }));
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
namespace VBAi.Tests.Unit
{
    using Microsoft.VisualStudio.TestTools.UnitTesting;
    using System;
    using System.Linq;
    using VBAi;

    public sealed partial class ToolbarProfilesTests
    {
        [TestMethod]
        public void ProfileCollectionRejectsMissingExcessNullAndCaseDuplicateEntries()
        {
            var invalid = new[] { (VbeToolbarProfiles.Bar[])null, Enumerable.Range(0, 33).Select(i => new VbeToolbarProfiles.Bar { Name = "VBAi - " + i }).ToArray(),
                new VbeToolbarProfiles.Bar[] { null }, new[] { ProfileBar(), new VbeToolbarProfiles.Bar { Name = "vbai - profile" } } };
            foreach (var bars in invalid)
                Assert.AreEqual("Invalid toolbar profile collection.", Assert.ThrowsException<InvalidOperationException>(() => ValidateProfile("Validate", bars)).Message);
            ValidateProfile("Validate", new VbeToolbarProfiles.Bar[0]);
            ValidateProfile("Validate", Enumerable.Range(0, 32).Select(i => new VbeToolbarProfiles.Bar { Name = "VBAi - " + i, Commands = new VbeToolbarProfiles.Command[0] }).ToArray());
        }

        [TestMethod]
        public void ProfileContentsValidateEveryBoundAndCommandIdentityBeforeUse()
        {
            var mutations = new Action<VbeToolbarProfiles.Bar>[] {
                b => b.Name = null, b => b.Name = "", b => b.Name = "vbai - Other", b => b.Name = "VBAi - " + new string('x',65), b => b.Name = "VBAi - \n",
                b => b.Position = -1, b => b.Position = 5, b => b.Left = -32769, b => b.Left = 32768, b => b.Top = -32769, b => b.Top = 32768,
                b => b.RowIndex = 0, b => b.Commands = null, b => b.Commands = Enumerable.Range(0,129).Select(i => ProfileCommand()).ToArray(),
                b => b.Commands = new VbeToolbarProfiles.Command[] { null }, b => b.Commands[0].Id = 1,
                b => b.Commands[0].Caption = null, b => b.Commands[0].Caption = "  ", b => b.Commands[0].Caption = new string('x',1025),
                b => b.Commands[0].Tag = null, b => b.Commands[0].Tag = "", b => b.Commands[0].Tag = "Other", b => b.Commands[0].Tag = PersistentTag + new string('x',97),
                b => b.Commands = new[] { ProfileCommand(), ProfileCommand() }
            };
            Assert.ThrowsException<InvalidOperationException>(() => ValidateProfile("ValidateContents", null));
            foreach (var mutate in mutations)
            {
                var bar = ProfileBar(); mutate(bar);
                Assert.AreEqual("Invalid toolbar profile contents.", Assert.ThrowsException<InvalidOperationException>(() => ValidateProfile("ValidateContents", bar)).Message);
                using (var scope = new ProfileScope())
                {
                    Assert.ThrowsException<InvalidOperationException>(() => scope.Profiles.Update(bar.Name ?? "VBAi - Missing", bar));
                    Assert.AreEqual(0, scope.Profiles.Read().Length, "Rejected state must not commit.");
                }
            }
            foreach (var edge in new[] { -32768, 32767 })
            {
                var bar = ProfileBar(); bar.Name = "VBAi - " + new string('x', 64); bar.Position = edge < 0 ? 0 : 4;
                bar.Left = edge; bar.Top = edge; bar.RowIndex = 1; bar.Commands[0].Id = 2; bar.Commands[0].Caption = new string('x', 1024);
                bar.Commands[0].Tag = PersistentTag + new string('x', 96 - PersistentTag.Length); ValidateProfile("ValidateContents", bar);
            }
            var many = ProfileBar(); many.Commands = Enumerable.Range(0, 128).Select(i => new VbeToolbarProfiles.Command { Id = i + 2, Caption = "C", Tag = PersistentTag + i }).ToArray();
            ValidateProfile("ValidateContents", many);
        }
    }
}

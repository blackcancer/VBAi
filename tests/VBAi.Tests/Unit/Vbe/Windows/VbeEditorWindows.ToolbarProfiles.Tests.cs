namespace VBAi.Tests.Unit
{
    using System;
    using System.Linq;
    using VBAi;
    using Microsoft.VisualStudio.TestTools.UnitTesting;

    public sealed partial class ToolbarProfilesTests
    {
        [TestMethod]
        public void RestorationRejectsEachNativeCollisionAndContinuesWithOtherProfiles()
        {
            var faults=new Action<ToolbarCustomizationTests.Host,ToolbarCustomizationTests.Bar>[] {
                (h,b)=>h.CommandBars.Add(b.Name,1,false,true), (h,b)=>b.BuiltIn=true, (h,b)=>b.Type=1,
                (h,b)=>b.Enabled=false, (h,b)=>b.Protection=1,
                (h,b)=>{var c=b.Controls.Add(1,42,Type.Missing,1,true);c.Tag=PersistentTag+"owned";var d=b.Controls.Add(1,42,Type.Missing,2,true);d.Tag=c.Tag;},
                (h,b)=>{var c=b.Controls.Add(1,43,Type.Missing,1,true);c.Tag=PersistentTag+"owned";},
                (h,b)=>{var c=b.Controls.Add(2,42,Type.Missing,1,true);c.Tag=PersistentTag+"owned";},
                (h,b)=>{var c=b.Controls.Add(1,42,Type.Missing,1,true);c.Tag=PersistentTag+"owned";c.BuiltIn=false;},
                (h,b)=>h.CommandBars[0].Controls.Clear(),
                (h,b)=>h.CommandBars.SourceOverride=new ToolbarCustomizationTests.Button { BuiltIn=false,Type=1 },
                (h,b)=>h.CommandBars.SourceOverride=new ToolbarCustomizationTests.Button { BuiltIn=true,Type=2 },
                (h,b)=>b.Controls.FailAdd=true, (h,b)=>b.Controls.FailTag=true,
                (h,b)=>b.FailVisibility=true
            };
            foreach(var fault in faults)
            using(var scope=new ProfileScope())
            {
                var profile=ProfileBar();scope.Profiles.Update(profile.Name,profile);
                var host=NativeHost();var bar=host.CommandBars.Add(profile.Name,1,false,true);fault(host,bar);
                var service=new VbeEditorWindows(host) { ToolbarProfiles=scope.Profiles };
                service.RestoreToolbarProfiles();Assert.AreEqual(1,service.ToolbarProfileErrors.Count);
                StringAssert.StartsWith(service.ToolbarProfileErrors[0],profile.Name+": ");
                Assert.AreEqual(1,scope.Profiles.Read().Length,"A native refusal must not mutate the persisted profile.");
            }
            using(var scope=new ProfileScope())
            {
                var profile=ProfileBar();scope.Profiles.Update(profile.Name,profile);
                var next=ProfileBar();next.Name="VBAi - Second";next.Commands=new VbeToolbarProfiles.Command[0];scope.Profiles.Update(next.Name,next);
                var host=NativeHost();host.CommandBars.Add(profile.Name,1,false,true).BuiltIn=true;
                var service=new VbeEditorWindows(host){ToolbarProfiles=scope.Profiles};service.RestoreToolbarProfiles();
                Assert.AreEqual(1,service.ToolbarProfileErrors.Count);Assert.IsTrue(host.CommandBars.Any(b=>b.Name==next.Name));
            }
        }

        [TestMethod]
        public void RestorationReportsStorageAndAddFaultsWithoutTouchingNativeCommands()
        {
            var host=NativeHost();new VbeEditorWindows(host).RestoreToolbarProfiles();Assert.AreEqual(1,host.CommandBars.Count);
            using(var scope=new ProfileScope())
            {
                var service=new VbeEditorWindows(host){ToolbarProfiles=scope.Profiles};service.RestoreToolbarProfiles();Assert.AreEqual(0,service.ToolbarProfileErrors.Count);
                using(var store=new ChatSessionStore(scope.Path))
                    store.UpdateToolbarProfile("Invalid",new VbeToolbarProfiles.Bar { Name="Invalid",Commands=new VbeToolbarProfiles.Command[0] },bars=>{});
                service.RestoreToolbarProfiles();Assert.AreEqual(1,service.ToolbarProfileErrors.Count);Assert.AreEqual(1,host.CommandBars.Count);
            }
            using(var scope=new ProfileScope())
            {
                var profile=ProfileBar();scope.Profiles.Update(profile.Name,profile);host.CommandBars.FailAdd=true;
                var service=new VbeEditorWindows(host){ToolbarProfiles=scope.Profiles};service.RestoreToolbarProfiles();
                Assert.AreEqual(1,service.ToolbarProfileErrors.Count);StringAssert.Contains(service.ToolbarProfileErrors[0],"native add rejected");
            }
        }

        [TestMethod]
        public void RestorationHandlesOptionalCoordinatesDockingAndExactExistingIdentities()
        {
            foreach(var geometry in new[] { new int?[] {4,10,20,3},new int?[] {4,null,20,null},new int?[] {4,10,null,null},new int?[] {1,null,null,3},new int?[] {1,null,null,null} })
            using(var scope=new ProfileScope())
            {
                var profile=ProfileBar();profile.Position=geometry[0].Value;profile.Left=geometry[1];profile.Top=geometry[2];profile.RowIndex=geometry[3];profile.Visible=false;
                scope.Profiles.Update(profile.Name,profile);
                var host=NativeHost();var service=new VbeEditorWindows(host){ToolbarProfiles=scope.Profiles};service.RestoreToolbarProfiles();service.RestoreToolbarProfiles();
                Assert.AreEqual(0,service.ToolbarProfileErrors.Count,string.Join(";",service.ToolbarProfileErrors));
                var bar=host.CommandBars[1];Assert.AreEqual(profile.Position,bar.Position);Assert.IsFalse(bar.Visible);Assert.AreEqual(1,bar.Controls.Count);
                Assert.IsTrue(bar.Controls[1].CopiedFromSource);Assert.AreEqual(profile.Commands[0].Tag,bar.Controls[1].Tag);
                Assert.AreEqual(profile.Position==4&&profile.Left.HasValue&&profile.Top.HasValue?profile.Left.Value:0,bar.Left);
                Assert.AreEqual(profile.Position==4&&profile.Left.HasValue&&profile.Top.HasValue?profile.Top.Value:0,bar.Top);
                Assert.AreEqual(profile.Position!=4&&profile.RowIndex.HasValue?profile.RowIndex.Value:1,bar.RowIndex);
            }
        }

        [TestMethod]
        public void TemporaryCleanupPreservesPersistentForeignMalformedAndMenuControls()
        {
            var host=NativeHost();var bar=host.CommandBars[0];bar.Controls[1].Tag=null;
            string owned="VBAi.ToolbarCommand."+new string('a',32);
            foreach(string tag in new[] { owned,owned, PersistentTag+"owned","ThirdParty", "VBAi.ToolbarCommand."+new string('A',32),"VBAi.ToolbarCommand.short" })
            { var c=bar.Controls.Add(1,42,Type.Missing,bar.Controls.Count+1,true);c.Tag=tag; }
            var menu=host.CommandBars.Add("Menu",1,true,true);menu.Type=1;menu.Controls.Add(1,42,Type.Missing,1,true).Tag=owned;
            new VbeEditorWindows(host).RemoveTemporaryToolbarCommands();
            Assert.AreEqual(5,bar.Controls.Count);Assert.IsFalse(bar.Controls.Any(c=>c.Tag==owned));Assert.AreEqual(1,menu.Controls.Count);
        }

        [TestMethod]
        public void ProfileSavingTracksGeometryVisibilityAndOnlyPersistentButtons()
        {
            using(var scope=new ProfileScope())
            {
                var host=NativeHost();var service=new VbeEditorWindows(host){ToolbarProfiles=scope.Profiles};
                var created=Data(service.CreateToolbar(new Request { ObjectName="Profile",Temporary=false,ExpectedToolbarCollectionVersion=(string)Data(service.Toolbars())["ToolbarCollectionVersion"] }));
                Assert.IsTrue((bool)created["Verified"]);var bar=host.CommandBars[1];
                var foreign=host.CommandBars.Add("VBAi - Temporary",1,false,true);SaveProfile(service,foreign,false);Assert.AreEqual(1,scope.Profiles.Read().Length);
                SaveProfile(new VbeEditorWindows(host),bar,false);
                foreach(int position in new[] {1,4})
                {
                    bar.Position=position;bar.Left=-12;bar.Top=37;bar.RowIndex=2;bar.Visible=false;
                    var persistent=bar.Controls.Add(1,42,Type.Missing,bar.Controls.Count+1,true);persistent.Tag=PersistentTag+position;
                    var temp=bar.Controls.Add(1,42,Type.Missing,bar.Controls.Count+1,true);temp.Tag="VBAi.ToolbarCommand."+new string('a',32);
                    SaveProfile(service,bar,false);var saved=scope.Profiles.Read()[0];
                    Assert.AreEqual(position,saved.Position);Assert.IsFalse(saved.Visible);Assert.AreEqual(position==4?(int?)-12:null,saved.Left);
                    Assert.AreEqual(position==4?(int?)37:null,saved.Top);Assert.AreEqual(position==4?null:(int?)2,saved.RowIndex);
                    Assert.AreEqual(position==4?2:1,saved.Commands.Length);Assert.IsTrue(saved.Commands.All(c=>c.Tag.StartsWith(PersistentTag,StringComparison.Ordinal)));
                }
                var request=ProfileRequest(service,foreign.Name);request.Temporary=false;request.ControlId=42;request.ControlCaption="Native command";
                Assert.ThrowsException<InvalidOperationException>(()=>service.AddToolbarCommand(request));Assert.AreEqual(0,foreign.Controls.Count);
            }
        }

        [TestMethod]
        public void PublicToolbarMutationsPersistNativeReadbackAndReportStorageFailure()
        {
            using(var scope=new ProfileScope())
            {
                var profile=ProfileBar();profile.Commands=new VbeToolbarProfiles.Command[0];scope.Profiles.Update(profile.Name,profile);
                var host=NativeHost();var bar=host.CommandBars.Add(profile.Name,1,false,true);var service=new VbeEditorWindows(host){ToolbarProfiles=scope.Profiles};
                Assert.IsTrue((bool)Data(service.SetToolbarVisibility(LayoutProfileRequest(service,bar.Name,"hide")))["Verified"]);
                Assert.IsFalse(scope.Profiles.Read()[0].Visible);
                Assert.IsTrue((bool)Data(service.SetToolbarPosition(LayoutProfileRequest(service,bar.Name,"float")))["Verified"]);
                var request=LayoutProfileRequest(service,bar.Name,"float");request.ToolbarLeft=17;request.ToolbarTop=31;
                Assert.IsTrue((bool)Data(service.SetToolbarPlacement(request))["Verified"]);Assert.AreEqual(17,scope.Profiles.Read()[0].Left);
                Assert.IsTrue((bool)Data(service.SetToolbarPosition(LayoutProfileRequest(service,bar.Name,"top")))["Verified"]);
                request=LayoutProfileRequest(service,bar.Name,"row");request.RowIndex=3;
                Assert.IsTrue((bool)Data(service.SetToolbarPlacement(request))["Verified"]);Assert.AreEqual(3,scope.Profiles.Read()[0].RowIndex);
                using(var store=new ChatSessionStore(scope.Path))store.UpdateToolbarProfile("Invalid",new VbeToolbarProfiles.Bar {Name="Invalid",Commands=new VbeToolbarProfiles.Command[0]},bars=>{});
                var failed=Data(service.SetToolbarVisibility(LayoutProfileRequest(service,bar.Name,"show")));
                Assert.IsTrue(bar.Visible);Assert.IsFalse((bool)failed["Verified"]);StringAssert.Contains((string)failed["NativeError"],"Invalid toolbar profile");
                var add=ProfileRequest(service,bar.Name);add.ControlId=42;add.ControlCaption="Native command";add.Temporary=true;
                service.AddToolbarCommand(add);var remove=ProfileRequest(service,bar.Name);remove.ControlId=42;remove.ControlCaption="Native command";remove.InsertIndex=1;
                var removed=Data(service.RemoveToolbarCommand(remove));Assert.IsTrue((bool)removed["Removed"]);Assert.IsFalse((bool)removed["Verified"]);
            }
        }

        [TestMethod]
        public void NativeDeletionRefusalRetainsProfileAndChangedCopyCaptionIsNotVerified()
        {
            using(var scope=new ProfileScope())
            {
                var profile=ProfileBar();profile.Commands=new VbeToolbarProfiles.Command[0];scope.Profiles.Update(profile.Name,profile);
                var host=NativeHost();var bar=host.CommandBars.Add(profile.Name,1,false,true);var service=new VbeEditorWindows(host){ToolbarProfiles=scope.Profiles};
                bar.IgnoreDelete=true;Assert.IsFalse((bool)Data(service.RemoveToolbar(ProfileRequest(service,bar.Name)))["Removed"]);Assert.AreEqual(1,scope.Profiles.Read().Length);
                bar.Controls.Copied=c=>c.Caption="Unexpected caption";
                var request=ProfileRequest(service,bar.Name);request.ControlId=42;request.ControlCaption="Native command";request.Temporary=true;
                var result=Data(service.AddToolbarCommand(request));Assert.IsTrue((bool)result["Added"]);Assert.IsFalse((bool)result["Verified"]);
                Assert.AreEqual("Unexpected caption",bar.Controls[1].Caption);
                bar.Controls.Clear();bar.IgnoreDelete=false;Assert.IsTrue((bool)Data(service.RemoveToolbar(ProfileRequest(service,bar.Name)))["Verified"]);Assert.AreEqual(0,scope.Profiles.Read().Length);
            }
        }
    }
}

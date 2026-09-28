namespace CodexVBE.Tests.Unit
{
    using System;
    using System.Collections.Generic;
    using System.Linq;
    using System.Web.Script.Serialization;
    using CodexVBE;
    using Microsoft.VisualStudio.TestTools.UnitTesting;

    [TestClass, TestCategory("Unit")]
    public sealed partial class ToolbarCustomizationTests
    {
        private static Request CommandRequest(VbeEditorWindows service,string bar="Standard")
        {return new Request{ObjectName=bar,ControlId=42,ControlCaption="Native command",ExpectedToolbarControlsVersion=(string)Data(service.ToolbarControls(new Request{ObjectName=bar}))["ToolbarControlsVersion"]};}
        [TestMethod]
        public void ToolbarCreationRequiresCurrentCollectionPrintableUniqueNameAndReportsNativePartialFailure()
        {
            var host=new Host();var service=new VbeEditorWindows(host);
            foreach(string expected in new[]{null,"stale"})Assert.ThrowsException<InvalidOperationException>(()=>service.CreateToolbar(new Request{ExpectedToolbarCollectionVersion=expected}));
            foreach(string name in new[]{null," ",new string('x',65),"bad\nname"})
                Assert.ThrowsException<ArgumentException>(()=>service.CreateToolbar(new Request{ObjectName=name,ExpectedToolbarCollectionVersion=(string)Data(service.Toolbars())["ToolbarCollectionVersion"]}));
            for(int scenario=0;scenario<4;scenario++)
            {
                var native=new Host();var client=new VbeEditorWindows(native);var request=new Request{ObjectName="Matrix",Temporary=false,ExpectedToolbarCollectionVersion=(string)Data(client.Toolbars())["ToolbarCollectionVersion"]};
                native.CommandBars.FailAdd=scenario==0;native.CommandBars.NullAdd=scenario==1;if(scenario==2)native.CommandBars.Added=bar=>bar.FailVisibility=true;
                var created=Data(client.CreateToolbar(request));Assert.AreEqual(scenario>=2,(bool)created["Created"]);Assert.AreEqual(scenario==3,(bool)created["Verified"]);Assert.IsFalse((bool)created["Temporary"]);
                if(scenario<3)Assert.IsNotNull(created["NativeError"]);
                if(scenario==3){request.ExpectedToolbarCollectionVersion=(string)created["ToolbarCollectionVersion"];Assert.ThrowsException<InvalidOperationException>(()=>client.CreateToolbar(request));}
            }
        }
        [TestMethod]
        public void ToolbarRemovalProtectsPrefixContentsProtectionAndVerifiesActualDisappearance()
        {
            for(int scenario=0;scenario<5;scenario++)
            {
                var host=new Host();var bar=host.CommandBars.Add(scenario==0?"Foreign":"VBAi - Matrix",1,false,true);var service=new VbeEditorWindows(host);
                if(scenario==1)bar.Controls.Add(1,42,Type.Missing,1,true);if(scenario==2)bar.Protection=1;if(scenario==3)bar.IgnoreDelete=true;
                var request=CommandRequest(service,bar.Name);request.ExpectedToolbarCollectionVersion=(string)Data(service.Toolbars())["ToolbarCollectionVersion"];
                if(scenario<=2)Assert.ThrowsException<InvalidOperationException>(()=>service.RemoveToolbar(request));
                else Assert.AreEqual(scenario==4,(bool)Data(service.RemoveToolbar(request))["Verified"]);
            }
        }
        [TestMethod]
        public void ToolbarCommandPreflightRejectsDisabledProtectedStaleMissingAndWrongNativeIdentities()
        {
            for(int scenario=0;scenario<10;scenario++)
            {
                var host=new Host();var bar=host.CommandBars.Add("Standard",1,false,true);var source=bar.Controls.Add(1,42,Type.Missing,1,false);var service=new VbeEditorWindows(host);var request=CommandRequest(service);
                if(scenario==0)request.ExpectedToolbarControlsVersion=null;if(scenario==1){bar.Enabled=false;request=CommandRequest(service);}if(scenario==2){bar.Protection=1;request=CommandRequest(service);}
                if(scenario==3)request.ControlId=1;if(scenario==4)request.ControlCaption=" ";if(scenario==5)request.ControlId=99;
                if(scenario==6){source.BuiltIn=false;request=CommandRequest(service);}if(scenario==7){host.CommandBars.SourceOverride=new Button{Type=2,Id=42,BuiltIn=true,Caption="Native command"};}
                if(scenario==8)request.InsertIndex=0;if(scenario==9)request.InsertIndex=3;
                Exception failure=null;try{service.AddToolbarCommand(request);}catch(Exception error){failure=error;}
                Assert.IsNotNull(failure,"Boundary "+scenario);Assert.AreEqual(1,bar.Controls.Count);
            }
        }
        [TestMethod]
        public void ToolbarCommandNativeFailuresAndReadbackNeverClaimVerifiedAddOrRemoval()
        {
            for(int scenario=0;scenario<4;scenario++)
            {
                var host=new Host();var bar=host.CommandBars.Add("Standard",1,false,true);bar.Controls.Add(1,42,Type.Missing,1,false);var service=new VbeEditorWindows(host);var request=CommandRequest(service);
                bar.Controls.FailAdd=scenario==0;bar.Controls.FailTag=scenario==1;bar.Controls.IgnoreTag=scenario==2;if(scenario==3)bar.Controls.AddedId=43;
                var result=Data(service.AddToolbarCommand(request));Assert.IsFalse((bool)result["Verified"]);Assert.AreEqual(scenario==3,(bool)result["Added"]);
                Assert.AreEqual(scenario==0?1:2,bar.Controls.Count);
            }
            for(int scenario=0;scenario<6;scenario++)
            {
                var host=new Host();var bar=host.CommandBars.Add("Standard",1,false,true);bar.Controls.Add(1,42,Type.Missing,1,false);var service=new VbeEditorWindows(host);
                service.AddToolbarCommand(CommandRequest(service));var request=CommandRequest(service);request.InsertIndex=2;
                if(scenario==0)request.InsertIndex=null;if(scenario==1)request.InsertIndex=3;if(scenario==2)request.ControlId=43;if(scenario==3)request.ControlCaption="other";
                if(scenario==4)bar.Controls[2].FailDelete=true;if(scenario==5)bar.Controls[2].IgnoreDelete=true;
                if(scenario<2)Assert.ThrowsException<ArgumentOutOfRangeException>(()=>service.RemoveToolbarCommand(request));
                else if(scenario<4)Assert.ThrowsException<InvalidOperationException>(()=>service.RemoveToolbarCommand(request));
                else {var result=Data(service.RemoveToolbarCommand(request));Assert.IsFalse((bool)result["Verified"]);Assert.IsFalse((bool)result["Removed"]);}
                Assert.AreEqual(2,bar.Controls.Count);
            }
        }
        [TestMethod]
        public void CollectionRevisionReadErrorsAreReportedWithoutLosingToolbarObservations()
        {
            var host=new VbeToolbarCoverageTests.Host();var bar=new VbeToolbarCoverageTests.NativeBar();host.CommandBars.Add(bar);bar.FailReads.Add("Name");
            var result=Data(new VbeEditorWindows(host).Toolbars());Assert.IsNull(result["ToolbarCollectionVersion"]);Assert.AreEqual(1,((System.Collections.ICollection)result["Errors"]).Count);Assert.AreEqual(1,((System.Collections.ICollection)result["Toolbars"]).Count);
        }
        private static readonly JavaScriptSerializer Json = new JavaScriptSerializer();
        private static Dictionary<string, object> Data(object value) => Json.Deserialize<Dictionary<string, object>>(Json.Serialize(value));
        [TestMethod]
        public void NativeCommandCustomizationRoundTripProtectsOriginalButtonsAndRevisions()
        {
            var host = new Host(); var bar = host.CommandBars.Add("Standard", 1, false, true); bar.BuiltIn = true;
            var original = bar.Controls.Add(1, 42, Type.Missing, 1, false);
            original.Caption = "Native command"; original.BuiltIn = true;
            original.Tag=null;
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
            Assert.IsTrue(bar.Controls[1].CopiedFromSource); Assert.IsFalse((bool)persistent["PersistenceVerified"]);
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
    }
}

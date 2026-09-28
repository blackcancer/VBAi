namespace CodexVBE.Tests.Unit
{
    using System;
    using System.Collections.Generic;
    using System.Linq;
    using CodexVBE;
    using Microsoft.VisualStudio.TestTools.UnitTesting;

    [TestClass, TestCategory("Unit")]
    public sealed class WritableOptionsTests
    {
        [TestMethod]
        public void WritableOptionsGuardEveryRequiredFieldTabCatalogueAndExactVisibleControl()
        {
            foreach(int scenario in Enumerable.Range(0,5))
            {
                var p=new WritableOptionsMatrixProbe();var r=p.Request();
                if(scenario==0)r=null;if(scenario==1)r.Pane=" ";if(scenario==2)r.Property=null;if(scenario==3)r.ExpectedOptionsVersion="";if(scenario==4)p.Open=false;
                if(scenario<4)Assert.ThrowsException<ArgumentException>(()=>VbeDebugWindows.SetVbeOption(r,p));
                else Assert.ThrowsException<InvalidOperationException>(()=>VbeDebugWindows.SetVbeOption(r,p));
                Assert.AreEqual(0,p.Writes);
            }
            foreach(int scenario in Enumerable.Range(0,9))
            {
                var p=new WritableOptionsMatrixProbe();var r=p.Request();
                if(scenario==0)p.Names.Clear();if(scenario==1)p.Names.AddRange(Enumerable.Repeat("Editor",8));if(scenario==2)p.Names[0]=" ";
                if(scenario==3)p.Items.AddRange(Enumerable.Repeat(p.Items[0],2000));
                if(scenario==4)r.Pane="Missing";
                if(scenario==5){int calls=0;p.OnTabs=()=>{if(++calls==2)p.Names.Add("Editor");};}
                if(scenario==6)r.Property="Missing";
                if(scenario==7){p.Items[0].Error="unreadable";r=p.Request();}
                if(scenario==8){p.Items.Add(p.Items[0]);r=p.Request();}
                Assert.ThrowsException<InvalidOperationException>(()=>VbeDebugWindows.SetVbeOption(r,p));Assert.AreEqual(0,p.Writes);Assert.AreEqual(1,p.Closes);
            }
        }
        [TestMethod]
        public void WritableOptionsReadbackAndAcceptFailureAlwaysRequestCancelWhenStillOpen()
        {
            foreach(int scenario in Enumerable.Range(0,7))
            {
                var p=new WritableOptionsMatrixProbe();var r=p.Request();
                if(scenario==0)p.OnWrite=()=>p.Items.Clear();
                if(scenario==1)p.OnWrite=()=>p.Items.Add(new VbeDebugWindows.OptionsControl{Name=r.Property,Type="ControlType.CheckBox",Value="Off"});
                if(scenario==2)p.OnWrite=()=>p.Items[0].Error="readback failed";
                if(scenario==3)p.OnWrite=()=>p.Items[0].Type="other";
                if(scenario==4)p.OnWrite=()=>p.Items[0].Visible=false;
                if(scenario==5)p.OnWrite=()=>p.Items[0].Enabled=false;
                if(scenario==6)p.OnAccept=()=>{throw new InvalidOperationException("OK rejected");};
                p.IgnoreWrite=scenario<6;
                Assert.ThrowsException<InvalidOperationException>(()=>VbeDebugWindows.SetVbeOption(r,p));Assert.AreEqual(1,p.Closes);Assert.IsFalse(p.Open);
            }
            var open=new WritableOptionsMatrixProbe{KeepOpen=true};var request=open.Request();
            dynamic result=VbeDebugWindows.SetVbeOption(request,open);Assert.IsFalse((bool)result.DialogClosed);Assert.AreEqual(1,open.Closes);
            foreach(string type in new[]{"ControlType.RadioButton","ControlType.Edit"})
            {
                var p=new WritableOptionsMatrixProbe();p.Names[0]=type.EndsWith("Edit")?"Editor":"General";
                p.Items[0].Type=type;p.Items[0].Name=type.EndsWith("Edit")?"Tab Width":"Break on All Errors";p.Items[0].Value=type.EndsWith("Edit")?(object)"4":false;
                var r=p.Request();r.Value=type.EndsWith("Edit")?(object)8:true;
                Assert.IsTrue((bool)((dynamic)VbeDebugWindows.SetVbeOption(r,p)).ControlValueVerified);
            }
        }
        [TestMethod]
        public void PreferenceAllowlistCoversLocalizedTabsTypesAndNullableOrInvalidValues()
        {
            Assert.ThrowsException<ArgumentException>(()=>VbeDebugWindows.SetVbeOption(null));
            var p=new WritableOptionsMatrixProbe();
            p.Items.Add(new VbeDebugWindows.OptionsControl{Name="Hidden",Type="ControlType.CheckBox",Visible=false,Enabled=true});
            p.Items.Add(new VbeDebugWindows.OptionsControl{Name="Disabled",Type="ControlType.CheckBox",Visible=true,Enabled=false});
            p.Items.Add(new VbeDebugWindows.OptionsControl{Name="Other",Type="ControlType.CheckBox",Visible=true,Enabled=true});
            p.Items.Add(new VbeDebugWindows.OptionsControl{Name=" ",Type="ControlType.Text",Visible=true,Enabled=true});
            var on=p.Request();on.Value=true;
            Assert.AreEqual("On",(string)((dynamic)VbeDebugWindows.SetVbeOption(on,p)).After);
            var check=new VbeDebugWindows.OptionsControl{Type="ControlType.CheckBox",Name="&Auto Syntax Check:"};
            foreach(string tab in new[]{"Editor","éditeur","editeur"})Assert.AreEqual(true,VbeDebugWindows.ValidateEditableOption(tab,check,true));
            check.Name="Compile on demand";foreach(string tab in new[]{"General","général"})Assert.AreEqual(false,VbeDebugWindows.ValidateEditableOption(tab,check,false));
            var radio=new VbeDebugWindows.OptionsControl{Type="ControlType.RadioButton",Name="Break on Unhandled Errors"};
            Assert.ThrowsException<ArgumentException>(()=>VbeDebugWindows.ValidateEditableOption("General",radio,"true"));
            foreach(string tab in new[]{null,"Editor Format"})Assert.ThrowsException<InvalidOperationException>(()=>VbeDebugWindows.ValidateEditableOption(tab,new VbeDebugWindows.OptionsControl(),true));
            Assert.ThrowsException<InvalidOperationException>(()=>VbeDebugWindows.ValidateEditableOption("Editor",radio,true));
            Assert.ThrowsException<InvalidOperationException>(()=>VbeDebugWindows.ValidateEditableOption("General",new VbeDebugWindows.OptionsControl{Type="ControlType.Edit",Name="Tab Width"},4));
        }
        [TestMethod]
        public void OptionsRevisionProtectsMutationAndRequestedValueIsReadBeforeCommit()
        {
            var fake = new Fake();
            dynamic read = VbeDebugWindows.ReadVbeOptions(fake); fake.Open = true;
            var request = new Request { Pane = "Editor", Property = "Auto Syntax Check", Value = false, ExpectedOptionsVersion = (string)read.OptionsVersion };
            dynamic result = VbeDebugWindows.SetVbeOption(request, fake);
            Assert.IsTrue((bool)result.ControlValueVerified); Assert.IsTrue((bool)result.DialogClosed);
            Assert.IsFalse((bool)result.PersistenceVerified); Assert.AreEqual(1, fake.Writes); Assert.AreEqual(1, fake.Accepts);
            Assert.AreEqual("On", (string)result.Before); Assert.AreEqual("Off", (string)result.After);
            fake.Open = true;
            Assert.ThrowsException<InvalidOperationException>(() => VbeDebugWindows.SetVbeOption(request, fake));
            Assert.AreEqual(1, fake.Writes); Assert.AreEqual(1, fake.Accepts); Assert.IsFalse(fake.Open);
        }
        [TestMethod]
        public void OptionsRefuseUnsupportedPreferencesAndCancelFailedReadback()
        {
            var fake = new Fake(); dynamic read = VbeDebugWindows.ReadVbeOptions(fake); fake.Open = true;
            var request = new Request { Pane="Editor",Property="Auto Syntax Check",Value=false,ExpectedOptionsVersion=(string)read.OptionsVersion };
            fake.RetainWrite = false;
            Assert.ThrowsException<InvalidOperationException>(() => VbeDebugWindows.SetVbeOption(request, fake));
            Assert.AreEqual(0, fake.Accepts); Assert.IsFalse(fake.Open);
            var check = new VbeDebugWindows.OptionsControl { Name="Auto Syntax Check",Type="ControlType.CheckBox",Value="On" };
            Assert.ThrowsException<InvalidOperationException>(()=>VbeDebugWindows.ValidateEditableOption("Editor Format",check,false));
            Assert.ThrowsException<ArgumentException>(()=>VbeDebugWindows.ValidateEditableOption("Editor",check,"false"));
            check.Name="Unknown preference";
            Assert.ThrowsException<InvalidOperationException>(()=>VbeDebugWindows.ValidateEditableOption("Editor",check,true));
            var width = new VbeDebugWindows.OptionsControl { Name="Tab Width",Type="ControlType.Edit",Value="4" };
            Assert.AreEqual("8",VbeDebugWindows.ValidateEditableOption("Editor",width,8));
            foreach(object invalid in new object[] {0,33,1.5,"4: Stop",true})
                Assert.ThrowsException<ArgumentException>(()=>VbeDebugWindows.ValidateEditableOption("Editor",width,invalid));
            var radio = new VbeDebugWindows.OptionsControl { Name="Break on All Errors",Type="ControlType.RadioButton",Value=false };
            Assert.AreEqual(true,VbeDebugWindows.ValidateEditableOption("General",radio,true));
            Assert.ThrowsException<ArgumentException>(()=>VbeDebugWindows.ValidateEditableOption("General",radio,false));
        }
        private sealed class Fake : VbeDebugWindows.IWritableOptionsProbe
        {
            public bool Open = true, RetainWrite = true;
            public int Writes, Accepts;
            private readonly VbeDebugWindows.OptionsControl choice = new VbeDebugWindows.OptionsControl { Name="Auto Syntax Check",Type="ControlType.CheckBox",Value="On" };
            public IntPtr Dialog() => Open ? new IntPtr(1) : IntPtr.Zero;
            public IList<string> Tabs(IntPtr dialog) => new[] { "Editor" };
            public IList<VbeDebugWindows.OptionsControl> Controls(IntPtr dialog,int index) => new[] { choice };
            public IList<VbeDebugWindows.OptionsChoice> ErrorChoices(IntPtr dialog) => new VbeDebugWindows.OptionsChoice[0];
            public void Close(IntPtr dialog) { Open=false; }
            public void Pause(int milliseconds) { }
            public void Write(IntPtr dialog,int index,string name,string type,object value) { Writes++; if(RetainWrite)choice.Value=(bool)value ? "On":"Off"; }
            public void Accept(IntPtr dialog) { Accepts++;Open=false; }
        }
    }
    public sealed partial class VbeDebugWindowsSystemTests
    {
        [TestMethod]
        public void NativeOptionsWriteRealUiaPatternsAndRefuseMissingPatternsOrReadOnlyValues()
        {
            var root=new AutomationNode{Name="Options",Kind=System.Windows.Automation.ControlType.Window};
            root.Add(new AutomationNode{Name="Editor",Kind=System.Windows.Automation.ControlType.TabItem}.With(System.Windows.Automation.SelectionItemPattern.Pattern));
            var check=root.Add(new AutomationNode{Name="Auto Syntax Check",Kind=System.Windows.Automation.ControlType.CheckBox}.With(System.Windows.Automation.TogglePattern.Pattern));
            var radio=root.Add(new AutomationNode{Name="Break on All Errors",Kind=System.Windows.Automation.ControlType.RadioButton}.With(System.Windows.Automation.SelectionItemPattern.Pattern));
            var edit=root.Add(new AutomationNode{Name="Tab Width",Kind=System.Windows.Automation.ControlType.Edit,Text="4"}.With(System.Windows.Automation.ValuePattern.Pattern));
            using(var host=new AutomationHost(root))
            using(var scene=new SystemScene())
            {
                var native=Native<VbeDebugWindows.IWritableOptionsProbe>("NativeOptionsProbe");native.Tabs(host.Handle);
                native.Write(host.Handle,0,check.Name,"ControlType.CheckBox",true);Assert.AreEqual(System.Windows.Automation.ToggleState.On,check.ToggleState);
                native.Write(host.Handle,0,check.Name,"ControlType.CheckBox",true);Assert.AreEqual(System.Windows.Automation.ToggleState.On,check.ToggleState);
                native.Write(host.Handle,0,check.Name,"ControlType.CheckBox",false);Assert.AreEqual(System.Windows.Automation.ToggleState.Off,check.ToggleState);
                native.Write(host.Handle,0,radio.Name,"ControlType.RadioButton",true);Assert.IsTrue(radio.Selected);
                native.Write(host.Handle,0,edit.Name,"ControlType.Edit",8);Assert.AreEqual("8",edit.Text);
                check.ToggleState=System.Windows.Automation.ToggleState.Indeterminate;Assert.ThrowsException<InvalidOperationException>(()=>native.Write(host.Handle,0,check.Name,"ControlType.CheckBox",true));check.ToggleState=System.Windows.Automation.ToggleState.Off;
                edit.ReadOnly=true;Assert.ThrowsException<InvalidOperationException>(()=>native.Write(host.Handle,0,edit.Name,"ControlType.Edit",8));edit.ReadOnly=false;
                foreach(var node in new[]{check,radio,edit})
                {
                    var pattern=node==check?System.Windows.Automation.TogglePattern.Pattern:node==radio?System.Windows.Automation.SelectionItemPattern.Pattern:System.Windows.Automation.ValuePattern.Pattern;
                    node.Patterns.Clear();Assert.ThrowsException<InvalidOperationException>(()=>native.Write(host.Handle,0,node.Name,node.Kind.ProgrammaticName,true));node.Patterns.Add(pattern.Id);
                }
                edit.Password=true;Assert.ThrowsException<InvalidOperationException>(()=>native.Write(host.Handle,0,edit.Name,"ControlType.Edit",8));edit.Password=false;
                Assert.ThrowsException<InvalidOperationException>(()=>native.Write(host.Handle,0,"absent","ControlType.Edit",8));
                check.Offscreen=true;Assert.ThrowsException<InvalidOperationException>(()=>native.Write(host.Handle,0,check.Name,"ControlType.CheckBox",true));check.Offscreen=false;
                check.Enabled=false;Assert.ThrowsException<InvalidOperationException>(()=>native.Write(host.Handle,0,check.Name,"ControlType.CheckBox",true));check.Enabled=true;
                Assert.ThrowsException<InvalidOperationException>(()=>native.Write(host.Handle,0,check.Name,"ControlType.Edit",8));
                root.Add(new AutomationNode{Name=check.Name,Kind=System.Windows.Automation.ControlType.CheckBox});Assert.ThrowsException<InvalidOperationException>(()=>native.Write(host.Handle,0,check.Name,"ControlType.CheckBox",true));root.Children.RemoveAt(root.Children.Count-1);
                var slider=root.Add(new AutomationNode{Name="Slider",Kind=System.Windows.Automation.ControlType.Slider});Assert.ThrowsException<InvalidOperationException>(()=>native.Write(host.Handle,0,slider.Name,"ControlType.Slider",8));
            }
        }
        [TestMethod]
        public void NativeOptionsAcceptChecksEveryOkButtonBoundaryBeforeSendingOwnedClick()
        {
            var saved=VbeDebugWindows.OptionsWindowEnabled;
            try
            {
                using(var scene=new SystemScene())
                {
                    var native=Native<VbeDebugWindows.IWritableOptionsProbe>("NativeOptionsProbe");var dialog=scene.Add("Options");
                    Assert.ThrowsException<InvalidOperationException>(()=>native.Accept(dialog.Handle));
                    var ok=scene.Add("OK","Edit",dialog,1);Assert.ThrowsException<InvalidOperationException>(()=>native.Accept(dialog.Handle));
                    ok.Class="Button";VbeDebugWindows.OptionsWindowEnabled=handle=>false;Assert.ThrowsException<InvalidOperationException>(()=>native.Accept(dialog.Handle));
                    VbeDebugWindows.OptionsWindowEnabled=handle=>true;scene.PostSucceeds=false;Assert.ThrowsException<InvalidOperationException>(()=>native.Accept(dialog.Handle));
                    scene.PostSucceeds=true;int clicks=0;scene.OnMessage=(window,message)=>{if(window==ok&&message==0xF5)clicks++;};native.Accept(dialog.Handle);Assert.AreEqual(1,clicks);
                }
            }
            finally{VbeDebugWindows.OptionsWindowEnabled=saved;}
        }
    }
}

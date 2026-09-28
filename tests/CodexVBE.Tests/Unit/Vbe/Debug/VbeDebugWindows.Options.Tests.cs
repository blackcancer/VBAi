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
        public void DocumentedOptionsCompleteTheWholeVersionWriteReadbackCommitWorkflow()
        {
            foreach (var row in OptionsCompletionMatrix.Supported())
            {
                var probe = new WritableOptionsMatrixProbe();
                probe.Names[0] = row.Item1;
                var control = probe.Items[0];
                control.Name = row.Item2; control.Type = row.Item3;
                control.Value = row.Item3 == "ControlType.CheckBox" ? "On" : "Original";
                control.Choices = new[] { "Original", "Native choice" };
                var request = probe.Request(); request.Value = row.Item4;
                dynamic result = VbeDebugWindows.SetVbeOption(request, probe);
                Assert.IsTrue((bool)result.ControlValueVerified, row.Item2);
                Assert.IsTrue((bool)result.DialogClosed, row.Item2);
                Assert.IsFalse((bool)result.PersistenceVerified, row.Item2);
                Assert.AreEqual(1, probe.Writes); Assert.AreEqual(1, probe.Accepts);
                Assert.AreEqual(0, probe.Closes);
            }
        }

        [TestMethod]
        public void GridDimensionsRejectInvalidValuesAndCancelBeforeWriting()
        {
            foreach (string name in new[] { "Width", "Height", "Largeur", "Hauteur" })
                foreach (object value in OptionsCompletionMatrix.InvalidGridValues)
                {
                    var probe = new WritableOptionsMatrixProbe(); probe.Names[0] = "General";
                    probe.Items[0].Name = name; probe.Items[0].Type = "ControlType.Edit"; probe.Items[0].Value = "6";
                    var request = probe.Request(); request.Value = value;
                    Assert.ThrowsException<ArgumentException>(() => VbeDebugWindows.SetVbeOption(request, probe));
                    Assert.AreEqual(0, probe.Writes); Assert.AreEqual(0, probe.Accepts); Assert.AreEqual(1, probe.Closes);
                }
        }

        [TestMethod]
        public void FormatChoicesRejectUnknownAmbiguousUnreadableOrArbitraryValues()
        {
            foreach (string scenario in OptionsCompletionMatrix.InvalidChoiceScenarios)
            {
                var probe = new WritableOptionsMatrixProbe(); probe.Names[0] = "Editor Format";
                var control = probe.Items[0]; control.Name = "Font"; control.Type = "ControlType.ComboBox";
                control.Value = "Consolas"; control.Choices = new[] { "Consolas", "Courier New" };
                object value = "Courier New";
                if (scenario == "absent") value = "invented font";
                if (scenario == "duplicate") control.Choices = new[] { "Courier New", "Courier New" };
                if (scenario == "null catalogue") control.Choices = null;
                if (scenario == "empty catalogue") control.Choices = new string[0];
                if (scenario == "null value") value = null;
                if (scenario == "numeric value") value = 12;
                if (scenario == "wrong tab") probe.Names[0] = "General";
                if (scenario == "unknown name") control.Name = "Unknown";
                if (scenario == "wrong type") control.Type = "ControlType.Edit";
                if (scenario == "unreadable") control.Error = "Selection unavailable";
                var request = probe.Request(); request.Value = value;
                if (scenario == "null value" || scenario == "numeric value")
                    Assert.ThrowsException<ArgumentException>(() => VbeDebugWindows.SetVbeOption(request, probe));
                else Assert.ThrowsException<InvalidOperationException>(() => VbeDebugWindows.SetVbeOption(request, probe));
                Assert.AreEqual(0, probe.Writes); Assert.AreEqual(0, probe.Accepts); Assert.AreEqual(1, probe.Closes);
            }
        }

        [TestMethod]
        public void ChoiceCatalogueSelectionAndColorCategoryAreAllVersioned()
        {
            foreach (string scenario in OptionsCompletionMatrix.RevisionScenarios)
            {
                var probe = new WritableOptionsMatrixProbe(); probe.Names[0] = "Editor Format";
                var color = probe.Items[0]; color.Name = "Foreground"; color.Type = "ControlType.ComboBox";
                color.Value = "Black"; color.Choices = new[] { "Black", "Red" };
                var category = new VbeDebugWindows.OptionsControl { Name = "Code Colors", Type = "ControlType.List",
                    Value = "Normal Text", Choices = new[] { "Normal Text", "Comment Text" } };
                probe.Items.Add(category);
                var request = probe.Request(); request.Value = "Red";
                if (scenario == "choices") color.Choices = new[] { "Black", "Blue" };
                if (scenario == "selection") color.Value = "Red";
                if (scenario == "category") category.Value = "Comment Text";
                Assert.ThrowsException<InvalidOperationException>(() => VbeDebugWindows.SetVbeOption(request, probe));
                Assert.AreEqual(0, probe.Writes); Assert.AreEqual(0, probe.Accepts); Assert.AreEqual(1, probe.Closes);
            }
        }

        [TestMethod]
        public void FormatChoiceReadbackFailureCancelsWithoutAccepting()
        {
            var probe = new WritableOptionsMatrixProbe { IgnoreWrite = true }; probe.Names[0] = "Editor Format";
            probe.Items[0].Name = "Size"; probe.Items[0].Type = "ControlType.ComboBox";
            probe.Items[0].Value = "10"; probe.Items[0].Choices = new[] { "10", "12" };
            var request = probe.Request(); request.Value = "12";
            Assert.ThrowsException<InvalidOperationException>(() => VbeDebugWindows.SetVbeOption(request, probe));
            Assert.AreEqual(1, probe.Writes); Assert.AreEqual(0, probe.Accepts); Assert.AreEqual(1, probe.Closes);
        }

        [TestMethod]
        public void SameNamedLabelsAreIgnoredAndOnlyEditableControlsCanBeWritten()
        {
            foreach (string kind in new[] { "ControlType.Text", "ControlType.Button", null })
            {
                var probe = new WritableOptionsMatrixProbe();
                var option = probe.Items[0];
                probe.Items.Add(new VbeDebugWindows.OptionsControl { Name = option.Name, Type = kind,
                    Visible = true, Enabled = true, Value = option.Value });
                var request = probe.Request();
                Assert.IsTrue((bool)((dynamic)VbeDebugWindows.SetVbeOption(request, probe)).ControlValueVerified);
                Assert.AreEqual(1, probe.Writes);

                var labelOnly = new WritableOptionsMatrixProbe();
                labelOnly.Items[0].Type = kind;
                var absent = labelOnly.Request();
                Assert.ThrowsException<InvalidOperationException>(() => VbeDebugWindows.SetVbeOption(absent, labelOnly));
                Assert.AreEqual(0, labelOnly.Writes);
                Assert.AreEqual(1, labelOnly.Closes);
            }
        }

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
        [TestMethod]
        public void FrenchNativeLabelsAndTabWidthLabelDoNotPreventEdits()
        {
            foreach (string name in new[] { "Complément automatique des instructions", "Info express automatique" })
                Assert.AreEqual(false, VbeDebugWindows.ValidateEditableOption("Éditeur",
                    new VbeDebugWindows.OptionsControl { Name = name, Type = "ControlType.CheckBox" }, false));
            Assert.AreEqual(false, VbeDebugWindows.ValidateEditableOption("Général",
                new VbeDebugWindows.OptionsControl { Name = "Compilation sur demande", Type = "ControlType.CheckBox" }, false));
            var fake = new Fake { Width = true };
            dynamic before = VbeDebugWindows.ReadVbeOptions(fake); fake.Open = true;
            dynamic after = VbeDebugWindows.SetVbeOption(new Request { Pane = "Editor", Property = "Largeur de la tabulation :",
                Value = 32, ExpectedOptionsVersion = (string)before.OptionsVersion }, fake);
            Assert.AreEqual("32", (string)after.After); Assert.AreEqual(1, fake.Accepts);
        }
        [TestMethod]
        public void DuplicateDebugObservationsKeepDistinctPathsAndValues()
        {
            string first = VbeDebugWindows.DebugRowIdentity("Expression child Value 3 Type Long", new[] { "values", "child" });
            Assert.AreEqual(first, VbeDebugWindows.DebugRowIdentity("Expression child Value 3 Type Long", new[] { "values", "child" }));
            Assert.AreNotEqual(first, VbeDebugWindows.DebugRowIdentity("Expression child Value 6 Type Long", new[] { "values", "child" }));
            Assert.AreNotEqual(first, VbeDebugWindows.DebugRowIdentity("Expression child Value 3 Type Long", new[] { "other", "child" }));
            Assert.AreNotEqual(VbeDebugWindows.DebugRowIdentity("row", new[] { "a/b", "c" }), VbeDebugWindows.DebugRowIdentity("row", new[] { "a", "b/c" }));
        }
        private sealed class Fake : VbeDebugWindows.IWritableOptionsProbe
        {
            public bool Open = true, RetainWrite = true, Width;
            public int Writes, Accepts;
            private readonly VbeDebugWindows.OptionsControl choice = new VbeDebugWindows.OptionsControl { Name="Auto Syntax Check",Type="ControlType.CheckBox",Value="On" };
            public IntPtr Dialog() => Open ? new IntPtr(1) : IntPtr.Zero;
            public IList<string> Tabs(IntPtr dialog) => new[] { "Editor" };
            public IList<VbeDebugWindows.OptionsControl> Controls(IntPtr dialog,int index) {
                if (!Width) return new[] { choice };
                choice.Name = "Largeur de la tabulation :"; choice.Type = "ControlType.Edit";
                if (choice.Value is string text && (text == "On" || text == "Off")) choice.Value = "4";
                return new[] { new VbeDebugWindows.OptionsControl { Name = choice.Name, Type = "ControlType.Text" }, choice };
            }
            public IList<VbeDebugWindows.OptionsChoice> ErrorChoices(IntPtr dialog) => new VbeDebugWindows.OptionsChoice[0];
            public void Close(IntPtr dialog) { Open=false; }
            public void Pause(int milliseconds) { }
            public void Write(IntPtr dialog,int index,string name,string type,object value) { Writes++; if(RetainWrite)choice.Value=Width ? value : (object)((bool)value ? "On":"Off"); }
            public void Accept(IntPtr dialog) { Accepts++;Open=false; }
        }
    }
    public sealed partial class VbeDebugWindowsSystemTests
    {
        [TestMethod]
        public void NativeOptionsEnumerateAndSelectExactChoicesWithoutTypingOrInventingColors()
        {
            var root = new AutomationNode { Name = "Options", Kind = System.Windows.Automation.ControlType.Window };
            root.Add(new AutomationNode { Name = "Editor Format", Kind = System.Windows.Automation.ControlType.TabItem }
                .With(System.Windows.Automation.SelectionItemPattern.Pattern));
            var combo = root.Add(new AutomationNode { Name = "Font", Kind = System.Windows.Automation.ControlType.ComboBox }
                .With(System.Windows.Automation.SelectionPattern.Pattern, System.Windows.Automation.ValuePattern.Pattern));
            var original = combo.Add(new AutomationNode { Name = "Consolas", Selected = true }.With(System.Windows.Automation.SelectionItemPattern.Pattern));
            var desired = combo.Add(new AutomationNode { Name = "Courier New", Offscreen = true }.With(System.Windows.Automation.SelectionItemPattern.Pattern));
            desired.SelectedAction = () => original.Selected = false;
            using (var host = new AutomationHost(root))
            using (var scene = new SystemScene())
            {
                var native = Native<VbeDebugWindows.IWritableOptionsProbe>("NativeOptionsProbe"); native.Tabs(host.Handle);
                var observed = native.Controls(host.Handle, 0).Single(x => x.Name == "Font");
                Assert.AreEqual("Consolas", observed.Value);
                CollectionAssert.AreEqual(new[] { "Consolas", "Courier New" }, observed.Choices.ToArray());
                native.Write(host.Handle, 0, "Font", "ControlType.ComboBox", "Courier New");
                Assert.AreEqual("Courier New", native.Controls(host.Handle, 0).Single(x => x.Name == "Font").Value);
                Assert.AreEqual("", combo.Text); Assert.AreEqual(0, combo.FocusCount);
                Assert.ThrowsException<InvalidOperationException>(() => native.Write(host.Handle, 0, "Font", "ControlType.ComboBox", "invented"));
                desired.Enabled = false;
                Assert.ThrowsException<InvalidOperationException>(() => native.Write(host.Handle, 0, "Font", "ControlType.ComboBox", "Courier New")); desired.Enabled = true;
                desired.Patterns.Clear();
                Assert.ThrowsException<InvalidOperationException>(() => native.Write(host.Handle, 0, "Font", "ControlType.ComboBox", "Courier New"));
                desired.Patterns.Add(System.Windows.Automation.SelectionItemPattern.Pattern.Id);
                var duplicate = combo.Add(new AutomationNode { Name = "Courier New" }.With(System.Windows.Automation.SelectionItemPattern.Pattern));
                Assert.ThrowsException<InvalidOperationException>(() => native.Write(host.Handle, 0, "Font", "ControlType.ComboBox", "Courier New"));
                combo.Children.Remove(duplicate);
                combo.Password = true;
                Assert.ThrowsException<InvalidOperationException>(() => native.Write(host.Handle, 0, "Font", "ControlType.ComboBox", "Courier New"));
                Assert.IsFalse(string.IsNullOrEmpty(native.Controls(host.Handle, 0).Single(x => x.Name == "Font").Error)); combo.Password = false;
                original.Selected = true;
                Assert.IsFalse(string.IsNullOrEmpty(native.Controls(host.Handle, 0).Single(x => x.Name == "Font").Error)); original.Selected = false;
                desired.Selected = false;
                Assert.IsFalse(string.IsNullOrEmpty(native.Controls(host.Handle, 0).Single(x => x.Name == "Font").Error)); desired.Selected = true;
                combo.Patterns.Remove(System.Windows.Automation.SelectionPattern.Pattern.Id); combo.Text = "Courier New";
                Assert.AreEqual("Courier New", native.Controls(host.Handle, 0).Single(x => x.Name == "Font").Value);
                combo.Patterns.Clear();
                Assert.IsFalse(string.IsNullOrEmpty(native.Controls(host.Handle, 0).Single(x => x.Name == "Font").Error));
                combo.Kind = System.Windows.Automation.ControlType.List;
                combo.Patterns.Add(System.Windows.Automation.SelectionPattern.Pattern.Id);
                native.Write(host.Handle, 0, "Font", "ControlType.List", "Courier New");
                Assert.AreEqual("Courier New", native.Controls(host.Handle, 0).Single(x => x.Name == "Font").Value);
            }
        }

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

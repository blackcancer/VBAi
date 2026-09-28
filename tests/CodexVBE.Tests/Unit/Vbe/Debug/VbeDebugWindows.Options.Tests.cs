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
}

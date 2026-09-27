using System;
using System.Collections.Generic;
using CodexVBE;
using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace CodexVBE.Tests.Unit
{
    [TestClass]
    [TestCategory("Unit")]
    public sealed class VbeDebugWindowsOptionsProbeTests
    {
        [TestMethod]
        public void OptionsRequireDialogAndBoundedReadableTabs()
        {
            var fake = new OptionsFake { Open = false };
            Assert.ThrowsException<InvalidOperationException>(() => VbeDebugWindows.ReadVbeOptions(fake));
            Assert.AreEqual(60, fake.Pauses);
            fake.Open = true;
            fake.Names.Clear();
            Assert.ThrowsException<InvalidOperationException>(() => VbeDebugWindows.ReadVbeOptions(fake));
            Assert.AreEqual(1, fake.Closes);
            fake = new OptionsFake();
            fake.Names.Add(null);
            Assert.ThrowsException<InvalidOperationException>(() => VbeDebugWindows.ReadVbeOptions(fake));
            fake = new OptionsFake();
            for (int index = 0; index < 8; index++) fake.Names.Add("Extra " + index);
            Assert.ThrowsException<InvalidOperationException>(() => VbeDebugWindows.ReadVbeOptions(fake));
        }

        [TestMethod]
        public void OptionsFiltersHiddenDisabledAndBlankTextWhileKeepingOtherValues()
        {
            var fake = new OptionsFake { CloseAfterRead = true };
            fake.Items.Add(new VbeDebugWindows.OptionsControl { Name = "Visible", Type = "ControlType.CheckBox",
                Value = "On" });
            fake.Items.Add(new VbeDebugWindows.OptionsControl { Name = "Hidden", Type = "ControlType.Edit",
                Visible = false });
            fake.Items.Add(new VbeDebugWindows.OptionsControl { Name = "Disabled", Type = "ControlType.Edit",
                Enabled = false });
            fake.Items.Add(new VbeDebugWindows.OptionsControl { Name = "", Type = "ControlType.Text" });
            fake.Items.Add(new VbeDebugWindows.OptionsControl { Name = "", Type = "ControlType.Edit",
                Value = "editable" });
            dynamic result = VbeDebugWindows.ReadVbeOptions(fake);
            Assert.AreEqual(1, (int)result.Count);
            Assert.AreEqual(2, (int)result.Tabs[0].Count);
            Assert.AreEqual("Visible", (string)result.Tabs[0].Controls[0].Name);
            Assert.AreEqual("editable", (string)result.Tabs[0].Controls[1].Value);
            Assert.IsTrue((bool)result.DialogClosed);
        }

        [TestMethod]
        public void OptionsRejectControlExplosionAndUnclosedDialog()
        {
            var fake = new OptionsFake();
            for (int index = 0; index < 2001; index++)
                fake.Items.Add(new VbeDebugWindows.OptionsControl { Name = "Item", Type = "ControlType.Text" });
            Assert.ThrowsException<InvalidOperationException>(() => VbeDebugWindows.ReadVbeOptions(fake));
            Assert.AreEqual(1, fake.Closes);
            fake = new OptionsFake();
            Assert.ThrowsException<InvalidOperationException>(() => VbeDebugWindows.ReadVbeOptions(fake));
            Assert.AreEqual(20, fake.ClosePolls);
        }

        private sealed class OptionsFake : VbeDebugWindows.IOptionsProbe
        {
            public readonly List<string> Names = new List<string> { "Editor" };
            public readonly List<VbeDebugWindows.OptionsControl> Items =
                new List<VbeDebugWindows.OptionsControl>();
            public bool Open = true, CloseAfterRead;
            public int Closes, Pauses, ClosePolls;
            public IntPtr Dialog()
            {
                if (Closes > 0) ClosePolls++;
                return Open && !(Closes > 0 && CloseAfterRead) ? new IntPtr(2) : IntPtr.Zero;
            }
            public IList<string> Tabs(IntPtr dialog) { return Names; }
            public IList<VbeDebugWindows.OptionsControl> Controls(IntPtr dialog, int tabIndex) { return Items; }
            public void Close(IntPtr dialog) { Closes++; }
            public void Pause(int milliseconds) { Pauses++; }
        }
    }
}

using System;
using CodexVBE;
using Microsoft.VisualStudio.TestTools.UnitTesting;
namespace CodexVBE.Tests.Unit
{
    [TestClass, TestCategory("Unit")]
    public sealed class VbeFormRunCommandTests
    {
        public sealed class Command { public int Id = 186; public bool Enabled = true; public string Caption; }
        public sealed class Bars
        {
            public Command Command = new Command();
            public object FindControl(int type, int id) { Assert.AreEqual(1, type); Assert.AreEqual(186, id); return Command; }
        }
        public sealed class Host { public Bars CommandBars = new Bars(); }
        [DataTestMethod]
        [DataRow("&Ausführen")]
        [DataRow("&Ejecutar")]
        [DataRow("実行")]
        public void ExactNativeIdentityAcceptsOtherLanguages(string caption)
        {
            var host = new Host(); host.CommandBars.Command.Caption = caption;
            Assert.AreSame(host.CommandBars.Command, new VbeForms(host).RequireFormRunCommand(caption));
        }
        [TestMethod]
        public void ChangedCaptionDisabledOrWrongNativeCommandIsRefused()
        {
            var host = new Host(); host.CommandBars.Command.Caption = "Run";
            var forms = new VbeForms(host);
            Assert.ThrowsException<InvalidOperationException>(() => forms.RequireFormRunCommand("Continue"));
            host.CommandBars.Command.Enabled = false;
            Assert.ThrowsException<InvalidOperationException>(() => forms.RequireFormRunCommand("Run"));
            host.CommandBars.Command.Enabled = true; host.CommandBars.Command.Id = 999;
            Assert.ThrowsException<InvalidOperationException>(() => forms.RequireFormRunCommand("Run"));
        }
    }
}

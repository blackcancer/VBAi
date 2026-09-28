using System.Web.Script.Serialization;
using CodexVBE;
using Microsoft.VisualStudio.TestTools.UnitTesting;
namespace CodexVBE.Tests.Unit
{
    [TestClass, TestCategory("Unit")]
    public sealed class FormCutChangeTests
    {
        [TestMethod]
        public void HistoryAndForkCannotRetainLiveRecoveryAuthority()
        {
            var tools = new LlmVbeTools(null, null, new LlmSettings());
            var change = new FormCutChange { Project="P", Form="F", RecoveryId="private-live-token", Owner=tools, ControlCount=2 };
            var serializer = new JavaScriptSerializer();
            string text = serializer.Serialize(new ChatEntry { FormCut=change });
            Assert.IsFalse(text.Contains("private-live-token"));
            var loaded = serializer.Deserialize<ChatEntry>(text).FormCut;
            Assert.IsNull(loaded.Owner); Assert.IsNull(loaded.RecoveryId); Assert.AreEqual(2, loaded.ControlCount);
            Assert.IsFalse(tools.CanRecoverFormCut(loaded));
            Assert.IsFalse(tools.RecoverFormCut(loaded).Ok);
        }
    }
}

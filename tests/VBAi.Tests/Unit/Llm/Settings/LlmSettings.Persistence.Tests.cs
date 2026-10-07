namespace VBAi.Tests.Unit
{
    using Microsoft.VisualStudio.TestTools.UnitTesting;
    using System.IO;
    using VBAi.Tests.Infrastructure;

    [TestClass, TestCategory("Unit")]
    public sealed class LlmSettingsDraftTests
    {
        [TestMethod]
        public void SavedDraftMergesExternalFieldsAndPublishesTheNewPersistenceBaseline()
        {
            using (var scope = new LlmBoundaryScope())
            {
                string previousPath = LlmSettings.StoragePathOverride;
                try
                {
                    LlmSettings.StoragePathOverride = Path.Combine(scope.Root, "draft-settings.json");
                    new LlmSettings { VbeEditApproval = "ReadOnly", CodexModel = "first" }.Save();
                    var shared = LlmSettings.Load();
                    var external = LlmSettings.Load();
                    var draft = shared.CreateDraft();
                    draft.VbeEditApproval = "Automatic";
                    draft.ManualModelLists["fixture"] = "draft";
                    Assert.IsFalse(shared.ManualModelLists.ContainsKey("fixture"));
                    external.CodexModel = "other-host";
                    external.Save();
                    draft.Save();
                    Assert.AreEqual("ReadOnly", shared.VbeEditApproval);
                    Assert.AreEqual("first", shared.CodexModel);
                    Assert.AreEqual("other-host", draft.CodexModel);
                    shared.PublishDraft(draft);
                    Assert.AreEqual("Automatic", shared.VbeEditApproval);
                    Assert.AreEqual("other-host", shared.CodexModel);
                    Assert.AreNotSame(draft.ManualModelLists, shared.ManualModelLists);
                    shared.VbeEditApproval = "AskEachTime";
                    shared.Save();
                    var persisted = LlmSettings.Load();
                    Assert.AreEqual("AskEachTime", persisted.VbeEditApproval);
                    Assert.AreEqual("other-host", persisted.CodexModel);
                    Assert.AreEqual("draft", persisted.ManualModelLists["fixture"]);
                }
                finally { LlmSettings.StoragePathOverride = previousPath; }
            }
        }
    }
}

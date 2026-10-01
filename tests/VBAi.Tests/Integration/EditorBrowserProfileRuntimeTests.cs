using System.IO;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using VBAi.Tests.Infrastructure;

namespace VBAi.Tests.Integration
{
    [TestClass, TestCategory("MonacoRuntime")]
    public sealed class EditorBrowserProfileRuntimeTests : EditorUiTestFixture
    {
        [STATestMethod]
        public void ClosingOneRealEditorDeletesOnlyItsExitedBrowserProfile()
        {
            using (var fixture = new EditorFixture())
            using (var first = new VBAi.ModernEditorWindow { Drafts = new VBAi.EditorDraftStore(fixture.Root) })
            using (var second = new VBAi.ModernEditorWindow { Drafts = new VBAi.EditorDraftStore(fixture.Root) })
            {
                MonacoRuntimeTests.Wait(first.OpenModule(fixture)); first.Show();
                MonacoRuntimeTests.Wait(() => first.Ready);
                MonacoRuntimeTests.Wait(second.OpenModule(fixture)); second.Show();
                MonacoRuntimeTests.Wait(() => second.Ready);
                var firstProfile = UiInvoke.Field<VBAi.EditorBrowserProfile>(first, "browserProfile");
                var secondProfile = UiInvoke.Field<VBAi.EditorBrowserProfile>(second, "browserProfile");
                Assert.AreNotEqual(firstProfile.Path, secondProfile.Path);
                Assert.IsTrue(Directory.Exists(firstProfile.Path)); Assert.IsTrue(Directory.Exists(secondProfile.Path));
                first.Close();
                MonacoRuntimeTests.Wait(() => first.IsDisposed && !Directory.Exists(firstProfile.Path));
                MonacoRuntimeTests.Wait(firstProfile.Cleanup);
                Assert.IsTrue(second.Ready); Assert.IsTrue(Directory.Exists(secondProfile.Path));
                StringAssert.Contains(MonacoRuntimeTests.Wait(second.Script("testInfo")), "pendingCommands");
                second.Close();
                MonacoRuntimeTests.Wait(() => second.IsDisposed && !Directory.Exists(secondProfile.Path));
                MonacoRuntimeTests.Wait(secondProfile.Cleanup);
            }
        }
    }
}

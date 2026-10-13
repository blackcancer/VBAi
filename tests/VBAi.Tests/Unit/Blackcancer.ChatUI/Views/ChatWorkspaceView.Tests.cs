using Blackcancer.ChatUI.Views;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using System;
using System.Globalization;
using System.Reflection;
using System.Windows.Forms;

namespace VBAi.Tests.Unit
{
    /// <summary>Exercises the shared workspace's actual localized resources and per-instance culture transitions.</summary>
    [TestClass, TestCategory("Unit")]
    public sealed class SharedChatWorkspaceResourceTests
    {
        [TestMethod]
        public void LegacyResourceReaderMapsToTheDeployedReader()
        {
            var resolve = typeof(ChatWorkspaceView).GetMethod("ResolveResourceReader", BindingFlags.Static | BindingFlags.NonPublic);
            var result = (Assembly)resolve.Invoke(null, new object[] { null, new ResolveEventArgs(
                "System.Resources.Extensions, Version=4.0.0.0, Culture=neutral, PublicKeyToken=cc7b13ffcd2ddd51") });
            Assert.AreEqual("System.Resources.Extensions", result.GetName().Name);
            Assert.IsNotNull(result.GetType("System.Resources.Extensions.DeserializingResourceReader"));
        }

        [TestMethod]
        [DataRow("Unrelated.Assembly")]
        [DataRow("System.Resources.Extensions, Version=8.0.0.0, Culture=neutral, PublicKeyToken=cc7b13ffcd2ddd51")]
        [DataRow("System.Resources.Extensions, Version=4.0.0.0, Culture=fr-FR, PublicKeyToken=cc7b13ffcd2ddd51")]
        [DataRow("System.Resources.Extensions, Version=4.0.0.0, Culture=neutral, PublicKeyToken=0000000000000000")]
        public void ResourceReaderResolutionLeavesOtherIdentitiesToTheHost(string identity)
        {
            var resolve = typeof(ChatWorkspaceView).GetMethod("ResolveResourceReader", BindingFlags.Static | BindingFlags.NonPublic);
            Assert.IsNull(resolve.Invoke(null, new object[] { null, new ResolveEventArgs(identity) }));
        }

        [STATestMethod]
        [DataRow("en-US"), DataRow("fr-FR"), DataRow("de-DE"), DataRow("es-ES")]
        [DataRow("it-IT"), DataRow("pt-BR"), DataRow("ru-RU"), DataRow("ar-SA")]
        [DataRow("hi-IN"), DataRow("ja-JP"), DataRow("ko-KR"), DataRow("zh-CN"), DataRow("zh-TW")]
        public void ArabicResourcesLoadAndDirectionTracksEachSubsequentCulture(string cultureName)
        {
            var original = CultureInfo.CurrentUICulture;
            using (var workspace = new ChatWorkspaceView(CultureInfo.GetCultureInfo("ar-SA")))
            {
                Assert.AreEqual(RightToLeft.Yes, workspace.RightToLeft);
                var target = CultureInfo.GetCultureInfo(cultureName);
                workspace.ApplyCulture(target);
                Assert.AreEqual(cultureName, workspace.UiCulture.Name);
                Assert.AreEqual(target.TextInfo.IsRightToLeft ? RightToLeft.Yes : RightToLeft.No, workspace.RightToLeft);
                Assert.IsFalse(string.IsNullOrWhiteSpace(workspace.modelSummary.AccessibleName));
                Assert.AreSame(original, CultureInfo.CurrentUICulture, "A per-instance UI change must not change the host thread culture.");
            }
        }
    }
}

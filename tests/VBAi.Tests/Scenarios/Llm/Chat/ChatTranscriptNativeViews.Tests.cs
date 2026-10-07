using Microsoft.VisualStudio.TestTools.UnitTesting;
using System.Windows.Controls;
using VBAi.Tests.Infrastructure;
namespace VBAi.Tests.Unit
{
    [TestClass, TestCategory("Unit")]
    public sealed class ChatTranscriptReleaseTests
    {
        /// <summary>Vérifie la libération d'un diff recyclé et les contenus sans hôte natif.</summary>
        [STATestMethod]
        public void ReleasingAnEntryDisposesItsNestedDiffAndHandlesNonVisualContent()
        {
            using (var window = new ChatWindow())
            {
                var entry = new ChatEntry { Speaker = "Agent" };
                using (var diff = new ChatDiffView("before", "after"))
                {
                    var child = diff.Child; var body = new StackPanel(); body.Children.Add(new TextBlock { Text = "header" }); body.Children.Add(diff);
                    var item = new TranscriptItem { RenderedContext = entry, Content = body };
                    UiInvoke.Call(typeof(ChatWindow), "ReleaseEntry", window, item);
                    Assert.IsTrue(child.IsDisposed);
                }
                foreach (object content in new object[] { null, "text", new ContentControl { Content = "literal" } })
                    UiInvoke.Call(typeof(ChatWindow), "ReleaseEntry", window, new TranscriptItem { RenderedContext = entry, Content = content });
            }
        }

    }
}

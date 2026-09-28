using System;
using System.Collections.Generic;
using System.Reflection;
using System.Windows;
using System.Windows.Controls;
using CodexVBE;
using CodexVBE.Tests.Infrastructure;
using Microsoft.VisualStudio.TestTools.UnitTesting;
namespace CodexVBE.Tests.Unit
{
    [TestClass, TestCategory("Unit")]
    public sealed class ChatTranscriptNativeViewsTests
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

        [STATestMethod]
        public void ReplacingTranscriptWindowDisposesItsHostedDesignerDiff()
        {
            using (var window = new ChatWindow())
            {
                var diff = new ChatDiffView("old", "new");
                var child = diff.Child;
                var body = new StackPanel(); body.Children.Add(diff);
                var views = (Dictionary<ChatEntry, FrameworkElement>)typeof(ChatWindow).GetField("entryViews", BindingFlags.Instance | BindingFlags.NonPublic).GetValue(window);
                views.Add(new ChatEntry { Speaker = "Agent" }, body);
                typeof(ChatWindow).GetMethod("RefreshTranscriptWindow", BindingFlags.Instance | BindingFlags.NonPublic).Invoke(window, new object[] { 0 });
                Assert.IsTrue(child.IsDisposed);
                Assert.AreEqual(0, views.Count);
            }
        }
    }
}

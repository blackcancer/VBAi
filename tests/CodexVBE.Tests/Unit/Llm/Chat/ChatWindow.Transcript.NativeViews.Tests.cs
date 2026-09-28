using System;
using System.Collections.Generic;
using System.Reflection;
using System.Windows;
using System.Windows.Controls;
using CodexVBE;
using Microsoft.VisualStudio.TestTools.UnitTesting;
namespace CodexVBE.Tests.Unit
{
    [TestClass, TestCategory("Unit")]
    public sealed class ChatTranscriptNativeViewsTests
    {
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

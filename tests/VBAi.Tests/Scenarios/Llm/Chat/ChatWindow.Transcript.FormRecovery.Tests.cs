using System;
using System.Collections.Generic;
using System.Windows;
using VBAi;
using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace VBAi.Tests.Unit
{
    public sealed partial class ChatWindowStateTests
    {
        [STATestMethod]
        public void VirtualizedRecoveryEntriesReleaseTheirButtonsWhenUnloaded()
        {
            using (var window = Surfaces())
            {
                var change = new FormCutChange { Form = "F" };
                var entry = new ChatEntry { FormCut = change };
                var container = new TranscriptItem { DataContext = entry };
                container.SetValue(TranscriptItem.RenderProperty, new Action<TranscriptItem>(item => Call(window, "RealizeEntry", item)));
                container.SetValue(TranscriptItem.ReleaseProperty, new Action<TranscriptItem>(item => Call(window, "ReleaseEntry", item)));
                container.RaiseEvent(new RoutedEventArgs(FrameworkElement.LoadedEvent));
                Assert.IsNotNull(container.Content);
                Assert.AreEqual(1, Get<Dictionary<FormCutChange, System.Windows.Forms.Button>>(window, "formCutButtons").Count);
                container.RaiseEvent(new RoutedEventArgs(FrameworkElement.UnloadedEvent));
                Assert.AreEqual(0, Get<Dictionary<FormCutChange, System.Windows.Forms.Button>>(window, "formCutButtons").Count);
                Assert.AreEqual(0, Get<Dictionary<ChatEntry, FrameworkElement>>(window, "entryViews").Count);
            }
        }

        [STATestMethod]
        public void ScopeClosureBetweenHostReadsIsRejectedByGitHubAndSelectionActions()
        {
            using (var runtime = new RuntimeScope())
            using (var window = new ChatWindow(runtime.Session))
            {
                var picker = Get<System.Windows.Forms.ComboBox>(window, "scopePicker");
                var saved = runtime.Host;
                runtime.Host = r => { var response = saved(r); if (r.Command == "code_panes") picker.SelectedIndex = -1; return response; };
                Call(window, "CaptureSelection");
                StringAssert.Contains(Get<System.Windows.Forms.Label>(window, "status").Text, UiText.Get("another document"));
                Set(window, "scopeSession", null);
                Call(window, "GitHub_Click", null, EventArgs.Empty);
                StringAssert.Contains(Get<System.Windows.Forms.Label>(window, "status").Text, UiText.Get("Save the document"));
            }
        }
    }
}

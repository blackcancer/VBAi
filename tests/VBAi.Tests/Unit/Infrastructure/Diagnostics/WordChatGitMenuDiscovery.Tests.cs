using System;
using System.Linq;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using VBAi.Tests.Integration;

namespace VBAi.Tests.Unit
{
    [TestClass, TestCategory("Unit")]
    public sealed class WordChatGitMenuDiscoveryTests
    {
        private const int ProcessId = 63920;
        private const uint ThreadId = 9780;
        private const long Owner = 46473454;

        [TestMethod]
        public void UniqueLocalizedVirtualGitItemBelongsToExactNewNativeMenuPopup()
        {
            var popup = Exact();
            Assert.AreSame(popup, WordChatGitMenuDiscovery.RequireUnique(new[] { popup }, ProcessId, ThreadId, Owner));
            // ToolStripMenuItem AutomationId and leaf HWND are absent on the observed product UIA surface.
            // The contract selects its exact localized Name inside a verified native Menu popup.
            Assert.IsTrue(WordChatGitMenuDiscovery.IsExactGitItem("GitHub · synchroniser le VBA…",
                "GitHub · synchroniser le VBA…", "ControlType.MenuItem", ProcessId, ProcessId));
            Assert.IsFalse(WordChatGitMenuDiscovery.IsExactGitItem("GitHub · synchronize VBA…",
                "GitHub · synchroniser le VBA…", "ControlType.MenuItem", ProcessId, ProcessId));
            Assert.IsFalse(WordChatGitMenuDiscovery.IsExactGitItem("GitHub · synchroniser le VBA…",
                "GitHub · synchroniser le VBA…", "ControlType.Text", ProcessId, ProcessId));
            Assert.IsFalse(WordChatGitMenuDiscovery.IsExactGitItem("GitHub · synchroniser le VBA…",
                "GitHub · synchroniser le VBA…", "ControlType.MenuItem", ProcessId + 1, ProcessId));
        }

        [TestMethod]
        public void ForeignStaleOrAmbiguousPopupNeverAuthorizesGitInvocation()
        {
            Action<WordChatGitMenuDiscovery.Candidate>[] changes = {
                row => row.PopupHandle = 0, row => row.OwnerHandle++, row => row.NativeProcessId++,
                row => row.UiProcessId++, row => row.GitItemProcessId++, row => row.NativeThreadId++,
                row => row.Visible = false, row => row.NewlyVisible = false,
                row => row.NativeClass = "OpusApp", row => row.UiType = "ControlType.Window",
                row => row.MenuItemCount = 0, row => row.MenuItemCount = 65,
                row => row.GitLabelMatches = 0, row => row.GitLabelMatches = 2,
                row => row.EnabledGitMatches = 0
            };
            foreach (var change in changes)
            {
                var popup = Exact(); change(popup);
                Assert.ThrowsException<InvalidOperationException>(() =>
                    WordChatGitMenuDiscovery.RequireUnique(new[] { popup }, ProcessId, ThreadId, Owner));
            }
            var another = Exact(); another.PopupHandle++;
            Assert.ThrowsException<InvalidOperationException>(() =>
                WordChatGitMenuDiscovery.RequireUnique(new[] { Exact(), another }, ProcessId, ThreadId, Owner));
            Assert.ThrowsException<InvalidOperationException>(() =>
                WordChatGitMenuDiscovery.RequireUnique(Enumerable.Repeat(Exact(), 65), ProcessId, ThreadId, Owner));
            Assert.ThrowsException<ArgumentNullException>(() =>
                WordChatGitMenuDiscovery.RequireUnique(null, ProcessId, ThreadId, Owner));
        }

        private static WordChatGitMenuDiscovery.Candidate Exact() => new WordChatGitMenuDiscovery.Candidate {
            PopupHandle = 50729498, OwnerHandle = Owner, NativeProcessId = ProcessId,
            UiProcessId = ProcessId, GitItemProcessId = ProcessId, NativeThreadId = ThreadId,
            Visible = true, NewlyVisible = true, NativeClass = "WindowsForms10.Window.20808",
            UiType = "ControlType.Menu", MenuItemCount = 17, GitLabelMatches = 1, EnabledGitMatches = 1
        };
    }
}

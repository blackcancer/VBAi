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
        public void LegacyToolbarPopupRetainsEveryExactOwnerAndVirtualItemConstraint()
        {
            var popup = HiddenOwner();
            popup.UiType = "ControlType.ToolBar";
            Assert.AreSame(popup, WordChatGitMenuDiscovery.RequireUnique(new[] { popup }, ProcessId, ThreadId, Owner));
            Action<WordChatGitMenuDiscovery.Candidate>[] changes = {
                row => row.NativeClass = "WindowsForms10.Window.0.app.0.3475548_r8_ad1",
                row => row.OwnerShape.ProcessId++, row => row.OwnerShape.ThreadId++,
                row => row.OwnerShape.Visible = true, row => row.NewlyVisible = false,
                row => row.GitLabelMatches = 2, row => row.EnabledGitMatches = 0,
                row => row.GitItemNativeAncestor++, row => row.GitItemProcessId++
            };
            foreach (var change in changes)
            {
                popup = HiddenOwner(); popup.UiType = "ControlType.ToolBar"; change(popup);
                Assert.ThrowsException<InvalidOperationException>(() =>
                    WordChatGitMenuDiscovery.RequireUnique(new[] { popup }, ProcessId, ThreadId, Owner));
            }
        }

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
        public void HiddenStandaloneWinFormsDropDownOwnerMatchesOnlyTheExactPopupDomain()
        {
            var popup = HiddenOwner();
            Assert.AreSame(popup, WordChatGitMenuDiscovery.RequireUnique(new[] { popup }, ProcessId, ThreadId, Owner));
            WordChatGitMenuDiscovery.RequireUnchangedPopupOwner(popup, popup.NativeClass,
                popup.OwnerHandle, Copy(popup.OwnerShape), ProcessId, ThreadId, Owner);
            Action<WordChatGitMenuDiscovery.Candidate>[] changes = {
                row => row.OwnerShape.Live = false,
                row => row.OwnerShape.ProcessId++, row => row.OwnerShape.ThreadId++,
                row => row.OwnerShape.Visible = true, row => row.OwnerShape.Parent = Owner,
                row => row.OwnerShape.Root = Owner, row => row.OwnerShape.Owner = Owner,
                row => row.OwnerShape.Style |= 0x40000000,
                row => row.OwnerShape.ExStyle &= ~0x00000080u,
                row => row.OwnerShape.ExStyle &= ~0x00000100u,
                row => row.OwnerShape.ClassName = "OpusApp",
                row => row.OwnerShape.ClassName = "WindowsForms10.Window.0.app.other",
                row => row.OwnerShape.Handle++
            };
            foreach (var change in changes)
            {
                popup = HiddenOwner(); change(popup);
                Assert.ThrowsException<InvalidOperationException>(() =>
                    WordChatGitMenuDiscovery.RequireUnique(new[] { popup }, ProcessId, ThreadId, Owner));
            }
        }

        [TestMethod]
        public void ExactVbeRootPopupOwnerAlsoRequiresLiveSameProcessAndThread()
        {
            var popup = Exact();
            WordChatGitMenuDiscovery.RequireUnchangedPopupOwner(popup, popup.NativeClass,
                popup.OwnerHandle, Copy(popup.OwnerShape), ProcessId, ThreadId, Owner);
            Action<WordChatGitMenuDiscovery.OwnerShape>[] changes = {
                row => row.Live = false, row => row.ProcessId++, row => row.ThreadId++,
                row => row.Root++
            };
            foreach (var change in changes)
            {
                popup = Exact(); change(popup.OwnerShape);
                Assert.ThrowsException<InvalidOperationException>(() =>
                    WordChatGitMenuDiscovery.RequireUnique(new[] { popup }, ProcessId, ThreadId, Owner));
            }
        }

        [TestMethod]
        public void FrozenPopupOwnerMetadataMustRemainExactBeforeInvocation()
        {
            var popup = HiddenOwner();
            Action<WordChatGitMenuDiscovery.OwnerShape>[] changes = {
                row => row.Style++, row => row.ExStyle++, row => row.ClassName += "changed",
                row => row.ProcessId++, row => row.ThreadId++, row => row.Visible = true,
                row => row.Parent = Owner, row => row.Owner = Owner, row => row.Root = Owner,
                row => row.Live = false, row => row.Handle++
            };
            foreach (var change in changes)
            {
                var now = Copy(popup.OwnerShape); change(now);
                Assert.ThrowsException<InvalidOperationException>(() =>
                    WordChatGitMenuDiscovery.RequireUnchangedPopupOwner(popup, popup.NativeClass,
                        popup.OwnerHandle, now, ProcessId, ThreadId, Owner));
            }
            Assert.ThrowsException<InvalidOperationException>(() =>
                WordChatGitMenuDiscovery.RequireUnchangedPopupOwner(popup, "WindowsForms10.Window.20808.app.other",
                    popup.OwnerHandle, Copy(popup.OwnerShape), ProcessId, ThreadId, Owner));
            Assert.ThrowsException<InvalidOperationException>(() =>
                WordChatGitMenuDiscovery.RequireUnchangedPopupOwner(popup, popup.NativeClass,
                    popup.OwnerHandle + 1, Copy(popup.OwnerShape), ProcessId, ThreadId, Owner));
        }

        [TestMethod]
        public void ForeignStaleOrAmbiguousPopupNeverAuthorizesGitInvocation()
        {
            Action<WordChatGitMenuDiscovery.Candidate>[] changes = {
                row => row.PopupHandle = 0, row => row.OwnerHandle = 0, row => row.OwnerHandle++,
                row => row.OwnerShape = null, row => row.NativeProcessId++,
                row => row.UiProcessId++, row => row.GitItemProcessId++, row => row.NativeThreadId++,
                row => row.GitItemNativeAncestor++,
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
            var another = Exact(); another.PopupHandle++; another.GitItemNativeAncestor++;
            Assert.ThrowsException<InvalidOperationException>(() =>
                WordChatGitMenuDiscovery.RequireUnique(new[] { Exact(), another }, ProcessId, ThreadId, Owner));
            Assert.ThrowsException<InvalidOperationException>(() =>
                WordChatGitMenuDiscovery.RequireUnique(Enumerable.Repeat(Exact(), 65), ProcessId, ThreadId, Owner));
            Assert.ThrowsException<ArgumentNullException>(() =>
                WordChatGitMenuDiscovery.RequireUnique(null, ProcessId, ThreadId, Owner));
        }

        private static WordChatGitMenuDiscovery.Candidate Exact() => new WordChatGitMenuDiscovery.Candidate {
            PopupHandle = 50729498, OwnerHandle = Owner, GitItemNativeAncestor = 50729498,
            NativeProcessId = ProcessId,
            UiProcessId = ProcessId, GitItemProcessId = ProcessId, NativeThreadId = ThreadId,
            Visible = true, NewlyVisible = true, NativeClass = "WindowsForms10.Window.20808",
            UiType = "ControlType.Menu", MenuItemCount = 17, GitLabelMatches = 1, EnabledGitMatches = 1,
            OwnerShape = new WordChatGitMenuDiscovery.OwnerShape {
                Handle = Owner, Live = true, ProcessId = ProcessId, ThreadId = ThreadId, Root = Owner
            }
        };

        private static WordChatGitMenuDiscovery.Candidate HiddenOwner()
        {
            var popup = Exact();
            popup.OwnerHandle = 115086426;
            popup.NativeClass = "WindowsForms10.Window.20808.app.0.3475548_r8_ad1";
            popup.OwnerShape = new WordChatGitMenuDiscovery.OwnerShape {
                Handle = popup.OwnerHandle, Live = true, ProcessId = ProcessId, ThreadId = ThreadId,
                ClassName = "WindowsForms10.Window.0.app.0.3475548_r8_ad1", Visible = false,
                Parent = 0, Root = popup.OwnerHandle, Owner = 0,
                Style = 79691776, ExStyle = 384
            };
            return popup;
        }

        private static WordChatGitMenuDiscovery.OwnerShape Copy(WordChatGitMenuDiscovery.OwnerShape value)
            => new WordChatGitMenuDiscovery.OwnerShape {
                Handle = value.Handle, Live = value.Live, ProcessId = value.ProcessId,
                ThreadId = value.ThreadId, ClassName = value.ClassName, Visible = value.Visible,
                Parent = value.Parent, Root = value.Root, Owner = value.Owner,
                Style = value.Style, ExStyle = value.ExStyle
            };
    }
}

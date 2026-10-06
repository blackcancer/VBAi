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
        public void RealDesktopPopupRequiresItsFrozenExactOptionsButtonOwner()
        {
            var button = OptionsButtonOwner();
            var popup = Exact(); popup.OwnerHandle = button.Handle; popup.OwnerShape = Copy(button);
            Assert.ThrowsException<InvalidOperationException>(() =>
                WordChatGitMenuDiscovery.RequireUnique(new[] { popup }, ProcessId, ThreadId, Owner));
            Assert.ThrowsException<InvalidOperationException>(() =>
                WordChatGitMenuDiscovery.RequireUnique(new[] { popup }, ProcessId, ThreadId, Owner, ScopePickerOwner()));
            Assert.AreSame(popup, WordChatGitMenuDiscovery.RequireUnique(new[] { popup }, ProcessId, ThreadId, Owner,
                exactOptionsButton: button));
            WordChatGitMenuDiscovery.RequireUnchangedPopupOwner(popup, popup.NativeClass, button.Handle,
                Copy(button), ProcessId, ThreadId, Owner, exactOptionsButton: button);
            var otherButton = Copy(button); otherButton.Handle++;
            Assert.ThrowsException<InvalidOperationException>(() =>
                WordChatGitMenuDiscovery.RequireUnique(new[] { popup }, ProcessId, ThreadId, Owner,
                    exactOptionsButton: otherButton));
        }

        [DataTestMethod]
        [DataRow("Handle"), DataRow("Parent"), DataRow("Root"), DataRow("Owner")]
        [DataRow("Process"), DataRow("Thread"), DataRow("Live"), DataRow("Visible")]
        [DataRow("Style"), DataRow("NoChildStyle"), DataRow("ExStyle")]
        [DataRow("ClassName"), DataRow("WrongClass"), DataRow("NullClass")]
        public void OptionsPopupRefusesChangedOwnerBeforeSelectionAndInvocation(string field)
        {
            var frozen = OptionsButtonOwner(); var now = Copy(frozen);
            switch (field)
            {
                case "Handle": now.Handle++; break;
                case "Parent": now.Parent++; break;
                case "Root": now.Root++; break;
                case "Owner": now.Owner = Owner; break;
                case "Process": now.ProcessId++; break;
                case "Thread": now.ThreadId++; break;
                case "Live": now.Live = false; break;
                case "Visible": now.Visible = false; break;
                case "Style": now.Style++; break;
                case "NoChildStyle": now.Style = 0; break;
                case "ExStyle": now.ExStyle++; break;
                case "ClassName": now.ClassName += "changed"; break;
                case "WrongClass": now.ClassName = "WindowsForms10.COMBOBOX.app.0.23dba96_r133_ad1"; break;
                case "NullClass": now.ClassName = null; break;
                default: Assert.Fail("Unknown owner mutation."); break;
            }
            var popup = Exact(); popup.OwnerHandle = now.Handle; popup.OwnerShape = now;
            Assert.ThrowsException<InvalidOperationException>(() =>
                WordChatGitMenuDiscovery.RequireUnique(new[] { popup }, ProcessId, ThreadId, Owner,
                    exactOptionsButton: frozen));
            popup = Exact(); popup.OwnerHandle = frozen.Handle; popup.OwnerShape = Copy(frozen);
            Assert.ThrowsException<InvalidOperationException>(() =>
                WordChatGitMenuDiscovery.RequireUnchangedPopupOwner(popup, popup.NativeClass, now.Handle,
                    now, ProcessId, ThreadId, Owner, exactOptionsButton: frozen));
        }

        [TestMethod]
        public void OptionsOwnerKeepsEveryPopupIdentityAndUniqueItemGuard()
        {
            var button = OptionsButtonOwner();
            Action<WordChatGitMenuDiscovery.Candidate>[] changes = {
                row => row.NativeProcessId++, row => row.UiProcessId++, row => row.NativeThreadId++,
                row => row.GitItemProcessId++, row => row.GitItemNativeAncestor++,
                row => row.Visible = false, row => row.NewlyVisible = false,
                row => row.NativeClass = "OpusApp", row => row.UiType = "ControlType.Window",
                row => row.MenuItemCount = 0, row => row.MenuItemCount = 65,
                row => row.GitLabelMatches = 0, row => row.GitLabelMatches = 2,
                row => row.EnabledGitMatches = 0
            };
            foreach (var change in changes)
            {
                var popup = Exact(); popup.OwnerHandle = button.Handle; popup.OwnerShape = Copy(button); change(popup);
                Assert.ThrowsException<InvalidOperationException>(() =>
                    WordChatGitMenuDiscovery.RequireUnique(new[] { popup }, ProcessId, ThreadId, Owner,
                        exactOptionsButton: button));
            }
            var one = Exact(); one.OwnerHandle = button.Handle; one.OwnerShape = Copy(button);
            var two = Exact(); two.OwnerHandle = button.Handle; two.OwnerShape = Copy(button);
            two.PopupHandle++; two.GitItemNativeAncestor++;
            Assert.ThrowsException<InvalidOperationException>(() =>
                WordChatGitMenuDiscovery.RequireUnique(new[] { one, two }, ProcessId, ThreadId, Owner,
                    exactOptionsButton: button));
            Assert.ThrowsException<InvalidOperationException>(() =>
                WordChatGitMenuDiscovery.RequireUnique(new[] { one, one }, ProcessId, ThreadId, Owner,
                    exactOptionsButton: button));
        }

        private static WordChatGitMenuDiscovery.OwnerShape OptionsButtonOwner()
            => new WordChatGitMenuDiscovery.OwnerShape {
                Handle = 21108174, Parent = 106566272, Root = Owner, Owner = 0,
                ProcessId = ProcessId, ThreadId = ThreadId, Live = true, Visible = true,
                ClassName = "WindowsForms10.BUTTON.app.0.23dba96_r133_ad1", Style = 1442906123, ExStyle = 0
            };

        [TestMethod]
        public void PrivatePopupMayUseOnlyItsFrozenExactScopePickerAsChildOwner()
        {
            var picker = ScopePickerOwner();
            var popup = Exact(); popup.OwnerHandle = picker.Handle; popup.OwnerShape = Copy(picker);
            Assert.ThrowsException<InvalidOperationException>(() =>
                WordChatGitMenuDiscovery.RequireUnique(new[] { popup }, ProcessId, ThreadId, Owner));
            Assert.AreSame(popup, WordChatGitMenuDiscovery.RequireUnique(new[] { popup }, ProcessId, ThreadId, Owner, picker));
            WordChatGitMenuDiscovery.RequireUnchangedPopupOwner(popup, popup.NativeClass, picker.Handle,
                Copy(picker), ProcessId, ThreadId, Owner, picker);
            Action<WordChatGitMenuDiscovery.OwnerShape>[] changes = {
                row => row.Handle++, row => row.Parent++, row => row.Root++, row => row.Owner = Owner,
                row => row.ProcessId++, row => row.ThreadId++, row => row.Live = false,
                row => row.Visible = false, row => row.Style = 0, row => row.ExStyle++,
                row => row.ClassName = "WindowsForms10.BUTTON.app.0.123_r6_ad1"
            };
            foreach (var change in changes)
            {
                var now = Copy(picker); change(now);
                popup = Exact(); popup.OwnerHandle = now.Handle; popup.OwnerShape = now;
                Assert.ThrowsException<InvalidOperationException>(() =>
                    WordChatGitMenuDiscovery.RequireUnique(new[] { popup }, ProcessId, ThreadId, Owner, picker));
                popup = Exact(); popup.OwnerHandle = picker.Handle; popup.OwnerShape = Copy(picker);
                Assert.ThrowsException<InvalidOperationException>(() =>
                    WordChatGitMenuDiscovery.RequireUnchangedPopupOwner(popup, popup.NativeClass, now.Handle,
                        now, ProcessId, ThreadId, Owner, picker));
            }
            popup = Exact(); popup.OwnerHandle = picker.Handle; popup.OwnerShape = Copy(picker);
            var anotherPicker = Copy(picker); anotherPicker.Handle++;
            Assert.ThrowsException<InvalidOperationException>(() =>
                WordChatGitMenuDiscovery.RequireUnique(new[] { popup }, ProcessId, ThreadId, Owner, anotherPicker));
        }

        private static WordChatGitMenuDiscovery.OwnerShape ScopePickerOwner()
            => new WordChatGitMenuDiscovery.OwnerShape {
                Handle = 53155074, Parent = Owner, Root = Owner, Owner = 0,
                ProcessId = ProcessId, ThreadId = ThreadId, Live = true, Visible = true,
                ClassName = "WindowsForms10.COMBOBOX.app.0.123_r6_ad1", Style = 0x40000000, ExStyle = 0
            };

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

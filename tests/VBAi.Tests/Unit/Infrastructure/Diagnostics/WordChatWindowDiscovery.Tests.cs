using System;
using System.Linq;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using VBAi.Tests.Integration;

namespace VBAi.Tests.Unit
{
    [TestClass, TestCategory("Unit")]
    public sealed class WordChatWindowDiscoveryTests
    {
        private const int ProcessId = 52188;
        private const uint ThreadId = 60720;

        [TestMethod]
        public void FormAutomationIdIsUnnecessaryWhenNativeChildShapeAndOwnerAreExact()
        {
            var chat = Exact();
            Assert.AreSame(chat, WordChatWindowDiscovery.RequireUnique(new[] { chat }, ProcessId, ThreadId));
            chat.ControlType = "ControlType.Pane"; // Docked WinForms form providers may expose a pane.
            Assert.AreSame(chat, WordChatWindowDiscovery.RequireUnique(new[] { chat }, ProcessId, ThreadId));
        }

        [TestMethod]
        public void EveryNativeOwnerAndLeafIdentityMustMatchBeforeAnyChatAction()
        {
            Action<WordChatWindowDiscovery.Candidate>[] changes = {
                row => row.Handle = 0, row => row.Visible = false, row => row.WithinOwnedVbe = false,
                row => row.NativeProcessId++, row => row.UiProcessId++, row => row.NativeThreadId++,
                row => row.NativeClass = "OpusApp", row => row.ControlType = "ControlType.Edit",
                row => row.FixedChatCaption = false,
                row => row.ScopePickerCount = 0, row => row.OptionsCount = 2,
                row => row.ScopePickerProcessId++, row => row.OptionsProcessId++,
                row => row.ScopePickerType = "ControlType.Edit", row => row.OptionsType = "ControlType.Text",
                row => row.ScopePickerHandle = 0, row => row.OptionsHandle = 0,
                row => row.ScopePickerThreadId++, row => row.OptionsThreadId++,
                row => row.ScopePickerWithinChat = false, row => row.OptionsWithinChat = false
            };
            foreach (var change in changes)
            {
                var candidate = Exact(); change(candidate);
                Assert.ThrowsException<InvalidOperationException>(() =>
                    WordChatWindowDiscovery.RequireUnique(new[] { candidate }, ProcessId, ThreadId));
            }
        }

        [TestMethod]
        public void AmbiguousOrUnboundedFormInventoryIsRefused()
        {
            Assert.ThrowsException<InvalidOperationException>(() => WordChatWindowDiscovery.RequireUnique(
                new[] { Exact(), Exact() }, ProcessId, ThreadId));
            var second = Exact(); second.Handle++;
            Assert.ThrowsException<InvalidOperationException>(() => WordChatWindowDiscovery.RequireUnique(
                new[] { Exact(), second }, ProcessId, ThreadId));
            Assert.ThrowsException<InvalidOperationException>(() => WordChatWindowDiscovery.RequireUnique(
                Enumerable.Range(0, 65).Select(_ => Exact()), ProcessId, ThreadId));
            Assert.ThrowsException<ArgumentNullException>(() => WordChatWindowDiscovery.RequireUnique(null, ProcessId, ThreadId));
        }

        [TestMethod]
        public void DockedChildModalUsesVerifiedTopLevelVbeRootAndRefusesChildAsOwner()
        {
            var docked = Docked();
            Assert.AreEqual(docked.VbeRoot, WordChatWindowDiscovery.RequireModalOwner(docked, ProcessId, ThreadId));
            WordChatWindowDiscovery.RequireUnchangedModalOwner(docked, Docked(), docked.VbeRoot, ProcessId, ThreadId);
            Assert.ThrowsException<InvalidOperationException>(() => WordChatWindowDiscovery.RequireUnchangedModalOwner(
                docked, Docked(), docked.ChatHandle, ProcessId, ThreadId));
            var foreign = Docked(); foreign.ChatRootProcessId++;
            Assert.ThrowsException<InvalidOperationException>(() => WordChatWindowDiscovery.RequireModalOwner(foreign, ProcessId, ThreadId));
            var changed = Docked(); changed.ChatRoot++;
            Assert.ThrowsException<InvalidOperationException>(() => WordChatWindowDiscovery.RequireUnchangedModalOwner(
                docked, changed, docked.VbeRoot, ProcessId, ThreadId));
        }

        [TestMethod]
        public void FloatingChatModalUsesExactOwnedChatRootAndRejectsOwnerChange()
        {
            var floating = Docked(); floating.ChatWithinVbe = false;
            floating.ChatRoot = floating.ChatHandle; floating.ChatOwner = floating.VbeRoot;
            Assert.AreEqual(floating.ChatHandle, WordChatWindowDiscovery.RequireModalOwner(floating, ProcessId, ThreadId));
            var changed = Docked(); changed.ChatWithinVbe = false;
            changed.ChatRoot = changed.ChatHandle; changed.ChatOwner = changed.VbeRoot;
            WordChatWindowDiscovery.RequireUnchangedModalOwner(floating, changed, floating.ChatHandle, ProcessId, ThreadId);
            changed.ChatOwner = 0;
            Assert.ThrowsException<InvalidOperationException>(() => WordChatWindowDiscovery.RequireUnchangedModalOwner(
                floating, changed, floating.ChatHandle, ProcessId, ThreadId));
            changed = Docked(); changed.ChatWithinVbe = false; changed.ChatRoot = changed.ChatHandle;
            changed.ChatOwner = changed.VbeRoot;
            Assert.ThrowsException<InvalidOperationException>(() => WordChatWindowDiscovery.RequireUnchangedModalOwner(
                floating, changed, floating.VbeRoot, ProcessId, ThreadId));
        }

        private static WordChatWindowDiscovery.OwnerIdentity Docked() => new WordChatWindowDiscovery.OwnerIdentity {
            VbeHandle = 18486114, VbeRoot = 18486114, ChatHandle = 18486120, ChatRoot = 18486114,
            ChatOwner = 0, ChatWithinVbe = true, VbeRootProcessId = ProcessId,
            ChatRootProcessId = ProcessId, VbeRootThreadId = ThreadId, ChatRootThreadId = ThreadId
        };

        private static WordChatWindowDiscovery.Candidate Exact() => new WordChatWindowDiscovery.Candidate {
            Handle = 18486120, NativeProcessId = ProcessId, UiProcessId = ProcessId, NativeThreadId = ThreadId,
            Visible = true, WithinOwnedVbe = true, FixedChatCaption = true,
            NativeClass = "WindowsForms10.Window.8.app.0.example",
            ControlType = "ControlType.Window", ScopePickerCount = 1, OptionsCount = 1,
            ScopePickerProcessId = ProcessId, OptionsProcessId = ProcessId,
            ScopePickerType = "ControlType.ComboBox", OptionsType = "ControlType.Button",
            ScopePickerHandle = 18486130, OptionsHandle = 18486140,
            ScopePickerThreadId = ThreadId, OptionsThreadId = ThreadId,
            ScopePickerWithinChat = true, OptionsWithinChat = true
        };
    }
}

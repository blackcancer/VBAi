using System;
using System.Collections;
using System.Collections.Generic;
using CodexVBE;
using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace CodexVBE.Tests.Unit
{
    [TestClass]
    [TestCategory("Unit")]
    public sealed class VbeMenuLifecycleTests
    {
        [TestMethod]
        public void MenuActionsUseTheirOwnCallbacksAndDisposeRemovesEveryCreatedButton()
        {
            var host = Host();
            var subscriptions = new Dictionary<object, Delegate>();
            var icons = new List<Type>();
            var actions = new List<string>();
            int removed = 0;
            var menu = new VbeMenu(host, () => actions.Add("assistant"),
                () => actions.Add("settings"), () => actions.Add("github"),
                command => actions.Add(command),
                (button, iid, dispid, handler) => { Assert.AreEqual(1, dispid); subscriptions.Add(button, handler); },
                (button, iid, dispid, handler) => { Assert.AreSame(handler, subscriptions[button]); removed++; },
                (button, iconType) => icons.Add(iconType));

            var view = host.CommandBars[0].Controls.Items[0].Controls.Items;
            var tools = host.CommandBars[0].Controls.Items[1].Controls.Items;
            var editor = host.CommandBars[1].Controls.Items;
            Assert.AreEqual("CodexVBE.Assistant", view[0].Tag);
            Assert.AreEqual("CodexVBE.GitHub", view[1].Tag);
            Assert.AreEqual("CodexVBE.Settings", tools[0].Tag);
            Assert.AreEqual(3, editor.Count);
            Assert.AreEqual("CodexVBE.expliquer", editor[0].Tag);
            Assert.AreEqual("CodexVBE.corriger", editor[1].Tag);
            Assert.AreEqual("CodexVBE.refactoriser", editor[2].Tag);
            Assert.AreEqual(6, icons.Count);

            foreach (var button in new[] { view[0], tools[0], view[1], editor[0], editor[1], editor[2] })
            {
                var arguments = new object[] { button, false };
                subscriptions[button].DynamicInvoke(arguments);
                Assert.AreEqual(true, arguments[1], "The VBE default click action must be canceled.");
            }
            CollectionAssert.AreEqual(new[] { "assistant", "settings", "github",
                "/expliquer", "/corriger", "/refactoriser" }, actions);

            menu.Dispose();
            Assert.AreEqual(6, removed);
            foreach (var button in new[] { view[0], tools[0], view[1], editor[0], editor[1], editor[2] })
                Assert.AreEqual(1, button.DeleteCount);
            menu.Dispose();
            Assert.AreEqual(6, removed, "Disposal must be idempotent.");
        }

        [TestMethod]
        public void FailedGitButtonCreationRemovesAssistantAndSettingsButtons()
        {
            var host = Host();
            var view = host.CommandBars[0].Controls.Items[0].Controls;
            var tools = host.CommandBars[0].Controls.Items[1].Controls;
            view.FailAtCount = 1;
            int removed = 0;
            Assert.ThrowsException<InvalidOperationException>(() => new VbeMenu(host,
                () => { }, () => { }, () => { }, null,
                (button, iid, dispid, handler) => { },
                (button, iid, dispid, handler) => removed++,
                (button, iconType) => { }));
            Assert.AreEqual(2, removed);
            Assert.AreEqual(1, view.Items[0].DeleteCount);
            Assert.AreEqual(1, tools.Items[0].DeleteCount);
        }

        [TestMethod]
        public void MissingEditorCommandBarDoesNotAffectMainMenuActions()
        {
            var host = Host();
            host.CommandBars[1].Name = "Immediate";
            int subscribed = 0;
            using (var menu = new VbeMenu(host, () => { }, () => { }, () => { },
                command => { }, (button, iid, dispid, handler) => subscribed++,
                (button, iid, dispid, handler) => { }, (button, iconType) => { }))
                Assert.AreEqual(3, subscribed);
        }

        [TestMethod]
        public void EditorSubscriptionFailureKeepsMainMenuAndCleansPartialEditorButton()
        {
            var host = Host();
            int subscriptions = 0;
            int removals = 0;
            using (var menu = new VbeMenu(host, () => { }, () => { }, () => { },
                command => { },
                (button, iid, dispid, handler) => {
                    subscriptions++;
                    if (subscriptions == 4) throw new InvalidOperationException("Editor event unavailable");
                },
                (button, iid, dispid, handler) => removals++,
                (button, iconType) => { }))
            {
                Assert.AreEqual(4, subscriptions);
                Assert.AreEqual(1, host.CommandBars[1].Controls.Items.Count);
                Assert.AreEqual(0, host.CommandBars[0].Controls.Items[0].Controls.Items[0].DeleteCount);
            }
            Assert.AreEqual(4, removals);
            Assert.AreEqual(1, host.CommandBars[1].Controls.Items[0].DeleteCount);
        }

        private static FakeHost Host()
        {
            var view = new FakeButton { Caption = "&View" };
            var tools = new FakeButton { Caption = "&Tools" };
            var main = new FakeBar { Type = 1, Controls = new FakeControls() };
            main.Controls.Items.Add(view);
            main.Controls.Items.Add(tools);
            return new FakeHost { CommandBars = new[] { main,
                new FakeBar { Type = 0, Name = "Code Window", Controls = new FakeControls() } } };
        }

        public sealed class FakeHost { public FakeBar[] CommandBars { get; set; } }
        public sealed class FakeBar
        {
            public int Type { get; set; }
            public string Name { get; set; }
            public FakeControls Controls { get; set; }
        }
        public sealed class FakeButton
        {
            public string Caption { get; set; }
            public string Tag { get; set; }
            public string TooltipText { get; set; }
            public FakeControls Controls { get; } = new FakeControls();
            public int DeleteCount { get; private set; }
            public void Delete() { DeleteCount++; }
        }
        public sealed class FakeControls : IEnumerable<FakeButton>
        {
            public List<FakeButton> Items { get; } = new List<FakeButton>();
            public int FailAtCount { get; set; } = int.MaxValue;
            public FakeButton Add(int type, object id, object parameter, object before, bool temporary)
            {
                if (Items.Count >= FailAtCount) throw new InvalidOperationException("Add failed");
                var button = new FakeButton();
                Items.Add(button);
                return button;
            }
            public IEnumerator<FakeButton> GetEnumerator() { return Items.GetEnumerator(); }
            IEnumerator IEnumerable.GetEnumerator() { return GetEnumerator(); }
        }
    }
}

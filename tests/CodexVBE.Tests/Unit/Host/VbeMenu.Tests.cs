namespace CodexVBE.Tests.Unit
{
    using System;
    using System.Reflection;
    using System.Windows.Forms;
    using CodexVBE;
    using Microsoft.VisualStudio.TestTools.UnitTesting;

    /// <summary>Vérifie la découverte des menus VBE et le traitement des barres COM.</summary>
    public sealed partial class HostSettingsCoverageTests
    {
        /// <summary>Ignore les barres et contrôles COM invalides et normalise les légendes localisées.</summary>
        [TestMethod]
        public void MenuLookupSkipsInvalidBarsAndNormalizesLocalizedCaptions()
        {
            var host = new FakeHost
            {
                CommandBars = new object[]
                {
                    new InvalidBar(),
                    new FakeBar
                    {
                        Type = 2,
                        Controls = new object[]
                        {
                            new FakeControl
                            {
                                Caption = "&Outils"
                            }
                        }
                    },
                    new FakeBar
                    {
                        Type = 1,
                        Controls = new object[]
                        {
                            new InvalidControl(),
                            new FakeControl
                            {
                                Caption = " &Affichage "
                            },
                            new FakeControl
                            {
                                Caption = "&Outils"
                            }
                        }
                    }
                }
            };
            var method = typeof(VbeMenu).GetMethod("FindMenu", BindingFlags.Static | BindingFlags.NonPublic);
            Assert.IsNotNull(method);
            Assert.AreSame(((FakeBar)host.CommandBars[2]).Controls[1], method.Invoke(null, new object[] { host, true }));
            Assert.AreSame(((FakeBar)host.CommandBars[2]).Controls[2], method.Invoke(null, new object[] { host, false }));
            var missing = Assert.ThrowsException<TargetInvocationException>(() => method.Invoke(null, new object[] { new FakeHost { CommandBars = new object[0] }, false }));
            StringAssert.Contains(missing.InnerException.Message, "VBE menu not found: Tools");
        }

        /// <summary>Signale l’absence du menu Tools lors de la création des commandes principales.</summary>
        [TestMethod]
        public void MenuConstructionReportsMissingHostMenu()
        {
            var host = new FakeHost
            {
                CommandBars = new object[]
                {
                    new FakeBar
                    {
                        Type = 1,
                        Controls = new object[]
                        {
                            new FakeControl
                            {
                                Caption = "&View"
                            }
                        }
                    }
                }
            };
            var thrown = Assert.ThrowsException<InvalidOperationException>(() => new VbeMenu(host, () =>
            {
            }, () =>
            {
            }, () =>
            {
            }));
            StringAssert.Contains(thrown.Message, "VBE menu not found: Tools");
        }
    }
}

namespace CodexVBE.Tests.Unit
{
    using System;
    using System.Collections.Generic;
    using System.Reflection;
    using System.Runtime.InteropServices;
    using System.Runtime.Serialization;
    using System.Windows.Forms;
    using CodexVBE;
    using Microsoft.VisualStudio.TestTools.UnitTesting;

    /// <summary>Vérifie la recherche des menus selon le type de barre et la légende normalisée.</summary>
    public sealed partial class HostSettingsWindowTests
    {
        /// <summary>Ignore les accélérateurs et les barres qui ne représentent pas des menus.</summary>
        [TestMethod]
        public void MenuDiscoveryIgnoresAcceleratorsAndNonMenuBars()
        {
            var method = typeof(VbeMenu).GetMethod("FindMenu", BindingFlags.NonPublic | BindingFlags.Static);
            Assert.IsNotNull(method);
            var host = new FakeMenus();
            host.CommandBars.Add(new FakeBar { Type = 2, Controls = new List<FakeMenu> { new FakeMenu { Caption = "&Affichage" } } });
            var expected = new FakeMenu
            {
                Caption = " &Affichage "
            };
            host.CommandBars.Add(new FakeBar { Type = 1, Controls = new List<FakeMenu> { expected, new FakeMenu { Caption = "&Outils" } } });
            Assert.AreSame(expected, method.Invoke(null, new object[] { host, true }));
            var error = Assert.ThrowsException<TargetInvocationException>(() => method.Invoke(null, new object[] { new FakeMenus(), true }));
            Assert.IsInstanceOfType(error.InnerException, typeof(InvalidOperationException));
            var incomplete = new FakeMenus();
            incomplete.CommandBars.Add(new FakeBar { Type = 1, Controls = new List<FakeMenu> { new FakeMenu { Caption = "View" } } });
            Assert.ThrowsException<InvalidOperationException>(() => new VbeMenu(incomplete, () =>
            {
            }, () =>
            {
            }, () =>
            {
            }));
        }
    }
}

namespace CodexVBE.Tests.Unit
{
    using System;
    using System.Collections;
    using System.Reflection;
    using System.Collections.Generic;
    using CodexVBE;
    using Microsoft.VisualStudio.TestTools.UnitTesting;

    /// <summary>Vérifie les abonnements, actions et nettoyages des menus VBE.</summary>
    [TestClass]
    [TestCategory("Unit")]
    public sealed partial class VbeMenuLifecycleTests
    {
                /// <summary>Vérifie les images natives, menus localisés, barres d’éditeur et erreurs de suppression.</summary>
[STATestMethod]
        public void NativeMenuPicturesLocalizedEditorBarsAndRemovalFailuresKeepAllGuards()
        {
            using (var theme = new CodexVBE.Tests.Infrastructure.ThemeScope()) using (var culture = new CodexVBE.Tests.Infrastructure.LocalizationScope())
            {
                var host = Host(); host.CommandBars[0].Controls.Items.Insert(0, new FakeButton { Caption = null }); host.CommandBars = new[] { host.CommandBars[0], new FakeBar { Type = 0, Name = "Code Window (Break)", Controls = new FakeControls() }, new FakeBar { Type = 0, Name = "Fenêtre code", Controls = new FakeControls() }, new FakeBar { Type = 0, Name = null, Controls = new FakeControls() } };
                using (var menu = new VbeMenu(host, () => { }, () => { }, () => { }, command => { }, (b, i, d, h) => { }, null, null))
                {
                    var view = host.CommandBars[0].Controls.Items[1].Controls.Items[0]; Assert.AreEqual(0, view.Style); menu.ReadIcon = t => (System.Drawing.Icon)System.Drawing.SystemIcons.Information.Clone(); typeof(VbeMenu).GetMethod("SetIcon", BindingFlags.NonPublic | BindingFlags.Instance).Invoke(menu, new object[] { view, typeof(ChatWindow) }); Assert.AreEqual(3, view.Style); Assert.IsNotNull(view.Picture); Assert.IsNotNull(view.Mask); Assert.AreEqual(3, host.CommandBars[1].Controls.Items.Count); Assert.AreEqual(3, host.CommandBars[2].Controls.Items.Count);
                    typeof(VbeMenu).GetMethod("SetIcon", BindingFlags.NonPublic | BindingFlags.Instance).Invoke(menu, new object[] { new object(), typeof(ChatWindow) });
                    typeof(VbeMenu).GetMethod("SetIcon", BindingFlags.NonPublic | BindingFlags.Instance).Invoke(menu, new object[] { new object(), typeof(VbeMenuLifecycleTests) });
                }
                var nested = typeof(VbeMenu).GetNestedType("MenuPicture", BindingFlags.NonPublic); var invalid = Assert.ThrowsException<TargetInvocationException>(() => Activator.CreateInstance(nested, true)); Assert.IsInstanceOfType<FormatException>(invalid.InnerException);
                host = Host(); Assert.ThrowsException<ArgumentException>(() => new VbeMenu(host, () => { }, () => { }, () => { }));
                var nativeSubscription = VbeMenu.SubscribeDefault; try { VbeMenu.SubscribeDefault = (b, i, d, h) => { }; using (var publicMenu = new VbeMenu(Host(), () => { }, () => { }, () => { })) Assert.IsNotNull(publicMenu); } finally { VbeMenu.SubscribeDefault = nativeSubscription; }
                host = Host(); var subscriptions = new List<object>(); using (var menu = new VbeMenu(host, () => { }, () => { }, () => { }, c => { }, (b, i, d, h) => subscriptions.Add(b), (b, i, d, h) => throw new InvalidOperationException("unsubscribe failed"), (b, t) => { })) { foreach (FakeButton button in subscriptions) button.RejectDelete = true; }
                Assert.AreEqual(6, subscriptions.Count);
            }
        }
                /// <summary>Vérifie le nettoyage partiel des menus face aux objets COM en lecture seule ou absents.</summary>
[TestMethod]
        public void PartialMenuDestructionPreservesReadonlyNullDefenses()
        {
            foreach (var name in new[] { "viewButton", "settingsButton", "viewHandler", "settingsHandler" })
            {
                var menu = new VbeMenu(Host(), () => { }, () => { }, () => { }, null, (b, i, d, h) => { }, (b, i, d, h) => { }, (b, t) => { }); var field = typeof(VbeMenu).GetField(name, BindingFlags.Instance | BindingFlags.NonPublic); var original = field.GetValue(menu);
                try { field.SetValue(menu, null); menu.Dispose(); }
                finally { field.SetValue(menu, original); menu.Dispose(); }
                Assert.AreSame(original, field.GetValue(menu));
            }
        }
        /// <summary>Associe chaque bouton à sa propre action et retire toutes les commandes au Dispose.</summary>
        [TestMethod]
        public void MenuActionsUseTheirOwnCallbacksAndDisposeRemovesEveryCreatedButton()
        {
            var host = Host();
            var subscriptions = new Dictionary<object, Delegate>();
            var icons = new List<Type>();
            var actions = new List<string>();
            int removed = 0;
            var menu = new VbeMenu(host, () => actions.Add("assistant"), () => actions.Add("settings"), () => actions.Add("github"), command => actions.Add(command), (button, iid, dispid, handler) =>
            {
                Assert.AreEqual(1, dispid);
                subscriptions.Add(button, handler);
            }, (button, iid, dispid, handler) =>
            {
                Assert.AreSame(handler, subscriptions[button]);
                removed++;
            }, (button, iconType) => icons.Add(iconType));
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
            foreach (var button in new[]
            {
                view[0],
                tools[0],
                view[1],
                editor[0],
                editor[1],
                editor[2]
            }

            )
            {
                var arguments = new object[]
                {
                    button,
                    false
                };
                subscriptions[button].DynamicInvoke(arguments);
                Assert.AreEqual(true, arguments[1], "The VBE default click action must be canceled.");
            }

            CollectionAssert.AreEqual(new[] { "assistant", "settings", "github", "/expliquer", "/corriger", "/refactoriser" }, actions);
            menu.Dispose();
            Assert.AreEqual(6, removed);
            foreach (var button in new[]
            {
                view[0],
                tools[0],
                view[1],
                editor[0],
                editor[1],
                editor[2]
            }

            )
                Assert.AreEqual(1, button.DeleteCount);
            menu.Dispose();
            Assert.AreEqual(6, removed, "Disposal must be idempotent.");
        }

        /// <summary>Supprime les boutons déjà créés lorsque la création du bouton Git échoue.</summary>
        [TestMethod]
        public void FailedGitButtonCreationRemovesAssistantAndSettingsButtons()
        {
            var host = Host();
            var view = host.CommandBars[0].Controls.Items[0].Controls;
            var tools = host.CommandBars[0].Controls.Items[1].Controls;
            view.FailAtCount = 1;
            int removed = 0;
            Assert.ThrowsException<InvalidOperationException>(() => new VbeMenu(host, () =>
            {
            }, () =>
            {
            }, () =>
            {
            }, null, (button, iid, dispid, handler) =>
            {
            }, (button, iid, dispid, handler) => removed++, (button, iconType) =>
            {
            }));
            Assert.AreEqual(2, removed);
            Assert.AreEqual(1, view.Items[0].DeleteCount);
            Assert.AreEqual(1, tools.Items[0].DeleteCount);
        }

        /// <summary>Préserve les commandes principales lorsque la barre de commandes éditeur manque.</summary>
        [TestMethod]
        public void MissingEditorCommandBarDoesNotAffectMainMenuActions()
        {
            var host = Host();
            host.CommandBars[1].Name = "Immediate";
            int subscribed = 0;
            using (var menu = new VbeMenu(host, () =>
            {
            }, () =>
            {
            }, () =>
            {
            }, command =>
            {
            }, (button, iid, dispid, handler) => subscribed++, (button, iid, dispid, handler) =>
            {
            }, (button, iconType) =>
            {
            }))
                Assert.AreEqual(3, subscribed);
        }

        /// <summary>Nettoie le bouton éditeur partiel sans retirer les commandes principales après échec d’abonnement.</summary>
        [TestMethod]
        public void EditorSubscriptionFailureKeepsMainMenuAndCleansPartialEditorButton()
        {
            var host = Host();
            int subscriptions = 0;
            int removals = 0;
            using (var menu = new VbeMenu(host, () =>
            {
            }, () =>
            {
            }, () =>
            {
            }, command =>
            {
            }, (button, iid, dispid, handler) =>
            {
                subscriptions++;
                if (subscriptions == 4)
                    throw new InvalidOperationException("Editor event unavailable");
            }, (button, iid, dispid, handler) => removals++, (button, iconType) =>
            {
            }))
            {
                Assert.AreEqual(4, subscriptions);
                Assert.AreEqual(1, host.CommandBars[1].Controls.Items.Count);
                Assert.AreEqual(0, host.CommandBars[0].Controls.Items[0].Controls.Items[0].DeleteCount);
            }

            Assert.AreEqual(4, removals);
            Assert.AreEqual(1, host.CommandBars[1].Controls.Items[0].DeleteCount);
        }
    }
}

namespace CodexVBE.Tests.Unit
{
using System;
using System.Collections.Generic;
using System.Linq;
using CodexVBE;
using Microsoft.VisualStudio.TestTools.UnitTesting;

    public sealed partial class VbeMenuLifecycleTests
    {
        [TestMethod]
        public void UpdatesHaveTheirOwnToolsCallbackAndCleanup()
        {
            var host = Host(); var subscriptions = new Dictionary<object, Delegate>(); int opened = 0, removed = 0;
            var menu = new VbeMenu(host, () => { }, () => { }, () => { }, null,
                (button, iid, dispid, callback) => subscriptions.Add(button, callback),
                (button, iid, dispid, callback) => removed++, (button, type) => { }, null, null, () => opened++);
            var updates = host.CommandBars[0].Controls.Items[1].Controls.Items.Single(x => x.Tag == "CodexVBE.Updates");
            var args = new object[] { updates, false }; subscriptions[updates].DynamicInvoke(args);
            Assert.AreEqual(true, args[1]); Assert.AreEqual(1, opened);
            menu.Dispose(); Assert.AreEqual(4, removed); Assert.AreEqual(1, updates.DeleteCount);
        }

        [TestMethod]
        public void CrashReportHasItsOwnToolsCallbackIconAndCleanup()
        {
            var host = Host(); var subscriptions = new Dictionary<object, Delegate>();
            var icons = new Dictionary<object, Type>(); int opened = 0, removed = 0;
            var menu = new VbeMenu(host, () => { }, () => { }, () => { }, null,
                (button, iid, dispid, callback) => subscriptions.Add(button, callback),
                (button, iid, dispid, callback) => removed++, (button, type) => icons.Add(button, type), null, () => opened++);
            var report = host.CommandBars[0].Controls.Items[1].Controls.Items.Single(x => x.Tag == "CodexVBE.CrashReport");
            Assert.AreEqual(typeof(CrashReportWindow), icons[report]);
            var args = new object[] { report, false }; subscriptions[report].DynamicInvoke(args);
            Assert.AreEqual(true, args[1]); Assert.AreEqual(1, opened);
            menu.Dispose(); menu.Dispose(); Assert.AreEqual(4, removed); Assert.AreEqual(1, report.DeleteCount);
        }

        [TestMethod]
        public void AboutHasItsOwnToolsCallbackIconAndCleanup()
        {
            var host = Host();
            var subscriptions = new Dictionary<object, Delegate>();
            var icons = new Dictionary<object, Type>();
            int opened = 0, removed = 0;
            var menu = new VbeMenu(host, () => { }, () => { }, () => { }, null,
                (button, iid, dispid, callback) => subscriptions.Add(button, callback),
                (button, iid, dispid, callback) => removed++,
                (button, type) => icons.Add(button, type), () => opened++);
            var about = host.CommandBars[0].Controls.Items[1].Controls.Items.Single(x => x.Tag == "CodexVBE.About");
            Assert.AreEqual(UiText.Get("About VBAi"), about.Caption);
            Assert.AreEqual(typeof(AboutWindow), icons[about]);
            var args = new object[] { about, false };
            subscriptions[about].DynamicInvoke(args);
            Assert.AreEqual(true, args[1]);
            Assert.AreEqual(1, opened);
            menu.Dispose(); menu.Dispose();
            Assert.AreEqual(4, removed);
            Assert.AreEqual(1, about.DeleteCount);
        }
    }
}

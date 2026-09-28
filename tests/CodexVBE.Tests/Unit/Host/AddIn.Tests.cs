namespace CodexVBE.Tests.Unit
{
    using System;
    using System.Reflection;
    using System.Windows.Forms;
    using CodexVBE;
    using Microsoft.VisualStudio.TestTools.UnitTesting;

    /// <summary>Vérifie l’arrêt de l’add-in et la propriété des fenêtres natives.</summary>
    [TestClass]
    [TestCategory("Unit")]
    public sealed partial class HostSettingsCoverageTests
    {
        /// <summary>Ferme la fenêtre native une fois et rend l’arrêt répétable.</summary>
        [TestMethod]
        [STATestMethod]
        public void AddInShutdownClosesNativeWindowAndCanRunTwice()
        {
            var addIn = new AddIn();
            var native = new FakeNativeWindow();
            var dispatcher = new Control();
            Set(addIn, "nativeChatWindow", native);
            Set(addIn, "dispatcher", dispatcher);
            object[] custom = null;
            addIn.OnBeginShutdown(ref custom);
            Assert.AreEqual(1, native.CloseCount);
            Assert.IsTrue(dispatcher.IsDisposed);
            Assert.IsNull(Field<object>(addIn, "nativeChatWindow"));
            Assert.IsNull(Field<object>(addIn, "dispatcher"));
            addIn.OnDisconnection(0, ref custom);
            Assert.AreEqual(1, native.CloseCount);
        }

        /// <summary>Utilise la fenêtre principale du VBE comme propriétaire des dialogues.</summary>
        [TestMethod]
        public void AddInUsesTheVbeMainWindowAsDialogOwner()
        {
            var addIn = new AddIn();
            Set(addIn, "vbe", new FakeOwnerHost { MainWindow = new FakeMainWindow { HWnd = 12345 } });
            var owner = (IWin32Window)Call(addIn, "VbeOwner");
            Assert.AreEqual(new IntPtr(12345), owner.Handle);
            object[] custom = null;
            addIn.OnDisconnection(0, ref custom);
            Assert.IsNull(Field<object>(addIn, "vbe"));
        }

        /// <summary>Libère les ressources de l’add-in même si la fenêtre native refuse la fermeture.</summary>
        [TestMethod]
        public void AddInShutdownContinuesWhenTheNativeWindowRejectsClose()
        {
            var addIn = new AddIn();
            var dispatcher = new Control();
            Set(addIn, "dispatcher", dispatcher);
            Set(addIn, "nativeChatWindow", new RejectingNativeWindow());
            Set(addIn, "addIn", new object());
            object[] custom = null;
            addIn.OnBeginShutdown(ref custom);
            Assert.IsTrue(dispatcher.IsDisposed);
            Assert.IsNull(Field<object>(addIn, "nativeChatWindow"));
            Assert.IsNull(Field<object>(addIn, "addIn"));
        }
    }
}

namespace CodexVBE.Tests.Unit
{
    using System;
    using System.IO;
    using System.Linq;
    using System.Reflection;
    using System.Runtime.InteropServices;
    using System.Windows.Forms;
    using CodexVBE;
    using CodexVBE.Tests.Infrastructure;
    using Microsoft.VisualStudio.TestTools.UnitTesting;

    [TestClass, TestCategory("Unit")]
    public sealed partial class AddInCoverageTests
    {
        [STATestMethod]
        public void ReopeningAChatWithinALiveOrPartiallyReleasedNativeSiteKeepsAttachGuards()
        {
            foreach (bool missingControl in new[] { false, true }) using (var scope = new HostUiScope())
            {
                var instance = scope.Connected(); var control = LlmBoundaryScope.Get<ChatToolWindow>(instance, "nativeChatControl"); var previous = Chat(instance);
                try { previous.Dispose(); LlmBoundaryScope.Set(instance, "chat", null); if (missingControl) LlmBoundaryScope.Set(instance, "nativeChatControl", null); Call(instance, "ShowChat"); Assert.AreNotSame(previous, Chat(instance)); Assert.AreEqual(missingControl, Chat(instance).TopLevel); }
                finally { LlmBoundaryScope.Set(instance, "nativeChatControl", control); scope.Close(instance); }
            }
        }
        [STATestMethod]
        public void RealAssistantCanDockUndockRecoverDisposedSiteAndReopenAfterClosing()
        {
            using (var scope = new HostUiScope())
            {
                var instance = scope.Connected(); try
                {
                    var chat = Chat(instance); Assert.IsFalse(chat.TopLevel); Assert.IsTrue(LlmBoundaryScope.Get<bool>(instance, "docked")); Call(instance, "ShowChat"); Assert.IsTrue(scope.Host.Windows.Window.FocusCount >= 2);
                    ((Action)typeof(ChatWindow).GetField("DockRequested", BindingFlags.Instance | BindingFlags.NonPublic).GetValue(chat))(); Assert.IsTrue(chat.TopLevel); Assert.IsFalse(LlmBoundaryScope.Get<bool>(instance, "docked")); Call(instance, "ShowChat");
                    Call(instance, "ToggleDock"); Assert.IsFalse(chat.TopLevel); scope.Host.Windows.Control.Dispose(); Call(instance, "ShowChat"); Assert.AreNotSame(chat, Chat(instance)); Assert.IsTrue(Chat(instance).TopLevel); Assert.IsNull(LlmBoundaryScope.Get<object>(instance, "nativeChatControl"));
                    var closed = Chat(instance); closed.Close(); Assert.IsNull(Chat(instance)); Call(instance, "ShowChat"); Assert.AreNotSame(closed, Chat(instance));
                    Chat(instance).Hide(); scope.Host.MainWindow.RejectHandle = true; Call(instance, "ShowChat"); Assert.IsTrue(Chat(instance).Visible); Assert.IsTrue(scope.Logs.Any(x => x.Contains("owner unavailable"))); scope.Host.MainWindow.RejectHandle = false;
                }
                finally { scope.Close(instance); }
                Assert.IsNull(Chat(instance)); Assert.IsNull(LlmBoundaryScope.Get<object>(instance, "server"));
            }
        }
        [STATestMethod]
        public void MenuCallbacksOpenOwnedDialogsAndPrepareActualComposerActions()
        {
            using (var scope = new HostUiScope())
            {
                Action chatAction = null, settingsAction = null, githubAction = null; Action<string> editor = null;
                AddIn.CreateMenu = (host, chat, settings, github, action) => { chatAction = chat; settingsAction = settings; githubAction = github; editor = action; return new VbeMenu(host, chat, settings, github, action, (b, i, d, h) => { }, (b, i, d, h) => { }, (b, t) => { }); };
                var instance = scope.Connected(); try { chatAction(); settingsAction(); githubAction(); editor("/expliquer"); Assert.AreEqual(2, scope.Dialogs.Count); Assert.AreEqual(typeof(LlmSettingsWindow), scope.Dialogs[0]); Assert.AreEqual(typeof(GitWindow), scope.Dialogs[1]); Assert.AreEqual("/expliquer ", LlmBoundaryScope.Get<System.Windows.Controls.TextBox>(Chat(instance), "prompt").Text); }
                finally { scope.Close(instance); }
            }
        }
        [STATestMethod]
        public void StartupNativeBoundariesLogNullComInstancesAndPropagateBridgeFailures()
        {
            using (var scope = new HostUiScope())
            {
                var instance = new AddIn(); object[] custom = null; instance.OnConnection(scope.Host, 0, null, ref custom); scope.Close(instance); Assert.IsTrue(scope.Logs.Any(x => x.Contains("AddInInst: null")));
                object picture = null; using (var image = new System.Drawing.Bitmap(1, 1)) { var nested = typeof(VbeMenu).GetNestedType("MenuPicture", BindingFlags.NonPublic); picture = nested.GetMethod("ToOle", BindingFlags.NonPublic | BindingFlags.Static).Invoke(null, new object[] { image }); }
                Assert.IsTrue(Marshal.IsComObject(picture)); try { instance = scope.Connected(picture); scope.Close(instance); Assert.IsTrue(scope.Logs.Any(x => x.Contains("COM=True"))); } finally { Marshal.ReleaseComObject(picture); }
                AddIn.StartBridge = bridge => throw new IOException("bridge unavailable"); instance = new AddIn(); Assert.ThrowsException<IOException>(() => instance.OnConnection(scope.Host, 0, scope.Host.AddIns.AddIn, ref custom)); Assert.IsNull(LlmBoundaryScope.Get<object>(instance, "dispatcher")); Assert.IsNull(LlmBoundaryScope.Get<object>(instance, "server"));
            }
            using (var scope = new HostUiScope()) { AddIn.CreateChat = session => throw new IOException("UI unavailable"); var instance = scope.Connected(); try { Assert.IsTrue(scope.Logs.Any(x => x.Contains("Assistant window failed"))); Assert.IsNull(Chat(instance)); } finally { scope.Close(instance); } }
        }
        [STATestMethod]
        public void DockingComFailuresRestoreFloatingChatAndRetainEveryCleanupDefense()
        {
            foreach (var failure in new[] { "lookup", "creation", "missing-control", "position", "focus", "close", "null-chat", "missing-attached-control" })
                using (var scope = new HostUiScope())
                {
                    AddIn instance = null; scope.Host.AddIns.Reject = failure == "lookup"; scope.Host.Windows.RejectCreation = failure == "creation"; scope.Host.Windows.MissingControl = failure == "missing-control" || failure == "missing-attached-control"; scope.Host.MainWindow.LinkedWindows.Reject = failure == "position";
                    scope.Host.Windows.AfterCreation = () => { var native = scope.Host.Windows.Window; native.RejectClose = failure == "close"; if (failure == "focus" || failure == "close" || failure == "null-chat") native.Focusing = () => { if (failure == "null-chat") { var current = Chat(instance); current.Dispose(); LlmBoundaryScope.Set(instance, "chat", null); } throw new IOException("focus unavailable"); }; if (failure == "missing-attached-control") { var external = new ChatToolWindow(); native.Form.Controls.Add(external); var handle = external.Handle; external.Attach(Chat(instance)); } };
                    instance = new AddIn(); object[] custom = null; try { instance.OnConnection(scope.Host, 0, scope.Host.AddIns.AddIn, ref custom); Assert.IsNotNull(Chat(instance)); if (failure == "lookup" || failure == "position") Assert.IsTrue(LlmBoundaryScope.Get<bool>(instance, "docked")); else { Assert.IsFalse(LlmBoundaryScope.Get<bool>(instance, "docked")); Assert.IsTrue(scope.Logs.Any(x => x.Contains("Native chat docking failed"))); } }
                    finally { scope.Close(instance); }
                }
        }
        [STATestMethod]
        public void SettingsAndGitHubValidationReportErrorsWithoutAnyRealUserDialog()
        {
            using (var scope = new HostUiScope())
            {
                var instance = new AddIn(); LlmBoundaryScope.Set(instance, "vbe", scope.Host); AddIn.ReadSettings = LlmSettings.Load; Call(instance, "ShowSettings"); Assert.AreEqual(typeof(LlmSettingsWindow), scope.Dialogs.Single());
                scope.Host.ActiveVBProject = null; Call(instance, "ShowGitHub"); scope.Host.ActiveVBProject = scope.Host.Project; foreach (var path in new[] { null, " ", "relative.xlsm" }) { scope.Host.Project.FileName = path; Call(instance, "ShowGitHub"); }
                Assert.AreEqual(4, scope.Notices.Count);
                scope.Host.Project.FileName = Path.Combine(Path.GetTempPath(), "missing-selector-" + Guid.NewGuid() + ".xlsm"); scope.Host.VBProjects.Clear(); Call(instance, "ShowGitHub"); Assert.AreEqual(5, scope.Notices.Count);
                AddIn.ReadSettings = () => throw new IOException("settings unavailable"); Call(instance, "ShowSettings"); Assert.AreEqual("settings unavailable", scope.Notices.Last()); scope.Host.MainWindow.RejectHandle = true; Call(instance, "ShowSettings"); Assert.IsTrue(scope.Notices.Count >= 7); scope.Close(instance);
            }
            using (var scope = new HostUiScope()) { var instance = scope.Connected(); var disposed = Chat(instance); disposed.Dispose(); LlmBoundaryScope.Set(instance, "chat", disposed); Call(instance, "ShowSettings"); Assert.AreEqual(typeof(LlmSettingsWindow), scope.Dialogs.Single()); scope.Close(instance); }
        }
        [STATestMethod]
        public void ShutdownHandlesLiveDisposedMissingAndUndockedChatStatesIdempotently()
        {
            foreach (var state in new[] { "live", "disposed", "missing", "no-control", "undocked", "no-window", "reject-close" }) using (var scope = new HostUiScope())
            {
                var instance = scope.Connected(); var chat = Chat(instance); if (state == "disposed") { chat.Dispose(); LlmBoundaryScope.Set(instance, "chat", chat); }
                if (state == "missing") { chat.Dispose(); LlmBoundaryScope.Set(instance, "chat", null); }
                if (state == "no-control") LlmBoundaryScope.Set(instance, "nativeChatControl", null); if (state == "undocked") Call(instance, "ToggleDock"); if (state == "no-window") LlmBoundaryScope.Set(instance, "nativeChatWindow", null); if (state == "reject-close") scope.Host.Windows.Window.RejectClose = true; scope.Close(instance); scope.Close(instance); Assert.IsTrue(chat.IsDisposed); Assert.IsNull(LlmBoundaryScope.Get<object>(instance, "vbe"));
            }
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

    /// <summary>Vérifie les métadonnées COM et les callbacks sans connexion active.</summary>
    [TestClass]
    [TestCategory("Unit")]
    public sealed partial class HostSettingsWindowTests
    {
        /// <summary>Conserve des callbacks inoffensifs lorsque l’objet AddIn n’a pas été initialisé par le VBE.</summary>
        [TestMethod]
        public void AddInMetadataAndNoOpLifecycleCallbacksRemainSafeWithoutConnection()
        {
            Assert.AreEqual("CodexVBE.AddIn", ((ProgIdAttribute)Attribute.GetCustomAttribute(typeof(AddIn), typeof(ProgIdAttribute))).Value);
            var addin = (AddIn)FormatterServices.GetUninitializedObject(typeof(AddIn));
            object[] custom = new object[0];
            addin.OnAddInsUpdate(ref custom);
            addin.OnStartupComplete(ref custom);
            addin.OnBeginShutdown(ref custom);
        }
    }
}

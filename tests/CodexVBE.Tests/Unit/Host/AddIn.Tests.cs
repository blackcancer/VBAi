using System.Linq;
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

        /// <summary>Vérifie le cycle de vie du complément et ses callbacks vers l’interface.</summary>
[TestClass, TestCategory("Unit")]
    public sealed partial class AddInCoverageTests
    {
        /// <summary>Conserve les sites illisibles et les dimensions minimales, puis répare uniquement les sites inutilisables.</summary>
        [STATestMethod]
        public void PlacementChecksNativeSiteAvailabilityAndEachMinimumDimension()
        {
            using(var scope=new HostUiScope())
            {
                var instance=scope.Connected();try
                {
                    var control=LlmBoundaryScope.Get<ChatToolWindow>(instance,"nativeChatControl");var window=scope.Host.Windows.Window;var before=window.Form.Bounds;
                    try { LlmBoundaryScope.Set(instance,"nativeChatControl",null);Call(instance,"EnsureUsableChatPlacement");Assert.AreEqual(before,window.Form.Bounds); }
                    finally { LlmBoundaryScope.Set(instance,"nativeChatControl",control); }
                    var parent=control.ParentReader;var reader=control.ClientReader;
                    try
                    {
                        control.ParentReader=h=>IntPtr.Zero;Call(instance,"EnsureUsableChatPlacement");Assert.AreEqual(before,window.Form.Bounds);
                        control.ParentReader=parent;control.ClientReader=(IntPtr h,out ChatToolWindow.NativeRect rect)=>{rect=default(ChatToolWindow.NativeRect);return false;};
                        Call(instance,"EnsureUsableChatPlacement");Assert.AreEqual(before,window.Form.Bounds);
                        var minimum=Chat(instance).MinimumSize;
                        foreach(var size in new[] {minimum,new System.Drawing.Size(minimum.Width-1,minimum.Height),new System.Drawing.Size(minimum.Width,minimum.Height-1)})
                        {
                            control.ClientReader=(IntPtr h,out ChatToolWindow.NativeRect rect)=>{rect=new ChatToolWindow.NativeRect {Right=size.Width,Bottom=size.Height};return true;};
                            window.LinkedWindowFrame=null;Call(instance,"EnsureUsableChatPlacement");
                            if(size==minimum)Assert.AreEqual(before,window.Form.Bounds);
                            else {var area=Screen.FromHandle(new IntPtr(scope.Host.MainWindow.HWnd)).WorkingArea;Assert.AreEqual(Math.Min(600,area.Width),window.Width);Assert.AreEqual(Math.Min(820,area.Height),window.Height);Assert.IsTrue(scope.Logs.Any(x=>x.Contains("Recovered unusable chat pane")));}
                        }
                        Assert.AreEqual(0,scope.Host.MainWindow.LinkedWindows.Removes);
                    }
                    finally {control.ParentReader=parent;control.ClientReader=reader;}
                }
                finally {scope.Close(instance);}
            }
        }

        /// <summary>Les deux points d'arrêt nettoient les copies temporaires, conservent les autres et continuent après refus natif.</summary>
        [STATestMethod]
        public void HostShutdownCleansOnlyTemporaryCommandsAndLogsNativeFailureBeforeDisposal()
        {
            foreach(bool shutdown in new[] {false,true})
            foreach(bool failDelete in new[] {false,true})
            using(var scope=new HostUiScope())
            {
                var host=new ToolbarCustomizationTests.Host();var bar=host.CommandBars.Add("Standard",1,false,true);
                var temp=bar.Controls.Add(1,42,Type.Missing,1,true);temp.Tag="VBAi.ToolbarCommand."+new string('a',32);temp.FailDelete=failDelete;
                var persistent=bar.Controls.Add(1,42,Type.Missing,2,true);persistent.Tag="VBAi.ToolbarCommand.Persistent.owned";
                var foreign=bar.Controls.Add(1,42,Type.Missing,3,true);foreign.Tag="ThirdParty";
                var instance=new AddIn();LlmBoundaryScope.Set(instance,"vbe",host);var dispatcher=new Control();LlmBoundaryScope.Set(instance,"dispatcher",dispatcher);
                object[] custom=null;if(shutdown)instance.OnBeginShutdown(ref custom);else instance.OnDisconnection(0,ref custom);
                Assert.AreEqual(failDelete?3:2,bar.Controls.Count);Assert.IsTrue(bar.Controls.Contains(persistent));Assert.IsTrue(bar.Controls.Contains(foreign));
                Assert.AreEqual(failDelete,scope.Logs.Any(x=>x.Contains("Temporary toolbar cleanup failed: native delete rejected")));
                Assert.IsTrue(dispatcher.IsDisposed);Assert.IsNull(LlmBoundaryScope.Get<object>(instance,"vbe"));
            }
        }
                /// <summary>Vérifie les garde-fous d’attachement lors de la réouverture d’une fenêtre native partiellement libérée.</summary>
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
                /// <summary>Vérifie l’attachement, le détachement et la réouverture de la fenêtre assistant.</summary>
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
        public void StartupRepairsCollapsedOrDefaultTinyPaneWithoutForcingMainFrameDocking()
        {
            foreach (var size in new[] { new System.Drawing.Size(1920, 6), new System.Drawing.Size(200, 100), new System.Drawing.Size(321, 766) })
            using (var scope = new HostUiScope())
            {
                scope.Host.Windows.AfterCreation = () => {
                    scope.Host.Windows.Window.Form.ClientSize = size;
                    scope.Host.Windows.Window.LinkedWindowFrame = scope.Host.MainWindow;
                };
                var instance = scope.Connected();
                try
                {
                    var window = scope.Host.Windows.Window;
                    var area = Screen.FromHandle(new IntPtr(scope.Host.MainWindow.HWnd)).WorkingArea;
                    Assert.AreEqual(0, scope.Host.MainWindow.LinkedWindows.Adds);
                    Assert.AreEqual(1, scope.Host.MainWindow.LinkedWindows.Removes);
                    Assert.AreEqual(Math.Min(600, area.Width), window.Width);
                    Assert.AreEqual(Math.Min(820, area.Height), window.Height);
                    Assert.IsTrue(area.Contains(window.Form.Bounds));
                    Assert.IsTrue(window.Visible);
                    Assert.IsTrue(LlmBoundaryScope.Get<bool>(instance, "docked"));
                }
                finally { scope.Close(instance); }
            }
        }

        [STATestMethod]
        public void FailedPlacementRecoveryFallsBackToVisibleStandaloneChat()
        {
            using (var scope = new HostUiScope())
            {
                scope.Host.Windows.AfterCreation = () => {
                    scope.Host.Windows.Window.Form.ClientSize = new System.Drawing.Size(1920, 6);
                    scope.Host.Windows.Window.LinkedWindowFrame = scope.Host.MainWindow;
                    scope.Host.MainWindow.LinkedWindows.Reject = true;
                };
                var instance = scope.Connected();
                try
                {
                    Assert.IsFalse(LlmBoundaryScope.Get<bool>(instance, "docked"));
                    Assert.IsTrue(Chat(instance).TopLevel);
                    Assert.IsTrue(Chat(instance).Visible);
                    Assert.IsTrue(scope.Logs.Any(x => x.Contains("Native chat docking failed")));
                }
                finally { scope.Close(instance); }
            }
        }

        [STATestMethod]
        public void OpeningUsablePanePreservesLayoutAndReopeningRepairsOnlyCollapsedPane()
        {
            using (var scope = new HostUiScope())
            {
                var instance = scope.Connected();
                try
                {
                    var window = scope.Host.Windows.Window;
                    var bounds = window.Form.Bounds;
                    Call(instance, "ShowChat");
                    Assert.AreEqual(bounds, window.Form.Bounds);
                    Assert.AreEqual(0, scope.Host.MainWindow.LinkedWindows.Adds);
                    Assert.AreEqual(0, scope.Host.MainWindow.LinkedWindows.Removes);
                    window.LinkedWindowFrame = scope.Host.MainWindow;
                    window.Form.ClientSize = new System.Drawing.Size(1000, 6);
                    Call(instance, "ShowChat");
                    Assert.AreEqual(1, scope.Host.MainWindow.LinkedWindows.Removes);
                    Assert.IsTrue(window.Form.ClientSize.Height > 200);
                }
                finally { scope.Close(instance); }
            }
        }
                /// <summary>Vérifie les menus, les dialogues possédés et les actions de préparation du compositeur.</summary>
[STATestMethod]
        public void MenuCallbacksOpenOwnedDialogsAndPrepareActualComposerActions()
        {
            using (var scope = new HostUiScope())
            {
                Action chatAction = null, settingsAction = null, githubAction = null; Action<string> editor = null;
                AddIn.CreateMenu = (host, chat, settings, github, action) => { chatAction = chat; settingsAction = settings; githubAction = github; editor = action; return new VbeMenu(host, chat, settings, github, action, (b, i, d, h) => { }, (b, i, d, h) => { }, (b, t) => { }); };
                var instance = scope.Connected(); try { chatAction(); settingsAction(); System.IO.File.WriteAllText(scope.Host.Project.FileName, string.Empty); githubAction(); editor("/expliquer"); Assert.AreEqual(2, scope.Dialogs.Count); Assert.AreEqual(typeof(LlmSettingsWindow), scope.Dialogs[0]); Assert.AreEqual(typeof(GitWindow), scope.Dialogs[1]); Assert.AreEqual("/expliquer ", LlmBoundaryScope.Get<System.Windows.Controls.TextBox>(Chat(instance), "prompt").Text); }
                finally { scope.Close(instance); }
            }
        }
                /// <summary>Vérifie les frontières natives au démarrage, la journalisation des valeurs nulles et les erreurs du pont.</summary>
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
                /// <summary>Vérifie que les erreurs COM d’attachement restaurent la fenêtre flottante et conservent les nettoyages.</summary>
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
                /// <summary>Vérifie les erreurs de validation des paramètres et de GitHub sans ouvrir de dialogue réel.</summary>
[STATestMethod]
        public void SettingsAndGitHubValidationReportErrorsWithoutAnyRealUserDialog()
        {
            using (var scope = new HostUiScope())
            {
                var instance = new AddIn(); LlmBoundaryScope.Set(instance, "vbe", scope.Host); AddIn.ReadSettings = LlmSettings.Load; Call(instance, "ShowSettings"); Assert.AreEqual(typeof(LlmSettingsWindow), scope.Dialogs.Single());
                scope.Host.ActiveVBProject = null; Call(instance, "ShowGitHub"); scope.Host.ActiveVBProject = scope.Host.Project; foreach (var path in new[] { null, " ", "relative.xlsm" }) { scope.Host.Project.FileName = path; Call(instance, "ShowGitHub"); }
                Assert.AreEqual(4, scope.Notices.Count);
                scope.Host.Project.ThrowDirectoryNotFound = true; Call(instance, "ShowGitHub"); scope.Host.Project.ThrowDirectoryNotFound = false;
                StringAssert.Contains(scope.Notices.Last(), "Save the macro");
                scope.Host.Project.FileName = Path.Combine(Path.GetTempPath(), "missing-selector-" + Guid.NewGuid() + ".xlsm"); scope.Host.VBProjects.Clear(); Call(instance, "ShowGitHub"); Assert.AreEqual(6, scope.Notices.Count);
                AddIn.ReadSettings = () => throw new IOException("settings unavailable"); Call(instance, "ShowSettings"); Assert.AreEqual("settings unavailable", scope.Notices.Last()); scope.Host.MainWindow.RejectHandle = true; Call(instance, "ShowSettings"); Assert.IsTrue(scope.Notices.Count >= 7); scope.Close(instance);
            }
            using (var scope = new HostUiScope()) { var instance = scope.Connected(); var disposed = Chat(instance); disposed.Dispose(); LlmBoundaryScope.Set(instance, "chat", disposed); Call(instance, "ShowSettings"); Assert.AreEqual(typeof(LlmSettingsWindow), scope.Dialogs.Single()); scope.Close(instance); }
        }
                /// <summary>Vérifie l’arrêt idempotent avec une fenêtre absente, libérée, détachée ou encore active.</summary>
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

namespace CodexVBE.Tests.Unit
{
    using System;
    using System.IO;
    using System.Runtime.InteropServices;
    using CodexVBE;
    using CodexVBE.Tests.Infrastructure;
    using Microsoft.VisualStudio.TestTools.UnitTesting;
    public sealed partial class AddInCoverageTests
    {
        [STATestMethod]
        public void GitHubLaunchHandlesComAndUnavailableDocumentErrorsAtBothIdentityStages()
        {
            foreach (Exception error in new Exception[] { new DirectoryNotFoundException("missing document"), new COMException("unavailable project") })
            using (var scope = new HostUiScope())
            {
                var instance = new AddIn(); var host = new GitHubLaunchHost { ProjectError = error };
                LlmBoundaryScope.Set(instance, "vbe", host);
                Call(instance, "ShowGitHub"); Assert.AreEqual(1, scope.Notices.Count);
                StringAssert.Contains(scope.Notices.Last(), "Select a saved VBA project");
                host.ProjectError = null; host.Project = new GitHubUnavailableProject { PathError = error };
                Call(instance, "ShowGitHub"); Assert.AreEqual(2, scope.Notices.Count);
                StringAssert.Contains(scope.Notices.Last(), "Save the macro");
                Assert.AreEqual(0, scope.Dialogs.Count); scope.Close(instance);
            }
        }
    }
}

namespace CodexVBE.Tests.Unit
{
    using System;
    using System.IO;
    using System.Linq;
    using System.Reflection;
    using CodexVBE;
    using CodexVBE.Tests.Infrastructure;
    using Microsoft.VisualStudio.TestTools.UnitTesting;

    [TestClass, TestCategory("Unit"), DoNotParallelize]
    public sealed class AddInModernEditorTests
    {
        [STATestMethod]
        public void ActiveEditorFollowsOnlyCodePanesInDesignModeAndRetainsComponentIdentity()
        {
            using (var fixture = new AddInModernEditorFixture())
            {
                Assert.IsNull(LlmBoundaryScope.Call(fixture.Instance, "ActiveEditorModule", true));
                fixture.Scope.Host.ActiveWindow = new AddInEditorActiveWindow { Type = 1 };
                Assert.IsNull(LlmBoundaryScope.Call(fixture.Instance, "ActiveEditorModule", true));
                fixture.Scope.Host.ActiveWindow = new AddInEditorActiveWindow();
                Assert.IsNull(LlmBoundaryScope.Call(fixture.Instance, "ActiveEditorModule", true));
                Assert.IsNull(LlmBoundaryScope.Call(fixture.Instance, "ActiveEditorModule", false));
                var component = fixture.Active(1);
                Assert.IsNull(LlmBoundaryScope.Call(fixture.Instance, "ActiveEditorModule", true));
                var module = (IEditorModule)LlmBoundaryScope.Call(fixture.Instance, "ActiveEditorModule", false);
                Assert.AreEqual("P · Module1", module.Name);
                fixture.Scope.Host.Project.Mode = 2;
                Assert.AreEqual(module.Name, ((IEditorModule)LlmBoundaryScope.Call(fixture.Instance, "ActiveEditorModule", true)).Name);
            }
        }

        [STATestMethod]
        public void FloatingEditorCreationVisibilityRecreationAndDisposedSiteAreOwned()
        {
            using (var fixture = new AddInModernEditorFixture())
            {
                Assert.IsNull(fixture.Get(false));
                var editor = fixture.Get(); Assert.IsTrue(editor.Visible); Assert.AreSame(editor, fixture.Get(false));
                Assert.AreSame(editor, fixture.Get()); editor.Hide(); Assert.AreSame(editor, fixture.Get()); Assert.IsTrue(editor.Visible);
                editor.Dispose(); Assert.IsNull(fixture.Get(false)); var replacement = fixture.Get(); Assert.AreNotSame(editor, replacement);
                fixture.Call("ToggleEditorDock"); var control = fixture.Scope.Host.Windows.EditorControl;
                control.Dispose(); Assert.IsTrue(replacement.IsDisposed);
                var recovered = fixture.Get(); Assert.AreNotSame(replacement, recovered); Assert.IsTrue(recovered.TopLevel);
                Assert.IsNull(LlmBoundaryScope.Get<object>(fixture.Instance, "nativeEditorControl"));
                Assert.IsNull(LlmBoundaryScope.Get<object>(fixture.Instance, "nativeEditorWindow"));
                Assert.IsFalse(LlmBoundaryScope.Get<bool>(fixture.Instance, "editorDocked"));
            }
        }

        [STATestMethod]
        public void DockRequestedCreatesDistinctEditorSiteUndocksAndReusesItsLiveControl()
        {
            using (var fixture = new AddInModernEditorFixture(true))
            {
                var editor = fixture.Get();
                var dock = (Action)typeof(ModernEditorWindow).GetField("DockRequested", BindingFlags.Instance | BindingFlags.NonPublic).GetValue(editor);
                dock(); var window = fixture.Scope.Host.Windows.EditorWindow;
                Assert.IsFalse(editor.TopLevel); Assert.AreEqual(900, window.Width); Assert.AreEqual(650, window.Height); Assert.AreEqual(1, window.FocusCount);
                Assert.IsTrue(LlmBoundaryScope.Get<bool>(fixture.Instance, "editorDocked"));
                Assert.AreSame(editor, fixture.Get()); Assert.AreEqual(2, window.FocusCount);
                dock(); Assert.IsTrue(editor.TopLevel); Assert.IsFalse(window.Visible); Assert.IsTrue(editor.Visible);
                dock(); Assert.AreSame(window, fixture.Scope.Host.Windows.EditorWindow); Assert.IsFalse(editor.TopLevel);
                Assert.IsTrue(fixture.Scope.Host.Windows.Window.Visible, "The assistant and editor must retain distinct native sites.");
            }
        }

        [STATestMethod]
        public void DockedEditorRecreationAttachesToExistingControlBeforeFocus()
        {
            using (var fixture = new AddInModernEditorFixture())
            {
                var editor = fixture.Get(); fixture.Call("ToggleEditorDock"); var window = fixture.Scope.Host.Windows.EditorWindow;
                editor.Dispose(); var replacement = fixture.Get(); Assert.IsFalse(replacement.TopLevel); Assert.AreSame(window, fixture.Scope.Host.Windows.EditorWindow);
                Assert.IsTrue(fixture.Scope.Host.Windows.EditorControl.Controls.Contains(replacement));
                Assert.AreEqual(2, window.FocusCount);
            }
        }

        [STATestMethod]
        public void EditorDockFailuresReportOwnedNoticeAndNativeCloseRefusalDoesNotLeakForms()
        {
            foreach (int failure in new[] { 0, 1, 2 })
            using (var fixture = new AddInModernEditorFixture())
            {
                var editor = fixture.Get();
                fixture.Scope.Host.Windows.RejectCreation = failure == 0;
                fixture.Scope.Host.Windows.MissingControl = failure == 1;
                fixture.Scope.Host.Windows.AfterCreation = () => { if (failure == 2) fixture.Scope.Host.Windows.EditorWindow.Focusing = () => throw new IOException("owned focus failure"); };
                fixture.Call("ToggleEditorDock"); Assert.AreEqual(1, fixture.Scope.Notices.Count);
                StringAssert.Contains(fixture.Scope.Logs.Last(), "VBE menu action failed:");
                if (fixture.Scope.Host.Windows.EditorWindow != null) fixture.Scope.Host.Windows.EditorWindow.RejectClose = true;
                fixture.Scope.Close(fixture.Instance); Assert.IsTrue(editor.IsDisposed); Assert.IsNull(fixture.Get(false));
            }
        }

        [STATestMethod]
        public void MenuEditorCommandAndNavigationOpenMemoryModulesOrReportTheirRealFailure()
        {
            using (var fixture = new AddInModernEditorFixture())
            using (var module = new EditorFixture())
            {
                fixture.Call("ShowModernEditor"); Assert.AreEqual(1, fixture.Editors.Count);
                LlmBoundaryScope.Call(fixture.Instance, "OpenModernModule", module);
                Assert.AreEqual(1, fixture.Get(false).Documents.Count());
                Assert.AreSame(module, fixture.Get(false).Documents.Single().Module);
                LlmBoundaryScope.Call(fixture.Instance, "OpenModernModule", new object[] { null }); Assert.AreEqual(1, fixture.Scope.Notices.Count);
                fixture.Active(); fixture.Call("ShowModernEditor"); Assert.AreEqual(2, fixture.Scope.Notices.Count);
                LlmBoundaryScope.Set(fixture.Instance, "vbe", new object()); fixture.Call("ShowModernEditor"); Assert.AreEqual(3, fixture.Scope.Notices.Count);
            }
            using (var fixture = new AddInModernEditorFixture())
            {
                Action<string> action = null;
                AddIn.CreateMenu = (host, chat, settings, github, editor) => { action = editor; return new VbeMenu(host, chat, settings, github, editor, (b, i, d, h) => { }, (b, i, d, h) => { }, (b, t) => { }); };
                object[] custom = null; fixture.Instance.OnConnection(fixture.Scope.Host, 0, fixture.Scope.Host.AddIns.AddIn, ref custom);
                action("/editor"); Assert.IsNotNull(fixture.Get(false)); Assert.AreEqual(0, fixture.Scope.Notices.Count);
            }
        }

        [STATestMethod]
        public void ShutdownRetainsEachDockedEditorNullAndDisposedGuard()
        {
            foreach (int state in new[] { 0, 1, 2, 3, 4 })
            using (var fixture = new AddInModernEditorFixture())
            {
                var editor = fixture.Get(); fixture.Call("ToggleEditorDock");
                if (state == 0) LlmBoundaryScope.Set(fixture.Instance, "editorDocked", false);
                if (state == 1) LlmBoundaryScope.Set(fixture.Instance, "modernEditor", null);
                if (state == 2) editor.Dispose();
                if (state == 3) LlmBoundaryScope.Set(fixture.Instance, "nativeEditorControl", null);
                fixture.Scope.Close(fixture.Instance);
                Assert.IsNull(LlmBoundaryScope.Get<object>(fixture.Instance, "modernEditor")); Assert.IsNull(LlmBoundaryScope.Get<object>(fixture.Instance, "nativeEditorWindow"));
                Assert.IsFalse(LlmBoundaryScope.Get<bool>(fixture.Instance, "editorDocked"));
                if (state != 1) Assert.IsTrue(editor.IsDisposed);
            }
        }

        [STATestMethod]
        public void ConnectionLogsSettingsAndNativeThemeFailuresAndEnabledThemeWithoutRealPalette()
        {
            foreach (int outcome in new[] { 0, 1, 2 })
            using (var theme = new NativeThemeFixture())
            using (var palette = new NativePaletteSchedulerFixture())
            using (var scope = new HostUiScope())
            {
                VbeNativeTheme.ApplyNativeTheme = window => 1;
                VbeNativeTheme.CreatePalette = (host, window) => palette.Service;
                scope.Settings.NativeVbeDarkTheme = outcome == 1;
                if (outcome == 0) AddIn.ReadSettings = () => throw new IOException("owned settings failure");
                if (outcome == 2) scope.Host.MainWindow.RejectHandle = true;
                var instance = scope.Connected();
                try
                {
                    StringAssert.Contains(string.Join("\n", scope.Logs), outcome == 0 ? "Native VBE theme setting unavailable" : outcome == 1 ? "Native VBE dark mode enabled." : "Native VBE dark mode unavailable:");
                }
                finally { scope.Host.MainWindow.RejectHandle = false; scope.Close(instance); }
            }
        }
        [STATestMethod]
        public void StartupFailureBeforeReporterAndConnectedMenuErrorPreserveTheirCleanupContracts()
        {
            using (var fixture = new AddInModernEditorFixture())
            {
                AddIn.StartUpdateCheck = () => throw new IOException("owned updater failure");
                object[] custom = null;
                Assert.ThrowsException<IOException>(() => fixture.Instance.OnConnection(fixture.Scope.Host, 0, fixture.Scope.Host.AddIns.AddIn, ref custom));
                Assert.IsNull(LlmBoundaryScope.Get<object>(fixture.Instance, "dispatcher"));
                Assert.IsNull(LlmBoundaryScope.Get<object>(fixture.Instance, "crashReporter"));
                StringAssert.Contains(string.Join("\n", fixture.Scope.Logs), "OnConnection failed:");
            }
            using (var fixture = new AddInModernEditorFixture(true))
            {
                LlmBoundaryScope.Call(fixture.Instance, "ReportMenuError", new IOException("owned menu failure"));
                Assert.AreEqual("owned menu failure", fixture.Scope.Notices.Single());
                Assert.IsNotNull(LlmBoundaryScope.Get<CrashReporter>(fixture.Instance, "crashReporter"));
                object[] custom = new object[] { "owned payload" };
                var original = custom;
                fixture.Instance.OnAddInsUpdate(ref custom); fixture.Instance.OnStartupComplete(ref custom);
                Assert.AreSame(original, custom); Assert.AreEqual("owned payload", custom[0]);
            }
        }

        [STATestMethod]
        public void ThemeCleanupRefusalAndFailureAreLoggedWhileAllHostReferencesAreReleased()
        {
            using (var renderer = new NativeRendererStopFixture())
            using (var theme = new NativeThemeFixture())
            using (var fixture = new AddInModernEditorFixture())
            {
                renderer.Result = 1444;
                fixture.Scope.Close(fixture.Instance);
                StringAssert.Contains(string.Join("\n", fixture.Scope.Logs), "Native VBE theme cleanup deferred: renderer still active.");
                Assert.IsTrue(VbeNativeRenderer.Active); Assert.IsNull(LlmBoundaryScope.Get<object>(fixture.Instance, "vbe"));
                renderer.Result = 0;
            }
            using (var theme = new NativeThemeFixture())
            using (var fixture = new AddInModernEditorFixture())
            {
                theme.AddBrush(); VbeNativeTheme.ReleaseBrush = brush => throw new InvalidOperationException("owned brush release failure");
                fixture.Scope.Close(fixture.Instance);
                StringAssert.Contains(string.Join("\n", fixture.Scope.Logs), "Native VBE theme cleanup failed:");
                StringAssert.Contains(string.Join("\n", fixture.Scope.Logs), "owned brush release failure");
                Assert.IsNull(LlmBoundaryScope.Get<object>(fixture.Instance, "vbe"));
            }
        }
    }
}
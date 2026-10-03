namespace VBAi.Tests.Unit
{
    using System;
    using System.Collections.Generic;
    using System.IO;
    using System.IO.Pipes;
    using System.Text;
    using System.Threading;
    using System.Threading.Tasks;
    using System.Web.Script.Serialization;
    using System.Windows.Forms;
    using VBAi;
    using Microsoft.VisualStudio.TestTools.UnitTesting;

    /// <summary>Vérifie le transport named-pipe du pont et la validation des commandes Immediate.</summary>
    [TestClass]
    [TestCategory("Unit")]
    public sealed partial class BridgeServerTests
    {
        [STATestMethod]
        public void PendingGeneralBlocksBridgeNativeRoutesBeforeAnyEntry()
        {
            AssertGeneralBlocksBridgeNativeRoutes("generalInFlight");
        }

        [STATestMethod]
        public void UncertainGeneralBlocksBridgeNativeRoutesBeforeAnyEntry()
        {
            AssertGeneralBlocksBridgeNativeRoutes("generalQuarantined");
        }

        private static void AssertGeneralBlocksBridgeNativeRoutes(string stateField)
        {
            using (var dispatcher = new Control())
            {
                var handle = dispatcher.Handle;
                var session = new VbeSession(new FakeVbe());
                int id = Guid.NewGuid().GetHashCode() & int.MaxValue;
                using (var server = new BridgeServer(dispatcher, session, id))
                {
                    int nativeEntries = 0, asyncEntries = 0, executeEntries = 0;
                    Func<object> entered = () => { nativeEntries++; return new { Native = true }; };
                    server.Native.Capture = _ => entered();
                    server.Native.ReadNavigationSurface = _ => entered();
                    server.Native.ChangeNavigationSurface = _ => entered();
                    server.Native.ListObjectBrowser = _ => entered();
                    server.Native.SelectObjectBrowser = _ => entered();
                    server.Native.ReadRuntimeForms = entered;
                    server.Native.ReadObjectBrowser = entered;
                    server.Native.ReadDebugDialog = entered;
                    server.Native.ChangeDebugItem = _ => entered();
                    server.Native.RespondDebugDialog = _ => entered();
                    server.Native.EnsureNoCompileDialog = () => { nativeEntries++; };
                    server.Native.SelectWatch = _ => entered();
                    server.Native.EnsureNoDebugOptionsDialog = () => { nativeEntries++; };
                    server.Native.EnsureNoProjectPropertiesDialog = () => { nativeEntries++; };
                    server.Native.EnsureNoSignatureDialog = () => { nativeEntries++; };
                    server.Native.ExecuteImmediate = (_, __) => entered();
                    server.ReadImmediateNative = _ => { asyncEntries++; return Task.FromResult<object>(null); };
                    server.InspectLocalScalarsNative = server.ReadImmediateNative;
                    server.SaveHostDocumentNative = server.ReadImmediateNative;
                    server.ProjectGeneralNative = (_, __) => { asyncEntries++; return Task.FromResult<object>(null); };
                    server.Execute = request => { executeEntries++; return session.Execute(request); };
                    server.Start();
                    Assert.AreEqual(true, SendWithMessagePump(id, "{\"Command\":\"debug_windows\"}")["Ok"]);
                    Assert.AreEqual(1, nativeEntries, "A settled session must retain the native route.");
                    nativeEntries = 0;
                    typeof(VbeSession).GetField(stateField, System.Reflection.BindingFlags.Instance | System.Reflection.BindingFlags.NonPublic).SetValue(session, true);
                    foreach (string command in new[] {
                        "debug_windows", "read_navigation_surface", "change_navigation_surface", "list_object_browser",
                        "select_object_browser", "read_runtime_forms", "read_object_browser", "debug_dialog", "debug_item",
                        "respond_debug_dialog", "immediate_execute", "compile_project", "add_watch", "edit_watch", "quick_watch",
                        "remove_watch", "read_debug_options", "read_vbe_options", "set_vbe_option", "read_project_protection",
                        "set_project_protection", "read_project_signature_dialog", "sign_project", "read_immediate",
                        "inspect_local_scalars", "save_host_document", "read_project_general", "set_project_general",
                        "list_projects", "diagnostic_path_visibility" })
                    {
                        var result = SendWithMessagePump(id, "{\"Command\":\"" + command + "\"}");
                        Assert.AreEqual(false, result["Ok"], command);
                        StringAssert.Contains((string)result["Error"], "pending or uncertain", command);
                    }
                    Assert.AreEqual(0, nativeEntries);
                    Assert.AreEqual(0, asyncEntries);
                    Assert.AreEqual(0, executeEntries, "Refusal must precede every native or session dispatch.");
                    var status = SendWithMessagePump(id, "{\"Command\":\"status\"}");
                    Assert.AreEqual(true, status["Ok"]);
                    Assert.AreEqual(true, ((IDictionary<string, object>)status["Data"])["Connected"]);
                    Assert.AreEqual(1, executeEntries, "Only managed status remains available.");
                }
            }
        }

        [STATestMethod]
        public void NativeBridgeAdmissionExcludesGeneralUntilWorkerReturnsOrThrows()
        {
            using (var dispatcher = new Control())
            {
                var handle = dispatcher.Handle;
                int owner = Thread.CurrentThread.ManagedThreadId;
                var session = new VbeSession(new FakeVbe());
                int id = Guid.NewGuid().GetHashCode() & int.MaxValue;
                using (var server = new BridgeServer(dispatcher, session, id))
                {
                    int entries = 0;
                    bool fail = false;
                    server.Native.Capture = _ => {
                        entries++;
                        Assert.AreNotEqual(owner, Thread.CurrentThread.ManagedThreadId,
                            "Native accessibility work must retain its existing worker thread.");
                        dispatcher.Invoke(new Action(() => {
                            var error = Assert.ThrowsException<InvalidOperationException>(() =>
                                session.ProjectGeneralAsync(null, false).GetAwaiter().GetResult());
                            StringAssert.Contains(error.Message, "bridge operation is pending");
                        }));
                        if (fail) throw new InvalidOperationException("native read failed after admission");
                        return new { Native = true };
                    };
                    server.Start();
                    foreach (bool throwFromNative in new[] { false, true })
                    {
                        fail = throwFromNative;
                        var result = SendWithMessagePump(id, "{\"Command\":\"debug_windows\"}");
                        Assert.AreEqual(!throwFromNative, result["Ok"]);
                        if (throwFromNative) StringAssert.Contains((string)result["Error"], "native read failed after admission");
                        // Admission is released on both paths: validation now reaches
                        // the ordinary missing-request guard, without touching COM.
                        Assert.ThrowsException<ArgumentException>(() =>
                            session.ProjectGeneralAsync(null, false).GetAwaiter().GetResult());
                    }
                    Assert.AreEqual(2, entries);
                }
            }
        }

        [STATestMethod]
        public void GeneralDispatchRetainsOwnerAuthorizationAcrossAwaitWithoutRetry()
        {
            using (var dispatcher = new Control())
            {
                var handle = dispatcher.Handle;
                int owner = Thread.CurrentThread.ManagedThreadId;
                int id = Guid.NewGuid().GetHashCode() & int.MaxValue;
                using (var server = new BridgeServer(dispatcher, null, id))
                {
                    int calls = 0;
                    Request retained = null;
                    server.ProjectGeneralNative = async (request, write) => {
                        retained = request;
                        calls++;
                        Assert.AreEqual(owner, Thread.CurrentThread.ManagedThreadId);
                        Assert.AreEqual(write, request.Command == "set_project_general");
                        Assert.IsNotNull(request.RevalidateProjectPropertyAuthorization);
                        request.RevalidateProjectPropertyAuthorization(true);
                        await Task.Yield();
                        Assert.AreEqual(owner, Thread.CurrentThread.ManagedThreadId);
                        request.RevalidateProjectPropertyAuthorization(false);
                        return new { Uncertain = write, MutationInvoked = write, RetryAllowed = false };
                    };
                    server.Execute = _ => { Assert.Fail("General must not use synchronous dispatch."); return null; };
                    server.Start();
                    foreach (string command in new[] { "read_project_general", "set_project_general" })
                    {
                        var result = SendWithMessagePump(id, "{\"Command\":\"" + command + "\",\"Project\":\"P\",\"ExpectedMode\":2,\"ExpectedProjectVersion\":\"v\"}");
                        Assert.AreEqual(true, result["Ok"]);
                        Assert.AreEqual(command == "set_project_general", ((IDictionary<string, object>)result["Data"])["Uncertain"]);
                        Assert.IsNull(retained.RevalidateProjectPropertyAuthorization);
                    }
                    Assert.AreEqual(2, calls);
                }
            }
        }

        [STATestMethod]
        public void GeneralBridgeRefusesChangedOriginalRequestBeforeNativeMutation()
        {
            using (var dispatcher = new Control())
            {
                var handle = dispatcher.Handle;
                int id = Guid.NewGuid().GetHashCode() & int.MaxValue;
                using (var server = new BridgeServer(dispatcher, null, id))
                {
                    int calls = 0;
                    Request retained = null;
                    server.ProjectGeneralNative = async (request, write) => {
                        retained = request; calls++;
                        await Task.Yield();
                        request.ControlCaption = "Another command";
                        request.RevalidateProjectPropertyAuthorization(false);
                        Assert.Fail("Changed native command caption must be refused.");
                        return null;
                    };
                    server.Start();
                    var result = SendWithMessagePump(id, "{\"Command\":\"set_project_general\",\"Project\":\"P\",\"ExpectedMode\":2,\"ControlCaption\":\"Original\"}");
                    Assert.AreEqual(false, result["Ok"]);
                    Assert.AreEqual(1, calls);
                    Assert.IsNull(retained.RevalidateProjectPropertyAuthorization);
                }
            }
        }

        [STATestMethod]
        public void HostSaveDispatchRunsOnOwnerThreadAndReturnsUncertainResultWithoutRetry()
        {
            using (var dispatcher = new Control())
            {
                var handle = dispatcher.Handle;
                int ownerThread = Thread.CurrentThread.ManagedThreadId;
                int id = Guid.NewGuid().GetHashCode() & int.MaxValue;
                using (var server = new BridgeServer(dispatcher, null, id))
                {
                    int saves = 0;
                    server.SaveHostDocumentNative = async request => {
                        Assert.AreEqual(ownerThread, Thread.CurrentThread.ManagedThreadId);
                        Assert.AreEqual("save_host_document", request.Command);
                        saves++;
                        await Task.Yield();
                        Assert.AreEqual(ownerThread, Thread.CurrentThread.ManagedThreadId);
                        Assert.AreEqual(ApartmentState.STA, Thread.CurrentThread.GetApartmentState());
                        return new { SaveInvoked = true, Verified = false, Uncertain = true };
                    };
                    server.Start();
                    var result = SendWithMessagePump(id, "{\"Command\":\"save_host_document\",\"Project\":\"P\",\"ExpectedProjectVersion\":\"version\",\"ExpectedHostPath\":\"C:\\\\fixture\\\\Owned.swp\"}");
                    Assert.AreEqual(true, result["Ok"]);
                    var data = (IDictionary<string, object>)result["Data"];
                    Assert.AreEqual(true, data["Uncertain"]); Assert.AreEqual(false, data["Verified"]);
                    Assert.AreEqual(1, saves);
                }
            }
        }

        [STATestMethod]
        public void ImmediateCopyYieldsOnOwningUiThreadAndReportsAsyncFailure()
        {
            using (var dispatcher = new Control())
            {
                var handle = dispatcher.Handle;
                int thread = Thread.CurrentThread.ManagedThreadId;
                int id = Guid.NewGuid().GetHashCode() & int.MaxValue;
                using (var server = new BridgeServer(dispatcher, null, id))
                {
                    server.ReadImmediateNative = async request => {
                        Assert.AreEqual("P", request.Project);
                        Assert.AreEqual(thread, Thread.CurrentThread.ManagedThreadId);
                        Assert.AreEqual(ApartmentState.STA, Thread.CurrentThread.GetApartmentState());
                        await Task.Yield();
                        Assert.AreEqual(thread, Thread.CurrentThread.ManagedThreadId);
                        if (request.ExpectedMode == 1) throw new InvalidOperationException("capture rejected");
                        return new { Text = "native output", ClipboardRestored = true };
                    };
                    server.Execute = _ => { Assert.Fail("Capture must use the asynchronous boundary."); return null; };
                    server.Start();
                    var rejected = SendWithMessagePump(id, "{\"Command\":\"read_immediate\",\"Project\":\"P\",\"ExpectedMode\":1}");
                    Assert.AreEqual(false, rejected["Ok"]);
                    Assert.AreEqual("capture rejected", rejected["Error"]);
                    var read = SendWithMessagePump(id, "{\"Command\":\"read_immediate\",\"Project\":\"P\",\"ExpectedMode\":2}");
                    Assert.AreEqual(true, read["Ok"]);
                    Assert.AreEqual("native output", ((IDictionary<string, object>)read["Data"])["Text"]);
                }
            }
        }

        [STATestMethod]
        public void LocalScalarInspectionUsesAsyncUiBoundary()
        {
            using (var dispatcher = new Control())
            {
                var handle = dispatcher.Handle;
                int thread = Thread.CurrentThread.ManagedThreadId;
                int id = Guid.NewGuid().GetHashCode() & int.MaxValue;
                using (var server = new BridgeServer(dispatcher, null, id))
                {
                    server.InspectLocalScalarsNative = async request => {
                        Assert.AreEqual("P", request.Project);
                        Assert.AreEqual("M", request.Module);
                        Assert.AreEqual("Run", request.Procedure);
                        Assert.AreEqual(new string('a', 64), request.ExpectedSha256);
                        Assert.AreEqual(1, request.ExpectedMode);
                        Assert.AreEqual(thread, Thread.CurrentThread.ManagedThreadId);
                        await Task.Yield();
                        Assert.AreEqual(thread, Thread.CurrentThread.ManagedThreadId);
                        if (request.Offset == 1) throw new InvalidOperationException("inspection rejected");
                        return new { Partial = true, Value = "41" };
                    };
                    server.Execute = _ => { Assert.Fail("Inspection must use the asynchronous boundary."); return null; };
                    server.Start();
                    const string prefix = "{\"Command\":\"inspect_local_scalars\",\"Project\":\"P\",\"Module\":\"M\",\"Procedure\":\"Run\",\"ExpectedSha256\":\"aaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaa\",\"ExpectedMode\":1,";
                    var rejected = SendWithMessagePump(id, prefix + "\"Offset\":1}");
                    Assert.AreEqual(false, rejected["Ok"]);
                    Assert.AreEqual("inspection rejected", rejected["Error"]);
                    var inspected = SendWithMessagePump(id, prefix + "\"Offset\":0}");
                    Assert.AreEqual(true, inspected["Ok"]);
                    Assert.AreEqual("41", ((IDictionary<string, object>)inspected["Data"])["Value"]);
                }
            }
        }

        /// <summary>Vérifie les connexions successives, les erreurs de requête et la reprise après une erreur.</summary>
        [TestMethod]
        [STATestMethod]
        public void PipeProcessesSuccessValidationAndMalformedJsonAcrossConnections()
        {
            using (var dispatcher = new Control())
            {
                var handle = dispatcher.Handle;
                var vbe = new FakeVbe();
                vbe.VBProjects.Add(new FakeProject { Name = "Disposable", FileName = @"C:\Temp\Disposable.xlsm" });
                int id = Guid.NewGuid().GetHashCode() & int.MaxValue;
                using (var server = new BridgeServer(dispatcher, new VbeSession(vbe), id))
                {
                    server.Start();
                    var status = SendWithMessagePump(id, "{\"Command\":\"status\"}");
                    Assert.AreEqual(true, status["Ok"]);
                    Assert.AreEqual(true, ((IDictionary<string, object>)status["Data"])["Connected"]);
                    var projects = SendWithMessagePump(id, "{\"Command\":\"list_projects\"}");
                    Assert.AreEqual(true, projects["Ok"]);
                    var first = ((object[])projects["Data"])[0] as IDictionary<string, object>;
                    Assert.IsNotNull(first);
                    Assert.AreEqual("Disposable", first["Name"]);
                    var missing = SendWithMessagePump(id, "{}");
                    Assert.AreEqual(false, missing["Ok"]);
                    StringAssert.Contains((string)missing["Error"], "command is required");
                    var malformed = SendWithMessagePump(id, "{invalid json}");
                    Assert.AreEqual(false, malformed["Ok"]);
                    Assert.IsFalse(string.IsNullOrWhiteSpace((string)malformed["Error"]));
                    var unknown = SendWithMessagePump(id, "{\"Command\":\"unknown_command\"}");
                    Assert.AreEqual(false, unknown["Ok"]);
                    StringAssert.Contains((string)unknown["Error"], "Unknown command");
                    var hostError = SendWithMessagePump(id, "{\"Command\":\"list_modules\",\"Project\":\"Disposable\"}");
                    Assert.AreEqual(false, hostError["Ok"]);
                    Assert.IsFalse(string.IsNullOrWhiteSpace((string)hostError["Error"]));
                    var recovered = SendWithMessagePump(id, "{\"Command\":\"status\"}");
                    Assert.AreEqual(true, recovered["Ok"]);
                    var badTree = SendWithMessagePump(id, "{\"Command\":\"debug_item\",\"Pane\":\"locals\",\"Action\":\"expand\"}");
                    Assert.AreEqual(false, badTree["Ok"]);
                    Assert.IsFalse(string.IsNullOrWhiteSpace((string)badTree["Error"]));
                    var badDialog = SendWithMessagePump(id, "{\"Command\":\"respond_debug_dialog\",\"Diagnostic\":\" \"}");
                    Assert.AreEqual(false, badDialog["Ok"]);
                    Assert.IsFalse(string.IsNullOrWhiteSpace((string)badDialog["Error"]));
                }
            }
        }

        /// <summary>Refuse une exécution Immediate sans projet avant tout accès à une fenêtre native.</summary>
        [TestMethod]
        [STATestMethod]
        public void ImmediateCommandRejectsMissingProjectBeforeTouchingNativeWindow()
        {
            using (var dispatcher = new Control())
            {
                var handle = dispatcher.Handle;
                int id = Guid.NewGuid().GetHashCode() & int.MaxValue;
                using (var server = new BridgeServer(dispatcher, new VbeSession(new FakeVbe()), id))
                {
                    server.Start();
                    var result = SendWithMessagePump(id, "{\"Command\":\"immediate_execute\",\"ExpectedMode\":2,\"Text\":\"Debug.Print 1\"}");
                    Assert.AreEqual(false, result["Ok"]);
                    StringAssert.Contains((string)result["Error"], "Project and ExpectedMode");
                }
            }
        }

        [TestMethod]
        [STATestMethod]
        public void ImmediateCommandRequiresSelectedProjectPathBeforeNativeExecution()
        {
            using (var dispatcher = new Control())
            {
                var handle = dispatcher.Handle;
                int id = Guid.NewGuid().GetHashCode() & int.MaxValue;
                using (var server = new BridgeServer(dispatcher, null, id))
                {
                    var state = new Infrastructure.VbeToolMode { Mode = 2, Project = "SameName", SelectedProject = "SameName",
                        SelectedProjectPath = @"C:\Temp\B.xlsm", ActiveModule = "Module1" };
                    int ownerThread = System.Threading.Thread.CurrentThread.ManagedThreadId;
                    server.Execute = request => {
                        Assert.AreEqual(ownerThread, System.Threading.Thread.CurrentThread.ManagedThreadId);
                        return Response.Success(state);
                    };
                    int executions = 0;
                    server.Native.ExecuteImmediate = (text, submit) => { submit(() => executions++); return new { Executed = true }; };
                    server.Start();
                    string requestJson = new JavaScriptSerializer().Serialize(new {
                        Command = "immediate_execute", Project = @"C:\Temp\A.xlsm", ExpectedMode = 2, Text = "Debug.Print 1" });
                    var refused = SendWithMessagePump(id, requestJson);
                    Assert.AreEqual(false, refused["Ok"]);
                    StringAssert.Contains((string)refused["Error"], "requested project must be active");
                    Assert.AreEqual(0, executions);
                    state.SelectedProjectPath = @"C:\Temp\A.xlsm";
                    Assert.AreEqual(true, SendWithMessagePump(id, requestJson)["Ok"]);
                    Assert.AreEqual(1, executions);
                    server.Native.ExecuteImmediate = (text, submit) => {
                        state.Mode = 1; // Native typing/echo completed while the approved mode changed.
                        submit(() => executions++);
                        return new { Executed = true };
                    };
                    var changed = SendWithMessagePump(id, requestJson);
                    Assert.AreEqual(false, changed["Ok"]);
                    StringAssert.Contains((string)changed["Error"], "mode changed");
                    Assert.AreEqual(1, executions);
                }
            }
        }

        /// <summary>Refuse un mode invalide, un projet absent et un mode modifié avant l’exécution.</summary>
        [TestMethod]
        [STATestMethod]
        public void ImmediateCommandRejectsInvalidModeMissingProjectAndChangedModeWithoutExecuting()
        {
            using (var dispatcher = new Control())
            {
                var handle = dispatcher.Handle;
                var vbe = new FakeVbe();
                vbe.VBProjects.Add(new FakeProject { Name = "Disposable", FileName = @"C:\Temp\Disposable.xlsm", Mode = 2 });
                int id = Guid.NewGuid().GetHashCode() & int.MaxValue;
                using (var server = new BridgeServer(dispatcher, new VbeSession(vbe), id))
                {
                    server.Start();
                    var invalidMode = SendWithMessagePump(id, "{\"Command\":\"immediate_execute\",\"Project\":\"Disposable\",\"ExpectedMode\":0}");
                    Assert.AreEqual(false, invalidMode["Ok"]);
                    StringAssert.Contains((string)invalidMode["Error"], "ExpectedMode (1 or 2)");
                    var missingProject = SendWithMessagePump(id, "{\"Command\":\"immediate_execute\",\"Project\":\"Absent\",\"ExpectedMode\":1}");
                    Assert.AreEqual(false, missingProject["Ok"]);
                    StringAssert.Contains((string)missingProject["Error"], "Project selector is absent or ambiguous");
                    var changedMode = SendWithMessagePump(id, "{\"Command\":\"immediate_execute\",\"Project\":\"Disposable\",\"ExpectedMode\":1,\"Text\":\"Debug.Print 1\"}");
                    Assert.AreEqual(false, changedMode["Ok"]);
                    StringAssert.Contains((string)changedMode["Error"], "Project mode changed");
                    var recovered = SendWithMessagePump(id, "{\"Command\":\"status\"}");
                    Assert.AreEqual(true, recovered["Ok"]);
                }
            }
        }
        [TestMethod]
                /// <summary>Vérifie que le dispatch natif conserve les résultats de requête et les échecs renvoyés par l’hôte.</summary>
[STATestMethod]
        public void NativeDispatchMatrixPreservesRequestResultsAndHostFailures()
        {
            using (var dispatcher = new Control())
            {
                var handle = dispatcher.Handle;
                int id = Guid.NewGuid().GetHashCode() & int.MaxValue;
                using (var server = new BridgeServer(dispatcher, null, id))
                {
                    Infrastructure.VbeToolBoundaryFixture.Configure(server.Native);
                    server.Execute = Infrastructure.VbeToolBoundaryFixture.Execute;
                    server.PersistSignature = project => new VBAi.Tests.Infrastructure.VbeToolPersistence { Saved = true };
                    server.Start();
                    var json = new JavaScriptSerializer();
                    foreach (string command in new[] { "debug_windows", "debug_dialog", "debug_item", "respond_debug_dialog",
                        "immediate_execute", "add_watch", "edit_watch", "quick_watch", "read_debug_options", "read_vbe_options",
                        "read_project_signature_dialog", "remove_watch", "status", "read_navigation_surface", "change_navigation_surface", "read_project_protection", "set_project_protection" })
                    {
                        var response = SendWithMessagePump(id, json.Serialize(new { Command = command, Project = "P", ExpectedMode = 2, Text = "Debug.Print 1" }));
                        Assert.AreEqual(true, response["Ok"], command);
                        Assert.IsNotNull(response["Data"], command);
                    }
                    foreach (string command in new[] { "add_watch", "edit_watch", "quick_watch", "read_debug_options", "read_vbe_options",
                        "read_project_signature_dialog", "remove_watch", "sign_project", "immediate_execute", "read_project_protection", "set_project_protection" })
                    {
                        server.Execute = request => Response.Failure("host rejected " + request.Command);
                        var response = SendWithMessagePump(id, json.Serialize(new { Command = command, Project = "P", ExpectedMode = 2 }));
                        Assert.AreEqual(false, response["Ok"], command);
                        StringAssert.Contains((string)response["Error"], "host rejected");
                    }
                    server.Execute = Infrastructure.VbeToolBoundaryFixture.Execute;
                    Assert.AreEqual(false, SendWithMessagePump(id, "{\"Command\":\"immediate_execute\",\"Project\":\"P\",\"ExpectedMode\":1}")["Ok"]);
                    Assert.AreEqual(false, SendWithMessagePump(id, "{\"Command\":\"immediate_execute\",\"Project\":\"P\",\"ExpectedMode\":3}")["Ok"]);
                    Assert.AreEqual(false, SendWithMessagePump(id, "null")["Ok"]);
                    server.Native.ReadDebugDialog = () => { throw new InvalidOperationException("native unavailable"); };
                    var unavailable = SendWithMessagePump(id, "{\"Command\":\"debug_dialog\"}");
                    Assert.AreEqual(false, unavailable["Ok"]);
                    Assert.AreEqual("native unavailable", unavailable["Error"]);
                }
            }
        }

        [TestMethod]
                /// <summary>Vérifie les diagnostics, erreurs hôte, exceptions et callbacks retardés de la compilation.</summary>
[STATestMethod]
        public void CompileMatrixReportsDiagnosticsHostFailuresExceptionsAndDelayedCallbacks()
        {
            using (var dispatcher = new Control())
            {
                var handle = dispatcher.Handle;
                int id = Guid.NewGuid().GetHashCode() & int.MaxValue;
                using (var server = new BridgeServer(dispatcher, null, id))
                {
                    Infrastructure.VbeToolBoundaryFixture.Configure(server.Native);
                    server.Execute = Infrastructure.VbeToolBoundaryFixture.Execute;
                    server.Start();
                    const string request = "{\"Command\":\"compile_project\",\"Project\":\"P\"}";
                    var success = SendWithMessagePump(id, request);
                    Assert.AreEqual(true, ((IDictionary<string, object>)success["Data"])["Compiled"]);
                    server.Native.AwaitCompileDialog = completed => { Assert.IsTrue(completed.Wait(5000)); return "Syntax error"; };
                    var diagnostic = (IDictionary<string, object>)SendWithMessagePump(id, request)["Data"];
                    Assert.AreEqual(false, diagnostic["Compiled"]);
                    Assert.AreEqual("NativeDiagnosticCaptured", diagnostic["Verification"]);
                    Assert.IsNotNull(diagnostic["NextRead"]);
                    server.Execute = r => Response.Failure("compile declined");
                    Assert.AreEqual("compile declined", SendWithMessagePump(id, request)["Error"]);
                    server.Execute = r => { throw new InvalidOperationException("compile threw"); };
                    Assert.AreEqual("compile threw", SendWithMessagePump(id, request)["Error"]);
                    server.Native.AwaitCompileDialog = completed => null;
                    // Holding the UI callback in the message queue models the native timeout.
                    var delayed = SendWithoutMessagePump(id, request);
                    StringAssert.Contains((string)delayed["Error"], "did not return");
                    Application.DoEvents();
                }
            }
        }

        [TestMethod]
                /// <summary>Vérifie le statut de sauvegarde de signature et les nouvelles tentatives bornées en cas d’occupation.</summary>
[STATestMethod]
        public void SignatureMatrixRetainsSaveStatusAndBoundedBusyRetries()
        {
            using (var dispatcher = new Control())
            {
                var handle = dispatcher.Handle;
                int id = Guid.NewGuid().GetHashCode() & int.MaxValue;
                using (var server = new BridgeServer(dispatcher, null, id))
                {
                    Infrastructure.VbeToolBoundaryFixture.Configure(server.Native);
                    server.Execute = Infrastructure.VbeToolBoundaryFixture.Execute;
                    server.Start();
                    const string request = "{\"Command\":\"sign_project\",\"Project\":\"P\",\"CertificateThumbprint\":\"fixture\"}";
                    foreach (bool saved in new[] { true, false })
                    {
                        server.PersistSignature = p => new VBAi.Tests.Infrastructure.VbeToolPersistence { Saved = saved };
                        var result = (IDictionary<string, object>)SendWithMessagePump(id, request)["Data"];
                        Assert.AreEqual(!saved, result["SaveRequired"]);
                        Assert.IsNull(result["PersistenceError"]);
                        Assert.IsNotNull(result["HostStatus"]);
                    }
                    server.PersistSignature = p => null;
                    Assert.AreEqual(true, ((IDictionary<string, object>)SendWithMessagePump(id, request)["Data"])["SaveRequired"]);
                    server.PersistSignature = p => { throw new InvalidOperationException("save refused"); };
                    server.Execute = r => r.Command == "project_signature_status" ? Response.Failure("status refused") : Infrastructure.VbeToolBoundaryFixture.Execute(r);
                    var failure = (IDictionary<string, object>)SendWithMessagePump(id, request)["Data"];
                    Assert.AreEqual("save refused", failure["PersistenceError"]);
                    Assert.AreEqual("status refused", failure["HostStatusError"]);
                    Assert.IsNull(failure["HostStatus"]);
                    int attempts = 0;
                    server.PersistSignature = p => { if (++attempts == 2) return new VBAi.Tests.Infrastructure.VbeToolPersistence { Saved = true }; throw new InvalidOperationException("0x800AC472 busy"); };
                    Assert.AreEqual(false, ((IDictionary<string, object>)SendWithMessagePump(id, request)["Data"])["SaveRequired"]);
                    Assert.AreEqual(2, attempts);
                    attempts = 0;
                    server.PersistSignature = p => { attempts++; throw new InvalidOperationException("0x800AC472 busy"); };
                    failure = (IDictionary<string, object>)SendWithMessagePump(id, request)["Data"];
                    Assert.AreEqual(12, attempts);
                    StringAssert.Contains((string)failure["PersistenceError"], "busy");
                }
            }
        }

        [TestMethod]
                /// <summary>Vérifie la fermeture du serveur lors d’une destruction avant démarrage ou pendant l’attente inactive.</summary>
[STATestMethod]
        public void DisposalBeforeStartAndDuringIdleWaitClosesServer()
        {
            using (var dispatcher = new Control())
            {
                var idle = new BridgeServer(dispatcher, null, Guid.NewGuid().GetHashCode() & int.MaxValue);
                idle.Dispose(); idle.Start();
                var worker = (Thread)typeof(BridgeServer).GetField("worker", System.Reflection.BindingFlags.NonPublic | System.Reflection.BindingFlags.Instance).GetValue(idle);
                Assert.IsTrue(worker.Join(5000));
                int id = Guid.NewGuid().GetHashCode() & int.MaxValue;
                var server = new BridgeServer(dispatcher, null, id);
                server.Start();
                using (var client = new NamedPipeClientStream(".", "VBAi." + id, PipeDirection.InOut))
                {
                    client.Connect(5000);
                    server.Dispose();
                }
                worker = (Thread)typeof(BridgeServer).GetField("worker", System.Reflection.BindingFlags.NonPublic | System.Reflection.BindingFlags.Instance).GetValue(server);
                Assert.IsTrue(worker.Join(5000));
            }
        }

        [TestMethod]
                /// <summary>Vérifie que l’échec de création du canal est retenté sauf si l’arrêt est déjà demandé.</summary>
[STATestMethod]
        public void PipeCreationIoFailureRetriesUnlessShutdownWasRequested()
        {
            using (var dispatcher = new Control())
            {
                var handle = dispatcher.Handle;
                int id = Guid.NewGuid().GetHashCode() & int.MaxValue;
                using (var server = new BridgeServer(dispatcher, null, id))
                {
                    var open = server.OpenPipe; int attempts = 0;
                    server.OpenPipe = security => { if (Interlocked.Increment(ref attempts)==1) throw new IOException("transient pipe failure"); return open(security); };
                    server.Execute = Infrastructure.VbeToolBoundaryFixture.Execute;
                    server.Start();
                    Assert.AreEqual(true,SendWithMessagePump(id,"{\"Command\":\"status\"}")["Ok"]);
                    Assert.IsTrue(attempts>=2);
                }
                var stopping = new BridgeServer(dispatcher,null,Guid.NewGuid().GetHashCode() & int.MaxValue);
                stopping.OpenPipe = security => { stopping.Dispose(); throw new IOException("shutdown during pipe creation"); };
                stopping.Start();
                var worker = (Thread)typeof(BridgeServer).GetField("worker",System.Reflection.BindingFlags.NonPublic | System.Reflection.BindingFlags.Instance).GetValue(stopping);
                Assert.IsTrue(worker.Join(5000));
                using (var disposed = new BridgeServer(dispatcher,null,Guid.NewGuid().GetHashCode() & int.MaxValue))
                {
                    disposed.OpenPipe = security => { throw new ObjectDisposedException("pipe creation"); };
                    disposed.Start();
                    worker = (Thread)typeof(BridgeServer).GetField("worker",System.Reflection.BindingFlags.NonPublic | System.Reflection.BindingFlags.Instance).GetValue(disposed);
                    Assert.IsTrue(worker.Join(5000),"A disposed transport must stop the worker.");
                }
            }
        }


        [STATestMethod]
        public void OptionMutationDispatchUsesNativeResultOnlyAfterSuccessfulScheduling()
        {
            using (var dispatcher = new Control())
            {
                var handle = dispatcher.Handle; int id = Guid.NewGuid().GetHashCode() & int.MaxValue;
                using (var server = new BridgeServer(dispatcher, null, id))
                {
                    Infrastructure.VbeToolBoundaryFixture.Configure(server.Native);
                    server.Execute = Infrastructure.VbeToolBoundaryFixture.Execute;
                    int calls = 0;
                    server.Native.SetVbeOption = request => { calls++; Assert.AreEqual("editor", request.Pane); return new { Changed = true }; };
                    server.Start();
                    var response = SendWithMessagePump(id, "{\"Command\":\"set_vbe_option\",\"Pane\":\"editor\"}");
                    Assert.AreEqual(true, response["Ok"]); Assert.AreEqual(1, calls);
                    server.Execute = request => Response.Failure("schedule rejected");
                    response = SendWithMessagePump(id, "{\"Command\":\"set_vbe_option\"}");
                    Assert.AreEqual(false, response["Ok"]); Assert.AreEqual(1, calls);
                    server.Execute = Infrastructure.VbeToolBoundaryFixture.Execute;
                    server.Native.SetVbeOption = request => throw new InvalidOperationException("native rejected");
                    response = SendWithMessagePump(id, "{\"Command\":\"set_vbe_option\"}");
                    Assert.AreEqual(false, response["Ok"]); Assert.AreEqual("native rejected", response["Error"]);
                }
            }
        }
        [STATestMethod]
        public void StopRequestedBeforeListenerPublicationExitsAfterOwnedConnectionWithoutDispatch()
        {
            using (var dispatcher = new Control())
            using (var created = new ManualResetEventSlim())
            {
                int id = Guid.NewGuid().GetHashCode() & int.MaxValue; int requests = 0;
                using (var server = new BridgeServer(dispatcher, null, id))
                {
                    var open = server.OpenPipe; server.Execute = request => { requests++; throw new AssertFailedException("A stopped listener must not dispatch"); };
                    server.OpenPipe = security => { var pipe = open(security); server.Dispose(); created.Set(); return pipe; };
                    server.Start(); Assert.IsTrue(created.Wait(5000));
                    using (var client = new NamedPipeClientStream(".", "VBAi." + id, PipeDirection.InOut)) client.Connect(5000);
                    var worker = (Thread)typeof(BridgeServer).GetField("worker", System.Reflection.BindingFlags.Instance | System.Reflection.BindingFlags.NonPublic).GetValue(server);
                    Assert.IsTrue(worker.Join(5000)); Assert.AreEqual(0, requests);
                }
            }
        }
    }
}

using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Runtime.InteropServices;
using System.Threading;
using System.Web.Script.Serialization;
using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace VBAi.Tests.Integration
{
    /// <summary>Provides durable adapter-only evidence without compilation or a post-trial save helper.</summary>
    internal sealed partial class OfficeVbeFixture
    {
        private bool adapterOnlyCleanup;
        private readonly List<object> retainedDiagnosticReferences = new List<object>();

        /// <summary>Requires discard-on-close for Access objects, including cleanup after a failed trial.</summary>
        internal void RequireAdapterOnlyCleanup() { adapterOnlyCleanup = true; }

        /// <summary>Records an assertion failure before cleanup so it survives a second cleanup exception.</summary>
        internal void RecordAdapterFailure(Exception error)
        {
            steps.Add(new { AdapterOnlyFailure = error.ToString() });
            FlushAdapterEvidence();
        }

        /// <summary>Persists a phase and its synthetic inputs before native dispatch.</summary>
        internal void RecordAdapterStage(string phase, object inputs)
        {
            steps.Add(new { AdapterStage = phase, Inputs = inputs, Utc = DateTime.UtcNow.ToString("O") });
            FlushAdapterEvidence();
        }

        /// <summary>Writes incremental evidence before entering any potentially blocking native operation.</summary>
        internal void FlushAdapterEvidence()
        {
            File.WriteAllText(Path.Combine(Root, "adapter-only-progress.json"),
                new JavaScriptSerializer { MaxJsonLength = 20 * 1024 * 1024 }.Serialize(
                    new { Host = Kind, HostProgId = hostProgId, ProcessId, DocumentPath, Project,
                        ExpectedMvid = typeof(VbeSession).Module.ModuleVersionId.ToString("D"),
                        PendingCommand = commandContainment.Command, CommandPending = commandContainment.Pending,
                        DeliveryUncertain = commandContainment.Uncertain,
                        AdapterOnlyCleanup = adapterOnlyCleanup, Failures, Steps = steps }));
        }

        /// <summary>Establishes the disposable baseline before final edits; never used after adapter invocation.</summary>
        internal void SaveAdapterBaseline(string[] modules)
        {
            commandContainment.RequireTerminal();
            steps.Add(new { AdapterBaselineSaveStarting = true, Modules = modules });
            FlushAdapterEvidence();
            if (Kind == "Access")
            {
                RequireOwnedDocument();
                object commands = null;
                try
                {
                    commands = ((dynamic)application).DoCmd;
                    foreach (string name in modules)
                    {
                        RequireOwnedDocument();
                        ((dynamic)commands).Save(5, name); // acModule: each explicitly named open synthetic object.
                        steps.Add(new { BaselineObjectSaved = name, ObjectType = 5 });
                        FlushAdapterEvidence();
                    }
                }
                finally { ReleaseAdapterReference(commands); }
            }
            else SaveNative();
            StopAccessSaveDialogHandler();
            RecordAdapterObservation("BaselineSaved");
        }

        /// <summary>Captures native PID, selection and per-component Saved observations without changing them.</summary>
        internal IDictionary<string, object> RecordAdapterObservation(string phase)
        {
            commandContainment.RequireTerminal();
            var observation = new Dictionary<string, object> {
                ["AdapterOnlyPhase"] = phase, ["Utc"] = DateTime.UtcNow.ToString("O"),
                ["ProcessId"] = ProcessId, ["DocumentPath"] = DocumentPath };
            steps.Add(observation);
            FlushAdapterEvidence();
            object window = null;
            try
            {
                observation["ApplicationVersion"] = Convert.ToString(((dynamic)application).Version);
                using (var process = Process.GetProcessById(ProcessId))
                {
                    observation["HostExecutable"] = process.MainModule.FileName;
                    observation["HostFileVersion"] = FileVersionInfo.GetVersionInfo(process.MainModule.FileName).FileVersion;
                }
                if (Kind == "Access") observation["ApplicationHwnd"] = Convert.ToInt64(((dynamic)application).hWndAccessApp());
                else if (Kind == "PowerPoint") observation["ApplicationHwnd"] = PowerPointWindow.Read(application).ToInt64();
                else
                {
                    window = ((dynamic)application).ActiveWindow;
                    observation["ApplicationHwnd"] = Convert.ToInt64(((dynamic)window).Hwnd);
                }
                uint owner;
                GetWindowThreadProcessId(new IntPtr(Convert.ToInt64(observation["ApplicationHwnd"])), out owner);
                observation["ApplicationOwnerPid"] = owner;
                Assert.AreEqual((uint)ProcessId, owner, "The observed application window must belong to the owned process.");
                // The bridge reads the native project on its owning STA. Publisher does not
                // need an invented Application.VBE or Document.VBProject accessor here.
                var persistence = Data("project_persistence_status");
                observation["Persistence"] = persistence;
                observation["ProjectSaved"] = persistence["ProjectSaved"];
                observation["ActiveCodeComponent"] = Data("debug_state")["ActiveModule"];
                var components = new List<object>();
                observation["Components"] = components;
                foreach (var item in Items("list_modules"))
                {
                    string name = Convert.ToString(item["Name"]);
                    var state = new Dictionary<string, object> { ["Name"] = name, ["Type"] = item["Type"] };
                    var saved = Response("component_probe", "Module", name, "Action", "descriptor_value", "Query", "Saved");
                    if (Convert.ToBoolean(saved["Ok"])) state["Saved"] = VbeBridgeClient.Object(saved["Data"])["Value"];
                    else state["SavedError"] = saved["Error"];
                    components.Add(state);
                }
            }
            catch (Exception error) { observation["ObservationError"] = error.ToString(); throw; }
            finally
            {
                ReleaseAdapterReference(window);
                FlushAdapterEvidence();
            }
            return observation;
        }

        /// <summary>Records the unchanged adapter result while yielding for delayed read-only native observations.</summary>
        internal void ObserveAdapterOutcome(IDictionary<string, object> reply)
        {
            steps.Add(new { OriginalAdapterResponse = reply, SaveRetryInvoked = false, CompileInvoked = false, PostAdapterHelperSaveInvoked = false });
            FlushAdapterEvidence();
            RecordAdapterObservation("ImmediatelyAfterAdapter");
            int elapsed = 0;
            foreach (int delay in new[] { 100, 250, 1000 })
            {
                Thread.Sleep(delay); elapsed += delay;
                RecordAdapterObservation("ReadOnlyAfter" + elapsed + "ms");
            }
        }

        /// <summary>Balances only the reference acquired by this diagnostic, preserving retained fixture objects.</summary>
        private void ReleaseAdapterReference(object item)
        {
            if (commandContainment.Pending || commandContainment.Uncertain)
            {
                if (item != null) retainedDiagnosticReferences.Add(item);
                return;
            }
            if (item != null && Marshal.IsComObject(item)) Marshal.ReleaseComObject(item);
        }
    }
}

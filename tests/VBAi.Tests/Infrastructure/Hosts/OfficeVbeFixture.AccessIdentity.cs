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
    /// <summary>Observes transient Access document identities without editing or saving the project.</summary>
    internal sealed partial class OfficeVbeFixture
    {
        /// <summary>Keeps all acquired objects alive until every identity observation is durable.</summary>
        internal void RecordAccessIdentityProbe(string expectedExecutable)
        {
            Assert.AreEqual("Access", Kind);
            Assert.AreEqual(ApartmentState.STA, Thread.CurrentThread.GetApartmentState());
            commandContainment.RequireTerminal();
            var acquisitions = new List<object>();
            var observations = new List<object>();
            var evidence = new Dictionary<string, object> {
                ["Phase"] = "Prepared", ["Utc"] = DateTime.UtcNow.ToString("O"),
                ["ProcessId"] = ProcessId, ["DocumentPath"] = DocumentPath,
                ["Project"] = Project, ["ExpectedMvid"] = typeof(VbeSession).Module.ModuleVersionId.ToString("D"),
                ["Samples"] = observations, ["SampleLimit"] = 5,
                ["NativeMutationAllowed"] = false, ["SaveAllowed"] = false,
                ["ObjectsHeldAcrossSamples"] = true };
            Action<string> persist = phase => {
                evidence["Phase"] = phase;
                File.WriteAllText(Path.Combine(Root, "access-identity-probe.json"),
                    new JavaScriptSerializer().Serialize(evidence));
            };
            Func<object, object> acquired = value => { if (value != null) acquisitions.Add(value); return value; };
            try
            {
                persist("IdentityReadPending");
                Assert.IsTrue(owned && ownedProcess != null && !ownedProcess.HasExited);
                evidence["ProcessStartUtc"] = ownedProcess.StartTime.ToUniversalTime().ToString("O");
                string executable = ExcelOwnedProcessImage.Read(ownedProcess.Handle);
                evidence["Executable"] = executable;
                evidence["ExpectedExecutable"] = expectedExecutable;
                evidence["ExecutableVersion"] = FileVersionInfo.GetVersionInfo(executable).FileVersion;
                Assert.AreEqual("MSACCESS.EXE", Path.GetFileName(executable).ToUpperInvariant());
                Assert.AreEqual(Path.GetFullPath(expectedExecutable), Path.GetFullPath(executable), true,
                    "The retained Access executable must match installed HKLM App Paths exactly.");
                evidence["ApplicationVersion"] = Convert.ToString(((dynamic)application).Version);
                Assert.IsTrue(((string)evidence["ApplicationVersion"]).StartsWith("16.", StringComparison.Ordinal));
                long hwnd = Convert.ToInt64(((dynamic)application).hWndAccessApp());
                uint pid; GetWindowThreadProcessId(new IntPtr(hwnd), out pid);
                evidence["ApplicationHwnd"] = hwnd; evidence["WindowOwnerPid"] = pid;
                Assert.AreEqual((uint)ProcessId, pid);
                evidence["LoadedStatus"] = Data("status");
                var status = (IDictionary<string, object>)evidence["LoadedStatus"];
                Assert.AreEqual(ProcessId, Convert.ToInt32(status["HostProcessId"]));
                Assert.AreEqual(evidence["ExpectedMvid"], status["AssemblyModuleVersionId"]);
                object editor = acquired(((dynamic)application).VBE);
                object projects = acquired(((dynamic)editor).VBProjects);
                evidence["ApplicationIUnknown"] = Identity(application);
                evidence["VbeIUnknown"] = Identity(editor);
                for (int sample = 0; sample < 5; sample++)
                {
                    evidence["PendingSample"] = sample;
                    persist("CurrentProjectReadPending");
                    object current = acquired(((dynamic)application).CurrentProject);
                    string fullName = Convert.ToString(((dynamic)current).FullName);
                    Assert.AreEqual(Path.GetFullPath(DocumentPath), Path.GetFullPath(fullName), true);
                    var row = new Dictionary<string, object> {
                        ["Index"] = sample, ["Utc"] = DateTime.UtcNow.ToString("O"),
                        ["CurrentProjectIUnknown"] = Identity(current), ["CurrentProjectFullName"] = fullName };
                    observations.Add(row);
                    persist("MappedProjectReadPending");
                    int matches = 0;
                    for (int index = 1; index <= Convert.ToInt32(((dynamic)projects).Count); index++)
                    {
                        object candidate = acquired(((dynamic)projects).Item(index));
                        string path = VbeProjectHostPath.Read(candidate);
                        if (!string.Equals(Path.GetFullPath(path), Path.GetFullPath(DocumentPath), StringComparison.OrdinalIgnoreCase)) continue;
                        matches++;
                        row["MappedProjectIUnknown"] = Identity(candidate);
                        row["MappedProjectPath"] = path;
                        row["MappedProjectName"] = Convert.ToString(((dynamic)candidate).Name);
                    }
                    Assert.AreEqual(1, matches, "The owned database must map to one native VBProject.");
                    persist("SelectedProjectReadPending");
                    object selected = acquired(((dynamic)editor).ActiveVBProject);
                    row["SelectedProjectIUnknown"] = Identity(selected);
                    row["SelectedProjectPath"] = VbeProjectHostPath.Read(selected);
                    Assert.AreEqual(row["MappedProjectIUnknown"], row["SelectedProjectIUnknown"],
                        "The selected project must be the exact mapped disposable project.");
                    GetWindowThreadProcessId(new IntPtr(Convert.ToInt64(((dynamic)application).hWndAccessApp())), out pid);
                    Assert.AreEqual((uint)ProcessId, pid);
                    Assert.IsFalse(ownedProcess.HasExited);
                    persist("SampleObserved");
                }
                evidence["PendingSample"] = null;
                persist("ObservationsDurableBeforeRelease");
                // Balance each successful getter acquisition once. Never FinalRelease a shared RCW.
                for (int index = acquisitions.Count - 1; index >= 0; index--)
                {
                    object value = acquisitions[index];
                    if (Marshal.IsComObject(value)) Marshal.ReleaseComObject(value);
                    acquisitions.RemoveAt(index);
                }
                persist("BalancedReferencesReleased");
                RecordAdapterStage("ReadOnlyAccessIdentityComplete", new { Samples = observations.Count,
                    CurrentProjectEqualityRequired = false, SourceEdited = false, PropertiesEdited = false, AdapterSaveInvoked = false });
            }
            catch (Exception error)
            {
                evidence["Error"] = error.ToString();
                retainedDiagnosticReferences.AddRange(acquisitions);
                RetainUncertainOffice();
                try { persist("FailedPreserved"); }
                catch (Exception writeError) { throw new AggregateException("Identity observation and evidence persistence failed; host retained.", error, writeError); }
                throw;
            }
        }

        /// <summary>Balances only the reference obtained by GetIUnknownForObject.</summary>
        private static string Identity(object value)
        {
            IntPtr pointer = Marshal.GetIUnknownForObject(value);
            try { return "0x" + unchecked((ulong)pointer.ToInt64()).ToString("X16"); }
            finally { Marshal.Release(pointer); }
        }
    }
}

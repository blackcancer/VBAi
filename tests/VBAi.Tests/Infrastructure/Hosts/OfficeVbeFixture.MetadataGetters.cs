using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Security.Cryptography;
using System.Threading;
using System.Web.Script.Serialization;
using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace VBAi.Tests.Integration
{
    /// <summary>Scopes getter diagnostics to the exact owned Access database and its native VBProject identity.</summary>
    internal sealed partial class OfficeVbeFixture
    {
        internal void RecordMetadataGetterProbe(string phase, bool requireSuccessfulEqualGetters = false)
        {
            if (Environment.GetEnvironmentVariable("VBAi_RUN_OFFICE_METADATA_GETTER_PROBE") != "1" || Kind != "Access") return;
            Assert.IsTrue(new[] { "FreshBaseline", "BeforeExistingMutation", "AfterExistingMutation", "AfterExistingSave", "FreshDiskReopen" }.Contains(phase));
            Assert.AreEqual(ApartmentState.STA, Thread.CurrentThread.GetApartmentState());
            commandContainment.RequireTerminal();
            RequireOwnedDocument();
            var objects = new List<object>();
            var report = new Dictionary<string, object> {
                ["Phase"] = phase, ["State"] = "PENDING", ["Utc"] = DateTime.UtcNow.ToString("o"),
                ["HostProcessId"] = ProcessId, ["HostStartedUtc"] = ownedProcess.StartTime.ToUniversalTime().ToString("o"),
                ["DocumentPath"] = DocumentPath, ["ProjectSelector"] = Project,
                ["DocumentExtension"] = Path.GetExtension(DocumentPath),
                ["ExpectedAssemblyMvid"] = typeof(VbeSession).Module.ModuleVersionId.ToString("D"),
                ["ReferencedAssemblyPath"] = typeof(VbeSession).Assembly.Location,
                ["Apartment"] = Thread.CurrentThread.GetApartmentState().ToString(),
                ["GetterExecution"] = "External owned fixture STA; COM proxy dispatches to Office's owning apartment.",
                ["Scope"] = "Only HelpFile/HelpContextID values and their runtime type info on the mapped owned Access VBProject; no setters, methods, save, help invocation or macro execution.",
                ["AdditionalNativeMutations"] = 0, ["AdditionalSaves"] = 0 };
            string path = Path.Combine(Root, "metadata-getters-" + phase + ".json");
            var serializer = new JavaScriptSerializer { MaxJsonLength = 1024 * 1024 };
            Action persist = () => File.WriteAllText(path, serializer.Serialize(report));
            Func<object, object> hold = value => { if (value != null) objects.Add(value); return value; };
            Exception failure = null;
            try
            {
                persist();
                var status = Data("status"); report["BridgeStatus"] = status;
                Assert.AreEqual(ProcessId, Convert.ToInt32(status["HostProcessId"]));
                Assert.AreEqual(report["ExpectedAssemblyMvid"], status["AssemblyModuleVersionId"]);
                report["ReferencedAssemblySha256"] = MetadataAssemblyHash(typeof(VbeSession).Assembly.Location);
                report["HostAssemblyFileSha256"] = MetadataAssemblyHash(Convert.ToString(status["AssemblyPath"]));
                Assert.AreEqual(report["ReferencedAssemblySha256"], report["HostAssemblyFileSha256"], "The owned host's assembly file differs from the test reference.");
                object editor = hold(((dynamic)application).VBE);
                object mainWindow = hold(((dynamic)editor).MainWindow);
                uint owner; uint thread = GetWindowThreadProcessId(new IntPtr(Convert.ToInt64(((dynamic)mainWindow).HWnd)), out owner);
                report["VbeWindowOwnerPid"] = owner; report["VbeOwnerThreadId"] = thread;
                Assert.AreEqual((uint)ProcessId, owner);
                object projects = hold(((dynamic)editor).VBProjects);
                int count = Convert.ToInt32(((dynamic)projects).Count);
                Assert.IsTrue(count > 0 && count <= 8, "Disposable Access project inventory exceeded its bound.");
                object mapped = null; int matches = 0;
                for (int index = 1; index <= count; index++)
                {
                    object candidate = hold(((dynamic)projects).Item(index));
                    string hostPath = VbeProjectHostPath.Read(candidate);
                    if (!string.Equals(Path.GetFullPath(hostPath), Path.GetFullPath(DocumentPath), StringComparison.OrdinalIgnoreCase)) continue;
                    matches++; mapped = candidate;
                }
                Assert.AreEqual(1, matches, "The owned database must map to exactly one VBProject.");
                object selected = hold(((dynamic)editor).ActiveVBProject);
                string identity = Identity(mapped);
                Assert.AreEqual(identity, Identity(selected), "The selected VBProject is not the mapped owned database.");
                report["NativeProjectIUnknown"] = identity;
                report["ProjectHostPath"] = VbeProjectHostPath.Read(mapped);
                report["ModeBefore"] = Convert.ToInt32(((dynamic)mapped).Mode);
                report["ProtectionBefore"] = Convert.ToInt32(((dynamic)mapped).Protection);
                report["SavedBefore"] = Convert.ToBoolean(((dynamic)mapped).Saved);
                Assert.AreEqual(2, report["ModeBefore"]); Assert.AreEqual(0, report["ProtectionBefore"]);
                report["Reads"] = NativeMetadataGetterProbe.Read(mapped, pending => {
                    report["PendingRead"] = pending; persist();
                }, partial => report["Reads"] = partial);
                report["PendingRead"] = null;
                Assert.AreEqual(identity, Identity(mapped));
                Assert.AreEqual(report["ProjectHostPath"], VbeProjectHostPath.Read(mapped));
                report["ModeAfter"] = Convert.ToInt32(((dynamic)mapped).Mode);
                report["ProtectionAfter"] = Convert.ToInt32(((dynamic)mapped).Protection);
                report["SavedAfter"] = Convert.ToBoolean(((dynamic)mapped).Saved);
                Assert.AreEqual(report["ModeBefore"], report["ModeAfter"]);
                Assert.AreEqual(report["ProtectionBefore"], report["ProtectionAfter"]);
                Assert.AreEqual(report["SavedBefore"], report["SavedAfter"], "Read-only getters changed Saved state.");
                RequireOwnedDocument();
                report["State"] = "OBSERVED"; persist();
                if (requireSuccessfulEqualGetters)
                {
                    var reads = (IDictionary<string, object>)report["Reads"];
                    var typeInfo = (IDictionary<string, object>)reads["RuntimeTypeInfo"];
                    Assert.AreEqual("READ", typeInfo["State"], "Runtime type-info read failed; raw evidence is retained.");
                    Assert.AreEqual(true, typeInfo["BothGettersAndSettersObserved"], "Runtime type info must expose both getter/setter contracts for DISPIDs 116 and 117.");
                    foreach (IDictionary<string, object> property in (System.Collections.IEnumerable)reads["Properties"])
                        Assert.AreEqual(true, property["ExactGetterEquality"], "Fresh baseline getters disagree or failed for " + property["Name"]);
                }
                // Comparison results remain raw evidence. A mismatch must not
                // prevent the existing save/reopen qualification from recording
                // its original outcome; its exact metadata assertions stay intact.
                RecordAdapterStage("MetadataGetterProbeObserved", new { Phase = phase, EvidencePath = path,
                    AdditionalNativeMutations = 0, AdditionalSaves = 0 });
            }
            catch (Exception error)
            {
                failure = error; report["State"] = "FAILED"; report["Error"] = error.ToString();
                try { persist(); } catch (Exception evidence) { failure = new AggregateException("Getter probe and evidence persistence failed.", error, evidence); }
            }
            finally
            {
                // Each successful acquisition is balanced once; never invalidate
                // another holder of the same RCW with FinalReleaseComObject.
                for (int index = objects.Count - 1; index >= 0; index--)
                    try { if (Marshal.IsComObject(objects[index])) Marshal.ReleaseComObject(objects[index]); }
                    catch (Exception release) { failure = failure == null ? release : new AggregateException("Getter probe and reference cleanup failed.", failure, release); }
            }
            if (failure != null)
            {
                report["State"] = "FAILED"; report["Error"] = failure.ToString();
                try { persist(); }
                catch (Exception evidence) { failure = new AggregateException("Getter probe/reference cleanup and final evidence persistence failed.", failure, evidence); }
                System.Runtime.ExceptionServices.ExceptionDispatchInfo.Capture(failure).Throw();
            }
        }

        private static string MetadataAssemblyHash(string path)
        {
            using (var hash = SHA256.Create())
            using (var input = File.OpenRead(path))
                return BitConverter.ToString(hash.ComputeHash(input)).Replace("-", "");
        }
    }
}

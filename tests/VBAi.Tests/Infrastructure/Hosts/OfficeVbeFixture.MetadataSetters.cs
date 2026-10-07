using Microsoft.VisualStudio.TestTools.UnitTesting;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.IO;
using System.Runtime.ExceptionServices;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading;
using System.Web.Script.Serialization;

namespace VBAi.Tests.Integration
{
    /// <summary>Guards a single diagnostic setter on the exact fresh disposable Access project, from the external fixture STA.</summary>
    internal sealed partial class OfficeVbeFixture
    {
        internal void InvokeOwnedMetadataSetter(string property, object value, bool rawDispatch)
        {
            Assert.AreEqual("1", Environment.GetEnvironmentVariable("VBAi_RUN_OFFICE_METADATA_SETTER_PROBE"));
            Assert.AreEqual("1", Environment.GetEnvironmentVariable("VBAi_RUN_OFFICE_METADATA_GETTER_PROBE"));
            Assert.AreEqual("Access", Kind);
            Assert.AreEqual(ApartmentState.STA, Thread.CurrentThread.GetApartmentState());
            Assert.IsTrue(property == "HelpFile" && value is string || property == "HelpContextID" && value is int);
            commandContainment.RequireTerminal(); RequireOwnedDocument();
            var objects = new List<object>();
            var row = new Dictionary<string, object>
            {
                ["State"] = "PREPARED",
                ["Property"] = property,
                ["ExpectedValue"] = value,
                ["Path"] = rawDispatch ? "ExternalSta.RawDispatchPut" : "ExternalSta.ProductionClrSetNative",
                ["ExecutionScope"] = "External fixture STA on mapped RCW; compare these paths in the same proxy context, not as in-process bridge dispatch.",
                ["HostProcessId"] = ProcessId,
                ["HostStartedUtc"] = ownedProcess.StartTime.ToUniversalTime().ToString("o"),
                ["DocumentPath"] = DocumentPath,
                ["DocumentExtension"] = Path.GetExtension(DocumentPath),
                ["ProjectSelector"] = Project,
                ["ExpectedMvid"] = typeof(VbeSession).Module.ModuleVersionId.ToString("D"),
                ["MutationRetryAllowed"] = false,
                ["MaximumMetadataSetterEntries"] = 1,
                ["AdditionalSaveInvocations"] = 0,
                ["HelpFileCreated"] = false,
                ["HelpInvoked"] = false
            };
            string path = Path.Combine(Root, "metadata-setter.json");
            var serializer = new JavaScriptSerializer { MaxJsonLength = 1024 * 1024 };
            Action persist = () => File.WriteAllText(path, serializer.Serialize(row));
            Func<object, object> hold = item => { if (item != null) objects.Add(item); return item; };
            Exception failure = null;
            bool ownsReport = false;
            try
            {
                using (var initial = new FileStream(path, FileMode.CreateNew, FileAccess.Write, FileShare.Read))
                {
                    ownsReport = true;
                    using (var writer = new StreamWriter(initial, new UTF8Encoding(false))) writer.Write(serializer.Serialize(row));
                }
                var status = Data("status"); row["BridgeStatus"] = status;
                Assert.AreEqual(ProcessId, Convert.ToInt32(status["HostProcessId"]));
                Assert.AreEqual(row["ExpectedMvid"], status["AssemblyModuleVersionId"]);
                row["ReferencedAssemblySha256"] = MetadataAssemblyHash(typeof(VbeSession).Assembly.Location);
                row["HostAssemblyFileSha256"] = MetadataAssemblyHash(Convert.ToString(status["AssemblyPath"]));
                Assert.AreEqual(row["ReferencedAssemblySha256"], row["HostAssemblyFileSha256"]);
                object revision = Data("project_properties")["Version"]; row["ExpectedProjectVersion"] = revision;
                object editor = hold(((dynamic)application).VBE);
                object window = hold(((dynamic)editor).MainWindow);
                uint owner; uint ownerThread = GetWindowThreadProcessId(new IntPtr(Convert.ToInt64(((dynamic)window).HWnd)), out owner);
                row["VbeWindowOwnerPid"] = owner; row["VbeOwnerThreadId"] = ownerThread;
                Assert.AreEqual((uint)ProcessId, owner);
                object projects = hold(((dynamic)editor).VBProjects);
                int count = Convert.ToInt32(((dynamic)projects).Count);
                Assert.IsTrue(count > 0 && count <= 8);
                object mapped = null; int matches = 0;
                for (int index = 1; index <= count; index++)
                {
                    object candidate = hold(((dynamic)projects).Item(index));
                    if (!string.Equals(Path.GetFullPath(VbeProjectHostPath.Read(candidate)), Path.GetFullPath(DocumentPath), StringComparison.OrdinalIgnoreCase)) continue;
                    mapped = candidate; matches++;
                }
                Assert.AreEqual(1, matches);
                object selected = hold(((dynamic)editor).ActiveVBProject);
                string identity = Identity(mapped); row["NativeProjectIUnknown"] = identity;
                Assert.AreEqual(identity, Identity(selected));
                row["ModeBefore"] = Convert.ToInt32(((dynamic)mapped).Mode);
                row["ProtectionBefore"] = Convert.ToInt32(((dynamic)mapped).Protection);
                row["SavedBefore"] = Convert.ToBoolean(((dynamic)mapped).Saved);
                Assert.AreEqual(2, row["ModeBefore"]); Assert.AreEqual(0, row["ProtectionBefore"]);
                var descriptor = TypeDescriptor.GetProperties(mapped).Find(property, false);
                Assert.IsNotNull(descriptor); Assert.IsFalse(descriptor.IsReadOnly);
                row["DescriptorType"] = descriptor.PropertyType.FullName;
                Assert.AreEqual(property == "HelpFile" ? typeof(string) : typeof(int), descriptor.PropertyType);
                Assert.AreEqual(revision, Data("project_properties")["Version"], "Project revision changed before native setter dispatch.");
                RequireOwnedDocument();
                object dispatchSelection = hold(((dynamic)editor).ActiveVBProject);
                Assert.AreEqual(identity, Identity(dispatchSelection), "Selection changed before the one-shot native setter.");
                Assert.AreEqual(2, Convert.ToInt32(((dynamic)mapped).Mode));
                Assert.AreEqual(0, Convert.ToInt32(((dynamic)mapped).Protection));
                row["State"] = "DISPATCH_ATTEMPT_STARTING"; persist();
                if (rawDispatch)
                {
                    var raw = NativeMetadataSetterProbe.PutRaw(mapped, property, value); row["RawDispatch"] = raw;
                    row["State"] = "DISPATCH_RETURNED"; persist();
                    Assert.AreEqual(1, raw["InvokeEntries"]);
                    Assert.AreEqual(-3, raw["NamedDispidAfterInvoke"]);
                    var returned = (IDictionary<string, object>)raw["ReturnVariant"];
                    var argument = (IDictionary<string, object>)raw["ArgumentVariant"];
                    // Preserve the original native HRESULT before diagnosing buffer/read errors.
                    if (returned.TryGetValue("HResult", out object hr))
                        Marshal.ThrowExceptionForHR(unchecked((int)Convert.ToUInt32(Convert.ToString(hr).Substring(2), 16)));
                    Assert.AreEqual("READ", returned["State"]); Assert.AreEqual("READ", argument["State"]);
                    Assert.AreEqual(true, argument["CanaryIntact"]); Assert.AreEqual(true, returned["CanaryIntact"]);
                    Assert.AreEqual("0x00000000", argument["VariantClearHResult"]);
                    Assert.AreEqual("0x00000000", returned["VariantClearHResult"]);
                    Assert.AreEqual(value, argument["Value"], "Native setter altered its input value.");
                }
                else
                {
                    VbeScalarProperty.SetNative(mapped, property, value); // Exact existing implementation, no descriptor setter.
                    row["State"] = "DISPATCH_RETURNED"; row["ClrSetNativeReturned"] = true; persist();
                }
                Assert.AreEqual(identity, Identity(mapped));
                Assert.AreEqual(DocumentPath, VbeProjectHostPath.Read(mapped));
                Assert.AreEqual(2, Convert.ToInt32(((dynamic)mapped).Mode));
                Assert.AreEqual(0, Convert.ToInt32(((dynamic)mapped).Protection));
                RequireOwnedDocument();
                row["State"] = "RETURNED_AND_IDENTITY_VERIFIED"; persist();
                RecordAdapterStage("OneShotMetadataSetterReturned", row);
            }
            catch (Exception error)
            {
                failure = error; row["State"] = "FAILED"; row["OriginalError"] = error.ToString();
                row["OriginalHResult"] = "0x" + unchecked((uint)error.HResult).ToString("X8");
                row["MutationAfterState"] = "Not inferred from failure; separate read-only getter evidence required.";
            }
            finally
            {
                for (int index = objects.Count - 1; index >= 0; index--)
                    try { if (Marshal.IsComObject(objects[index])) Marshal.ReleaseComObject(objects[index]); }
                    catch (Exception release) { failure = failure == null ? release : new AggregateException("Setter outcome and reference cleanup failed.", failure, release); }
                if (failure != null) { row["State"] = "FAILED"; row["FinalError"] = failure.ToString(); }
                try { if (ownsReport) persist(); }
                catch (Exception evidence) { failure = failure == null ? evidence : new AggregateException("Setter outcome and durable evidence both failed.", failure, evidence); }
            }
            if (failure != null) ExceptionDispatchInfo.Capture(failure).Throw();
        }
    }
}

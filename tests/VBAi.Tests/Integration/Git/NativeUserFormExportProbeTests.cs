using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Security.AccessControl;
using System.Security.Cryptography;
using System.Security.Principal;
using System.Web.Script.Serialization;
using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace VBAi.Tests.Integration
{
    /// <summary>Separates native export destination and dispatch failures, using one export per owned Excel trial.</summary>
    [TestClass, TestCategory("Excel"), TestCategory("NativeUserFormExportProbe"), DoNotParallelize]
    public sealed class NativeUserFormExportProbeTests
    {
        public TestContext TestContext { get; set; }

        [STATestMethod]
        [DataRow("ExternalSta", "FixtureTemporary")]
        [DataRow("ExternalSta", "GitTemporary")]
        [DataRow("ExternalSta", "EvidenceRoot")]
        [DataRow("HostBridge", "FixtureTemporary")]
        [DataRow("HostBridge", "GitTemporary")]
        [DataRow("HostBridge", "EvidenceRoot")]
        public void SingleExportPreservesExactOwnedIdentityAndProducesFormAndResources(string dispatch, string location)
        { RunSingleExport(dispatch, location, false); }

        /// <summary>Compares fresh sibling directories differing only in EFS, without changing production storage.</summary>
        [STATestMethod]
        [DataRow("Plain")]
        [DataRow("Encrypted")]
        public void ControlledSiblingEfsExportPreservesIdentityAndRetainsRawEvidence(string variant)
        { RunSingleExport("HostBridge", variant, true); }

        /// <summary>Separates volume and fresh-parent EFS inheritance from encryption applied only to an export leaf.</summary>
        [STATestMethod]
        [DataRow("C", "Parent")]
        [DataRow("C", "Leaf")]
        [DataRow("E", "Parent")]
        [DataRow("E", "Leaf")]
        public void ControlledVolumeAndAncestorEfsExportRetainsOwnedIdentity(string volume, string encryptionScope)
        { RunSingleExport("HostBridge", encryptionScope, true, volume); }

        /// <summary>Locates an inherited storage boundary without altering any parent directory or encryption flag.</summary>
        [STATestMethod]
        [DataRow("LocalAppData")]
        [DataRow("VbaiLocalAppData")]
        [DataRow("GitTemporary")]
        [DataRow("UserTemporary")]
        public void InheritedStorageAncestorExportRetainsExactOwnedIdentity(string location)
        { RunSingleExport("HostBridge", location, false, null, true); }

        /// <summary>Crosses only copied parent DACLs on fresh GUID children, retaining native EFS and owner unchanged.</summary>
        [STATestMethod]
        [DataRow("LocalAppData", "UserTemporary")]
        [DataRow("UserTemporary", "LocalAppData")]
        public void CrossedSourceOnlyDaclExportRetainsIdentityAndShutdownEvidence(string location, string daclSource)
        { RunSingleExport("HostBridge", location, false, null, true, daclSource); }

        private void RunSingleExport(string dispatch, string location, bool controlledEfs, string volume = null, bool probeAncestor = false,
            string daclSource = null)
        {
            if (Environment.GetEnvironmentVariable("VBAi_RUN_USERFORM_EXPORT_PROBES") != "1" ||
                Environment.GetEnvironmentVariable("VBAi_RUN_EXCEL_TESTS") != "1")
                Assert.Inconclusive("Set VBAi_RUN_USERFORM_EXPORT_PROBES=1 and VBAi_RUN_EXCEL_TESTS=1 for one export per disposable Excel trial.");
            string root = Environment.GetEnvironmentVariable("VBAi_TEST_USERFORM_LOCAL_GIT_OUTPUT");
            if (string.IsNullOrWhiteSpace(root)) root = TestContext.TestRunResultsDirectory;
            if (string.IsNullOrWhiteSpace(root)) root = Path.Combine(Path.GetTempPath(), "VBAi-UserFormExports");
            Assert.IsTrue(Path.IsPathRooted(root));
            bool rootExisted = Directory.Exists(root);
            string trial = Guid.NewGuid().ToString("N");
            string output = Path.Combine(root, controlledEfs ? "efs-" + trial : dispatch + "-" + location + "-" + trial);
            Assert.IsFalse(Directory.Exists(output)); Directory.CreateDirectory(output);
            string reportPath = Path.Combine(output, "native-export.json");
            var report = new Dictionary<string, object> {
                ["Stage"] = "STARTED", ["Dispatch"] = dispatch, ["Location"] = location, ["ControlledEfs"] = controlledEfs,
                ["ControlledVolume"] = volume,
                ["DaclSource"] = daclSource,
                ["AssemblyMvid"] = typeof(VbeSession).Module.ModuleVersionId.ToString("D"),
                ["NativeExportRequests"] = 0, ["MacroExecutions"] = 0, ["RemoteOperations"] = 0,
                ["EvidenceRootExisted"] = rootExisted,
                ["Scope"] = "One native UserForm Export per owned Excel process; dispatch/path diagnosis only. No Git import, capture comparison or recovery acceptance."
            };
            var json = new JavaScriptSerializer();
            ExcelVbeFixture observedHost = null;
            try
            {
                ExcelVbeFixture.Run(host => {
                    observedHost = host;
                    const string form = "QualificationForm";
                    string workbook = host.File("native-export-probe.xlsm");
                    host.PrepareGitLayout(form, "LabelButton", workbook);
                    host.CaptureGitFormDesigner(form, Path.Combine(output, "source-designer.png"));
                    var before = host.ReadGitExportContext(form);
                    report["Before"] = before; report["NativeBefore"] = host.ReadGitLayout(form, "LabelButton");
                    var status = host.Command("status"); report["HostStatus"] = status;
                    var loaded = Data(status);
                    Assert.AreEqual(host.ProcessId, Convert.ToInt32(loaded["HostProcessId"]));
                    Assert.AreEqual(report["AssemblyMvid"], loaded["AssemblyModuleVersionId"]);

                    string baseDirectory = location == "GitTemporary" ? Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData), "VBAi", "GitTemporary") :
                        location == "VbaiLocalAppData" ? Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData), "VBAi") :
                        location == "LocalAppData" ? Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData) :
                        location == "UserTemporary" ? Path.GetTempPath() : location == "FixtureTemporary" ? host.Root : output;
                    string destinationDirectory;
                    if (controlledEfs)
                    {
                        report["Stage"] = "CONTROLLED_EFS_PREPARATION";
                        string controlledParent = output;
                        if (volume == "C")
                        {
                            Assert.AreEqual("C:\\", Path.GetPathRoot(Path.GetTempPath()), true,
                                "The C volume control must use the owned TEMP root on C.");
                            controlledParent = Path.Combine(Path.GetTempPath(), "vfq-efs-" + trial);
                            Assert.IsFalse(Directory.Exists(controlledParent)); Directory.CreateDirectory(controlledParent);
                        }
                        else if (volume == "E")
                            Assert.AreEqual("E:\\", Path.GetPathRoot(output), true,
                                "The E volume control requires the explicit evidence root on E.");
                        report["ControlledParent"] = controlledParent;
                        report["ParentAttributesBefore"] = System.IO.File.GetAttributes(controlledParent).ToString();
                        report["EnvironmentTempAttributes"] = System.IO.File.GetAttributes(Path.GetTempPath()).ToString();
                        Assert.AreEqual((FileAttributes)0, System.IO.File.GetAttributes(controlledParent) & FileAttributes.Encrypted,
                            "The fresh common parent must start unencrypted; do not remove inherited EFS to satisfy this test.");
                        bool encryptParent = volume != null && location == "Parent";
                        if (encryptParent) System.IO.File.Encrypt(controlledParent);
                        report["ParentAttributes"] = System.IO.File.GetAttributes(controlledParent).ToString();
                        Assert.AreEqual(encryptParent ? FileAttributes.Encrypted : (FileAttributes)0,
                            System.IO.File.GetAttributes(controlledParent) & FileAttributes.Encrypted);
                        string plain = Path.Combine(controlledParent, volume == null ? "plain0" : "peer00"),
                            encrypted = Path.Combine(controlledParent, volume == null ? "efs001" : "leaf00");
                        Assert.IsFalse(Directory.Exists(plain)); Assert.IsFalse(Directory.Exists(encrypted));
                        Directory.CreateDirectory(plain); Directory.CreateDirectory(encrypted);
                        Assert.AreEqual(encryptParent ? FileAttributes.Encrypted : (FileAttributes)0,
                            System.IO.File.GetAttributes(plain) & FileAttributes.Encrypted);
                        Assert.AreEqual(encryptParent ? FileAttributes.Encrypted : (FileAttributes)0,
                            System.IO.File.GetAttributes(encrypted) & FileAttributes.Encrypted);
                        // The parent strategy proves inherited EFS without applying
                        // Encrypt to the leaf. The leaf strategy encrypts only this
                        // empty sibling. Neither strategy decrypts or changes VBAi.
                        if (!encryptParent) System.IO.File.Encrypt(encrypted);
                        report["PeerAttributes"] = System.IO.File.GetAttributes(plain).ToString();
                        if (volume == null) report["PlainAttributes"] = report["PeerAttributes"];
                        report["EncryptedAttributes"] = System.IO.File.GetAttributes(encrypted).ToString();
                        Assert.AreEqual(encryptParent ? FileAttributes.Encrypted : (FileAttributes)0,
                            System.IO.File.GetAttributes(plain) & FileAttributes.Encrypted);
                        Assert.AreEqual(FileAttributes.Encrypted, System.IO.File.GetAttributes(encrypted) & FileAttributes.Encrypted,
                            "The requested EFS sibling must actually be encrypted before any export.");
                        destinationDirectory = volume != null || location == "Encrypted" ? encrypted : plain;
                        report["SiblingDirectories"] = new[] { plain, encrypted };
                    }
                    else
                    {
                        Assert.IsTrue(Directory.Exists(baseDirectory), "The probe must not create or alter an existing storage parent.");
                        report["StorageParent"] = DescribeProbeDirectory(baseDirectory);
                        destinationDirectory = Path.Combine(baseDirectory, trial);
                        Assert.IsFalse(Directory.Exists(destinationDirectory)); Directory.CreateDirectory(destinationDirectory);
                        if (daclSource != null)
                        {
                            report["Stage"] = "SOURCE_ONLY_DACL_PREPARATION";
                            string source = daclSource == "UserTemporary" ? Path.GetTempPath() :
                                Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData);
                            CopyProbeDacl(baseDirectory, destinationDirectory, source, report, json);
                        }
                    }
                    report["ExportDirectory"] = DescribeProbeDirectory(destinationDirectory);
                    report["CurrentUserIdentity"] = System.Security.Principal.WindowsIdentity.GetCurrent().Name;
                    if (probeAncestor)
                    {
                        using (var testProcess = Process.GetCurrentProcess())
                            report["ProcessTokens"] = new { TestHost = NativeProcessTokenObservation.Read(testProcess.Id),
                                ExactOwnedExcel = NativeProcessTokenObservation.Read(host.ProcessId) };
                        string accessFile = Path.Combine(destinationDirectory, "probe-access.txt");
                        const string synthetic = "VBAi disposable inherited-storage probe. No macro or credentials.";
                        Assert.IsFalse(System.IO.File.Exists(accessFile));
                        System.IO.File.WriteAllText(accessFile, synthetic);
                        Assert.AreEqual(synthetic, System.IO.File.ReadAllText(accessFile));
                        report["SyntheticWriteReadVerified"] = true;
                        report["SyntheticEfsMetadata"] = ReadProbeEfsMetadata(accessFile);
                    }
                    string destination = Path.Combine(destinationDirectory, form + ".frm");
                    report["Destination"] = destination; report["DestinationLength"] = destination.Length;
                    report["DestinationContainsNonAscii"] = destination.Any(character => character > 127);
                    Exception exportFailure = null;
                    try
                    {
                        if (dispatch == "HostBridge")
                        {
                            var state = Data(host.Command(new { Command = "component_properties", Project = before["ProjectName"], Module = form }));
                            report["ComponentBefore"] = state;
                            report["NativeExportRequests"] = 1; report["Stage"] = "ONE_NATIVE_EXPORT_PENDING";
                            System.IO.File.WriteAllText(reportPath, json.Serialize(report));
                            var response = host.Command(new { Command = "export_component", Project = before["ProjectName"], Module = form,
                                ExpectedComponentVersion = state["Version"], Path = destination });
                            report["ExportResponse"] = response;
                            Data(response);
                        }
                        else
                        {
                            report["NativeExportRequests"] = 1; report["Stage"] = "ONE_NATIVE_EXPORT_PENDING";
                            System.IO.File.WriteAllText(reportPath, json.Serialize(report));
                            host.ExportGitFormOnce(form, destination);
                        }
                        var after = host.ReadGitExportContext(form); report["After"] = after;
                        foreach (var item in before) Assert.AreEqual(item.Value, after[item.Key], "Export changed identity/state: " + item.Key);
                        var nativeAfter = host.ReadGitLayout(form, "LabelButton"); report["NativeAfter"] = nativeAfter;
                        var nativeBefore = (IDictionary<string, object>)report["NativeBefore"];
                        foreach (var item in nativeBefore) Assert.AreEqual(item.Value, nativeAfter[item.Key], "Export changed native form: " + item.Key);
                        Assert.IsTrue(System.IO.File.Exists(destination), "Native Export did not create the FRM.");
                        Assert.IsTrue(System.IO.File.Exists(Path.ChangeExtension(destination, ".frx")), "Baseline native resources are missing.");
                        report["Stage"] = "EXPORT_VERIFIED_AWAITING_NORMAL_SHUTDOWN";
                    }
                    catch (Exception error) { exportFailure = error; throw; }
                    finally
                    {
                        // Keep every partial/successful raw file. Do not retry an
                        // export that failed, switch destinations or repair it.
                        try
                        {
                            report["RawFiles"] = Directory.GetFiles(destinationDirectory).Select(file => {
                                byte[] bytes = System.IO.File.ReadAllBytes(file);
                                System.IO.File.WriteAllBytes(Path.Combine(output, "raw-" + Path.GetFileName(file)), bytes);
                                using (var sha = SHA256.Create()) return (object)new { Path = file, Bytes = bytes.Length,
                                    Sha256 = BitConverter.ToString(sha.ComputeHash(bytes)).Replace("-", "") };
                            }).ToArray();
                        }
                        catch (Exception evidenceFailure)
                        {
                            if (exportFailure != null) throw new AggregateException("Native export and raw evidence retention both failed.", exportFailure, evidenceFailure);
                            throw;
                        }
                        // Public EFS certificate hashes only; no key material or
                        // certificate export. Metadata denial is recorded as denial,
                        // never interpreted as an absent key or a successful read.
                        try
                        {
                            report["EfsMetadata"] = Directory.GetFiles(destinationDirectory)
                                .Select(file => ReadProbeEfsMetadata(file)).ToArray();
                        }
                        catch (Exception metadataFailure)
                        {
                            if (exportFailure != null) throw new AggregateException("Native export and EFS evidence readback both failed.", exportFailure, metadataFailure);
                            throw;
                        }
                    }
                }, host => {
                    report["ShutdownDiagnostics"] = host.ShutdownDiagnostics;
                    Assert.AreEqual(host.ProcessId, Convert.ToInt32(host.ShutdownDiagnostics["ProcessId"]));
                    Assert.AreEqual(true, host.ShutdownDiagnostics["Exited"]);
                    Assert.AreEqual("0x00000000", host.ShutdownDiagnostics["ExitCodeHex"]);
                    report["NormalShutdownVerified"] = true;
                    report["NormalShutdownScope"] = "Exact owned Excel PID exited with code zero; Dispose succeeded before rethrowing any original scenario failure.";
                });
                report["NormalShutdownVerified"] = true; report["Stage"] = "PASS";
            }
            catch (Exception error) { report["Failure"] = error.ToString(); throw; }
            finally
            {
                if (observedHost?.ShutdownDiagnostics != null)
                    report["ShutdownDiagnostics"] = observedHost.ShutdownDiagnostics;
                System.IO.File.WriteAllText(reportPath, json.Serialize(report));
                TestContext.AddResultFile(reportPath);
            }
        }

        private static void CopyProbeDacl(string parent, string child, string source,
            IDictionary<string, object> report, JavaScriptSerializer json)
        {
            // Only this freshly created GUID child is writable. No privilege
            // adjustment, elevation, owner/SACL/EFS change or fallback is permitted.
            Assert.AreEqual(Path.GetFullPath(parent).TrimEnd(Path.DirectorySeparatorChar),
                Path.GetFullPath(Path.GetDirectoryName(child)).TrimEnd(Path.DirectorySeparatorChar), true);
            Assert.IsTrue(Guid.TryParseExact(Path.GetFileName(child), "N", out _));
            Assert.AreEqual(0, Directory.GetFileSystemEntries(child).Length);
            using (var process = Process.GetCurrentProcess())
            {
                var token = NativeProcessTokenObservation.Read(process.Id);
                report["DaclPreparationToken"] = token;
                Assert.AreEqual("READ", token["State"]);
                Assert.AreNotEqual(2, Convert.ToInt32(token["ElevationType"]), "This probe refuses an elevated token; do not elevate to apply a DACL.");
                string integrity = Convert.ToString(token["IntegritySid"]);
                Assert.IsTrue(int.TryParse(integrity.Substring(integrity.LastIndexOf('-') + 1), out int level) && level < 12288,
                    "This probe refuses high-integrity/system tokens; do not elevate to apply a DACL.");
            }
            string parentBefore = json.Serialize(DescribeProbeDirectory(parent));
            string sourceBefore = json.Serialize(DescribeProbeDirectory(source));
            var before = Directory.GetAccessControl(child, AccessControlSections.Owner | AccessControlSections.Access);
            var owner = (SecurityIdentifier)before.GetOwner(typeof(SecurityIdentifier));
            using (var identity = WindowsIdentity.GetCurrent()) Assert.AreEqual(identity.User, owner);
            FileAttributes attributes = System.IO.File.GetAttributes(child);
            report["DaclChildBefore"] = DescribeProbeDirectory(child);
            report["DaclParentBefore"] = DescribeProbeDirectory(parent);
            report["DaclSourceBefore"] = DescribeProbeDirectory(source);
            report["DaclEfsBefore"] = ReadProbeEfsMetadata(child);
            string access = Path.Combine(child, "dacl-before-access.txt");
            const string content = "VBAi disposable DACL control; no user data.";
            System.IO.File.WriteAllText(access, content);
            Assert.AreEqual(content, System.IO.File.ReadAllText(access));
            report["DaclSyntheticBeforeVerified"] = true;
            report["DaclSyntheticEfsBefore"] = ReadProbeEfsMetadata(access);

            var sourceAcl = Directory.GetAccessControl(source, AccessControlSections.Access);
            var replacement = new DirectorySecurity();
            replacement.SetAccessRuleProtection(true, false);
            bool currentFullControl = false;
            var copied = new List<object>();
            foreach (FileSystemAccessRule rule in sourceAcl.GetAccessRules(true, true, typeof(SecurityIdentifier)))
            {
                var sid = (SecurityIdentifier)rule.IdentityReference;
                replacement.AddAccessRule(new FileSystemAccessRule(sid, rule.FileSystemRights,
                    rule.InheritanceFlags, rule.PropagationFlags, rule.AccessControlType));
                copied.Add(new { Sid = sid.Value, Rights = rule.FileSystemRights.ToString(),
                    Inheritance = rule.InheritanceFlags.ToString(), Propagation = rule.PropagationFlags.ToString(),
                    Type = rule.AccessControlType.ToString(), SourceWasInherited = rule.IsInherited });
                if (sid.Equals(owner) && rule.AccessControlType == AccessControlType.Allow &&
                    (rule.PropagationFlags & PropagationFlags.InheritOnly) == 0 &&
                    (rule.FileSystemRights & FileSystemRights.FullControl) == FileSystemRights.FullControl)
                    currentFullControl = true;
            }
            Assert.IsTrue(currentFullControl, "Source-only DACL must already allow current SID FullControl; no extra ACE will be inserted.");
            report["SourceCurrentSidFullControlVerified"] = true;
            report["DaclPermissionEscalations"] = 0;
            report["CopiedSourceRules"] = copied;
            report["ExpectedProtectedDaclSddl"] = replacement.GetSecurityDescriptorSddlForm(AccessControlSections.Access);
            Directory.SetAccessControl(child, replacement);
            var after = Directory.GetAccessControl(child, AccessControlSections.Owner | AccessControlSections.Access);
            report["DaclChildAfter"] = DescribeProbeDirectory(child);
            Assert.IsTrue(after.AreAccessRulesProtected);
            Assert.AreEqual(report["ExpectedProtectedDaclSddl"], after.GetSecurityDescriptorSddlForm(AccessControlSections.Access));
            Assert.AreEqual(owner, after.GetOwner(typeof(SecurityIdentifier)));
            Assert.AreEqual(attributes, System.IO.File.GetAttributes(child), "DACL copy must not change EFS or any other attributes.");
            Assert.AreEqual(parentBefore, json.Serialize(DescribeProbeDirectory(parent)), "Existing target parent changed.");
            Assert.AreEqual(sourceBefore, json.Serialize(DescribeProbeDirectory(source)), "Existing source parent changed.");
            report["DaclEfsAfter"] = ReadProbeEfsMetadata(child);
            report["DaclSyntheticEfsAfter"] = ReadProbeEfsMetadata(access);
            Assert.AreEqual(json.Serialize(report["DaclEfsBefore"]), json.Serialize(report["DaclEfsAfter"]));
            Assert.AreEqual(json.Serialize(report["DaclSyntheticEfsBefore"]), json.Serialize(report["DaclSyntheticEfsAfter"]));
            Assert.AreEqual(content, System.IO.File.ReadAllText(access));
            report["DaclSyntheticAfterVerified"] = true;
            report["DaclMutationScope"] = "Only fresh GUID child DACL: source ACEs copied as explicit, inheritance protected. Owner/SACL/token/parents/EFS unchanged. No elevation or permission retry.";
        }

        private static IDictionary<string, object> Data(IDictionary<string, object> response)
        {
            Assert.IsNotNull(response, "Owned host bridge did not answer.");
            object error; response.TryGetValue("Error", out error);
            Assert.AreEqual(true, response["Ok"], Convert.ToString(error));
            return VbeBridgeClient.Object(response["Data"]);
        }

        private static object DescribeProbeDirectory(string path)
        {
            var acl = Directory.GetAccessControl(path, AccessControlSections.Owner | AccessControlSections.Access);
            return new { Path = Path.GetFullPath(path), Attributes = System.IO.File.GetAttributes(path).ToString(),
                Sddl = acl.GetSecurityDescriptorSddlForm(AccessControlSections.Owner | AccessControlSections.Access),
                acl.AreAccessRulesProtected };
        }

        [StructLayout(LayoutKind.Sequential)] private struct EfsHashList { internal uint Count; internal IntPtr Users; }
        [StructLayout(LayoutKind.Sequential)] private struct EfsUserHash { internal uint Length; internal IntPtr Sid, Hash, Display; }
        [StructLayout(LayoutKind.Sequential)] private struct EfsHashBlob { internal uint Length; internal IntPtr Data; }
        [DllImport("advapi32.dll", CharSet = CharSet.Unicode)] private static extern uint QueryUsersOnEncryptedFile(string path, out IntPtr users);
        [DllImport("advapi32.dll")] private static extern void FreeEncryptionCertificateHashList(IntPtr users);

        private static object ReadProbeEfsMetadata(string path)
        {
            var result = new Dictionary<string, object> { ["Path"] = path, ["Attributes"] = System.IO.File.GetAttributes(path).ToString() };
            if ((System.IO.File.GetAttributes(path) & FileAttributes.Encrypted) == 0) { result["State"] = "NOT_ENCRYPTED"; return result; }
            IntPtr users;
            uint error = QueryUsersOnEncryptedFile(path, out users);
            result["Win32Error"] = error;
            if (error != 0) { result["State"] = "UNREADABLE"; return result; }
            try
            {
                var list = (EfsHashList)Marshal.PtrToStructure(users, typeof(EfsHashList));
                Assert.IsTrue(list.Count <= 16, "EFS metadata exceeded the owned-probe bound.");
                var hashes = new List<string>();
                for (int i = 0; i < list.Count; i++)
                {
                    var user = (EfsUserHash)Marshal.PtrToStructure(Marshal.ReadIntPtr(list.Users, i * IntPtr.Size), typeof(EfsUserHash));
                    var hash = (EfsHashBlob)Marshal.PtrToStructure(user.Hash, typeof(EfsHashBlob));
                    Assert.IsTrue(hash.Length <= 128, "EFS certificate hash exceeded the owned-probe bound.");
                    var bytes = new byte[hash.Length]; Marshal.Copy(hash.Data, bytes, 0, bytes.Length);
                    hashes.Add(BitConverter.ToString(bytes).Replace("-", ""));
                }
                result["State"] = "READ"; result["CertificateHashes"] = hashes;
            }
            finally { FreeEncryptionCertificateHashList(users); }
            return result;
        }
    }
}

using System;
using System.Collections;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text.RegularExpressions;
using System.Threading.Tasks;

namespace VBAi
{
    internal sealed partial class VbeProjectComponents
    {
        internal sealed class SolidWorksMacroPublicationResult
        {
            public bool Terminal, Uncertain, MutationInvoked, Verified, OriginalPreserved, IdentityChanged, DestinationCreated;
            public string OriginalProject, OriginalHostPath, OriginalProjectVersion, DestinationProject, HostPath,
                ProjectVersion, CollectionVersion, StagingPath, Error;
            public object Creation, Save, Components, References, AddedNativeHostReferences, NativeGeneratedDefaults;
            public PublicationImportMismatch ImportMismatch, FinalMismatch;
            private readonly List<MacroMutationClaim> claims = new List<MacroMutationClaim>();
            public MacroMutationClaim[] Claims => claims.ToArray();
            internal void AddClaim(MacroMutationClaim claim) { claims.Add(claim); }
            public bool RetryAllowed => false;
            public bool RollbackPerformed => false;
            public bool PersistenceReloadVerified => false;
            public string Limit => "Explicit publication to a new native SOLIDWORKS project; source identity is retained. Nondefault help and conditional-compilation metadata and project references are unsupported. Digital signatures are not transported or qualified; the new identity requires a separate signing decision. Reopen is required to qualify persisted contents.";
        }

        internal sealed class PublicationComponent
        {
            internal string Name, Code, FormVersion, Version, Attributes, ExportPath, ExportSha256, FrxSha256, DesignerJson, ComponentSnapshotJson;
            internal int Type;
            internal object Canonical;
        }
        // Only scalars cross the result boundary; full code/designer observations stay in owned staging.
        internal sealed class PublicationImportMismatch
        {
            public string Phase { get; }
            public int ComponentOrdinal { get; }
            public bool CanonicalEqual { get; }
            public bool VersionEqual { get; }
            public string ExpectedVersion { get; }
            public string ActualVersion { get; }
            public bool NameEqual { get; }
            public bool TypeEqual { get; }
            public bool CodeEqual { get; }
            public bool FormVersionEqual { get; }
            public int ExpectedType { get; }
            public int ActualType { get; }
            public int ExpectedCharacters { get; }
            public int ActualCharacters { get; }
            public int ExpectedLines { get; }
            public int ActualLines { get; }
            public int ExpectedLeadingEmptyLines { get; }
            public int ActualLeadingEmptyLines { get; }
            public int FirstDifferentCodeLine { get; }
            public string ExpectedCodeSha256 { get; }
            public string ActualCodeSha256 { get; }
            public string ExpectedFormVersion { get; }
            public string ActualFormVersion { get; }
            public string ExpectedRelativeFile { get; }
            public string ActualRelativeFile { get; }
            public string ExpectedFileSha256 { get; }
            public string ActualFileSha256 { get; }
            public string DiagnosticError { get; }
            internal PublicationImportMismatch(PublicationComponent expected, PublicationComponent actual,
                string expectedFile, string actualFile, string expectedHash, string actualHash, string error,
                string phase = null, int ordinal = 0, bool canonicalEqual = true)
            {
                Phase = phase; ComponentOrdinal = ordinal; CanonicalEqual = canonicalEqual;
                VersionEqual = expected.Version == actual.Version; ExpectedVersion = expected.Version; ActualVersion = actual.Version;
                NameEqual = string.Equals(expected.Name, actual.Name, StringComparison.Ordinal);
                TypeEqual = expected.Type == actual.Type;
                CodeEqual = string.Equals(expected.Code, actual.Code, StringComparison.Ordinal);
                FormVersionEqual = string.Equals(expected.FormVersion, actual.FormVersion, StringComparison.Ordinal);
                ExpectedType = expected.Type; ActualType = actual.Type;
                string left = expected.Code ?? "", right = actual.Code ?? "";
                ExpectedCharacters = left.Length; ActualCharacters = right.Length;
                string[] l = left.Length == 0 ? new string[0] : left.Split('\n');
                string[] r = right.Length == 0 ? new string[0] : right.Split('\n');
                ExpectedLines = l.Length; ActualLines = r.Length;
                ExpectedLeadingEmptyLines = l.TakeWhile(x => x == "" || x == "\r").Count();
                ActualLeadingEmptyLines = r.TakeWhile(x => x == "" || x == "\r").Count();
                FirstDifferentCodeLine = -1;
                for (int i = 0; i < Math.Max(l.Length, r.Length); i++)
                    if (i >= l.Length || i >= r.Length || !string.Equals(l[i], r[i], StringComparison.Ordinal))
                    { FirstDifferentCodeLine = i + 1; break; }
                ExpectedCodeSha256 = Hash(left); ActualCodeSha256 = Hash(right);
                ExpectedFormVersion = expected.FormVersion; ActualFormVersion = actual.FormVersion;
                ExpectedRelativeFile = expectedFile; ActualRelativeFile = actualFile;
                ExpectedFileSha256 = expectedHash; ActualFileSha256 = actualHash; DiagnosticError = error;
            }
        }
        internal static PublicationImportMismatch CapturePublicationImportMismatch(
            PublicationComponent expected, PublicationComponent actual, string staging, int ordinal,
            string phase = null, bool canonicalEqual = true)
        {
            if (expected == null || actual == null) throw new ArgumentNullException();
            string expectedFile = null, actualFile = null, expectedHash = null, actualHash = null, error = null;
            try
            {
                if (ordinal < 1 || ordinal > 4096 || string.IsNullOrWhiteSpace(staging) ||
                    !Path.IsPathRooted(staging) || !Directory.Exists(staging) ||
                    (File.GetAttributes(staging) & FileAttributes.ReparsePoint) != 0)
                    throw new IOException("Existing canonical owned staging is required.");
                var serializer = new System.Web.Script.Serialization.JavaScriptSerializer { MaxJsonLength = 16 * 1024 * 1024 };
                Func<PublicationComponent, string> observation = c => serializer.Serialize(new {
                    c.Name, c.Type, c.Code, c.FormVersion, c.Version,
                    ComponentSnapshot = c.ComponentSnapshotJson == null ? null : serializer.DeserializeObject(c.ComponentSnapshotJson),
                    Designer = c.DesignerJson == null ? null : serializer.DeserializeObject(c.DesignerJson) });
                if (phase != null && phase != "BeforeSave" && phase != "AfterSave") throw new ArgumentException("Unsupported diagnostic phase.");
                string prefix = (phase == null ? "import-mismatch-" : "final-mismatch-" + phase + "-") + ordinal.ToString("D4", System.Globalization.CultureInfo.InvariantCulture);
                expectedFile = prefix + "-expected.json";
                expectedHash = WritePublicationMismatchFile(staging, expectedFile, observation(expected));
                actualFile = prefix + "-actual.json";
                actualHash = WritePublicationMismatchFile(staging, actualFile, observation(actual));
            }
            catch (Exception failure)
            {
                // No file/code/project identifiers in the public diagnostic error or global claims.
                error = "Local import diagnostic could not be completed: " + failure.GetType().FullName;
            }
            return new PublicationImportMismatch(expected, actual, expectedFile, actualFile, expectedHash, actualHash, error, phase, ordinal, canonicalEqual);
        }
        private static string WritePublicationMismatchFile(string staging, string relative, string text)
        {
            const int maximum = 16 * 1024 * 1024;
            if (text == null || text.Length > maximum) throw new IOException("Diagnostic size limit exceeded.");
            byte[] bytes = new System.Text.UTF8Encoding(false, true).GetBytes(text);
            if (bytes.Length > maximum) throw new IOException("Diagnostic byte limit exceeded.");
            using (var output = new FileStream(Path.Combine(staging, relative), FileMode.CreateNew,
                FileAccess.Write, FileShare.Read)) { output.Write(bytes, 0, bytes.Length); output.Flush(true); }
            using (var sha = System.Security.Cryptography.SHA256.Create())
                return BitConverter.ToString(sha.ComputeHash(bytes)).Replace("-", "").ToLowerInvariant();
        }
        internal sealed class PublicationReference
        {
            public string Guid;
            public int Major, Minor;
            public bool BuiltIn;
            internal string Key => Guid.ToUpperInvariant() + "|" + Major + "|" + Minor + "|" + BuiltIn;
        }
        private sealed class PublicationSource
        {
            internal object Canonical;
            internal string Name, Path, DiskSha256, Version, Fingerprint, Description;
            internal bool Saved;
            internal List<PublicationComponent> Components;
            internal List<PublicationReference> References;
        }

        // Injection points preserve actual asynchronous creation/save contracts in focused fault tests.
        internal Func<Request, Action<bool>, Action<MacroMutationClaim>, Action, Task<SolidWorksMacroCreationResult>>
            PublicationCreate;
        internal Func<Request, Task<object>> PublicationSave;
        internal Func<Request, Task<object>> PublicationReadGeneral;
        internal Func<string, string, object> PublicationFormTree;

        internal async Task<object> PublishSolidWorksMacroAsync(Request request, Action<bool> authorization = null,
            Action<MacroMutationClaim> recordClaim = null, Action requireNativeContext = null)
        {
            if (request == null || request.ExpectedMode != 2 || string.IsNullOrWhiteSpace(request.Project))
                throw new ArgumentException("An explicit source project and design mode are required.");
            string destinationPath = RequireFreshSolidWorksMacroPath(request.Path);
            Action<bool> authorize = authorization ?? (_ => { });
            Action context = requireNativeContext ?? (() => { });
            Action guard = () => { context(); authorize(true); context(); };
            guard();
            var owner = SolidWorksSaveProbe();
            if (!owner.IsSolidWorks) throw new InvalidOperationException("Publication requires the current SOLIDWORKS process.");
            owner.RequireOwner((object)vbe);
            var frozen = new Request { Project = request.Project, ExpectedProjectVersion = request.ExpectedProjectVersion,
                ExpectedMode = 2, Path = destinationPath };
            object canonical = (object)GetDesignProject(frozen.Project);
            AssertProjectVersion(frozen, canonical);
            var source = ReadPublicationSource(canonical);
            var readGeneral = PublicationReadGeneral ?? throw new InvalidOperationException("An actual asynchronous native General reader is required before publication.");
            if (!string.Equals(source.Name, frozen.Project, StringComparison.Ordinal))
                throw new InvalidOperationException("An exact source name is required; aliases cannot select publication.");
            var beforeCollection = ReadLifecycleCollection();
            RequireLifecycleDesign(beforeCollection);
            string collectionVersion = LifecycleVersion(beforeCollection);
            string targetName = Path.GetFileNameWithoutExtension(destinationPath);
            if (beforeCollection.Any(p => string.Equals(p.Name, targetName, StringComparison.OrdinalIgnoreCase)))
                throw new InvalidOperationException("The destination basename collides with an existing project identity.");
            var result = new SolidWorksMacroPublicationResult { OriginalProject = source.Name, OriginalHostPath = source.Path,
                OriginalProjectVersion = source.Version, HostPath = destinationPath, OriginalPreserved = true };
            int ordinal = 0;
            bool callPending = false, sourceChecked = false;
            Action<string> claim = phase => {
                var entry = new MacroMutationClaim(phase, destinationPath, ++ordinal);
                result.AddClaim(entry); recordClaim?.Invoke(entry);
            };
            Action<Action> settledCall = action => { callPending = true; try { action(); } finally { callPending = false; } };
            Action sourceGuard = () => {
                sourceChecked = false; guard();
                if (!GeneralProjectIdentity(source.Canonical, (object)GetDesignProject(source.Name)))
                    throw new InvalidOperationException("Original canonical source project changed.");
                var current = ReadPublicationSource(source.Canonical);
                if (current.Fingerprint != source.Fingerprint || current.Components.Count != source.Components.Count ||
                    current.Components.Where((item, index) => !GeneralProjectIdentity(item.Canonical, source.Components[index].Canonical)).Any())
                    throw new InvalidOperationException("Original source contents, metadata, designer, references or persistence changed.");
                authorize(false); context(); sourceChecked = true;
            };
            string generalBaseline = null;
            Func<Task> generalGuard = async () => {
                sourceGuard(); sourceChecked = false;
                var generalRequest = new Request { Command = "read_project_general", Project = source.Name,
                    ExpectedProjectVersion = source.Version, ExpectedMode = 2 };
                callPending = true;
                object observed = await readGeneral(generalRequest);
                var fields = json.Deserialize<Dictionary<string, object>>(json.Serialize(observed));
                if (fields != null && (PublicationGeneralBoolean(fields, "MutationInvoked", true) ||
                    PublicationGeneralBoolean(fields, "CommittedRequested", true) ||
                    fields.ContainsKey("FieldAttempts") && fields["FieldAttempts"] is int && (int)fields["FieldAttempts"] > 0 ||
                    fields.ContainsKey("OkAttempts") && fields["OkAttempts"] is int && (int)fields["OkAttempts"] > 0))
                    result.MutationInvoked = true;
                // A returned Task is insufficient: the original modal Execute and Cancel must both settle.
                if (fields == null || !PublicationGeneralBoolean(fields, "Terminal", true) ||
                    !PublicationGeneralBoolean(fields, "OriginalExecuteReturned", true) ||
                    !PublicationGeneralBoolean(fields, "DialogClosed", true) ||
                    !PublicationGeneralBoolean(fields, "Uncertain", false))
                    throw new InvalidOperationException("Original native General inspection is unsettled or uncertain; no replay.");
                callPending = false;
                if (!PublicationGeneralBoolean(fields, "Available", true) ||
                    !PublicationGeneralBoolean(fields, "MutationInvoked", false) ||
                    !PublicationGeneralBoolean(fields, "CommittedRequested", false) ||
                    !fields.ContainsKey("Error") || fields["Error"] != null ||
                    !PublicationGeneralCount(fields, "OpenAttempts", 1) ||
                    !PublicationGeneralCount(fields, "CancelAttempts", 1) ||
                    !PublicationGeneralCount(fields, "FieldAttempts", 0) ||
                    !PublicationGeneralCount(fields, "OkAttempts", 0))
                    throw new InvalidOperationException("Native General inspection did not prove an unchanged read and one Cancel.");
                var metadata = new Dictionary<string, object>();
                foreach (string key in new[] { "OptionsVersion", "Name", "Description", "HelpFile", "HelpContextText", "ConditionalCompilation" })
                {
                    object value;
                    if (!fields.TryGetValue(key, out value) || !(value is string))
                        throw new InvalidOperationException("Native General metadata is missing or unreadable.");
                    metadata.Add(key, value);
                }
                if (string.IsNullOrWhiteSpace((string)metadata["OptionsVersion"]) ||
                    (string)metadata["Name"] != source.Name || (string)metadata["Description"] != (source.Description ?? "") ||
                    !string.IsNullOrEmpty((string)metadata["HelpFile"]) || RequireGeneralInt32(metadata["HelpContextText"]) != 0 ||
                    !string.IsNullOrEmpty((string)metadata["ConditionalCompilation"]))
                    throw new InvalidOperationException("Nondefault or inconsistent native General metadata cannot be published by this route.");
                string fingerprint = json.Serialize(metadata);
                if (generalBaseline != null && generalBaseline != fingerprint)
                    throw new InvalidOperationException("Original native General values or options version changed during publication.");
                sourceGuard();
                if (generalBaseline == null) generalBaseline = fingerprint;
            };
            try
            {
                await generalGuard();
                sourceGuard();
                string staging = Path.Combine(Path.GetDirectoryName(destinationPath),
                    ".vbai-publication-" + System.Guid.NewGuid().ToString("N"));
                if (Directory.Exists(staging) || File.Exists(staging)) throw new IOException("Fresh publication staging required.");
                claim("BeforeStaging");
                Directory.CreateDirectory(staging);
                result.StagingPath = staging; // Retain all exports after any outcome; never delete or retry.
                foreach (var component in source.Components)
                {
                    sourceGuard();
                    component.ExportPath = Path.Combine(staging, component.Name + PublicationExtension(component.Type));
                    claim("BeforeExport");
                    result.MutationInvoked = true; settledCall(() => ((dynamic)component.Canonical).Export(component.ExportPath));
                    if (!File.Exists(component.ExportPath)) throw new IOException("Export returned without its original file.");
                    component.ExportSha256 = PublicationFileHash(component.ExportPath);
                    component.Attributes = PublicationAttributes(File.ReadAllText(component.ExportPath));
                    string frx = Path.ChangeExtension(component.ExportPath, ".frx");
                    if (component.Type == 3)
                    {
                        if (!File.Exists(frx)) throw new IOException("A UserForm export requires its exact FRX companion.");
                        component.FrxSha256 = PublicationFileHash(frx);
                    }
                    sourceGuard();
                }
                // Export is read-only to the VBE source. Revalidate the complete collection immediately before A.
                await generalGuard();
                sourceGuard();
                if (LifecycleVersion(ReadLifecycleCollection()) != collectionVersion)
                    throw new InvalidOperationException("The original project collection changed before native creation.");
                RequireFreshSolidWorksMacroPath(destinationPath);
                claim("BeforeNativeCreation");
                var createRequest = new Request { Path = destinationPath, ExpectedMode = 2, ExpectedProjectVersion = collectionVersion };
                Action<bool> creationAuthorization = shared => { if (shared) sourceGuard(); authorize(shared); context(); };
                callPending = true; sourceChecked = false; result.OriginalPreserved = false;
                var creation = await (PublicationCreate == null
                    ? CreateSolidWorksMacroAsync(createRequest, creationAuthorization, c => claim("NativeCreation"), context)
                    : PublicationCreate(createRequest, creationAuthorization, c => claim("NativeCreation"), context));
                callPending = false;
                result.Creation = creation;
                result.MutationInvoked |= creation.CommandEntered;
                result.DestinationCreated = creation.DestinationCreated;
                result.DestinationProject = creation.Project;
                if (!creation.Terminal || creation.CommandEntered && !creation.OriginalCommandReturned)
                {
                    result.Terminal = false; result.Uncertain = true; result.Error = creation.Error;
                    return result; // No read/recovery following an unsettled native command.
                }
                if (!creation.Verified || creation.Uncertain)
                {
                    result.Terminal = true; result.Uncertain = creation.Uncertain; result.Error = creation.Error;
                    return result;
                }
                await generalGuard();
                sourceGuard();
                object destination = (object)GetDesignProject(creation.Project);
                Action destinationGuard = () => {
                    sourceGuard();
                    if (GeneralProjectIdentity(source.Canonical, destination) ||
                        !GeneralProjectIdentity(destination, (object)GetDesignProject(creation.Project)) ||
                        (int)((dynamic)destination).Type != 100 || (int)((dynamic)destination).Protection != 0 ||
                        !string.Equals(StandaloneAwareProjectPath(destination), destinationPath, StringComparison.OrdinalIgnoreCase))
                        throw new InvalidOperationException("New native destination identity, type or path changed.");
                    authorize(false); context();
                };
                destinationGuard();
                result.IdentityChanged = true;
                if (!string.IsNullOrWhiteSpace(creation.CollectionVersion) &&
                    LifecycleVersion(ReadLifecycleCollection()) != creation.CollectionVersion)
                    throw new InvalidOperationException("Fresh creation collection baseline changed before publication.");
                var initial = ReadPublicationDestinationComponents(destination);
                if (initial.Count(c => c.Name == "ThisLibrary" && c.Type == 100) != 1 ||
                    initial.Count(c => c.Name != "ThisLibrary") > 1 ||
                    initial.Any(c => c.Name != "ThisLibrary" && c.Type != 1) ||
                    initial.Select(c => c.Name).Distinct(StringComparer.OrdinalIgnoreCase).Count() != initial.Count)
                    throw new InvalidOperationException("Unsupported native creation baseline; nothing will be removed.");
                var defaults = initial.Where(c => c.Name != "ThisLibrary").Select(c =>
                    ReadPublicationComponent(creation.Project, (object)GetComponent(destination, c.Name))).ToArray();
                object library = (object)GetComponent(destination, "ThisLibrary");
                string libraryVersion = (string)((dynamic)ComponentSnapshot(creation.Project, (dynamic)library)).Version;
                var existingReferences = ReadPublicationReferences(destination);
                var nativeReferences = existingReferences.ToArray();
                foreach (var reference in source.References)
                {
                    var sameGuid = existingReferences.Where(r => r.Guid == reference.Guid).ToArray();
                    if (sameGuid.Length != 0 && sameGuid[0].Key != reference.Key ||
                        reference.BuiltIn && sameGuid.Length == 0)
                        throw new InvalidOperationException("Source and native reference identity/version/built-in flags conflict.");
                }
                result.AddedNativeHostReferences = nativeReferences.Where(r =>
                    !source.References.Any(s => s.Key == r.Key)).ToArray();
                foreach (var reference in source.References.Where(r => !r.BuiltIn))
                {
                    destinationGuard();
                    if (existingReferences.Any(r => r.Key == reference.Key)) continue;
                    claim("BeforeReference");
                    result.MutationInvoked = true; settledCall(() => ((dynamic)destination).References.AddFromGuid(reference.Guid, reference.Major, reference.Minor));
                    existingReferences = ReadPublicationReferences(destination);
                    if (!existingReferences.Any(r => r.Key == reference.Key))
                        throw new InvalidOperationException("The requested reference version was not retained.");
                }
                var unionReferences = nativeReferences.Concat(source.References).GroupBy(r => r.Key).Select(g => g.First()).ToArray();
                if (!ReferenceSet(existingReferences).SequenceEqual(ReferenceSet(unionReferences)))
                    throw new InvalidOperationException("Native destination differs from the baseline plus source references.");
                foreach (var generated in defaults)
                {
                    destinationGuard();
                    var liveDefault = ReadPublicationComponent(creation.Project, (object)GetComponent(destination, generated.Name));
                    if (!GeneralProjectIdentity(generated.Canonical, liveDefault.Canonical) ||
                        liveDefault.Version != generated.Version || liveDefault.Code != generated.Code)
                        throw new InvalidOperationException("Fresh generated default changed; it will not be removed.");
                    string baselinePath = Path.Combine(staging, "native-default-" + generated.Name + ".bas");
                    claim("BeforeGeneratedDefaultExport");
                    settledCall(() => ((dynamic)generated.Canonical).Export(baselinePath));
                    generated.ExportSha256 = PublicationFileHash(baselinePath);
                    generated.Attributes = PublicationAttributes(File.ReadAllText(baselinePath));
                    destinationGuard();
                    liveDefault = ReadPublicationComponent(creation.Project, (object)GetComponent(destination, generated.Name));
                    if (!GeneralProjectIdentity(generated.Canonical, liveDefault.Canonical) ||
                        liveDefault.Version != generated.Version || liveDefault.Code != generated.Code)
                        throw new InvalidOperationException("Generated default changed after baseline export.");
                    claim("BeforeGeneratedDefaultRemoval");
                    result.MutationInvoked = true;
                    settledCall(() => ((dynamic)destination).VBComponents.Remove((dynamic)generated.Canonical));
                    if (ReadPublicationDestinationComponents(destination).Any(c => c.Name == generated.Name))
                        throw new InvalidOperationException("Single generated default removal was not verified.");
                }
                var published = new List<object>();
                var importedComponents = new List<PublicationComponent>();
                foreach (var component in source.Components)
                {
                    destinationGuard();
                    VerifyPublicationExport(component);
                    if (ReadPublicationDestinationComponents(destination).Any(c => c.Name.Equals(component.Name, StringComparison.OrdinalIgnoreCase)))
                        throw new InvalidOperationException("A destination component collision prevents import.");
                    claim("BeforeImport");
                    result.MutationInvoked = true; object imported = null;
                    settledCall(() => imported = ((dynamic)destination).VBComponents.Import(component.ExportPath));
                    destinationGuard();
                    var actual = ReadPublicationComponent(creation.Project, imported);
                    string verifyPath = Path.Combine(staging, "verified-" + component.Name + PublicationExtension(component.Type));
                    claim("BeforeVerificationExport");
                    settledCall(() => ((dynamic)imported).Export(verifyPath));
                    destinationGuard();
                    string verifyFrx = Path.ChangeExtension(verifyPath, ".frx");
                    if (component.Type == 3 && !File.Exists(verifyFrx))
                        throw new IOException("Imported UserForm has no exported resource companion.");
                    bool designerEqual;
                    try { designerEqual = component.Type != 3 || PublicationDesignerContentEquals(component, actual, component.ExportPath, verifyPath); }
                    catch
                    {
                        result.ImportMismatch = CapturePublicationImportMismatch(component, actual, staging, importedComponents.Count + 1);
                        throw;
                    }
                    int surplus = component.Type == 3 ? PublicationSurplusImportedCrLfPrefix(component.Code, actual.Code) : 0;
                    if (actual.Name != component.Name || actual.Type != component.Type || !designerEqual ||
                        (actual.Code != component.Code && surplus == 0))
                    {
                        result.ImportMismatch = CapturePublicationImportMismatch(component, actual, staging, importedComponents.Count + 1);
                        throw new InvalidOperationException("Imported component code or complete designer/picture state differs.");
                    }
                    if (PublicationAttributes(File.ReadAllText(verifyPath)) != component.Attributes)
                        throw new InvalidOperationException("Imported hidden class/form attributes differ.");
                    if (surplus > 0)
                    {
                        destinationGuard(); VerifyPublicationExport(component);
                        Action requireImportedUnchanged = () => {
                            destinationGuard();
                            object selected = (object)GetComponent(destination, component.Name);
                            var current = ReadPublicationComponent(creation.Project, selected);
                            if (!GeneralProjectIdentity(imported, selected) || current.Name != actual.Name || current.Type != actual.Type ||
                                current.Code != actual.Code || current.Version != actual.Version || current.FormVersion != actual.FormVersion)
                                throw new InvalidOperationException("Imported canonical form or revision changed before prefix reconciliation.");
                        };
                        requireImportedUnchanged();
                        claim("BeforeImportedCodePrefixReconciliation");
                        requireImportedUnchanged();
                        result.MutationInvoked = true;
                        settledCall(() => ((dynamic)imported).CodeModule.DeleteLines(1, surplus));
                        destinationGuard();
                        var reconciled = ReadPublicationComponent(creation.Project, imported);
                        if (reconciled.Name != component.Name || reconciled.Type != component.Type || reconciled.Code != component.Code)
                            throw new InvalidOperationException("Imported form prefix reconciliation readback differs.");
                        actual = reconciled;
                        verifyPath = Path.Combine(staging, "verified-reconciled-" + component.Name + PublicationExtension(component.Type));
                        claim("BeforeReconciledVerificationExport");
                        settledCall(() => ((dynamic)imported).Export(verifyPath));
                        destinationGuard();
                        verifyFrx = Path.ChangeExtension(verifyPath, ".frx");
                        if (!PublicationDesignerContentEquals(component, actual, component.ExportPath, verifyPath) ||
                            PublicationAttributes(File.ReadAllText(verifyPath)) != component.Attributes)
                            throw new InvalidOperationException("Reconciled persisted form content or hidden attributes differ.");
                    }
                    importedComponents.Add(actual);
                    // Verify semantic Designer/resource fidelity; retain original and verification hashes without claiming byte-identical FRX.
                    published.Add(new { component.Name, component.Type, OriginalExportSha256 = component.ExportSha256,
                        OriginalFrxSha256 = component.FrxSha256, VerifiedExportSha256 = PublicationFileHash(verifyPath),
                        VerifiedFrxSha256 = component.Type == 3 ? PublicationFileHash(verifyFrx) : null,
                        DesignerAndPictureStateVerified = component.Type == 3, HiddenAttributesVerified = true });
                }
                destinationGuard();
                claim("BeforeDescription");
                result.MutationInvoked = true; settledCall(() => ((dynamic)destination).Description = source.Description);
                if ((string)((dynamic)destination).Description != source.Description)
                    throw new InvalidOperationException("Published description readback differs.");
                Action<bool> verifyPublished = afterSave => {
                    destinationGuard();
                    var inventory = ReadPublicationDestinationComponents(destination);
                    if (inventory.Count != source.Components.Count + 1 ||
                        !GeneralProjectIdentity(library, (object)GetComponent(destination, "ThisLibrary")) ||
                        (string)((dynamic)ComponentSnapshot(creation.Project, (dynamic)library)).Version != libraryVersion ||
                        !ReferenceSet(ReadPublicationReferences(destination)).SequenceEqual(ReferenceSet(unionReferences)) ||
                        (string)((dynamic)destination).Description != source.Description)
                        throw new InvalidOperationException("Final native library, inventory, references or description differs.");
                    for (int verifiedOrdinal = 0; verifiedOrdinal < importedComponents.Count; verifiedOrdinal++)
                    {
                        var imported = importedComponents[verifiedOrdinal];
                        var current = ReadPublicationComponent(creation.Project, (object)GetComponent(destination, imported.Name));
                        bool canonicalEqual = GeneralProjectIdentity(current.Canonical, imported.Canonical);
                        try
                        {
                            if (!canonicalEqual || current.Name != imported.Name || current.Type != imported.Type || current.Code != imported.Code ||
                                (!afterSave && (current.Version != imported.Version || current.FormVersion != imported.FormVersion)) ||
                                (afterSave && !PublicationPostSaveMetadataEquals(imported, current)))
                                throw new InvalidOperationException("A published component changed before final verification.");
                            if (afterSave)
                            {
                                var original = source.Components.Single(c => c.Name == imported.Name);
                                string verificationPath = Path.Combine(staging, "verified-after-save-" + imported.Name + PublicationExtension(imported.Type));
                                var saveObservation = current;
                                Action requireSavedObservationUnchanged = () => {
                                    destinationGuard();
                                    current = ReadPublicationComponent(creation.Project, (object)GetComponent(destination, imported.Name));
                                    canonicalEqual = GeneralProjectIdentity(current.Canonical, imported.Canonical);
                                    if (!GeneralProjectIdentity(current.Canonical, saveObservation.Canonical) || current.Name != saveObservation.Name ||
                                        current.Type != saveObservation.Type || current.Code != saveObservation.Code || current.Version != saveObservation.Version ||
                                        current.FormVersion != saveObservation.FormVersion)
                                        throw new InvalidOperationException("Saved component changed around verification export.");
                                };
                                destinationGuard(); claim("BeforePostSaveVerificationExport"); requireSavedObservationUnchanged();
                                settledCall(() => ((dynamic)current.Canonical).Export(verificationPath));
                                requireSavedObservationUnchanged();
                                if ((current.Type == 3 && !PublicationDesignerContentEquals(original, current, original.ExportPath, verificationPath)) ||
                                    PublicationAttributes(File.ReadAllText(verificationPath)) != original.Attributes)
                                    throw new InvalidOperationException("Saved component persisted content or hidden attributes differs.");
                            }
                        }
                        catch
                        {
                            result.FinalMismatch = CapturePublicationImportMismatch(imported, current, staging, verifiedOrdinal + 1,
                                afterSave ? "AfterSave" : "BeforeSave", canonicalEqual);
                            throw;
                        }
                    }
                };
                verifyPublished(false);
                await generalGuard();
                destinationGuard();
                dynamic destinationState = ProjectProperties(creation.Project);
                var saveRequest = new Request { Project = creation.Project, ExpectedMode = 2,
                    ExpectedProjectVersion = (string)destinationState.Version, ExpectedHostPath = destinationPath };
                claim("BeforeSave");
                result.MutationInvoked = true; callPending = true;
                object saved = await (PublicationSave == null
                    ? SaveSolidWorksMacroAsync(saveRequest, new PublicationSaveContext(SolidWorksSaveProbe(), destinationGuard))
                    : PublicationSave(saveRequest));
                callPending = false; result.Save = saved;
                var saveFields = json.Deserialize<Dictionary<string, object>>(json.Serialize(saved));
                if (!saveFields.ContainsKey("Verified") || !Convert.ToBoolean(saveFields["Verified"]) ||
                    saveFields.ContainsKey("Uncertain") && Convert.ToBoolean(saveFields["Uncertain"]))
                {
                    result.Terminal = true; result.Uncertain = true; result.Error = "Native publication save was not verified.";
                    return result;
                }
                verifyPublished(true);
                await generalGuard();
                result.OriginalPreserved = true; sourceChecked = true;
                result.Components = published; result.References = source.References.ToArray();
                result.NativeGeneratedDefaults = defaults.Select(c => new { c.Name, c.Code, c.Version,
                    OriginalExportSha256 = c.ExportSha256, HiddenAttributes = c.Attributes }).ToArray();
                result.ProjectVersion = (string)((dynamic)ProjectProperties(creation.Project)).Version;
                result.CollectionVersion = LifecycleVersion(ReadLifecycleCollection());
                result.Verified = true; result.Terminal = true;
                claim("Terminal");
                return result;
            }
            catch (Exception error)
            {
                result.Error = error.ToString(); result.Verified = false;
                result.Terminal = !callPending; result.Uncertain = result.MutationInvoked || callPending;
                result.OriginalPreserved = sourceChecked; // Unknown preservation must never be advertised as verified.
                return result;
            }
        }

        private static bool PublicationGeneralBoolean(Dictionary<string, object> fields, string key, bool expected)
        { object value; return fields.TryGetValue(key, out value) && value is bool && (bool)value == expected; }
        private static bool PublicationGeneralCount(Dictionary<string, object> fields, string key, int expected)
        { object value; return fields.TryGetValue(key, out value) && value is int && (int)value == expected; }

        private PublicationSource ReadPublicationSource(object canonical)
        {
            dynamic project = canonical;
            if ((int)project.Type != 101 || (int)project.Mode != 2 || (int)project.Protection != 0)
                throw new InvalidOperationException("Publication requires an unprotected standalone Type101 source in design mode.");
            string name = (string)project.Name;
            if (!string.IsNullOrEmpty((string)project.HelpFile) || Convert.ToInt32(project.HelpContextID) != 0)
                throw new InvalidOperationException("Nondefault help or conditional-compilation metadata cannot be published by this route.");
            var components = new List<PublicationComponent>();
            var names = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            foreach (dynamic component in project.VBComponents)
            {
                var item = ReadPublicationComponent(name, (object)component);
                if (!names.Add(item.Name) || item.Name.Equals("ThisLibrary", StringComparison.OrdinalIgnoreCase))
                    throw new InvalidOperationException("Duplicate or native-reserved component name.");
                components.Add(item);
            }
            var references = ReadPublicationReferences(canonical);
            string path = StandaloneAwareProjectPath(canonical);
            if (string.IsNullOrWhiteSpace(path)) path = null;
            string diskSha = path == null ? null : PublicationFileHash(path);
            string description = (string)project.Description;
            string version = (string)((dynamic)ProjectProperties(name)).Version;
            string fingerprint = Hash(json.Serialize(new { name, path, diskSha, version, Saved = (bool)project.Saved,
                description, Components = components.Select(c => new { c.Name, c.Type, c.Code, c.FormVersion, c.Version }),
                References = ReferenceSet(references) }));
            return new PublicationSource { Canonical = canonical, Name = name, Path = path, DiskSha256 = diskSha,
                Version = version, Saved = (bool)project.Saved, Description = description, Components = components,
                References = references, Fingerprint = fingerprint };
        }
        private PublicationComponent ReadPublicationComponent(string projectName, object canonical)
        {
            dynamic component = canonical;
            string name = (string)component.Name;
            int type = (int)component.Type;
            if (type != 1 && type != 2 && type != 3)
                throw new InvalidOperationException("Only standard modules, class modules and UserForms can be published.");
            if (!Regex.IsMatch(name ?? "", @"^[A-Za-z_][A-Za-z0-9_]*$") || name.Length > 31)
                throw new InvalidOperationException("An export-safe canonical component identifier is required.");
            dynamic code = component.CodeModule;
            int count = (int)code.CountOfLines;
            string formVersion = null, designerJson = null;
            if (type == 3)
            {
                object tree = PublicationFormTree == null ? forms.Tree(projectName, name) : PublicationFormTree(projectName, name);
                string serializedTree = json.Serialize(tree);
                var fields = json.Deserialize<Dictionary<string, object>>(serializedTree);
                RequirePublicationDesignerReadable(fields);
                if (!fields.TryGetValue("TreeVersion", out var treeVersion) || !(treeVersion is string) ||
                    string.IsNullOrWhiteSpace((string)treeVersion))
                    throw new InvalidOperationException("A complete designer tree version is required.");
                formVersion = (string)treeVersion;
                designerJson = serializedTree;
            }
            string capturedCode = count == 0 ? "" : (string)code.Lines(1, count);
            object snapshot;
            string version;
            if (type == 3)
            {
                snapshot = new { Name = name, Type = type, Code = capturedCode,
                    FormVersion = formVersion, Properties = ReadProperties(canonical) };
                version = Hash(json.Serialize(snapshot));
            }
            else
            {
                snapshot = ComponentSnapshot(projectName, component);
                version = (string)((dynamic)snapshot).Version;
            }
            return new PublicationComponent { Canonical = canonical, Name = name, Type = type,
                Code = capturedCode, FormVersion = formVersion, DesignerJson = designerJson,
                Version = version, ComponentSnapshotJson = json.Serialize(snapshot) };
        }
        internal static void RequirePublicationDesignerReadable(object value)
        {
            var dictionary = value as IDictionary<string, object>;
            if (dictionary != null)
            {
                foreach (var pair in dictionary)
                {
                    if ((pair.Key == "Error" || pair.Key == "ReadError") && pair.Value != null)
                        throw new InvalidOperationException("An unreadable designer/resource prevents complete publication.");
                    RequirePublicationDesignerReadable(pair.Value);
                }
            }
            else if (value is IEnumerable sequence && !(value is string))
                foreach (object item in sequence) RequirePublicationDesignerReadable(item);
        }
        private static List<PublicationReference> ReadPublicationReferences(object canonical)
        {
            var result = new List<PublicationReference>(); var identities = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            foreach (dynamic reference in ((dynamic)canonical).References)
            {
                string guid = (string)reference.GUID;
                if ((int)reference.Type != 0 || (bool)reference.IsBroken || !System.Guid.TryParse(guid, out var parsed))
                    throw new InvalidOperationException("Only identified, unbroken type-library references can be published; project references are unsupported.");
                var item = new PublicationReference { Guid = parsed.ToString("B"), Major = (int)reference.Major,
                    Minor = (int)reference.Minor, BuiltIn = (bool)reference.BuiltIn };
                if (item.Major < 0 || item.Minor < 0 || !identities.Add(item.Guid))
                    throw new InvalidOperationException("Ambiguous reference identity.");
                result.Add(item);
            }
            return result;
        }
        private static string[] ReferenceSet(IEnumerable<PublicationReference> references) =>
            references.Select(r => r.Key).OrderBy(s => s, StringComparer.Ordinal).ToArray();
        private static List<PublicationComponent> ReadPublicationDestinationComponents(object canonical)
        {
            var result = new List<PublicationComponent>();
            foreach (dynamic component in ((dynamic)canonical).VBComponents)
                result.Add(new PublicationComponent { Name = (string)component.Name, Type = (int)component.Type });
            return result;
        }
        private static string PublicationExtension(int type) => type == 1 ? ".bas" : type == 2 ? ".cls" : ".frm";
        internal static string PublicationAttributes(string text) =>
            string.Join("\n", Regex.Matches(text ?? "", @"(?m)^Attribute\s+[^\r\n]+").Cast<Match>()
                .Select(m => m.Value.Trim()).OrderBy(s => s, StringComparer.Ordinal));
        private static string PublicationFileHash(string path)
        {
            if (!File.Exists(path)) throw new IOException("Frozen publication file is absent.");
            using (var sha = System.Security.Cryptography.SHA256.Create())
            using (var stream = File.Open(path, FileMode.Open, FileAccess.Read, FileShare.Read))
                return BitConverter.ToString(sha.ComputeHash(stream)).Replace("-", "");
        }
        private static void VerifyPublicationExport(PublicationComponent component)
        {
            if (PublicationFileHash(component.ExportPath) != component.ExportSha256 ||
                component.Type == 3 && PublicationFileHash(Path.ChangeExtension(component.ExportPath, ".frx")) != component.FrxSha256)
                throw new IOException("Frozen publication exports/resources changed before import.");
        }

        private sealed class PublicationSaveContext : ISolidWorksSaveProbe
        {
            private readonly ISolidWorksSaveProbe inner;
            private readonly Action guard;
            internal PublicationSaveContext(ISolidWorksSaveProbe inner, Action guard) { this.inner = inner; this.guard = guard; }
            public bool IsSolidWorks { get { guard(); return inner.IsSolidWorks; } }
            public int ProcessId { get { guard(); return inner.ProcessId; } }
            public SolidWorksSaveSelection Selection { get { guard(); return inner.Selection; } }
            public void RequireOwner(object editor) { guard(); inner.RequireOwner(editor); }
            public bool SameProject(object first, object second) { guard(); return inner.SameProject(first, second); }
            public object SelectComponent(object editor, object project) { guard(); return inner.SelectComponent(editor, project); }
            public void RestoreSelection(object editor, object project, object component) { guard(); inner.RestoreSelection(editor, project, component); }
            public bool SelectionMatches(object editor, object project, object component) { guard(); return inner.SelectionMatches(editor, project, component); }
            public object SaveControl(object editor) { guard(); return inner.SaveControl(editor); }
            public void Save(object control) { guard(); inner.Save(control); }
            public bool FileExists(string path) { guard(); return inner.FileExists(path); }
            public bool FileReadOnly(string path) { guard(); return inner.FileReadOnly(path); }
            public long FileLength(string path) { guard(); return inner.FileLength(path); }
        }
    }
}

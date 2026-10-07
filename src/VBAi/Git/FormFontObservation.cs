using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Runtime.InteropServices.ComTypes;
using System.Security.Cryptography;
using System.Text;
using System.Text.RegularExpressions;
using System.Web.Script.Serialization;

namespace VBAi
{

    /// <summary>One explicitly armed, disposable root-font import observation on the native VBE owner.</summary>
    internal sealed partial class FormFontObservation
    {

        /// <summary>Environment-variable name that opts one disposable import into root-font observation.</summary>
        internal const string ManifestVariable = "VBAi_TEST_ROOT_FONT_OBSERVATION_MANIFEST";

        /// <summary>Mode that records font writes during the selected form import.</summary>
        internal const string ObserveWrites = "ObserveWrites";

        /// <summary>Mode that observes an explicitly distinct child-control font name.</summary>
        internal const string DistinctChildName = "DistinctChildName";

        /// <summary>Mode that permits one deferred root-font transfer after an initial capture gate.</summary>
        internal const string AfterInitialCapture = "AfterInitialCapture";

        /// <summary>Exact opt-in identity, target hashes, nonce-bound output location, and observation mode.</summary>
        internal sealed class Manifest
        {

            /// <summary>Gets or sets the disposable host document path for this observation.</summary>
            /// <value>Absolute project path validated against the active operation.</value>
            public string ProjectPath { get; set; }

            /// <summary>Gets or sets the single changed form component observed by this run.</summary>
            /// <value>Validated VBIDE component name.</value>
            public string FormName { get; set; }

            /// <summary>Gets or sets the SHA-256 of the expected exported form source.</summary>
            /// <value>64-character hexadecimal digest.</value>
            public string TargetFormSha256 { get; set; }

            /// <summary>Gets or sets the expected root StdFont descriptor bytes as hexadecimal.</summary>
            /// <value>Hex digest of the exact descriptor, compared before transfer.</value>
            public string TargetDescriptorHex { get; set; }

            /// <summary>Gets or sets the assembly MVID that must equal the currently loaded candidate.</summary>
            /// <value>GUID in parseable form.</value>
            public string CandidateMvid { get; set; }

            /// <summary>Gets or sets the fresh output directory whose final path component is the nonce.</summary>
            /// <value>Absolute, not-yet-existing output path.</value>
            public string OutputRoot { get; set; }

            /// <summary>Gets or sets the one-use random nonce binding the receipt directory to this trial.</summary>
            /// <value>32 hexadecimal characters representing a GUID without separators.</value>
            public string Nonce { get; set; }

            /// <summary>Gets or sets one of the supported observation modes.</summary>
            /// <value>Mode constant validated before the import starts.</value>
            public string Mode { get; set; }

            /// <summary>Gets or sets the required temporary face name for the distinct-child-name experiment.</summary>
            /// <value>Must be the declared test value only in that mode; otherwise empty.</value>
            public string TemporaryName { get; set; }
        }

        /// <summary>Validated manifest governing this one observation instance.</summary>
        private readonly Manifest manifest;

        /// <summary>Monotonic receipt sequence used to create unique evidence files.</summary>
        private int receipt;

        /// <summary>Prevents a second deferred native font transfer in the same observation.</summary>
        private bool transferStarted;

        /// <summary>Creates the observation from a manifest already validated against the active target and candidate.</summary>
        /// <param name="value">Manifest already validated against the active target and candidate.</param>
        private FormFontObservation(Manifest value) { manifest = value; }

        /// <summary>Gets the single form component named by the validated manifest.</summary>
        /// <value>Validated component name.</value>
        internal string FormName { get { return manifest.FormName; } }

        /// <summary>Gets whether the observation uses the distinct-child-name experiment.</summary>
        /// <value>True only for <see cref="DistinctChildName"/> mode.</value>
        internal bool HasDistinctName { get { return manifest.Mode == DistinctChildName; } }

        /// <summary>Gets whether the operation is in the deferred-after-capture mode.</summary>
        /// <value>True only for <see cref="AfterInitialCapture"/> mode.</value>
        internal bool IsAfterInitialCapture { get { return manifest.Mode == AfterInitialCapture; } }

        /// <summary>Gets the temporary child font name only for the distinct-name experiment.</summary>
        /// <value>Validated temporary name, or null for other modes.</value>
        internal string TemporaryName { get { return HasDistinctName ? manifest.TemporaryName : null; } }

        /// <summary>Loads and validates the environment-selected manifest before arming one disposable import observation.</summary>
        /// <param name="projectPath">Active disposable host document path.</param>
        /// <param name="target">Expected Git snapshot containing the target form and descriptor.</param>
        /// <param name="changed">Changed component names; must identify only the manifest's form.</param>
        /// <param name="project">Live project object whose owning thread is verified.</param>
        /// <returns>Armed observation, or null when the opt-in environment variable is absent.</returns>
        internal static FormFontObservation TryBegin(string projectPath, VbaGitSnapshot target,
            ISet<string> changed, object project)
        {
            return TryBeginAtPath(Environment.GetEnvironmentVariable(ManifestVariable), projectPath, target, changed, project);
        }

        /// <summary>Validates and arms one manifest read from the specified path; stale or unsafe trials are refused.</summary>
        /// <param name="path">Absolute manifest path, bounded to 16 KiB and rejected if it traverses a reparse point.</param>
        /// <param name="projectPath">Expected active disposable project path.</param>
        /// <param name="target">Expected source snapshot used to verify form bytes and descriptor.</param>
        /// <param name="changed">Changed component set, required to contain exactly the selected form.</param>
        /// <param name="project">Live project used for owner-thread validation.</param>
        /// <returns>Armed observation after creating its fresh nonce-bound output directory.</returns>
        internal static FormFontObservation TryBeginAtPath(string path, string projectPath, VbaGitSnapshot target,
            ISet<string> changed, object project)
        {
            if (string.IsNullOrEmpty(path)) return null;
            if (!Path.IsPathRooted(path) || !File.Exists(path) || new FileInfo(path).Length > 16384)
                throw new InvalidOperationException("Root font observation manifest is missing, relative or oversized.");
            RequireSafeLocalPath(path, true);
            var value = new JavaScriptSerializer { MaxJsonLength = 16384 }.Deserialize<Manifest>(
                File.ReadAllText(path, new UTF8Encoding(false, true)));
            ValidateManifest(value, projectPath, target, changed);
            RequireSafeLocalPath(value.ProjectPath, true);
            RequireSafeLocalPath(value.OutputRoot, false);
            FormFontRestoration.RequireOwner(project);
            // A trial owns one fresh output root. Neither a stale claim nor a completed
            // run is ever resumed, even if a previous native call returned an error.
            Directory.CreateDirectory(value.OutputRoot);
            if ((File.GetAttributes(value.OutputRoot) & FileAttributes.ReparsePoint) != 0)
                throw new InvalidOperationException("Root font observation output root became a reparse point.");
            var observation = new FormFontObservation(value);
            observation.Write("armed", new
            {
                value.Mode,
                value.FormName,
                value.Nonce,
                value.CandidateMvid,
                value.TargetFormSha256,
                value.TargetDescriptorHex
            });
            return observation;
        }

        /// <summary>Checks candidate MVID, project, changed form, fresh output root, and exact target resource identity.</summary>
        /// <param name="value">Manifest to validate.</param>
        /// <param name="projectPath">Currently selected disposable project path.</param>
        /// <param name="target">Expected Git snapshot and resource bytes.</param>
        /// <param name="changed">Changed component set from the operation plan.</param>
        /// <exception cref="InvalidOperationException">Any identity, mode, path, resource hash, descriptor, or freshness condition fails.</exception>
        internal static void ValidateManifest(Manifest value, string projectPath, VbaGitSnapshot target, ISet<string> changed)
        {
            if (value == null || value.Mode != ObserveWrites && value.Mode != DistinctChildName &&
                value.Mode != AfterInitialCapture ||
                !Guid.TryParseExact(value.Nonce, "N", out _) || !Guid.TryParse(value.CandidateMvid, out Guid mvid) ||
                mvid != typeof(FormFontObservation).Module.ModuleVersionId)
                throw new InvalidOperationException("Root font observation identity or mode does not match this assembly.");
            VbaGitSnapshot.ValidateName(value.FormName);
            if (!Path.IsPathRooted(value.ProjectPath) ||
                !string.Equals(Path.GetFullPath(value.ProjectPath), Path.GetFullPath(projectPath), StringComparison.OrdinalIgnoreCase) ||
                changed == null || changed.Count != 1 || !changed.Contains(value.FormName))
                throw new InvalidOperationException("Root font observation does not match the one changed disposable project form.");
            if (!Path.IsPathRooted(value.OutputRoot) || Directory.Exists(value.OutputRoot) ||
                File.Exists(value.OutputRoot) || !string.Equals(Path.GetFileName(
                    Path.GetFullPath(value.OutputRoot).TrimEnd(Path.DirectorySeparatorChar)), value.Nonce, StringComparison.OrdinalIgnoreCase))
                throw new InvalidOperationException("Root font observation output root is not fresh or nonce-bound.");
            var form = target.Manifest.Components.SingleOrDefault(item => item.Name == value.FormName && item.Type == 3);
            if (form == null || !form.HasResources || !IsSha256(value.TargetFormSha256) ||
                !string.Equals(Sha256(target.Files[form.FileName]), value.TargetFormSha256, StringComparison.OrdinalIgnoreCase))
                throw new InvalidOperationException("Root font observation form source hash differs from the target.");
            var bindings = target.FormFonts(form);
            var roots = bindings == null ? new FormStreamPadding.FormFontBinding[0] :
                bindings.Where(item => item.OwnerPath == "" && item.Type == 7).ToArray();
            if (roots.Length != 1 || !IsHex(value.TargetDescriptorHex) ||
                !string.Equals(Hex(roots[0].Descriptor), value.TargetDescriptorHex, StringComparison.OrdinalIgnoreCase))
                throw new InvalidOperationException("Root font observation descriptor differs from the target.");
            string targetName = Encoding.ASCII.GetString(roots[0].Descriptor, 11, roots[0].Descriptor[10]);
            if (value.Mode == DistinctChildName &&
                (!string.Equals(value.TemporaryName, "Arial", StringComparison.Ordinal) ||
                 string.Equals(value.TemporaryName, targetName, StringComparison.OrdinalIgnoreCase)))
                throw new InvalidOperationException("Root font observation temporary face is invalid or not distinct.");
            if (value.Mode != DistinctChildName && !string.IsNullOrEmpty(value.TemporaryName))
                throw new InvalidOperationException("This root font observation mode cannot request a temporary face.");
        }

        /// <summary>Requires every initial source, identity and reference byte except the selected FRX to match.</summary>
        /// <param name="observed">Post-import project snapshot.</param>
        /// <param name="target">Expected source snapshot.</param>
        /// <param name="formName">Only form whose FRX may differ.</param>
        /// <returns>True when package identities and every other comparison file match exactly.</returns>
        internal static bool SameExceptSelectedFrx(VbaGitSnapshot observed, VbaGitSnapshot target, string formName)
        {
            if (observed == null || target == null || string.IsNullOrEmpty(formName)) return false;
            var actual = observed.ComparisonFiles(); var desired = target.ComparisonFiles();
            string resource = formName + ".frx";
            return actual.Count == desired.Count && actual.ContainsKey(resource) && desired.ContainsKey(resource) &&
                actual.Where(item => item.Key != resource).All(item => desired.TryGetValue(item.Key, out byte[] bytes) &&
                    item.Value.SequenceEqual(bytes));
        }

        /// <summary>Captures and records the first post-import state before allowing the one deferred transfer.</summary>
        /// <param name="capture">Fresh live-project snapshot operation.</param>
        /// <param name="target">Expected imported source snapshot.</param>
        /// <param name="transfer">Single deferred font restoration action authorized by the gate.</param>
        internal void RunAfterInitialCapture(Func<VbaGitSnapshot> capture, VbaGitSnapshot target, Action transfer)
        {
            if (!IsAfterInitialCapture) throw new InvalidOperationException("Initial capture belongs only to the deferred mode.");
            GateAfterInitialCapture(capture, target, FormName,
                (exact, onlySelectedFrx) => Write("initial-post-import-capture",
                    new { Exact = exact, OnlySelectedFrxDiffers = onlySelectedFrx }), transfer);
        }

        /// <summary>Records the first post-import capture before any deferred native transfer.</summary>
        /// <param name="capture">Operation that observes the live project immediately before transfer.</param>
        /// <param name="target">Expected snapshot for the import.</param>
        /// <param name="formName">Form whose FRX is the only permitted difference at this stage.</param>
        /// <param name="record">Receipt callback receiving exact-match and only-selected-FRX-differs flags.</param>
        /// <param name="transfer">Native font delivery invoked only after the capture gate passes.</param>
        internal static void GateAfterInitialCapture(Func<VbaGitSnapshot> capture, VbaGitSnapshot target, string formName,
            Action<bool, bool> record, Action transfer)
        {
            VbaGitSnapshot observed = capture();
            bool exact = observed.SameAs(target);
            bool onlySelectedFrx = SameExceptSelectedFrx(observed, target, formName);
            record(exact, onlySelectedFrx);
            if (exact || !onlySelectedFrx)
                throw new InvalidOperationException("Deferred font transfer requires one nonexact post-import capture with only the selected FRX differing.");
            transfer();
        }

        /// <summary>Records the single permitted deferred delivery after matching the planned root-font descriptor.</summary>
        /// <param name="descriptor">Exact serialized descriptor about to be transferred.</param>
        internal void BeforeDeferredTransfer(byte[] descriptor)
        {
            if (!IsAfterInitialCapture || transferStarted ||
                !string.Equals(Hex(descriptor), manifest.TargetDescriptorHex, StringComparison.OrdinalIgnoreCase))
                throw new InvalidOperationException("Deferred root font transfer is not the single declared target delivery.");
            transferStarted = true;
            Write("before-deferred-transfer", new
            {
                Owner = FormName,
                TargetDescriptorHex = Hex(descriptor),
                Delivery = "one fresh StdFont IPersistStream.Load then Designer.Font.put"
            });
        }

        /// <summary>Records that the native deferred setter returned after a transfer was armed.</summary>
        internal void DeferredTransferReturned()
        {
            if (!transferStarted) throw new InvalidOperationException("No deferred font transfer was started.");
            Write("deferred-transfer-returned", new { Owner = FormName, NativeSetterReturned = true });
        }

        /// <summary>Persists the requested child-font property immediately before its native setter.</summary>
        /// <param name="field">Property name being delivered.</param>
        /// <param name="requested">Value passed to the native property setter.</param>
        internal void BeforeDelivery(string field, object requested)
        {
            Write("before-put", new { Field = field, Requested = requested });
        }

        /// <summary>Reads back a delivered font property and fails the observation when it differs from the requested value.</summary>
        /// <param name="field">Property name to read from the child object.</param>
        /// <param name="requested">Value expected after the setter returns.</param>
        /// <param name="child">Native property wrapper whose Value is read back.</param>
        internal void AfterDelivery(string field, object requested, object child)
        {
            object actual = NativeRead("VBIDE.Property.Value.get(Font." + field + ")", () => ((dynamic)child).Value);
            Write("after-put", new
            {
                Field = field,
                Requested = requested,
                Actual = actual,
                Equal = object.Equals(requested, actual)
            });
            if (!object.Equals(requested, actual))
                throw new InvalidOperationException("Root font child " + field + " did not read back the planned value.");
        }

        /// <summary>Captures exported resource bytes, child font values, root font descriptors, and COM identity for one observation phase.</summary>
        /// <param name="phase">Receipt phase name.</param>
        /// <param name="component">Form component exported for evidence.</param>
        /// <param name="designer">Designer whose root Font object is compared.</param>
        /// <param name="rootProperty">VBIDE root font property wrapper.</param>
        /// <param name="children">Ordered child-property wrappers for the observed font attributes.</param>
        /// <param name="revalidate">Authorization/source check repeated immediately before native reads.</param>
        internal void Observe(string phase, object component, object designer, object rootProperty, object[] children,
            Action revalidate)
        {
            revalidate();
            Write("before-observe", new { TargetPhase = phase });
            // Child preflight getters have already run. This export precedes only
            // the attached-font getters in this observation phase.
            var (FrxBytes, FrxSha256, RootDescriptors) = Export(component, phase);
            var values = new Dictionary<string, object>(StringComparer.Ordinal);
            string[] names = { "Name", "Size", "Bold", "Italic", "Underline", "Strikethrough", "Weight", "Charset" };
            for (int i = 0; i < names.Length; i++)
            {
                int index = i;
                values.Add(names[i], NativeRead("VBIDE.Property.Value.get(Font." + names[i] + ")",
                    () => ((dynamic)children[index]).Value));
            }
            object propertyFont = null, designerFont = null;
            Exception primary = null;
            try
            {
                propertyFont = NativeRead("VBIDE.Property.Object.get(Font)", () => ((dynamic)rootProperty).Object);
                designerFont = NativeRead("MSForms.Font.get", () => ((dynamic)designer).Font);
                bool same = NativeRead("Font.IUnknown.compare", () => SameComIdentity(propertyFont, designerFont));
                var propertyDescriptor = SaveDescriptor(propertyFont, "VBIDE.Font.Object");
                var designerDescriptor = SaveDescriptor(designerFont, "Designer.Font");
                Write(phase, new
                {
                    Children = values,
                    PropertyFontDescriptor = Hex(propertyDescriptor),
                    DesignerFontDescriptor = Hex(designerDescriptor),
                    SameFontIUnknown = same,
                    FrxBytes,
                    FrxSha256,
                    RootDescriptors
                });
            }
            catch (Exception error) { primary = error; throw; }
            finally { ReleaseBoth(designerFont, propertyFont, primary); }
        }

        /// <summary>Writes a bounded failure receipt containing exception type, HRESULT, message, and inner stage.</summary>
        /// <param name="error">Failure raised by the native observation operation.</param>
        internal void Failure(Exception error)
        {
            Write("throw", new
            {
                ExceptionType = error.GetType().FullName,
                error.HResult,
                error.Message,
                Stage = error.InnerException?.GetType().FullName
            });
        }

        /// <summary>Writes the final exact-snapshot comparison result.</summary>
        /// <param name="exact">Whether the observed project exactly matches the target snapshot.</param>
        internal void Complete(bool exact)
        {
            Write("final-comparison", new { Exact = exact });
        }

        /// <summary>Revalidates authority and exports the form before any post-import root-font getters run.</summary>
        /// <param name="component">Imported form component.</param>
        /// <param name="revalidate">Fresh project/source validation callback.</param>
        internal void BeforeFontGetters(object component, Action revalidate)
        {
            revalidate();
            Write("before-export", new { TargetPhase = "post-import-before-font-getters" });
            var (FrxBytes, FrxSha256, RootDescriptors) = Export(component, "post-import-before-font-getters");
            Write("post-import-before-font-getters", new { FrxBytes, FrxSha256, RootDescriptors });
        }

        /// <summary>Exports the form into a unique receipt directory and extracts resource length, hash, and root-font descriptors.</summary>
        /// <param name="component">Native VBComponent being observed.</param>
        /// <param name="phase">Receipt phase used in the evidence path and native-read diagnostics.</param>
        /// <returns>FRX byte length, SHA-256 digest, and root descriptors; empty values when no FRX file was exported.</returns>
        private (int FrxBytes, string FrxSha256, string[] RootDescriptors) Export(object component, string phase)
        {
            string folder = Path.Combine(manifest.OutputRoot, receipt.ToString("D3", CultureInfo.InvariantCulture) + "-" + phase);
            Directory.CreateDirectory(folder);
            string baseName = Path.Combine(folder, manifest.FormName);
            NativeRead("VBComponent.Export(" + phase + ")", () => { ((dynamic)component).Export(baseName + ".frm"); return true; });
            string frx = baseName + ".frx";
            if (!File.Exists(frx)) return (0, null, new string[0]);
            byte[] bytes = File.ReadAllBytes(frx);
            string source = File.ReadAllText(baseName + ".frm", NativeEncoding).Replace("\r\n", "\n").Replace("\r", "\n");
            var form = new VbaGitComponent { Name = manifest.FormName, Type = 3, HasResources = true };
            // Exported files use the actual form name only inside the validated snapshot,
            // while the phase-prefixed files remain durable trial evidence.
            var files = new Dictionary<string, byte[]>
            {
                [form.FileName] = VbaGitSnapshot.Utf8.GetBytes(source),
                [form.Name + ".frx"] = bytes
            };
            var snapshot = new VbaGitSnapshot(new VbaGitManifest
            {
                References = "",
                Components = new[] { form }
            }, files);
            string[] roots = snapshot.FormFonts(form)?.Where(item => item.OwnerPath == "")
                .Select(item => Hex(item.Descriptor)).ToArray() ?? new string[0];
            return (bytes.Length, Sha256(bytes), roots);
        }

        /// <summary>Creates a unique durable JSON receipt using UTF-8 without BOM and flushes it to disk.</summary>
        /// <param name="phase">Receipt phase included in its filename and payload.</param>
        /// <param name="data">Bounded evidence payload for this phase.</param>
        private void Write(string phase, object data)
        {
            string file = Path.Combine(manifest.OutputRoot,
                (++receipt).ToString("D3", CultureInfo.InvariantCulture) + "-" + phase + ".json");
            byte[] bytes = new UTF8Encoding(false).GetBytes(
                new JavaScriptSerializer().Serialize(new { Phase = phase, manifest.Nonce, Data = data }) + "\n");
            using (var stream = new FileStream(file, FileMode.CreateNew, FileAccess.Write, FileShare.None))
            { stream.Write(bytes, 0, bytes.Length); stream.Flush(true); }
        }

        /// <summary>Serializes one font through IPersistStream and requires a complete bounded descriptor read.</summary>
        /// <param name="font">Native font COM object implementing IPersistStream.</param>
        /// <param name="owner">Diagnostic label identifying the property or designer source.</param>
        /// <returns>Serialized descriptor between 11 and 4096 bytes.</returns>
        private static byte[] SaveDescriptor(object font, string owner)
        {
            IStream stream = null;
            IntPtr count = IntPtr.Zero;
            Exception primary = null;
            try
            {
                NativeRead("CreateStreamOnHGlobal", () =>
                {
                    Marshal.ThrowExceptionForHR(CreateStreamOnHGlobal(IntPtr.Zero, true, out stream)); return true;
                });
                NativeRead("IPersistStream.Save(" + owner + ")", () => { ((PersistStream)font).Save(stream, false); return true; });
                var stat = NativeRead("IStream.Stat(" + owner + ")", () =>
                {
                    stream.Stat(out System.Runtime.InteropServices.ComTypes.STATSTG result, 1); return result;
                });
                if (stat.cbSize < 11 || stat.cbSize > 4096)
                    throw new InvalidOperationException("Observed font descriptor size is outside bounded profile.");
                byte[] data = new byte[checked((int)stat.cbSize)];
                NativeRead("IStream.Seek(" + owner + ")", () => { stream.Seek(0, 0, IntPtr.Zero); return true; });
                count = Marshal.AllocHGlobal(sizeof(int));
                NativeRead("IStream.Read(" + owner + ")", () => { stream.Read(data, data.Length, count); return true; });
                if (Marshal.ReadInt32(count) != data.Length)
                    throw new InvalidOperationException("Observed font descriptor was not read completely.");
                return data;
            }
            catch (Exception error) { primary = error; throw; }
            finally
            {
                if (count != IntPtr.Zero) Marshal.FreeHGlobal(count);
                try { Release(stream); }
                catch (Exception cleanup)
                {
                    if (primary != null) throw new AggregateException("Font observation and stream release both failed.", primary, cleanup);
                    throw;
                }
            }
        }

        /// <summary>Labels selected COM failures with the native operation and HRESULT while preserving the inner exception.</summary>
        /// <typeparam name="T">Return type of the native read.</typeparam>
        /// <param name="operation">Operation label written into the diagnostic.</param>
        /// <param name="action">Single native read to execute.</param>
        /// <returns>Value returned by the read.</returns>
        private static T NativeRead<T>(string operation, Func<T> action)
        {
            try { return action(); }
            catch (Exception error) when (error is COMException || error is NotSupportedException || error is InvalidCastException)
            {
                throw new InvalidOperationException("Root font observation failed at " + operation +
                    " (HRESULT 0x" + error.HResult.ToString("X8", CultureInfo.InvariantCulture) + ").", error);
            }
        }

        /// <summary>Releases two owned COM references and aggregates cleanup failures with the primary observation error.</summary>
        /// <param name="first">First owned interface reference.</param>
        /// <param name="second">Second owned interface reference.</param>
        /// <param name="primary">Earlier failure to preserve, if any.</param>
        private static void ReleaseBoth(object first, object second, Exception primary)
        {
            var failures = new List<Exception>();
            if (primary != null) failures.Add(primary);
            try { Release(first); } catch (Exception error) { failures.Add(error); }
            try { Release(second); } catch (Exception error) { failures.Add(error); }
            if (failures.Count > (primary == null ? 0 : 1))
                throw new AggregateException("Font observation and COM release provenance.", failures);
        }

        /// <summary>Compares IUnknown identity and releases both temporary interface pointers in all cases.</summary>
        /// <param name="first">First native COM object.</param>
        /// <param name="second">Second native COM object.</param>
        /// <returns>True when both objects expose the same controlling IUnknown pointer.</returns>
        private static bool SameComIdentity(object first, object second)
        {
            IntPtr a = IntPtr.Zero, b = IntPtr.Zero;
            try
            {
                a = Marshal.GetIUnknownForObject(first);
                b = Marshal.GetIUnknownForObject(second);
                return a == b;
            }
            finally
            {
                if (b != IntPtr.Zero) Marshal.Release(b);
                if (a != IntPtr.Zero) Marshal.Release(a);
            }
        }

        /// <summary>Checks for exactly 64 hexadecimal characters.</summary>
        /// <param name="text">Candidate SHA-256 string.</param>
        /// <returns>True when the candidate has the expected digest syntax.</returns>
        private static bool IsSha256(string text) { return Regex.IsMatch(text ?? "", "^[0-9A-Fa-f]{64}$"); }

        /// <summary>Checks for an even-length hexadecimal descriptor between 11 and 64 bytes.</summary>
        /// <param name="text">Candidate descriptor string.</param>
        /// <returns>True when the value is syntactically valid hexadecimal within the supported size bound.</returns>
        private static bool IsHex(string text) { return Regex.IsMatch(text ?? "", "^(?:[0-9A-Fa-f]{2}){11,64}$"); }

        /// <summary>Formats bytes as uppercase hexadecimal without separators.</summary>
        /// <param name="bytes">Bytes to encode.</param>
        /// <returns>Two hexadecimal characters per byte.</returns>
        private static string Hex(byte[] bytes) { return BitConverter.ToString(bytes).Replace("-", ""); }

        /// <summary>Computes SHA-256 and returns the digest as uppercase hexadecimal.</summary>
        /// <param name="bytes">Content to hash.</param>
        /// <returns>64 hexadecimal characters.</returns>
        private static string Sha256(byte[] bytes) { using (var hash = SHA256.Create()) return Hex(hash.ComputeHash(bytes)); }

        /// <summary>Requires an exact local drive path without alternate streams or reparse points in any existing ancestor.</summary>
        /// <param name="path">Manifest, project, or output path to validate.</param>
        /// <param name="mustExist">When true, the path itself must already exist as a regular file.</param>
        /// <exception cref="InvalidOperationException">The path is nonlocal, noncanonical, missing when required, or crosses a reparse point.</exception>
        internal static void RequireSafeLocalPath(string path, bool mustExist)
        {
            string full;
            try { full = Path.GetFullPath(path); }
            catch (Exception error) when (error is ArgumentException || error is NotSupportedException || error is PathTooLongException)
            { throw new InvalidOperationException("Root font observation path is invalid.", error); }
            string root = Path.GetPathRoot(full);
            if (!Regex.IsMatch(root ?? "", "^[A-Za-z]:\\\\$") ||
                !string.Equals(full, path, StringComparison.OrdinalIgnoreCase) ||
                full.Substring(root.Length).Contains(":"))
                throw new InvalidOperationException("Root font observation requires an exact local path without alternate streams.");
            if (mustExist && !File.Exists(full))
                throw new InvalidOperationException("Root font observation required file is missing.");
            for (string part = mustExist ? full : Path.GetDirectoryName(full); part != null; part = Path.GetDirectoryName(part))
            {
                if (File.Exists(part) || Directory.Exists(part))
                {
                    FileAttributes attributes = File.GetAttributes(part);
                    if ((attributes & FileAttributes.ReparsePoint) != 0 ||
                        (part == full && mustExist && (attributes & FileAttributes.Directory) != 0))
                        throw new InvalidOperationException("Root font observation path crosses a reparse point or non-file.");
                }
            }
        }

        /// <summary>Releases one owned runtime-callable COM wrapper exactly once.</summary>
        /// <param name="value">Object whose COM reference was acquired by this observation.</param>
        private static void Release(object value) { if (value != null && Marshal.IsComObject(value)) Marshal.ReleaseComObject(value); }

        /// <summary>Gets strict text encoding for VBIDE exports using the active Windows ANSI code page.</summary>
        /// <value>Encoding configured to throw on invalid input or output.</value>
        private static Encoding NativeEncoding
        {
            get
            {
                return Encoding.GetEncoding((int)GetACP(),
            EncoderFallback.ExceptionFallback, DecoderFallback.ExceptionFallback);
            }
        }

        /// <summary>Reads the active Windows ANSI code page identifier.</summary>
        /// <returns>Windows code-page number.</returns>
        [DllImport("kernel32.dll")] private static extern uint GetACP();

        /// <summary>Creates a COM IStream backed by a movable global-memory block.</summary>
        /// <param name="handle">Existing HGLOBAL, or zero to allocate a new one.</param>
        /// <param name="free">Whether stream release frees the HGLOBAL.</param>
        /// <param name="stream">Receives the created stream interface.</param>
        /// <returns>HRESULT from the OLE API.</returns>
        [DllImport("ole32.dll", ExactSpelling = true)]
        private static extern int CreateStreamOnHGlobal(IntPtr handle, [MarshalAs(UnmanagedType.Bool)] bool free, out IStream stream);

        /// <summary>COM vtable contract for serializing a font object's persisted descriptor without altering it.</summary>
        [ComImport, Guid("00000109-0000-0000-C000-000000000046"), InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
        private interface PersistStream
        {

            /// <summary>Returns the class identifier of the persisted object.</summary>
            /// <param name="clsid">Receives the object's CLSID.</param>
            void GetClassID(out Guid clsid);

            /// <summary>Reports whether the object has changes not represented by its persisted stream.</summary>
            /// <returns>HRESULT indicating dirty state.</returns>
            [PreserveSig] int IsDirty();

            /// <summary>Loads object state from a stream.</summary>
            /// <param name="stream">Serialized object state.</param>
            void Load(IStream stream);

            /// <summary>Saves object state to a stream.</summary>
            /// <param name="stream">Destination for serialized state.</param>
            /// <param name="clearDirty">Whether saving clears the object's dirty flag.</param>
            void Save(IStream stream, [MarshalAs(UnmanagedType.Bool)] bool clearDirty);

            /// <summary>Returns the maximum stream size required to save the object.</summary>
            /// <param name="size">Receives the maximum byte count.</param>
            void GetSizeMax(out long size);
        }
    }
}

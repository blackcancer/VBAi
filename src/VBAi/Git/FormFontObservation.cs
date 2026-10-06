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
    internal sealed class FormFontObservation
    {

        /// <summary>Maintains the manifest variable state for form font observation.</summary>
        internal const string ManifestVariable = "VBAi_TEST_ROOT_FONT_OBSERVATION_MANIFEST";

        /// <summary>Maintains the observe writes state for form font observation.</summary>
        internal const string ObserveWrites = "ObserveWrites";

        /// <summary>Maintains the distinct child name state for form font observation.</summary>
        internal const string DistinctChildName = "DistinctChildName";

        /// <summary>Maintains the after initial capture state for form font observation.</summary>
        internal const string AfterInitialCapture = "AfterInitialCapture";

        /// <summary>Owns the manifest state and operations.</summary>
        internal sealed class Manifest
        {

            /// <summary>Gets or sets the project path.</summary>
            /// <value>Current project path exposed by manifest.</value>
            public string ProjectPath { get; set; }

            /// <summary>Gets or sets the form name.</summary>
            /// <value>Current form name exposed by manifest.</value>
            public string FormName { get; set; }

            /// <summary>Gets or sets the target form sha256.</summary>
            /// <value>Current target form sha256 exposed by manifest.</value>
            public string TargetFormSha256 { get; set; }

            /// <summary>Gets or sets the target descriptor hex.</summary>
            /// <value>Current target descriptor hex exposed by manifest.</value>
            public string TargetDescriptorHex { get; set; }

            /// <summary>Gets or sets the candidate mvid.</summary>
            /// <value>Current candidate mvid exposed by manifest.</value>
            public string CandidateMvid { get; set; }

            /// <summary>Gets or sets the output root.</summary>
            /// <value>Current output root exposed by manifest.</value>
            public string OutputRoot { get; set; }

            /// <summary>Gets or sets the nonce.</summary>
            /// <value>Current nonce exposed by manifest.</value>
            public string Nonce { get; set; }

            /// <summary>Gets or sets the mode.</summary>
            /// <value>Current mode exposed by manifest.</value>
            public string Mode { get; set; }

            /// <summary>Gets or sets the temporary name.</summary>
            /// <value>Current temporary name exposed by manifest.</value>
            public string TemporaryName { get; set; }
        }

        /// <summary>Maintains the manifest state for form font observation.</summary>
        private readonly Manifest manifest;

        /// <summary>Maintains the receipt state for form font observation.</summary>
        private int receipt;

        /// <summary>Maintains the transfer started state for form font observation.</summary>
        private bool transferStarted;

        /// <summary>Initializes a FormFontObservation instance with the supplied state.</summary>
        /// <param name="value">manifest that supplies the value for this operation.</param>
        private FormFontObservation(Manifest value) { manifest = value; }

        /// <summary>Gets the form name.</summary>
        /// <value>Current form name exposed by form font observation.</value>
        internal string FormName { get { return manifest.FormName; } }

        /// <summary>Gets the has distinct name.</summary>
        /// <value>Current has distinct name exposed by form font observation.</value>
        internal bool HasDistinctName { get { return manifest.Mode == DistinctChildName; } }

        /// <summary>Gets the is after initial capture.</summary>
        /// <value>Current is after initial capture exposed by form font observation.</value>
        internal bool IsAfterInitialCapture { get { return manifest.Mode == AfterInitialCapture; } }

        /// <summary>Gets the temporary name.</summary>
        /// <value>Current temporary name exposed by form font observation.</value>
        internal string TemporaryName { get { return HasDistinctName ? manifest.TemporaryName : null; } }

        /// <summary>Attempts to begin for form font observation.</summary>
        /// <param name="projectPath">Path used for the project path being processed.</param>
        /// <param name="target">vba git snapshot that supplies the target for this operation.</param>
        /// <param name="changed">i set&lt;string&gt; that supplies the changed for this operation.</param>
        /// <param name="project">object that supplies the project for this operation.</param>
        /// <returns>form font observation produced by the operation for try begin on form font observation.</returns>
        internal static FormFontObservation TryBegin(string projectPath, VbaGitSnapshot target,
            ISet<string> changed, object project)
        {
            return TryBeginAtPath(Environment.GetEnvironmentVariable(ManifestVariable), projectPath, target, changed, project);
        }

        /// <summary>Attempts to begin at path for form font observation.</summary>
        /// <param name="path">Path used for the path being processed.</param>
        /// <param name="projectPath">Path used for the project path being processed.</param>
        /// <param name="target">vba git snapshot that supplies the target for this operation.</param>
        /// <param name="changed">i set&lt;string&gt; that supplies the changed for this operation.</param>
        /// <param name="project">object that supplies the project for this operation.</param>
        /// <returns>form font observation produced by the operation for try begin at path on form font observation.</returns>
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
            observation.Write("armed", new { value.Mode, value.FormName, value.Nonce,
                value.CandidateMvid, value.TargetFormSha256, value.TargetDescriptorHex });
            return observation;
        }

        /// <summary>Validates manifest for form font observation.</summary>
        /// <param name="value">manifest that supplies the value for this operation.</param>
        /// <param name="projectPath">Path used for the project path being processed.</param>
        /// <param name="target">vba git snapshot that supplies the target for this operation.</param>
        /// <param name="changed">i set&lt;string&gt; that supplies the changed for this operation.</param>
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
        /// <param name="observed">vba git snapshot that supplies the observed for this operation.</param>
        /// <param name="target">vba git snapshot that supplies the target for this operation.</param>
        /// <param name="formName">Text that supplies the form name value. Use the format required by the calling operation.</param>
        /// <returns>Boolean indicating the result of the check for same except selected frx on form font observation.</returns>
        internal static bool SameExceptSelectedFrx(VbaGitSnapshot observed, VbaGitSnapshot target, string formName)
        {
            if (observed == null || target == null || string.IsNullOrEmpty(formName)) return false;
            var actual = observed.ComparisonFiles(); var desired = target.ComparisonFiles();
            string resource = formName + ".frx";
            return actual.Count == desired.Count && actual.ContainsKey(resource) && desired.ContainsKey(resource) &&
                actual.Where(item => item.Key != resource).All(item => desired.TryGetValue(item.Key, out byte[] bytes) &&
                    item.Value.SequenceEqual(bytes));
        }

        /// <summary>Runs after initial capture for form font observation.</summary>
        /// <param name="capture">func&lt;vba git snapshot&gt; that supplies the capture for this operation.</param>
        /// <param name="target">vba git snapshot that supplies the target for this operation.</param>
        /// <param name="transfer">action that supplies the transfer for this operation.</param>
        internal void RunAfterInitialCapture(Func<VbaGitSnapshot> capture, VbaGitSnapshot target, Action transfer)
        {
            if (!IsAfterInitialCapture) throw new InvalidOperationException("Initial capture belongs only to the deferred mode.");
            GateAfterInitialCapture(capture, target, FormName,
                (exact, onlySelectedFrx) => Write("initial-post-import-capture",
                    new { Exact = exact, OnlySelectedFrxDiffers = onlySelectedFrx }), transfer);
        }

        /// <summary>Records the first post-import capture before any deferred native transfer.</summary>
        /// <param name="capture">func&lt;vba git snapshot&gt; that supplies the capture for this operation.</param>
        /// <param name="target">vba git snapshot that supplies the target for this operation.</param>
        /// <param name="formName">Text that supplies the form name value. Use the format required by the calling operation.</param>
        /// <param name="record">action&lt;bool, bool&gt; that supplies the record for this operation.</param>
        /// <param name="transfer">action that supplies the transfer for this operation.</param>
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

        /// <summary>Handles before deferred transfer for form font observation.</summary>
        /// <param name="descriptor">byte[] that supplies the descriptor for this operation.</param>
        internal void BeforeDeferredTransfer(byte[] descriptor)
        {
            if (!IsAfterInitialCapture || transferStarted ||
                !string.Equals(Hex(descriptor), manifest.TargetDescriptorHex, StringComparison.OrdinalIgnoreCase))
                throw new InvalidOperationException("Deferred root font transfer is not the single declared target delivery.");
            transferStarted = true;
            Write("before-deferred-transfer", new { Owner = FormName, TargetDescriptorHex = Hex(descriptor),
                Delivery = "one fresh StdFont IPersistStream.Load then Designer.Font.put" });
        }

        /// <summary>Handles deferred transfer returned for form font observation.</summary>
        internal void DeferredTransferReturned()
        {
            if (!transferStarted) throw new InvalidOperationException("No deferred font transfer was started.");
            Write("deferred-transfer-returned", new { Owner = FormName, NativeSetterReturned = true });
        }

        /// <summary>Handles before delivery for form font observation.</summary>
        /// <param name="field">Text that supplies the field value. Use the format required by the calling operation.</param>
        /// <param name="requested">object that supplies the requested for this operation.</param>
        internal void BeforeDelivery(string field, object requested)
        {
            Write("before-put", new { Field = field, Requested = requested });
        }

        /// <summary>Handles after delivery for form font observation.</summary>
        /// <param name="field">Text that supplies the field value. Use the format required by the calling operation.</param>
        /// <param name="requested">object that supplies the requested for this operation.</param>
        /// <param name="child">object that supplies the child for this operation.</param>
        internal void AfterDelivery(string field, object requested, object child)
        {
            object actual = NativeRead("VBIDE.Property.Value.get(Font." + field + ")", () => ((dynamic)child).Value);
            Write("after-put", new { Field = field, Requested = requested, Actual = actual,
                Equal = object.Equals(requested, actual) });
            if (!object.Equals(requested, actual))
                throw new InvalidOperationException("Root font child " + field + " did not read back the planned value.");
        }

        /// <summary>Observes  for form font observation.</summary>
        /// <param name="phase">Text that supplies the phase value. Use the format required by the calling operation.</param>
        /// <param name="component">object that supplies the component for this operation.</param>
        /// <param name="designer">object that supplies the designer for this operation.</param>
        /// <param name="rootProperty">object that supplies the root property for this operation.</param>
        /// <param name="children">object[] that supplies the children for this operation.</param>
        /// <param name="revalidate">action that supplies the revalidate for this operation.</param>
        internal void Observe(string phase, object component, object designer, object rootProperty, object[] children,
            Action revalidate)
        {
            revalidate();
            Write("before-observe", new { TargetPhase = phase });
            // Child preflight getters have already run. This export precedes only
            // the attached-font getters in this observation phase.
            var export = Export(component, phase);
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
                Write(phase, new { Children = values, PropertyFontDescriptor = Hex(propertyDescriptor),
                    DesignerFontDescriptor = Hex(designerDescriptor), SameFontIUnknown = same,
                    export.FrxBytes, export.FrxSha256, export.RootDescriptors });
            }
            catch (Exception error) { primary = error; throw; }
            finally { ReleaseBoth(designerFont, propertyFont, primary); }
        }

        /// <summary>Handles failure for form font observation.</summary>
        /// <param name="error">Exception describing the error failure.</param>
        internal void Failure(Exception error)
        {
            Write("throw", new { ExceptionType = error.GetType().FullName, error.HResult,
                Message = error.Message, Stage = error.InnerException?.GetType().FullName });
        }

        /// <summary>Handles complete for form font observation.</summary>
        /// <param name="exact">Indicates whether exact is enabled.</param>
        internal void Complete(bool exact)
        {
            Write("final-comparison", new { Exact = exact });
        }

        /// <summary>Handles before font getters for form font observation.</summary>
        /// <param name="component">object that supplies the component for this operation.</param>
        /// <param name="revalidate">action that supplies the revalidate for this operation.</param>
        internal void BeforeFontGetters(object component, Action revalidate)
        {
            revalidate();
            Write("before-export", new { TargetPhase = "post-import-before-font-getters" });
            var export = Export(component, "post-import-before-font-getters");
            Write("post-import-before-font-getters", new { export.FrxBytes, export.FrxSha256, export.RootDescriptors });
        }

        /// <summary>Handles export for form font observation.</summary>
        /// <param name="component">object that supplies the component for this operation.</param>
        /// <param name="phase">Text that supplies the phase value. Use the format required by the calling operation.</param>
        /// <returns>(int frx bytes, string frx sha256, string[] root descriptors) produced by the operation for export on form font observation.</returns>
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
            var files = new Dictionary<string, byte[]> {
                [form.FileName] = VbaGitSnapshot.Utf8.GetBytes(source),
                [form.Name + ".frx"] = bytes
            };
            var snapshot = new VbaGitSnapshot(new VbaGitManifest {
                References = "", Components = new[] { form }
            }, files);
            string[] roots = snapshot.FormFonts(form)?.Where(item => item.OwnerPath == "")
                .Select(item => Hex(item.Descriptor)).ToArray() ?? new string[0];
            return (bytes.Length, Sha256(bytes), roots);
        }

        /// <summary>Writes  for form font observation.</summary>
        /// <param name="phase">Text that supplies the phase value. Use the format required by the calling operation.</param>
        /// <param name="data">object that supplies the data for this operation.</param>
        private void Write(string phase, object data)
        {
            string file = Path.Combine(manifest.OutputRoot,
                (++receipt).ToString("D3", CultureInfo.InvariantCulture) + "-" + phase + ".json");
            byte[] bytes = new UTF8Encoding(false).GetBytes(
                new JavaScriptSerializer().Serialize(new { Phase = phase, Nonce = manifest.Nonce, Data = data }) + "\n");
            using (var stream = new FileStream(file, FileMode.CreateNew, FileAccess.Write, FileShare.None))
            { stream.Write(bytes, 0, bytes.Length); stream.Flush(true); }
        }

        /// <summary>Saves descriptor for form font observation.</summary>
        /// <param name="font">object that supplies the font for this operation.</param>
        /// <param name="owner">Text that supplies the owner value. Use the format required by the calling operation.</param>
        /// <returns>byte[] produced by the operation for save descriptor on form font observation.</returns>
        private static byte[] SaveDescriptor(object font, string owner)
        {
            IStream stream = null;
            IntPtr count = IntPtr.Zero;
            Exception primary = null;
            try
            {
                NativeRead("CreateStreamOnHGlobal", () => {
                    Marshal.ThrowExceptionForHR(CreateStreamOnHGlobal(IntPtr.Zero, true, out stream)); return true; });
                NativeRead("IPersistStream.Save(" + owner + ")", () => { ((PersistStream)font).Save(stream, false); return true; });
                var stat = NativeRead("IStream.Stat(" + owner + ")", () => {
                    stream.Stat(out System.Runtime.InteropServices.ComTypes.STATSTG result, 1); return result; });
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
            finally {
                if (count != IntPtr.Zero) Marshal.FreeHGlobal(count);
                try { Release(stream); }
                catch (Exception cleanup) {
                    if (primary != null) throw new AggregateException("Font observation and stream release both failed.", primary, cleanup);
                    throw;
                }
            }
        }

        /// <summary>Handles native read for form font observation.</summary>
        /// <typeparam name="T">The type used for t.</typeparam>
        /// <param name="operation">Text that supplies the operation value. Use the format required by the calling operation.</param>
        /// <param name="action">func&lt;t&gt; that supplies the action for this operation.</param>
        /// <returns>t produced by the operation for native read on form font observation.</returns>
        private static T NativeRead<T>(string operation, Func<T> action)
        {
            try { return action(); }
            catch (Exception error) when (error is COMException || error is NotSupportedException || error is InvalidCastException)
            {
                throw new InvalidOperationException("Root font observation failed at " + operation +
                    " (HRESULT 0x" + error.HResult.ToString("X8", CultureInfo.InvariantCulture) + ").", error);
            }
        }

        /// <summary>Releases both for form font observation.</summary>
        /// <param name="first">object that supplies the first for this operation.</param>
        /// <param name="second">object that supplies the second for this operation.</param>
        /// <param name="primary">Exception describing the primary failure.</param>
        private static void ReleaseBoth(object first, object second, Exception primary)
        {
            var failures = new List<Exception>();
            if (primary != null) failures.Add(primary);
            try { Release(first); } catch (Exception error) { failures.Add(error); }
            try { Release(second); } catch (Exception error) { failures.Add(error); }
            if (failures.Count > (primary == null ? 0 : 1))
                throw new AggregateException("Font observation and COM release provenance.", failures);
        }

        /// <summary>Compares com identity for form font observation.</summary>
        /// <param name="first">object that supplies the first for this operation.</param>
        /// <param name="second">object that supplies the second for this operation.</param>
        /// <returns>Boolean indicating the result of the check for same com identity on form font observation.</returns>
        private static bool SameComIdentity(object first, object second)
        {
            IntPtr a = IntPtr.Zero, b = IntPtr.Zero;
            try
            {
                a = Marshal.GetIUnknownForObject(first);
                b = Marshal.GetIUnknownForObject(second);
                return a == b;
            }
            finally {
                if (b != IntPtr.Zero) Marshal.Release(b);
                if (a != IntPtr.Zero) Marshal.Release(a);
            }
        }

        /// <summary>Determines whether sha256 for form font observation.</summary>
        /// <param name="text">Text that supplies the text value. Use the format required by the calling operation.</param>
        /// <returns>Boolean indicating the result of the check for is sha256 on form font observation.</returns>
        private static bool IsSha256(string text) { return Regex.IsMatch(text ?? "", "^[0-9A-Fa-f]{64}$"); }

        /// <summary>Determines whether hex for form font observation.</summary>
        /// <param name="text">Text that supplies the text value. Use the format required by the calling operation.</param>
        /// <returns>Boolean indicating the result of the check for is hex on form font observation.</returns>
        private static bool IsHex(string text) { return Regex.IsMatch(text ?? "", "^(?:[0-9A-Fa-f]{2}){11,64}$"); }

        /// <summary>Handles hex for form font observation.</summary>
        /// <param name="bytes">byte[] that supplies the bytes for this operation.</param>
        /// <returns>Text produced by the operation for hex on form font observation.</returns>
        private static string Hex(byte[] bytes) { return BitConverter.ToString(bytes).Replace("-", ""); }

        /// <summary>Handles sha256 for form font observation.</summary>
        /// <param name="bytes">byte[] that supplies the bytes for this operation.</param>
        /// <returns>Text produced by the operation for sha256 on form font observation.</returns>
        private static string Sha256(byte[] bytes) { using (var hash = SHA256.Create()) return Hex(hash.ComputeHash(bytes)); }

        /// <summary>Requires safe local path for form font observation.</summary>
        /// <param name="path">Path used for the path being processed.</param>
        /// <param name="mustExist">Indicates whether must exist is enabled.</param>
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

        /// <summary>Releases  for form font observation.</summary>
        /// <param name="value">object that supplies the value for this operation.</param>
        private static void Release(object value) { if (value != null && Marshal.IsComObject(value)) Marshal.ReleaseComObject(value); }

        /// <summary>Gets the native encoding.</summary>
        /// <value>Current native encoding exposed by form font observation.</value>
        private static Encoding NativeEncoding { get { return Encoding.GetEncoding((int)GetACP(),
            EncoderFallback.ExceptionFallback, DecoderFallback.ExceptionFallback); } }

        /// <summary>Returns acp for form font observation.</summary>
        /// <returns>uint produced by the operation for get acp on form font observation.</returns>
        [DllImport("kernel32.dll")] private static extern uint GetACP();

        /// <summary>Creates stream on h global for form font observation.</summary>
        /// <param name="handle">Native handle that supplies the handle for this operation.</param>
        /// <param name="free">Indicates whether free is enabled.</param>
        /// <param name="stream">i stream that supplies the stream for this operation.</param>
        /// <returns>int produced by the operation for create stream on h global on form font observation.</returns>
        [DllImport("ole32.dll", ExactSpelling = true)]
        private static extern int CreateStreamOnHGlobal(IntPtr handle, [MarshalAs(UnmanagedType.Bool)] bool free, out IStream stream);

        /// <summary>Defines the persist stream contract.</summary>
        [ComImport, Guid("00000109-0000-0000-C000-000000000046"), InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
        private interface PersistStream
        {

            /// <summary>Returns class id for persist stream.</summary>
            /// <param name="clsid">Identifier that supplies the clsid for this operation.</param>
            void GetClassID(out Guid clsid);

/// <summary>Determines whether dirty for persist stream.</summary>
/// <returns>int produced by the operation for is dirty on persist stream.</returns>
[PreserveSig] int IsDirty();

            /// <summary>Loads  for persist stream.</summary>
            /// <param name="stream">i stream that supplies the stream for this operation.</param>
            void Load(IStream stream);

/// <summary>Saves  for persist stream.</summary>
/// <param name="stream">i stream that supplies the stream for this operation.</param>
/// <param name="clearDirty">Indicates whether clear dirty is enabled.</param>
void Save(IStream stream, [MarshalAs(UnmanagedType.Bool)] bool clearDirty);

/// <summary>Returns size max for persist stream.</summary>
/// <param name="size">long that supplies the size for this operation.</param>
void GetSizeMax(out long size);
        }
    }
}

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
        internal const string ManifestVariable = "VBAi_TEST_ROOT_FONT_OBSERVATION_MANIFEST";
        internal const string ObserveWrites = "ObserveWrites";
        internal const string DistinctChildName = "DistinctChildName";
        internal const string AfterInitialCapture = "AfterInitialCapture";

        internal sealed class Manifest
        {
            public string ProjectPath { get; set; }
            public string FormName { get; set; }
            public string TargetFormSha256 { get; set; }
            public string TargetDescriptorHex { get; set; }
            public string CandidateMvid { get; set; }
            public string OutputRoot { get; set; }
            public string Nonce { get; set; }
            public string Mode { get; set; }
            public string TemporaryName { get; set; }
        }

        private readonly Manifest manifest;
        private int receipt;
        private bool transferStarted;
        private FormFontObservation(Manifest value) { manifest = value; }
        internal string FormName { get { return manifest.FormName; } }
        internal bool HasDistinctName { get { return manifest.Mode == DistinctChildName; } }
        internal bool IsAfterInitialCapture { get { return manifest.Mode == AfterInitialCapture; } }
        internal string TemporaryName { get { return HasDistinctName ? manifest.TemporaryName : null; } }

        internal static FormFontObservation TryBegin(string projectPath, VbaGitSnapshot target,
            ISet<string> changed, object project)
        {
            return TryBeginAtPath(Environment.GetEnvironmentVariable(ManifestVariable), projectPath, target, changed, project);
        }

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
        internal static bool SameExceptSelectedFrx(VbaGitSnapshot observed, VbaGitSnapshot target, string formName)
        {
            if (observed == null || target == null || string.IsNullOrEmpty(formName)) return false;
            var actual = observed.ComparisonFiles(); var desired = target.ComparisonFiles();
            string resource = formName + ".frx";
            return actual.Count == desired.Count && actual.ContainsKey(resource) && desired.ContainsKey(resource) &&
                actual.Where(item => item.Key != resource).All(item => desired.TryGetValue(item.Key, out byte[] bytes) &&
                    item.Value.SequenceEqual(bytes));
        }

        internal void RunAfterInitialCapture(Func<VbaGitSnapshot> capture, VbaGitSnapshot target, Action transfer)
        {
            if (!IsAfterInitialCapture) throw new InvalidOperationException("Initial capture belongs only to the deferred mode.");
            GateAfterInitialCapture(capture, target, FormName,
                (exact, onlySelectedFrx) => Write("initial-post-import-capture",
                    new { Exact = exact, OnlySelectedFrxDiffers = onlySelectedFrx }), transfer);
        }

        /// <summary>Records the first post-import capture before any deferred native transfer.</summary>
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

        internal void BeforeDeferredTransfer(byte[] descriptor)
        {
            if (!IsAfterInitialCapture || transferStarted ||
                !string.Equals(Hex(descriptor), manifest.TargetDescriptorHex, StringComparison.OrdinalIgnoreCase))
                throw new InvalidOperationException("Deferred root font transfer is not the single declared target delivery.");
            transferStarted = true;
            Write("before-deferred-transfer", new { Owner = FormName, TargetDescriptorHex = Hex(descriptor),
                Delivery = "one fresh StdFont IPersistStream.Load then Designer.Font.put" });
        }

        internal void DeferredTransferReturned()
        {
            if (!transferStarted) throw new InvalidOperationException("No deferred font transfer was started.");
            Write("deferred-transfer-returned", new { Owner = FormName, NativeSetterReturned = true });
        }

        internal void BeforeDelivery(string field, object requested)
        {
            Write("before-put", new { Field = field, Requested = requested });
        }

        internal void AfterDelivery(string field, object requested, object child)
        {
            object actual = NativeRead("VBIDE.Property.Value.get(Font." + field + ")", () => ((dynamic)child).Value);
            Write("after-put", new { Field = field, Requested = requested, Actual = actual,
                Equal = object.Equals(requested, actual) });
            if (!object.Equals(requested, actual))
                throw new InvalidOperationException("Root font child " + field + " did not read back the planned value.");
        }

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

        internal void Failure(Exception error)
        {
            Write("throw", new { ExceptionType = error.GetType().FullName, error.HResult,
                Message = error.Message, Stage = error.InnerException?.GetType().FullName });
        }

        internal void Complete(bool exact)
        {
            Write("final-comparison", new { Exact = exact });
        }

        internal void BeforeFontGetters(object component, Action revalidate)
        {
            revalidate();
            Write("before-export", new { TargetPhase = "post-import-before-font-getters" });
            var export = Export(component, "post-import-before-font-getters");
            Write("post-import-before-font-getters", new { export.FrxBytes, export.FrxSha256, export.RootDescriptors });
        }

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

        private void Write(string phase, object data)
        {
            string file = Path.Combine(manifest.OutputRoot,
                (++receipt).ToString("D3", CultureInfo.InvariantCulture) + "-" + phase + ".json");
            byte[] bytes = new UTF8Encoding(false).GetBytes(
                new JavaScriptSerializer().Serialize(new { Phase = phase, Nonce = manifest.Nonce, Data = data }) + "\n");
            using (var stream = new FileStream(file, FileMode.CreateNew, FileAccess.Write, FileShare.None))
            { stream.Write(bytes, 0, bytes.Length); stream.Flush(true); }
        }

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

        private static T NativeRead<T>(string operation, Func<T> action)
        {
            try { return action(); }
            catch (Exception error) when (error is COMException || error is NotSupportedException || error is InvalidCastException)
            {
                throw new InvalidOperationException("Root font observation failed at " + operation +
                    " (HRESULT 0x" + error.HResult.ToString("X8", CultureInfo.InvariantCulture) + ").", error);
            }
        }

        private static void ReleaseBoth(object first, object second, Exception primary)
        {
            var failures = new List<Exception>();
            if (primary != null) failures.Add(primary);
            try { Release(first); } catch (Exception error) { failures.Add(error); }
            try { Release(second); } catch (Exception error) { failures.Add(error); }
            if (failures.Count > (primary == null ? 0 : 1))
                throw new AggregateException("Font observation and COM release provenance.", failures);
        }

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

        private static bool IsSha256(string text) { return Regex.IsMatch(text ?? "", "^[0-9A-Fa-f]{64}$"); }
        private static bool IsHex(string text) { return Regex.IsMatch(text ?? "", "^(?:[0-9A-Fa-f]{2}){11,64}$"); }
        private static string Hex(byte[] bytes) { return BitConverter.ToString(bytes).Replace("-", ""); }
        private static string Sha256(byte[] bytes) { using (var hash = SHA256.Create()) return Hex(hash.ComputeHash(bytes)); }
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
        private static void Release(object value) { if (value != null && Marshal.IsComObject(value)) Marshal.ReleaseComObject(value); }
        private static Encoding NativeEncoding { get { return Encoding.GetEncoding((int)GetACP(),
            EncoderFallback.ExceptionFallback, DecoderFallback.ExceptionFallback); } }
        [DllImport("kernel32.dll")] private static extern uint GetACP();
        [DllImport("ole32.dll", ExactSpelling = true)]
        private static extern int CreateStreamOnHGlobal(IntPtr handle, [MarshalAs(UnmanagedType.Bool)] bool free, out IStream stream);
        [ComImport, Guid("00000109-0000-0000-C000-000000000046"), InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
        private interface PersistStream
        {
            void GetClassID(out Guid clsid); [PreserveSig] int IsDirty();
            void Load(IStream stream); void Save(IStream stream, [MarshalAs(UnmanagedType.Bool)] bool clearDirty); void GetSizeMax(out long size);
        }
    }
}

using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Runtime.InteropServices.ComTypes;
using System.Text;
using System.Web.Script.Serialization;

namespace VBAi
{
    internal sealed partial class FormFontObservation
    {
        /// <summary>Separate opt-in so attached observations never select the experimental restoration modes.</summary>
        internal const string AttachedManifestVariable = "VBAi_TEST_ATTACHED_FONT_OBSERVATION_MANIFEST";
        /// <summary>Known stop after baseline observation, before recovery or native import admission.</summary>
        internal const string AttachedBaselineStop = "Attached font baseline observation completed; native import intentionally not admitted.";

        /// <summary>Frozen scope for one synthetic owner-font observation.</summary>
        internal sealed class AttachedManifest
        {
            /// <summary>Exact disposable workbook path.</summary>
            public string ProjectPath { get; set; }
            /// <summary>Exact changed resource-bearing form.</summary>
            public string FormName { get; set; }
            /// <summary>Loaded candidate module identity.</summary>
            public string CandidateMvid { get; set; }
            /// <summary>Exact source digest in the frozen selected snapshot.</summary>
            public string FormSha256 { get; set; }
            /// <summary>Exact FRX digest in the frozen selected snapshot.</summary>
            public string ResourceSha256 { get; set; }
            /// <summary>Nonce-bound fresh evidence directory.</summary>
            public string OutputRoot { get; set; }
            /// <summary>One-use nonce in N GUID format.</summary>
            public string Nonce { get; set; }
            /// <summary>Either baseline-only or font-delivery-returned.</summary>
            public string Stage { get; set; }
            /// <summary>One IFontDisp property to read on each declared owner.</summary>
            public string Getter { get; set; }
            /// <summary>Exactly the root and one declared Frame owner, in that order.</summary>
            public string[] OwnerPaths { get; set; }
            /// <summary>Explicit permission for the bounded detached Frame clone/load/Size prototype.</summary>
            public bool DetachedFrameClone { get; set; }
        }

        /// <summary>One armed observation; no metric getter, setter, clone, or load is executed while arming.</summary>
        internal sealed class Attached
        {
            private readonly AttachedManifest value;
            private readonly string path, hash;
            private readonly byte[] frameDescriptor;
            private int receipt;
            private bool observed;
            /// <summary>Frozen stage selected before any observation.</summary>
            internal string Stage => value.Stage;
            /// <summary>Frozen form selected before any observation.</summary>
            internal string FormName => value.FormName;
            private Attached(AttachedManifest value, string path, string hash, byte[] frameDescriptor)
            { this.value = value; this.path = path; this.hash = hash; this.frameDescriptor = frameDescriptor; }

            /// <summary>Arms a fresh, exact candidate/snapshot/path/nonce observation, or returns before all access when disabled.</summary>
            internal static Attached TryBegin(string path, string projectPath, VbaGitSnapshot expected, VbaGitSnapshot target, string evidenceRoot = null)
            {
                if (string.IsNullOrWhiteSpace(path)) return null;
                if (!string.IsNullOrWhiteSpace(Environment.GetEnvironmentVariable(ManifestVariable)))
                    throw new InvalidOperationException("Attached and experimental root-font observations cannot be combined.");
                RequireSafeLocalPath(path, true);
                byte[] bytes = ReadAttachedManifestBytes(path);
                if (bytes.Length == 0 || bytes.Length > 16384) throw new InvalidOperationException("Attached font manifest exceeds its bound.");
                var value = ParseAttachedManifest(bytes);
                ValidateAttachedManifest(value, projectPath, expected, target);
                if (Path.GetFileName(path) != value.Nonce + ".attached-font.json")
                    throw new InvalidOperationException("Nonce-named attached font manifest required.");
                if (evidenceRoot != null)
                {
                    OwnerGitQualificationManifest.RequireChild(evidenceRoot, path, directory: false);
                    OwnerGitQualificationManifest.RequireChild(evidenceRoot, value.OutputRoot, directory: true);
                }
                RequireSafeLocalPath(value.ProjectPath, true);
                RequireSafeLocalPath(value.OutputRoot, false);
                var snapshot = value.Stage == "baseline-only" ? expected : target;
                var form = snapshot.Manifest.Components.Single(item => item.Name == value.FormName);
                var descriptor = snapshot.FormFonts(form).Single(item => item.OwnerPath == value.OwnerPaths[1]).Descriptor;
                var result = new Attached(value, path, Sha256(bytes), (byte[])descriptor.Clone());
                Directory.CreateDirectory(value.OutputRoot);
                RequireSafeLocalPath(value.OutputRoot, false);
                result.Write("armed", new { value.Stage, value.Getter, value.FormName, value.OwnerPaths, value.FormSha256,
                    value.ResourceSha256, value.CandidateMvid, ManifestSha256 = result.hash,
                    PotentialFontCacheMutation = true, SnapshotResigningPermitted = false });
                return result;
            }

            /// <summary>Observes one real root and Frame font, with durable intent before each potentially realizing operation.</summary>
            internal void Observe(object component, Action guard)
            {
                if (observed) throw new InvalidOperationException("Attached font observation was already consumed.");
                guard(); RequireManifest(); observed = true;
                var owned = new List<object>(); Exception primary = null;
                try
                {
                    Write("designer-get-intent", new { PotentialFontCacheMutation = true });
                    guard(); RequireManifest();
                    object designer = ((dynamic)component).Designer; owned.Add(designer);
                    object controls = ((dynamic)designer).Controls; owned.Add(controls);
                    string name = value.OwnerPaths[1].Substring("Controls/".Length);
                    object frame = ((dynamic)controls).Item(name); owned.Add(frame);
                    object parent = ((dynamic)frame).Parent; owned.Add(parent);
                    if (!string.Equals(Convert.ToString(((dynamic)frame).Name), name, StringComparison.Ordinal) ||
                        !FormFontRestoration.SameFontOwnerParent(parent, designer, true) ||
                        !string.Equals(TypeDescriptor.GetClassName(frame), "Frame", StringComparison.OrdinalIgnoreCase))
                        throw new InvalidOperationException("Attached font owner hierarchy differs from its frozen declaration.");
                    var fonts = new object[2];
                    for (int i = 0; i < fonts.Length; i++)
                    {
                        guard(); RequireManifest();
                        Write("font-object-get-intent", new { OwnerPath = value.OwnerPaths[i] });
                        guard(); RequireManifest();
                        fonts[i] = ((dynamic)(i == 0 ? designer : frame)).Font; owned.Add(fonts[i]);
                    }
                    Write("font-identities", new { RootAndFrameSameIUnknown = SameComIdentity(fonts[0], fonts[1]),
                        RootIFont = Supports(fonts[0], "BEF6E002-A874-101A-8BBA-00AA00300CAB"),
                        FrameIFont = Supports(fonts[1], "BEF6E002-A874-101A-8BBA-00AA00300CAB"),
                        RootIFontDisp = Supports(fonts[0], "BEF6E003-A874-101A-8BBA-00AA00300CAB"),
                        FrameIFontDisp = Supports(fonts[1], "BEF6E003-A874-101A-8BBA-00AA00300CAB") });
                    // Persist BOTH owners before any metric getter, so shared identity cannot hide the initial state.
                    var before = new byte[2][]; var classes = new Guid[2];
                    for (int i = 0; i < fonts.Length; i++)
                    {
                        guard(); RequireManifest();
                        Write("persist-before-intent", new { OwnerPath = value.OwnerPaths[i] });
                        guard(); RequireManifest();
                        ((PersistStream)fonts[i]).GetClassID(out Guid classId); classes[i] = classId;
                        guard(); RequireManifest();
                        before[i] = SaveDescriptor(fonts[i], value.OwnerPaths[i]);
                        Write("persist-before", new { OwnerPath = value.OwnerPaths[i], ClassId = classId.ToString("D"),
                            BytesHex = Hex(before[i]), BytesSha256 = Sha256(before[i]) });
                    }
                    for (int i = 0; i < fonts.Length; i++)
                    {
                        int index = i; guard(); RequireManifest();
                        RunAttachedGetter(() => ReadAttachedGetter((AttachedFontDisp)fonts[index], value.Getter),
                            () => SaveDescriptor(fonts[index], value.OwnerPaths[index]),
                            (phase, data) => Write(phase, new { OwnerPath = value.OwnerPaths[index], value.Getter, Data = data }),
                            () => { guard(); RequireManifest(); });
                    }
                    // Save both again after the two isolated reads; no additional property getter is introduced.
                    for (int i = 0; i < fonts.Length; i++)
                    {
                        guard(); RequireManifest();
                        Write("persist-final-intent", new { OwnerPath = value.OwnerPaths[i] });
                        guard(); RequireManifest();
                        byte[] after = SaveDescriptor(fonts[i], value.OwnerPaths[i]);
                        Write("persist-final", new { OwnerPath = value.OwnerPaths[i], BytesHex = Hex(after),
                            BytesSha256 = Sha256(after), ExactBefore = before[i].SequenceEqual(after) });
                    }
                    if (value.DetachedFrameClone)
                        ObserveDetachedClone(fonts, classes[1], guard);
                    Write("completed", new { PotentialFontCacheMutation = true, NativeImportsAdded = 0,
                        OwnerFontSetters = 0, OriginalFontLoads = 0, DetachedPrototypeRequested = value.DetachedFrameClone });
                }
                catch (Exception error) { primary = error; throw; }
                finally { FormFontRestoration.ReleaseOwnedReferences(owned, Release, primary); }
            }

            private void ObserveDetachedClone(object[] originals, Guid frameClass, Action guard)
            {
                Guid stdFont = new Guid("0BE35203-8F91-11CE-9DE3-00AA004BB851");
                if (frameClass != stdFont || !Supports(originals[1], "BEF6E002-A874-101A-8BBA-00AA00300CAB") ||
                    !Supports(originals[1], "00000109-0000-0000-C000-000000000046"))
                {
                    Write("detached-unsupported", new { FrameClassId = frameClass.ToString("D"), Loads = 0, Clones = 0 }); return;
                }
                FormFontRestoration.ValidateDescriptor(frameDescriptor);
                AttachedFont clone = null; IStream stream = null; Exception primary = null;
                try
                {
                    Write("detached-clone-intent", new { FrameClassId = frameClass.ToString("D"), DescriptorHex = Hex(frameDescriptor) });
                    guard(); RequireManifest(); ((AttachedFont)originals[1]).Clone(out clone);
                    if (clone == null || SameComIdentity(clone, originals[0]) || SameComIdentity(clone, originals[1]))
                        throw new InvalidOperationException("Detached font clone has no independent COM identity.");
                    if (!Supports(clone, "00000109-0000-0000-C000-000000000046") ||
                        !Supports(clone, "BEF6E003-A874-101A-8BBA-00AA00300CAB"))
                    { Write("detached-unsupported", new { Clones = 1, Loads = 0 }); return; }
                    ((PersistStream)clone).GetClassID(out Guid cloneClass);
                    if (cloneClass != stdFont)
                    { Write("detached-unsupported", new { CloneClassId = cloneClass.ToString("D"), Clones = 1, Loads = 0 }); return; }
                    Write("detached-persist-before-intent", null);
                    guard(); RequireManifest();
                    byte[] before = SaveDescriptor(clone, "detached-before-load");
                    Write("detached-clone-before", new { ClassId = cloneClass.ToString("D"), BytesHex = Hex(before),
                        IndependentOfRootAndFrame = true });
                    Marshal.ThrowExceptionForHR(CreateStreamOnHGlobal(IntPtr.Zero, true, out stream));
                    stream.Write(frameDescriptor, frameDescriptor.Length, IntPtr.Zero); stream.Seek(0, 0, IntPtr.Zero);
                    Write("detached-load-intent", new { DescriptorHex = Hex(frameDescriptor), OriginalFontLoads = 0, OwnerFontSetters = 0 });
                    guard(); RequireManifest(); ((PersistStream)clone).Load(stream);
                    Write("detached-loaded-persist-intent", null);
                    guard(); RequireManifest(); byte[] loaded = SaveDescriptor(clone, "detached-loaded");
                    Write("detached-load-returned", new { BytesHex = Hex(loaded), ExactDeclaredDescriptor = loaded.SequenceEqual(frameDescriptor) });
                    RunAttachedGetter(() => ((AttachedFontDisp)clone).Size, () => SaveDescriptor(clone, "detached-after-Size"),
                        (phase, data) => Write("detached-" + phase, data), () => { guard(); RequireManifest(); });
                    // Observe any unexpected original persistence change, without another font getter.
                    for (int i = 0; i < originals.Length; i++)
                    {
                        Write("original-after-clone-intent", new { OwnerPath = value.OwnerPaths[i] });
                        guard(); RequireManifest(); byte[] bytes = SaveDescriptor(originals[i], value.OwnerPaths[i]);
                        Write("original-after-clone", new { OwnerPath = value.OwnerPaths[i], BytesHex = Hex(bytes) });
                    }
                }
                catch (Exception error) { primary = error; throw; }
                finally { FormFontRestoration.ReleaseOwnedReferences(new object[] { clone, stream }, Release, primary); }
            }

            private void RequireManifest()
            {
                RequireSafeLocalPath(path, true); RequireSafeLocalPath(value.OutputRoot, false);
                byte[] bytes = ReadAttachedManifestBytes(path);
                if (bytes.Length > 16384 || !string.Equals(Sha256(bytes), hash, StringComparison.Ordinal))
                    throw new InvalidOperationException("Attached font manifest changed after arming.");
            }
            private void Write(string phase, object data)
            {
                if (++receipt > 40) throw new InvalidOperationException("Attached font receipt budget exceeded.");
                RequireSafeLocalPath(value.OutputRoot, false);
                byte[] bytes = new UTF8Encoding(false).GetBytes(new JavaScriptSerializer().Serialize(
                    new { Phase = phase, value.Nonce, Data = data, ObservedUtc = DateTime.UtcNow.ToString("o") }) + "\n");
                if (bytes.Length > 16384) throw new InvalidOperationException("Attached font receipt size exceeded.");
                string file = Path.Combine(value.OutputRoot, receipt.ToString("D3") + "-" + phase + ".json");
                using (var stream = new FileStream(file, FileMode.CreateNew, FileAccess.Write, FileShare.None))
                { stream.Write(bytes, 0, bytes.Length); stream.Flush(true); }
            }
        }

        /// <summary>Validates only managed frozen metadata before any COM acquisition or evidence write.</summary>
        internal static void ValidateAttachedManifest(AttachedManifest value, string projectPath, VbaGitSnapshot expected, VbaGitSnapshot target)
        {
            if (value == null || (value.Stage != "baseline-only" && value.Stage != "font-delivery-returned") ||
                !new[] { "Name", "Size", "Bold", "Italic", "Underline", "Strikethrough", "Weight", "Charset" }.Contains(value.Getter) ||
                !Guid.TryParseExact(value.Nonce, "N", out _) || !Guid.TryParse(value.CandidateMvid, out Guid mvid) ||
                mvid != typeof(FormFontObservation).Module.ModuleVersionId ||
                !string.Equals(value.ProjectPath, projectPath, StringComparison.OrdinalIgnoreCase) ||
                !IsSha256(value.FormSha256) || !IsSha256(value.ResourceSha256) ||
                value.OutputRoot == null || !Path.IsPathRooted(value.OutputRoot) ||
                Path.GetFileName(value.OutputRoot) != value.Nonce || Directory.Exists(value.OutputRoot) || File.Exists(value.OutputRoot))
                throw new InvalidOperationException("Invalid attached font observation manifest.");
            var snapshot = value.Stage == "baseline-only" ? expected : target;
            var form = snapshot?.Manifest.Components?.SingleOrDefault(item => item.Name == value.FormName);
            if (form == null || form.Type != 3 || !form.HasResources ||
                !snapshot.Files.TryGetValue(form.FileName, out byte[] source) || !snapshot.Files.TryGetValue(form.Name + ".frx", out byte[] resource) ||
                !string.Equals(Sha256(source), value.FormSha256, StringComparison.OrdinalIgnoreCase) ||
                !string.Equals(Sha256(resource), value.ResourceSha256, StringComparison.OrdinalIgnoreCase))
                throw new InvalidOperationException("Attached font snapshot declaration differs from the frozen form.");
            var bindings = snapshot.FormFonts(form);
            if (value.OwnerPaths == null || value.OwnerPaths.Length != 2 || value.OwnerPaths[0] != "" ||
                value.OwnerPaths[1] == null || !value.OwnerPaths[1].StartsWith("Controls/", StringComparison.Ordinal) ||
                value.OwnerPaths[1].Split('/').Length != 2 ||
                bindings == null || !value.OwnerPaths.All(owner => bindings.Count(item => item.OwnerPath == owner &&
                    item.Type == (owner.Length == 0 ? 7 : 14)) == 1))
                throw new InvalidOperationException("Attached font owners must be the declared root and direct Frame.");
        }

        /// <summary>One intent/getter/save/result sequence; failures never retry the getter or add a later save.</summary>
        internal static void RunAttachedGetter(Func<object> getter, Func<byte[]> save, Action<string, object> write, Action guard = null)
        {
            write("single-getter-intent", null);
            guard?.Invoke();
            object value = getter();
            write("single-getter-returned", new { Value = value });
            write("persist-after-intent", null);
            guard?.Invoke();
            byte[] bytes = save();
            write("persist-after", new { BytesHex = Hex(bytes), BytesSha256 = Sha256(bytes) });
        }

        /// <summary>Requires the exact private diagnostic schema without accepting extra or case-shifted fields.</summary>
        internal static AttachedManifest ParseAttachedManifest(byte[] bytes)
        {
            if (bytes == null || bytes.Length == 0 || bytes.Length > 16384)
                throw new InvalidOperationException("Attached font manifest exceeds its bound.");
            var json = new JavaScriptSerializer { MaxJsonLength = 16384 };
            string text = new UTF8Encoding(false, true).GetString(bytes);
            var fields = json.DeserializeObject(text) as IDictionary<string, object>;
            string[] names = { "ProjectPath", "FormName", "CandidateMvid", "FormSha256", "ResourceSha256", "OutputRoot", "Nonce", "Stage", "Getter", "OwnerPaths", "DetachedFrameClone" };
            if (fields == null || fields.Count != names.Length || fields.Keys.Except(names, StringComparer.Ordinal).Any() ||
                !(fields["DetachedFrameClone"] is bool))
                throw new InvalidOperationException("Invalid attached font manifest schema.");
            return json.Deserialize<AttachedManifest>(text);
        }

        private static byte[] ReadAttachedManifestBytes(string path)
        {
            if (new FileInfo(path).Length > 16384) throw new InvalidOperationException("Attached font manifest exceeds its bound.");
            using (var stream = new FileStream(path, FileMode.Open, FileAccess.Read, FileShare.Read))
            {
                if (stream.Length == 0 || stream.Length > 16384) throw new InvalidOperationException("Attached font manifest exceeds its bound.");
                byte[] bytes = new byte[checked((int)stream.Length)]; int offset = 0;
                while (offset < bytes.Length)
                {
                    int count = stream.Read(bytes, offset, bytes.Length - offset);
                    if (count == 0) throw new IOException("Attached font manifest read was incomplete.");
                    offset += count;
                }
                if (stream.ReadByte() != -1) throw new IOException("Attached font manifest changed during its read.");
                return bytes;
            }
        }

        private static bool Supports(object font, string iid)
        {
            IntPtr unknown = IntPtr.Zero, queried = IntPtr.Zero;
            try
            {
                unknown = Marshal.GetIUnknownForObject(font); Guid id = new Guid(iid);
                int result = Marshal.QueryInterface(unknown, ref id, out queried);
                if (result == unchecked((int)0x80004002)) return false;
                Marshal.ThrowExceptionForHR(result); return true;
            }
            finally { if (queried != IntPtr.Zero) Marshal.Release(queried); if (unknown != IntPtr.Zero) Marshal.Release(unknown); }
        }
        private static object ReadAttachedGetter(AttachedFontDisp font, string getter)
        {
            switch (getter)
            {
                case "Name": return font.Name; case "Size": return font.Size; case "Bold": return font.Bold;
                case "Italic": return font.Italic; case "Underline": return font.Underline;
                case "Strikethrough": return font.Strikethrough; case "Weight": return font.Weight; case "Charset": return font.Charset;
                default: throw new InvalidOperationException("Undeclared font getter.");
            }
        }
        // Exact IFont vtable prefix through Clone from Windows SDK ocidl.h; no property setter is called.
        [ComImport, Guid("BEF6E002-A874-101A-8BBA-00AA00300CAB"), InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
        private interface AttachedFont
        {
            void GetName([MarshalAs(UnmanagedType.BStr)] out string value); void SetName([MarshalAs(UnmanagedType.BStr)] string value);
            void GetSize([MarshalAs(UnmanagedType.Currency)] out decimal value); void SetSize([MarshalAs(UnmanagedType.Currency)] decimal value);
            void GetBold([MarshalAs(UnmanagedType.Bool)] out bool value); void SetBold([MarshalAs(UnmanagedType.Bool)] bool value);
            void GetItalic([MarshalAs(UnmanagedType.Bool)] out bool value); void SetItalic([MarshalAs(UnmanagedType.Bool)] bool value);
            void GetUnderline([MarshalAs(UnmanagedType.Bool)] out bool value); void SetUnderline([MarshalAs(UnmanagedType.Bool)] bool value);
            void GetStrikethrough([MarshalAs(UnmanagedType.Bool)] out bool value); void SetStrikethrough([MarshalAs(UnmanagedType.Bool)] bool value);
            void GetWeight(out short value); void SetWeight(short value); void GetCharset(out short value); void SetCharset(short value);
            void GetHFont(out IntPtr value); void Clone([MarshalAs(UnmanagedType.Interface)] out AttachedFont value);
        }

        [ComImport, Guid("BEF6E003-A874-101A-8BBA-00AA00300CAB"), InterfaceType(ComInterfaceType.InterfaceIsIDispatch)]
        private interface AttachedFontDisp
        {
            [DispId(0)] string Name { [return: MarshalAs(UnmanagedType.BStr)] get; }
            [DispId(2)] decimal Size { [return: MarshalAs(UnmanagedType.Currency)] get; }
            [DispId(3)] bool Bold { [return: MarshalAs(UnmanagedType.VariantBool)] get; }
            [DispId(4)] bool Italic { [return: MarshalAs(UnmanagedType.VariantBool)] get; }
            [DispId(5)] bool Underline { [return: MarshalAs(UnmanagedType.VariantBool)] get; }
            [DispId(6)] bool Strikethrough { [return: MarshalAs(UnmanagedType.VariantBool)] get; }
            [DispId(7)] short Weight { get; }
            [DispId(8)] short Charset { get; }
        }
    }
}

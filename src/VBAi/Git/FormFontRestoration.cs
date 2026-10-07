using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Diagnostics;
using System.Globalization;
using System.Runtime.InteropServices;
using System.Runtime.InteropServices.ComTypes;
using System.Text;
using System.Threading;

namespace VBAi
{

    /// <summary>Restores only preflighted, declared standard fonts on the actual native VBE owner thread.</summary>
    internal static partial class FormFontRestoration
    {

        /// <summary>Refuses external-STA font transfer before any project mutation is started.</summary>
        /// <param name="project">Live project whose VBE main-window HWND must belong to this process and STA thread.</param>
        internal static void RequireOwner(object project)
        {
            object editor = null, main = null;
            Exception primary = null;
            try
            {
                editor = ((dynamic)project).VBE; main = ((dynamic)editor).MainWindow;
                IntPtr handle = new IntPtr(Convert.ToInt64(((dynamic)main).HWnd));
                uint thread = GetWindowThreadProcessId(handle, out uint pid);
                if (handle == IntPtr.Zero || pid != (uint)Process.GetCurrentProcess().Id || thread == 0 ||
                    thread != GetCurrentThreadId() || Thread.CurrentThread.GetApartmentState() != ApartmentState.STA)
                    throw new InvalidOperationException("UserForm font restoration requires the actual owning VBE process and STA thread.");
            }
            catch (Exception error) { primary = error; throw; }
            finally { ReleaseOwnedReferences(new object[] { editor, main }, Release, primary); }
        }

        /// <summary>Preflights every declared font owner and root child before any delivery, with identity guards.</summary>
        /// <param name="component">Imported UserForm component whose designer and properties are inspected.</param>
        /// <param name="bindings">Fully preflighted root/nested standard-font descriptors from the matching FRX.</param>
        /// <param name="revalidate">Project and imported-component identity check repeated before native mutation.</param>
        /// <param name="observation">Optional one-shot diagnostic receipt collector for a selected form.</param>
        internal static void Restore(object component, FormStreamPadding.FormFontBinding[] bindings, Action revalidate,
            FormFontObservation observation = null)
        {
            if (bindings == null || bindings.Length == 0) return;
            var references = new List<object>();
            var owners = new List<object>();
            var rootProperties = new List<object[]>();
            Exception primary = null;
            try
            {
                revalidate();
                if (observation != null && !observation.IsAfterInitialCapture)
                    observation.BeforeFontGetters(component, revalidate);
                object designer = NativeRead<object>("VBComponent.Designer.get", () => ((dynamic)component).Designer); references.Add(designer);
                foreach (var binding in bindings)
                {
                    ValidateDescriptor(binding.Descriptor);
                    object current = designer;
                    if (binding.OwnerPath.Length == 0)
                    {
                        if (binding.Type != 7)
                            throw new InvalidOperationException("The root persisted font owner is not a UserForm.");
                        object properties = NativeRead<object>("VBComponent.Properties.get", () => ((dynamic)component).Properties); references.Add(properties);
                        current = NativeRead<object>("VBIDE.Properties.Item(Font)", () => ((dynamic)properties).Item("Font")); references.Add(current);
                        if (!string.Equals(NativeRead<string>("VBIDE.Property.Name.get", () => Convert.ToString(((dynamic)current).Name)), "Font", StringComparison.Ordinal) ||
                            NativeRead<int>("VBIDE.Property.NumIndices.get", () => Convert.ToInt32(((dynamic)current).NumIndices)) != 0)
                            throw new InvalidOperationException("The imported UserForm Font property changed.");
                        object children = NativeRead<object>("VBIDE.Property.Value.get(Font)", () => ((dynamic)current).Value);
                        references.Add(children);
                        var propertiesByName = new object[ChildNames.Length];
                        for (int index = 0; index < ChildNames.Length; index++)
                        {
                            string name = ChildNames[index];
                            object child = NativeRead<object>("VBIDE.Properties.Item(Font." + name + ")", () => ((dynamic)children).Item(name));
                            references.Add(child);
                            if (!string.Equals(NativeRead<string>("VBIDE.Property.Name.get(Font." + name + ")", () => Convert.ToString(((dynamic)child).Name)), name, StringComparison.Ordinal) ||
                                NativeRead<int>("VBIDE.Property.NumIndices.get(Font." + name + ")", () => Convert.ToInt32(((dynamic)child).NumIndices)) != 0)
                                throw new InvalidOperationException("The imported UserForm Font." + name + " property changed.");
                            object value = NativeRead<object>("VBIDE.Property.Value.get(Font." + name + ")", () => ((dynamic)child).Value);
                            if (value == null || value.GetType() != ChildTypes[index])
                                throw new InvalidOperationException("The imported UserForm Font." + name + " value type changed.");
                            propertiesByName[index] = child;
                        }
                        rootProperties.Add(propertiesByName);
                    }
                    else
                    {
                        string[] parts = binding.OwnerPath.Split('/');
                        if (parts.Length % 2 != 0 || parts.Length > 128)
                            throw new InvalidOperationException("Invalid persisted font owner hierarchy.");
                        for (int i = 0; i < parts.Length; i += 2)
                        {
                            if (parts[i] != "Controls" && parts[i] != "Pages")
                                throw new InvalidOperationException("Unsupported persisted font owner collection.");
                            object collection = parts[i] == "Controls" ? ((dynamic)current).Controls : ((dynamic)current).Pages;
                            references.Add(collection);
                            object child = ((dynamic)collection).Item(parts[i + 1]); references.Add(child);
                            object parent = ((dynamic)child).Parent; references.Add(parent);
                            if (!string.Equals(Convert.ToString(((dynamic)child).Name), parts[i + 1], StringComparison.Ordinal) ||
                                !SameFontOwnerParent(parent, current, ReferenceEquals(current, designer)))
                                throw new InvalidOperationException("The imported font owner hierarchy changed.");
                            current = child;
                        }
                        string expectedClass = binding.Type == 14 ? "Frame" : binding.Type == 57 ? "MultiPage" : "Page";
                        if (!string.Equals(TypeDescriptor.GetClassName(current), expectedClass, StringComparison.OrdinalIgnoreCase))
                            throw new InvalidOperationException("The imported font owner type differs from its resource storage.");
                        rootProperties.Add(null);
                    }
                    owners.Add(current);
                }
                for (int i = 0; i < bindings.Length; i++)
                {
                    if (rootProperties[i] == null)
                    {
                        revalidate();
                        AssignNestedClone(owners[i], bindings[i].Descriptor, revalidate);
                    }
                    else
                    {
                        if (observation != null && observation.IsAfterInitialCapture)
                        {
                            DeliverDeferredRoot(bindings[i].Descriptor, revalidate,
                                observation.BeforeDeferredTransfer,
                                descriptor => AssignNested(designer, descriptor, "Designer.Font.set", revalidate),
                                observation.DeferredTransferReturned);
                            observation.Observe("after-deferred-transfer", component, designer, owners[i], rootProperties[i], revalidate);
                        }
                        else
                        {
                            observation?.Observe("before-write-after-preflight", component, designer, owners[i], rootProperties[i], revalidate);
                            AssignRoot(rootProperties[i], bindings[i].Descriptor, revalidate, observation,
                                () => observation.Observe("after-temporary-name", component, designer, owners[i], rootProperties[i], revalidate));
                            observation?.Observe("after-root", component, designer, owners[i], rootProperties[i], revalidate);
                        }
                    }
                }
            }
            catch (Exception error) { primary = error; throw; }
            finally { ReleaseOwnedReferences(references, Release, primary); }
        }

        /// <summary>One predeclared attached-font delivery after the initial capture; no fallback or replay.</summary>
        /// <param name="descriptor">Exact target root-font descriptor.</param>
        /// <param name="revalidate">Fresh authorization and identity check immediately before delivery.</param>
        /// <param name="intent">Records the declared delivery intent before any native call.</param>
        /// <param name="deliver">Single native transfer; never retried if it throws.</param>
        /// <param name="returned">Records that the setter returned successfully.</param>
        internal static void DeliverDeferredRoot(byte[] descriptor, Action revalidate, Action<byte[]> intent,
            Action<byte[]> deliver, Action returned)
        {
            revalidate();
            intent(descriptor);
            deliver(descriptor);
            returned();
        }

        /// <summary>Releases every acquired reference and retains the native failure before any cleanup failures.</summary>
        /// <param name="references">Owned COM wrappers or pointers acquired during preflight and transfer.</param>
        /// <param name="release">Type-aware release callback for each acquired value.</param>
        /// <param name="primary">Primary operation failure to preserve before cleanup failures.</param>
        internal static void ReleaseOwnedReferences(IList<object> references, Action<object> release, Exception primary)
        {
            var failures = new List<Exception>();
            if (primary != null) failures.Add(primary);
            for (int i = references.Count - 1; i >= 0; i--)
            {
                try { release(references[i]); }
                catch (Exception error) { failures.Add(error); }
            }
            if (failures.Count > (primary == null ? 0 : 1))
                throw new AggregateException("UserForm font restoration and COM cleanup provenance.", failures);
        }

        /// <summary>Accepts the bounded MS-OFORMS StdFont profile; no decoder or COM activation occurs.</summary>
        /// <param name="data">Serialized StdFont descriptor bytes from the validated FRX storage.</param>
        /// <exception cref="InvalidOperationException">The profile, flags, size, or ASCII font name falls outside the supported bounds.</exception>
        internal static void ValidateDescriptor(byte[] data)
        {
            if (data == null || data.Length < 11 || data[0] != 1 || (data[3] & ~14) != 0 ||
                BitConverter.ToUInt16(data, 4) > 1000 || BitConverter.ToUInt32(data, 6) == 0 ||
                BitConverter.ToUInt32(data, 6) > 655350000 || data[10] >= 32 || data.Length != 11 + data[10])
                throw new InvalidOperationException("Unsupported persisted standard font descriptor.");
            for (int i = 11; i < data.Length; i++)
                if (data[i] >= 128) throw new InvalidOperationException("Unsupported persisted standard font name.");
        }

        /// <summary>Root UserForm Font child property names in descriptor-delivery order.</summary>
        private static readonly string[] ChildNames = { "Name", "Size", "Bold", "Italic", "Underline", "Strikethrough", "Weight", "Charset" };

        /// <summary>Required managed value types returned by corresponding VBIDE font properties.</summary>
        private static readonly Type[] ChildTypes = { typeof(string), typeof(decimal), typeof(bool), typeof(bool), typeof(bool), typeof(bool), typeof(short), typeof(short) };

        /// <summary>Delivers each declared scalar once on the owning STA; Weight follows all style fields.</summary>
        /// <param name="children">Preflighted VBIDE property wrappers indexed by <see cref="ChildNames"/>.</param>
        /// <param name="data">Validated persisted StdFont descriptor.</param>
        /// <param name="revalidate">Project/component revalidation before each setter.</param>
        /// <param name="observation">Optional receipt callbacks for each requested and read-back value.</param>
        /// <param name="afterTemporary">Optional observation after a temporary distinct Name was written.</param>
        private static void AssignRoot(object[] children, byte[] data, Action revalidate,
            FormFontObservation observation, Action afterTemporary)
        {
            AssignRootCore(children, data, revalidate, observation?.TemporaryName,
                observation == null ? null : new Action<string, object>(observation.BeforeDelivery),
                observation == null ? null : new Action<string, object, object>(observation.AfterDelivery), afterTemporary);
        }

        /// <summary>One predeclared delivery per field; the diagnostic may prepend one distinct Name value.</summary>
        /// <param name="children">Preflighted root-font property wrappers.</param>
        /// <param name="data">Validated serialized descriptor supplying the target values.</param>
        /// <param name="revalidate">Check repeated before each irreversible property setter.</param>
        /// <param name="temporaryName">Optional distinct experimental Name value written before the descriptor's name.</param>
        /// <param name="beforeDelivery">Receipt callback invoked before each setter.</param>
        /// <param name="afterDelivery">Readback callback invoked after each setter returns.</param>
        /// <param name="afterTemporary">Optional hook after the temporary Name setter and readback.</param>
        internal static void AssignRootCore(object[] children, byte[] data, Action revalidate, string temporaryName,
            Action<string, object> beforeDelivery, Action<string, object, object> afterDelivery, Action afterTemporary)
        {
            object[] values = {
                Encoding.ASCII.GetString(data, 11, data[10]),
                BitConverter.ToUInt32(data, 6) / 10000m,
                null,
                (data[3] & 2) != 0,
                (data[3] & 4) != 0,
                (data[3] & 8) != 0,
                (short)BitConverter.ToUInt16(data, 4),
                BitConverter.ToInt16(data, 1)
            };
            if (temporaryName != null)
            {
                revalidate();
                beforeDelivery?.Invoke("Name.temporary", temporaryName);
                try { ((dynamic)children[0]).Value = temporaryName; }
                catch (Exception error) when (error is COMException || error is NotSupportedException)
                { throw NativeFailure("VBIDE.Property.Value.set(Font.Name.temporary)", error); }
                afterDelivery?.Invoke("Name.temporary", temporaryName, children[0]);
                afterTemporary?.Invoke();
            }
            foreach (int index in new[] { 0, 1, 7, 3, 4, 5, 6 })
            {
                revalidate();
                string operation = "VBIDE.Property.Value.set(Font." + ChildNames[index] + ")";
                beforeDelivery?.Invoke(ChildNames[index], values[index]);
                try { ((dynamic)children[index]).Value = values[index]; }
                catch (Exception error) when (error is COMException || error is NotSupportedException)
                { throw NativeFailure(operation, error); }
                afterDelivery?.Invoke(ChildNames[index], values[index], children[index]);
            }
        }

        /// <summary>Validates a parent alias without requiring Name on the design-time UserForm wrapper.</summary>
        /// <param name="parent">Parent acquired from the child selected in the validated owner collection.</param>
        /// <param name="expected">Container whose collection supplied that child.</param>
        /// <param name="root">Whether the expected container is the current component's exact designer.</param>
        /// <returns>True for an identical COM object or the matching typed container alias.</returns>
        internal static bool SameFontOwnerParent(object parent, object expected, bool root)
        {
            if (parent == null || expected == null) return false;
            if (VbeProjectHostPath.SameProject(parent, expected)) return true;
            string parentType = TypeDescriptor.GetClassName(parent);
            string expectedType = TypeDescriptor.GetClassName(expected);
            if (string.IsNullOrWhiteSpace(parentType) ||
                !string.Equals(parentType, expectedType, StringComparison.OrdinalIgnoreCase)) return false;
            // The owning project/component is revalidated before this traversal,
            // and the child comes from this designer's Controls collection. Its
            // design-time parent may be a distinct UserForm wrapper without Name.
            if (root) return string.Equals(expectedType, "UserForm", StringComparison.OrdinalIgnoreCase);
            string parentName = Convert.ToString(((dynamic)parent).Name);
            string expectedName = Convert.ToString(((dynamic)expected).Name);
            return !string.IsNullOrWhiteSpace(parentName) &&
                string.Equals(parentName, expectedName, StringComparison.Ordinal);
        }

        /// <summary>Loads the exact descriptor into a new local font, then transfers it once to a nested owner.</summary>
        /// <param name="owner">Preflighted nested Page, Frame, or MultiPage font owner.</param>
        /// <param name="data">Exact validated persisted font descriptor.</param>
        /// <param name="setter">Diagnostic label for the final native Font assignment.</param>
        /// <param name="revalidate">Optional identity and policy check before assignment.</param>
        private static void AssignNested(object owner, byte[] data, string setter = "MSForms.Font.set", Action revalidate = null)
        {
            var description = new FontDescription
            {
                StructureSize = (uint)Marshal.SizeOf(typeof(FontDescription)),
                Name = Encoding.ASCII.GetString(data, 11, data[10]),
                Size = BitConverter.ToUInt32(data, 6),
                Weight = (short)BitConverter.ToUInt16(data, 4),
                Charset = BitConverter.ToInt16(data, 1),
                Italic = (data[3] & 2) != 0,
                Underline = (data[3] & 4) != 0,
                Strikethrough = (data[3] & 8) != 0
            };
            IntPtr pointer = IntPtr.Zero; object font = null; IStream stream = null;
            string operation = "OleCreateFontIndirect";
            Exception primary = null;
            try
            {
                var iid = new Guid("BEF6E003-A874-101A-8BBA-00AA00300CAB");
                Marshal.ThrowExceptionForHR(OleCreateFontIndirect(ref description, ref iid, out pointer));
                if (pointer == IntPtr.Zero) throw new InvalidOperationException("Native font factory returned no interface.");
                operation = "Marshal.GetObjectForIUnknown";
                font = Marshal.GetObjectForIUnknown(pointer);
                operation = "CreateStreamOnHGlobal";
                Marshal.ThrowExceptionForHR(CreateStreamOnHGlobal(IntPtr.Zero, true, out stream));
                operation = "IStream.Write";
                stream.Write(data, data.Length, IntPtr.Zero);
                operation = "IStream.Seek";
                stream.Seek(0, 0, IntPtr.Zero);
                operation = "IPersistStream.Load";
                ((PersistStream)font).Load(stream);
                // Native outcome is never retried. The caller's complete snapshot
                // comparison remains authoritative, including all font bytes.
                operation = setter;
                revalidate?.Invoke();
                ((dynamic)owner).Font = font;
            }
            catch (Exception error) when (error is COMException || error is NotSupportedException)
            {
                // COM interop can map an unsupported native operation to a managed
                // NotSupportedException. Preserve the stage and original exception
                // without attempting a second delivery.
                primary = NativeFailure(operation, error);
                throw primary;
            }
            catch (Exception error) { primary = error; throw; }
            finally
            {
                ReleaseOwnedReferences(new object[] { pointer, font, stream }, value =>
                {
                    if (value is IntPtr acquired && acquired != IntPtr.Zero) Marshal.Release(acquired);
                    else Release(value);
                }, primary);
            }
        }

        /// <summary>Retains the exact failed native getter without exposing project content or retrying it.</summary>
        /// <typeparam name="T">Native getter result type.</typeparam>
        /// <param name="operation">Constant operation label included in failure diagnostics.</param>
        /// <param name="read">Single native getter to invoke.</param>
        /// <returns>Getter result.</returns>
        private static T NativeRead<T>(string operation, Func<T> read)
        {
            try { return read(); }
            catch (Exception error) when (error is COMException || error is NotSupportedException)
            { throw NativeFailure(operation, error); }
        }

        /// <summary>Preserves the original native-operation failure and HRESULT with its constant operation name.</summary>
        /// <param name="operation">Constant native operation label.</param>
        /// <param name="error">Original COM or unsupported-operation exception.</param>
        /// <returns>Invalid-operation wrapper preserving the original exception and HRESULT.</returns>
        private static InvalidOperationException NativeFailure(string operation, Exception error)
        {
            return new InvalidOperationException("UserForm font restoration failed at " + operation +
                " (HRESULT 0x" + error.HResult.ToString("X8", CultureInfo.InvariantCulture) + "): " + error.Message, error);
        }

        /// <summary>Releases one owned runtime-callable COM wrapper.</summary>
        /// <param name="value">Acquired COM object.</param>
        private static void Release(object value) { if (value != null && Marshal.IsComObject(value)) Marshal.ReleaseComObject(value); }

        /// <summary>Reads the owning thread and process IDs for a native window handle.</summary>
        /// <param name="handle">Window handle being checked.</param>
        /// <param name="process">Receives the owning process ID.</param>
        /// <returns>Owning thread ID, or zero on failure.</returns>
        [DllImport("user32.dll")] private static extern uint GetWindowThreadProcessId(IntPtr handle, out uint process);

        /// <summary>Reads the native identifier of the current OS thread.</summary>
        /// <returns>Current thread ID.</returns>
        [DllImport("kernel32.dll")] private static extern uint GetCurrentThreadId();

        /// <summary>Creates an IStream backed by HGLOBAL for the exact persisted font descriptor.</summary>
        /// <param name="handle">Existing HGLOBAL, or zero to allocate one.</param>
        /// <param name="free">Whether releasing the stream frees the global block.</param>
        /// <param name="stream">Receives the stream interface.</param>
        /// <returns>OLE HRESULT.</returns>
        [DllImport("ole32.dll", ExactSpelling = true)]
        private static extern int CreateStreamOnHGlobal(IntPtr handle, [MarshalAs(UnmanagedType.Bool)] bool free, out IStream stream);

        /// <summary>Creates a local standard font object from the bounded descriptor metadata.</summary>
        /// <param name="description">LOGFONT-like structure decoded from the persisted descriptor.</param>
        /// <param name="iid">Requested IFont interface identifier.</param>
        /// <param name="font">Receives the created interface pointer.</param>
        /// <returns>OLE HRESULT.</returns>
        [DllImport("oleaut32.dll", ExactSpelling = true)]
        private static extern int OleCreateFontIndirect(ref FontDescription description, ref Guid iid, out IntPtr font);

        /// <summary>Carries the font description values passed between operations.</summary>
        [StructLayout(LayoutKind.Sequential, CharSet = CharSet.Unicode)]
        private struct FontDescription
        {

            /// <summary>Byte size of this native structure, required by OleCreateFontIndirect.</summary>
            internal uint StructureSize;

            /// <summary>Unicode face name decoded from the persisted ASCII profile.</summary>
            [MarshalAs(UnmanagedType.LPWStr)] internal string Name;

            /// <summary>Font height in native HIMETRIC units.</summary>
            internal long Size;

            /// <summary>Native font weight and character-set identifiers.</summary>
            internal short Weight, Charset;

            /// <summary>Whether the persisted font is italic.</summary>
            [MarshalAs(UnmanagedType.Bool)] internal bool Italic;

            /// <summary>Whether the persisted font is underlined.</summary>
            [MarshalAs(UnmanagedType.Bool)] internal bool Underline;

            /// <summary>Whether the persisted font is struck through.</summary>
            [MarshalAs(UnmanagedType.Bool)] internal bool Strikethrough;
        }

        /// <summary>COM vtable contract for loading a font from its persisted descriptor stream.</summary>
        [ComImport, Guid("00000109-0000-0000-C000-000000000046"), InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
        private interface PersistStream
        {

            /// <summary>Returns the class ID of the persisted font object.</summary>
            /// <param name="clsid">Receives the class ID.</param>
            void GetClassID(out Guid clsid);

            /// <summary>Reports whether the font has state not reflected by its persisted stream.</summary>
            /// <returns>HRESULT indicating dirty state.</returns>
            [PreserveSig] int IsDirty();

            /// <summary>Loads font state from the supplied persisted descriptor stream.</summary>
            /// <param name="stream">Descriptor source.</param>
            void Load(IStream stream);

            /// <summary>Saves font state to a stream.</summary>
            /// <param name="stream">Destination stream.</param>
            /// <param name="clearDirty">Whether saving clears the object's dirty state.</param>
            void Save(IStream stream, [MarshalAs(UnmanagedType.Bool)] bool clearDirty);

            /// <summary>Returns the maximum stream size required to save this object.</summary>
            /// <param name="size">Receives the maximum byte count.</param>
            void GetSizeMax(out long size);
        }
    }
}

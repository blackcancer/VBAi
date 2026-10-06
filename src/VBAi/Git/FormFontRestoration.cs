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
    internal static class FormFontRestoration
    {

        /// <summary>Refuses external-STA font transfer before any project mutation is started.</summary>
        /// <param name="project">object that supplies the project for this operation.</param>
        internal static void RequireOwner(object project)
        {
            object editor = null, main = null;
            Exception primary = null;
            try
            {
                editor = ((dynamic)project).VBE; main = ((dynamic)editor).MainWindow;
                IntPtr handle = new IntPtr(Convert.ToInt64(((dynamic)main).HWnd));
                uint pid; uint thread = GetWindowThreadProcessId(handle, out pid);
                if (handle == IntPtr.Zero || pid != (uint)Process.GetCurrentProcess().Id || thread == 0 ||
                    thread != GetCurrentThreadId() || Thread.CurrentThread.GetApartmentState() != ApartmentState.STA)
                    throw new InvalidOperationException("UserForm font restoration requires the actual owning VBE process and STA thread.");
            }
            catch (Exception error) { primary = error; throw; }
            finally { ReleaseOwnedReferences(new object[] { editor, main }, Release, primary); }
        }

        /// <summary>Preflights every declared font owner and root child before any delivery, with identity guards.</summary>
        /// <param name="component">object that supplies the component for this operation.</param>
        /// <param name="bindings">form font binding[] that supplies the bindings for this operation.</param>
        /// <param name="revalidate">action that supplies the revalidate for this operation.</param>
        /// <param name="observation">form font observation that supplies the observation for this operation.</param>
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
                        AssignNested(owners[i], bindings[i].Descriptor);
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
                            if (observation != null)
                                observation.Observe("before-write-after-preflight", component, designer, owners[i], rootProperties[i], revalidate);
                            AssignRoot(rootProperties[i], bindings[i].Descriptor, revalidate, observation,
                                () => observation.Observe("after-temporary-name", component, designer, owners[i], rootProperties[i], revalidate));
                            if (observation != null)
                                observation.Observe("after-root", component, designer, owners[i], rootProperties[i], revalidate);
                        }
                    }
                }
            }
            catch (Exception error) { primary = error; throw; }
            finally { ReleaseOwnedReferences(references, Release, primary); }
        }

        /// <summary>One predeclared attached-font delivery after the initial capture; no fallback or replay.</summary>
        /// <param name="descriptor">byte[] that supplies the descriptor for this operation.</param>
        /// <param name="revalidate">action that supplies the revalidate for this operation.</param>
        /// <param name="intent">action&lt;byte[]&gt; that supplies the intent for this operation.</param>
        /// <param name="deliver">action&lt;byte[]&gt; that supplies the deliver for this operation.</param>
        /// <param name="returned">action that supplies the returned for this operation.</param>
        internal static void DeliverDeferredRoot(byte[] descriptor, Action revalidate, Action<byte[]> intent,
            Action<byte[]> deliver, Action returned)
        {
            revalidate();
            intent(descriptor);
            deliver(descriptor);
            returned();
        }

        /// <summary>Releases every acquired reference and retains the native failure before any cleanup failures.</summary>
        /// <param name="references">i list&lt;object&gt; that supplies the references for this operation.</param>
        /// <param name="release">action&lt;object&gt; that supplies the release for this operation.</param>
        /// <param name="primary">Exception describing the primary failure.</param>
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
        /// <param name="data">byte[] that supplies the data for this operation.</param>
        internal static void ValidateDescriptor(byte[] data)
        {
            if (data == null || data.Length < 11 || data[0] != 1 || (data[3] & ~14) != 0 ||
                BitConverter.ToUInt16(data, 4) > 1000 || BitConverter.ToUInt32(data, 6) == 0 ||
                BitConverter.ToUInt32(data, 6) > 655350000 || data[10] >= 32 || data.Length != 11 + data[10])
                throw new InvalidOperationException("Unsupported persisted standard font descriptor.");
            for (int i = 11; i < data.Length; i++)
                if (data[i] >= 128) throw new InvalidOperationException("Unsupported persisted standard font name.");
        }

        /// <summary>Maintains the child names state for form font restoration.</summary>
        private static readonly string[] ChildNames = { "Name", "Size", "Bold", "Italic", "Underline", "Strikethrough", "Weight", "Charset" };

        /// <summary>Maintains the child types state for form font restoration.</summary>
        private static readonly Type[] ChildTypes = { typeof(string), typeof(decimal), typeof(bool), typeof(bool), typeof(bool), typeof(bool), typeof(short), typeof(short) };

        /// <summary>Delivers each declared scalar once on the owning STA; Weight follows all style fields.</summary>
        /// <param name="children">object[] that supplies the children for this operation.</param>
        /// <param name="data">byte[] that supplies the data for this operation.</param>
        /// <param name="revalidate">action that supplies the revalidate for this operation.</param>
        /// <param name="observation">form font observation that supplies the observation for this operation.</param>
        /// <param name="afterTemporary">action that supplies the after temporary for this operation.</param>
        private static void AssignRoot(object[] children, byte[] data, Action revalidate,
            FormFontObservation observation, Action afterTemporary)
        {
            AssignRootCore(children, data, revalidate, observation?.TemporaryName,
                observation == null ? null : new Action<string, object>(observation.BeforeDelivery),
                observation == null ? null : new Action<string, object, object>(observation.AfterDelivery), afterTemporary);
        }

        /// <summary>One predeclared delivery per field; the diagnostic may prepend one distinct Name value.</summary>
        /// <param name="children">object[] that supplies the children for this operation.</param>
        /// <param name="data">byte[] that supplies the data for this operation.</param>
        /// <param name="revalidate">action that supplies the revalidate for this operation.</param>
        /// <param name="temporaryName">Text that supplies the temporary name value. Use the format required by the calling operation.</param>
        /// <param name="beforeDelivery">action&lt;string, object&gt; that supplies the before delivery for this operation.</param>
        /// <param name="afterDelivery">action&lt;string, object, object&gt; that supplies the after delivery for this operation.</param>
        /// <param name="afterTemporary">action that supplies the after temporary for this operation.</param>
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
        /// <param name="owner">object that supplies the owner for this operation.</param>
        /// <param name="data">byte[] that supplies the data for this operation.</param>
        /// <param name="setter">Text that supplies the setter value. Use the format required by the calling operation.</param>
        /// <param name="revalidate">action that supplies the revalidate for this operation.</param>
        private static void AssignNested(object owner, byte[] data, string setter = "MSForms.Font.set", Action revalidate = null)
        {
            var description = new FontDescription {
                StructureSize = (uint)Marshal.SizeOf(typeof(FontDescription)),
                Name = Encoding.ASCII.GetString(data, 11, data[10]), Size = BitConverter.ToUInt32(data, 6),
                Weight = (short)BitConverter.ToUInt16(data, 4), Charset = BitConverter.ToInt16(data, 1),
                Italic = (data[3] & 2) != 0, Underline = (data[3] & 4) != 0, Strikethrough = (data[3] & 8) != 0
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
                ReleaseOwnedReferences(new object[] { pointer, font, stream }, value => {
                    if (value is IntPtr acquired && acquired != IntPtr.Zero) Marshal.Release(acquired);
                    else Release(value);
                }, primary);
            }
        }

        /// <summary>Retains the exact failed native getter without exposing project content or retrying it.</summary>
        /// <typeparam name="T">The type used for t.</typeparam>
        /// <param name="operation">Text that supplies the operation value. Use the format required by the calling operation.</param>
        /// <param name="read">func&lt;t&gt; that supplies the read for this operation.</param>
        /// <returns>t produced by the operation for native read on form font restoration.</returns>
        private static T NativeRead<T>(string operation, Func<T> read)
        {
            try { return read(); }
            catch (Exception error) when (error is COMException || error is NotSupportedException)
            { throw NativeFailure(operation, error); }
        }

        /// <summary>Preserves the original native-operation failure and HRESULT with its constant operation name.</summary>
        /// <param name="operation">Text that supplies the operation value. Use the format required by the calling operation.</param>
        /// <param name="error">Exception describing the error failure.</param>
        /// <returns>invalid operation exception produced by the operation for native failure on form font restoration.</returns>
        private static InvalidOperationException NativeFailure(string operation, Exception error)
        {
            return new InvalidOperationException("UserForm font restoration failed at " + operation +
                " (HRESULT 0x" + error.HResult.ToString("X8", CultureInfo.InvariantCulture) + "): " + error.Message, error);
        }

        /// <summary>Releases  for form font restoration.</summary>
        /// <param name="value">object that supplies the value for this operation.</param>
        private static void Release(object value) { if (value != null && Marshal.IsComObject(value)) Marshal.ReleaseComObject(value); }

        /// <summary>Returns window thread process id for form font restoration.</summary>
        /// <param name="handle">Native handle that supplies the handle for this operation.</param>
        /// <param name="process">uint that supplies the process for this operation.</param>
        /// <returns>uint produced by the operation for get window thread process id on form font restoration.</returns>
        [DllImport("user32.dll")] private static extern uint GetWindowThreadProcessId(IntPtr handle, out uint process);

        /// <summary>Returns current thread id for form font restoration.</summary>
        /// <returns>uint produced by the operation for get current thread id on form font restoration.</returns>
        [DllImport("kernel32.dll")] private static extern uint GetCurrentThreadId();

        /// <summary>Creates stream on h global for form font restoration.</summary>
        /// <param name="handle">Native handle that supplies the handle for this operation.</param>
        /// <param name="free">Indicates whether free is enabled.</param>
        /// <param name="stream">i stream that supplies the stream for this operation.</param>
        /// <returns>int produced by the operation for create stream on h global on form font restoration.</returns>
        [DllImport("ole32.dll", ExactSpelling = true)]
        private static extern int CreateStreamOnHGlobal(IntPtr handle, [MarshalAs(UnmanagedType.Bool)] bool free, out IStream stream);

        /// <summary>Handles ole create font indirect for form font restoration.</summary>
        /// <param name="description">font description that supplies the description for this operation.</param>
        /// <param name="iid">Identifier that supplies the iid for this operation.</param>
        /// <param name="font">Native handle that supplies the font for this operation.</param>
        /// <returns>int produced by the operation for ole create font indirect on form font restoration.</returns>
        [DllImport("oleaut32.dll", ExactSpelling = true)]
        private static extern int OleCreateFontIndirect(ref FontDescription description, ref Guid iid, out IntPtr font);

        /// <summary>Carries the font description values passed between operations.</summary>
        [StructLayout(LayoutKind.Sequential, CharSet = CharSet.Unicode)]
        private struct FontDescription
        {

            /// <summary>Maintains the structure size state for font description.</summary>
            internal uint StructureSize;

            /// <summary>Maintains the name state for font description.</summary>
            [MarshalAs(UnmanagedType.LPWStr)] internal string Name;

            /// <summary>Maintains the size state for font description.</summary>
            internal long Size;

            /// <summary>Maintains the weight and charset state for font description.</summary>
            internal short Weight, Charset;

            /// <summary>Maintains the italic state for font description.</summary>
            [MarshalAs(UnmanagedType.Bool)] internal bool Italic;

            /// <summary>Maintains the underline state for font description.</summary>
            [MarshalAs(UnmanagedType.Bool)] internal bool Underline;

            /// <summary>Maintains the strikethrough state for font description.</summary>
            [MarshalAs(UnmanagedType.Bool)] internal bool Strikethrough;
        }

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

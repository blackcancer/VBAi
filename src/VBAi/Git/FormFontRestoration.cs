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
        internal static void RequireOwner(object project)
        {
            object editor = null, main = null;
            try
            {
                editor = ((dynamic)project).VBE; main = ((dynamic)editor).MainWindow;
                IntPtr handle = new IntPtr(Convert.ToInt64(((dynamic)main).HWnd));
                uint pid; uint thread = GetWindowThreadProcessId(handle, out pid);
                if (handle == IntPtr.Zero || pid != (uint)Process.GetCurrentProcess().Id || thread == 0 ||
                    thread != GetCurrentThreadId() || Thread.CurrentThread.GetApartmentState() != ApartmentState.STA)
                    throw new InvalidOperationException("UserForm font restoration requires the actual owning VBE process and STA thread.");
            }
            finally { Release(main); Release(editor); }
        }

        /// <summary>Preflights every declared font owner and root child before any delivery, with identity guards.</summary>
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
                observation?.BeforeFontGetters(component, revalidate);
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
                                !string.Equals(Convert.ToString(((dynamic)parent).Name), Convert.ToString(((dynamic)current).Name), StringComparison.Ordinal))
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
                        if (observation != null)
                            observation.Observe("before-write-after-preflight", component, designer, owners[i], rootProperties[i], revalidate);
                        AssignRoot(rootProperties[i], bindings[i].Descriptor, revalidate, observation,
                            () => observation.Observe("after-temporary-name", component, designer, owners[i], rootProperties[i], revalidate));
                        if (observation != null)
                            observation.Observe("after-root", component, designer, owners[i], rootProperties[i], revalidate);
                    }
                }
            }
            catch (Exception error) { primary = error; throw; }
            finally { ReleaseOwnedReferences(references, Release, primary); }
        }

        /// <summary>Releases every acquired reference and retains the native failure before any cleanup failures.</summary>
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
        internal static void ValidateDescriptor(byte[] data)
        {
            if (data == null || data.Length < 11 || data[0] != 1 || (data[3] & ~14) != 0 ||
                BitConverter.ToUInt16(data, 4) > 1000 || BitConverter.ToUInt32(data, 6) == 0 ||
                BitConverter.ToUInt32(data, 6) > 655350000 || data[10] >= 32 || data.Length != 11 + data[10])
                throw new InvalidOperationException("Unsupported persisted standard font descriptor.");
            for (int i = 11; i < data.Length; i++)
                if (data[i] >= 128) throw new InvalidOperationException("Unsupported persisted standard font name.");
        }

        private static readonly string[] ChildNames = { "Name", "Size", "Bold", "Italic", "Underline", "Strikethrough", "Weight", "Charset" };
        private static readonly Type[] ChildTypes = { typeof(string), typeof(decimal), typeof(bool), typeof(bool), typeof(bool), typeof(bool), typeof(short), typeof(short) };

        /// <summary>Delivers each declared scalar once on the owning STA; Weight follows all style fields.</summary>
        private static void AssignRoot(object[] children, byte[] data, Action revalidate,
            FormFontObservation observation, Action afterTemporary)
        {
            AssignRootCore(children, data, revalidate, observation?.TemporaryName,
                observation == null ? null : new Action<string, object>(observation.BeforeDelivery),
                observation == null ? null : new Action<string, object, object>(observation.AfterDelivery), afterTemporary);
        }

        /// <summary>One predeclared delivery per field; the diagnostic may prepend one distinct Name value.</summary>
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

        /// <summary>Loads the exact descriptor into a new local font, then transfers it once to a nested owner.</summary>
        private static void AssignNested(object owner, byte[] data)
        {
            var description = new FontDescription {
                StructureSize = (uint)Marshal.SizeOf(typeof(FontDescription)),
                Name = Encoding.ASCII.GetString(data, 11, data[10]), Size = BitConverter.ToUInt32(data, 6),
                Weight = (short)BitConverter.ToUInt16(data, 4), Charset = BitConverter.ToInt16(data, 1),
                Italic = (data[3] & 2) != 0, Underline = (data[3] & 4) != 0, Strikethrough = (data[3] & 8) != 0
            };
            IntPtr pointer = IntPtr.Zero; object font = null; IStream stream = null;
            string operation = "OleCreateFontIndirect";
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
                operation = "MSForms.Font.set";
                ((dynamic)owner).Font = font;
            }
            catch (Exception error) when (error is COMException || error is NotSupportedException)
            {
                // COM interop can map an unsupported native operation to a managed
                // NotSupportedException. Preserve the stage and original exception
                // without attempting a second delivery.
                throw NativeFailure(operation, error);
            }
            finally
            {
                try { Release(stream); } finally { try { Release(font); } finally { if (pointer != IntPtr.Zero) Marshal.Release(pointer); } }
            }
        }

        /// <summary>Retains the exact failed native getter without exposing project content or retrying it.</summary>
        private static T NativeRead<T>(string operation, Func<T> read)
        {
            try { return read(); }
            catch (Exception error) when (error is COMException || error is NotSupportedException)
            { throw NativeFailure(operation, error); }
        }

        /// <summary>Preserves the original native-operation failure and HRESULT with its constant operation name.</summary>
        private static InvalidOperationException NativeFailure(string operation, Exception error)
        {
            return new InvalidOperationException("UserForm font restoration failed at " + operation +
                " (HRESULT 0x" + error.HResult.ToString("X8", CultureInfo.InvariantCulture) + "): " + error.Message, error);
        }

        private static void Release(object value) { if (value != null && Marshal.IsComObject(value)) Marshal.ReleaseComObject(value); }
        [DllImport("user32.dll")] private static extern uint GetWindowThreadProcessId(IntPtr handle, out uint process);
        [DllImport("kernel32.dll")] private static extern uint GetCurrentThreadId();
        [DllImport("ole32.dll", ExactSpelling = true)]
        private static extern int CreateStreamOnHGlobal(IntPtr handle, [MarshalAs(UnmanagedType.Bool)] bool free, out IStream stream);
        [DllImport("oleaut32.dll", ExactSpelling = true)]
        private static extern int OleCreateFontIndirect(ref FontDescription description, ref Guid iid, out IntPtr font);
        [StructLayout(LayoutKind.Sequential, CharSet = CharSet.Unicode)]
        private struct FontDescription
        {
            internal uint StructureSize;
            [MarshalAs(UnmanagedType.LPWStr)] internal string Name;
            internal long Size;
            internal short Weight, Charset;
            [MarshalAs(UnmanagedType.Bool)] internal bool Italic;
            [MarshalAs(UnmanagedType.Bool)] internal bool Underline;
            [MarshalAs(UnmanagedType.Bool)] internal bool Strikethrough;
        }
        [ComImport, Guid("00000109-0000-0000-C000-000000000046"), InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
        private interface PersistStream
        {
            void GetClassID(out Guid clsid); [PreserveSig] int IsDirty();
            void Load(IStream stream); void Save(IStream stream, [MarshalAs(UnmanagedType.Bool)] bool clearDirty); void GetSizeMax(out long size);
        }
    }
}

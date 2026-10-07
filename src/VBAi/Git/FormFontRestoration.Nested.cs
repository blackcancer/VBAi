using System;
using System.Runtime.InteropServices;
using System.Runtime.InteropServices.ComTypes;
using System.Threading;

namespace VBAi
{
    internal static partial class FormFontRestoration
    {
        /// <summary>Clones the actual attached font state before loading and transferring one nested descriptor.</summary>
        /// <param name="owner">Preflighted native Frame, Page, or MultiPage owner.</param>
        /// <param name="data">Exact validated persisted standard-font descriptor.</param>
        /// <param name="revalidate">Current project, mode and imported-component identity guard.</param>
        private static void AssignNestedClone(object owner, byte[] data, Action revalidate)
        {
            if (Thread.CurrentThread.GetApartmentState() != ApartmentState.STA)
                throw new InvalidOperationException("Nested font transfer requires the actual owning STA thread.");
            AssignNestedCloneCore(owner, data, revalidate, new NativeNestedFontTransfer());
        }

        /// <summary>Preserves native clone ownership and single-delivery ordering without observing font metrics.</summary>
        /// <param name="owner">Preflighted nested font owner.</param>
        /// <param name="data">Exact target descriptor; never loaded into the original attached font.</param>
        /// <param name="revalidate">Identity and authorization guard before acquisition and each transfer stage.</param>
        /// <param name="native">Native transfer boundary, including acquired-reference release.</param>
        internal static void AssignNestedCloneCore(object owner, byte[] data, Action revalidate, INestedFontTransfer native)
        {
            ValidateDescriptor(data);
            if (revalidate == null) throw new ArgumentNullException(nameof(revalidate));
            if (native == null) throw new ArgumentNullException(nameof(native));
            object original = null, clone = null;
            Exception primary = null;
            string operation = "MSForms.Font.get";
            try
            {
                revalidate();
                original = native.ReadFont(owner);
                operation = "AttachedFont.interfaces";
                if (original == null || !native.SupportsTransfer(original))
                    throw new InvalidOperationException("The attached nested font does not support standard-font transfer.");
                operation = "AttachedFont.IPersistStream.GetClassID";
                Guid stdFont = new Guid("0BE35203-8F91-11CE-9DE3-00AA004BB851");
                if (native.ClassId(original) != stdFont)
                    throw new InvalidOperationException("The attached nested font is not a persisted standard font.");
                operation = "IFont.Clone";
                revalidate();
                clone = native.Clone(original);
                operation = "ClonedFont.IUnknown.identity";
                if (clone == null || ReferenceEquals(clone, original) || native.SameIdentity(clone, original))
                    throw new InvalidOperationException("The nested font clone has no independent COM identity.");
                operation = "ClonedFont.interfaces";
                if (!native.SupportsTransfer(clone))
                    throw new InvalidOperationException("The nested font clone does not support standard-font transfer.");
                operation = "ClonedFont.IPersistStream.GetClassID";
                if (native.ClassId(clone) != stdFont)
                    throw new InvalidOperationException("The nested font clone is not a persisted standard font.");
                operation = "IPersistStream.Load";
                revalidate();
                native.Load(clone, data);
                // Only the detached clone is loaded. A failed or uncertain native
                // assignment is never replayed or replaced by a factory fallback.
                operation = "MSForms.Font.set";
                revalidate();
                native.Assign(owner, clone);
            }
            catch (Exception error) when (error is COMException || error is NotSupportedException)
            {
                primary = NativeFailure(operation, error);
                throw primary;
            }
            catch (Exception error) { primary = error; throw; }
            finally
            {
                // Each returned font carries one acquisition, even when native
                // Clone wrongly aliases its source. Balance both acquisitions.
                ReleaseOwnedReferences(new[] { original, clone }, value =>
                {
                    if (value != null) native.Release(value);
                }, primary);
            }
        }

        /// <summary>Bounded native operations for acquiring, cloning and transferring one nested font.</summary>
        internal interface INestedFontTransfer
        {
            /// <summary>Acquires the owner's current attached font once without reading metrics.</summary>
            object ReadFont(object owner);
            /// <summary>Requires native IFont, IFontDisp and IPersistStream support.</summary>
            bool SupportsTransfer(object font);
            /// <summary>Reads the persisted font object's class without a property getter.</summary>
            Guid ClassId(object font);
            /// <summary>Creates one independent, acquired font with the source's native state.</summary>
            object Clone(object font);
            /// <summary>Compares native IUnknown identities while balancing temporary pointers.</summary>
            bool SameIdentity(object first, object second);
            /// <summary>Loads one exact target descriptor into the detached clone only.</summary>
            void Load(object font, byte[] data);
            /// <summary>Performs one native owner Font assignment.</summary>
            void Assign(object owner, object font);
            /// <summary>Balances one acquired font reference without final-releasing borrowed owners.</summary>
            void Release(object font);
        }

        /// <summary>Actual COM implementation; retains native font context without reading or changing scaling.</summary>
        private sealed class NativeNestedFontTransfer : INestedFontTransfer
        {
            public object ReadFont(object owner) { return ((dynamic)owner).Font; }
            public bool SupportsTransfer(object font)
            {
                return Marshal.IsComObject(font) && font is NestedNativeFont &&
                    font is PersistStream && font is NestedNativeFontDisp;
            }
            public Guid ClassId(object font) { ((PersistStream)font).GetClassID(out Guid value); return value; }
            public object Clone(object font)
            {
                NestedNativeFont value = null;
                try { ((NestedNativeFont)font).Clone(out value); return value; }
                catch (Exception error)
                {
                    // A failed HRESULT can still populate the out acquisition.
                    // Release it once without repeating the native clone call.
                    Exception primary = error is COMException || error is NotSupportedException
                        ? NativeFailure("IFont.Clone", error) : error;
                    ReleaseOwnedReferences(new object[] { value }, FormFontRestoration.Release, primary);
                    throw primary;
                }
            }
            public bool SameIdentity(object first, object second) { return VbeProjectHostPath.SameProject(first, second); }
            public void Assign(object owner, object font) { ((dynamic)owner).Font = font; }
            public void Release(object font) { FormFontRestoration.Release(font); }
            public void Load(object font, byte[] data)
            {
                IStream stream = null; Exception primary = null;
                string operation = "CreateStreamOnHGlobal";
                try
                {
                    Marshal.ThrowExceptionForHR(CreateStreamOnHGlobal(IntPtr.Zero, true, out stream));
                    operation = "IStream.Write"; stream.Write(data, data.Length, IntPtr.Zero);
                    operation = "IStream.Seek"; stream.Seek(0, 0, IntPtr.Zero);
                    operation = "IPersistStream.Load"; ((PersistStream)font).Load(stream);
                }
                catch (Exception error) when (error is COMException || error is NotSupportedException)
                { primary = NativeFailure(operation, error); throw primary; }
                catch (Exception error) { primary = error; throw; }
                finally { ReleaseOwnedReferences(new object[] { stream }, FormFontRestoration.Release, primary); }
            }
        }

        // Exact IFont vtable prefix through Clone, verified against Windows SDK
        // ocidl.h. The getter/setter slots are required for layout, never invoked.
        [ComImport, Guid("BEF6E002-A874-101A-8BBA-00AA00300CAB"), InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
        private interface NestedNativeFont
        {
            void GetName([MarshalAs(UnmanagedType.BStr)] out string value); void SetName([MarshalAs(UnmanagedType.BStr)] string value);
            void GetSize([MarshalAs(UnmanagedType.Currency)] out decimal value); void SetSize([MarshalAs(UnmanagedType.Currency)] decimal value);
            void GetBold([MarshalAs(UnmanagedType.Bool)] out bool value); void SetBold([MarshalAs(UnmanagedType.Bool)] bool value);
            void GetItalic([MarshalAs(UnmanagedType.Bool)] out bool value); void SetItalic([MarshalAs(UnmanagedType.Bool)] bool value);
            void GetUnderline([MarshalAs(UnmanagedType.Bool)] out bool value); void SetUnderline([MarshalAs(UnmanagedType.Bool)] bool value);
            void GetStrikethrough([MarshalAs(UnmanagedType.Bool)] out bool value); void SetStrikethrough([MarshalAs(UnmanagedType.Bool)] bool value);
            void GetWeight(out short value); void SetWeight(short value); void GetCharset(out short value); void SetCharset(short value);
            void GetHFont(out IntPtr value); void Clone([MarshalAs(UnmanagedType.Interface)] out NestedNativeFont value);
        }

        /// <summary>IFontDisp QI marker only; no metric or property operation is dispatched.</summary>
        [ComImport, Guid("BEF6E003-A874-101A-8BBA-00AA00300CAB"), InterfaceType(ComInterfaceType.InterfaceIsIDispatch)]
        private interface NestedNativeFontDisp { }
    }
}

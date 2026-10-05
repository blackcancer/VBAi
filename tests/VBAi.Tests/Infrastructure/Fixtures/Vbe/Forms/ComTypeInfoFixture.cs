namespace VBAi.Tests.Unit
{
    using System;
    using System.Collections.Generic;
    using System.Runtime.InteropServices;
    using System.Runtime.InteropServices.ComTypes;
    using VBAi;
    using TYPEATTR = System.Runtime.InteropServices.ComTypes.TYPEATTR;
    using FUNCDESC = System.Runtime.InteropServices.ComTypes.FUNCDESC;
    using TYPEKIND = System.Runtime.InteropServices.ComTypes.TYPEKIND;
    using INVOKEKIND = System.Runtime.InteropServices.ComTypes.INVOKEKIND;
    using IMPLTYPEFLAGS = System.Runtime.InteropServices.ComTypes.IMPLTYPEFLAGS;
    using DISPPARAMS = System.Runtime.InteropServices.ComTypes.DISPPARAMS;

    internal sealed class ComClassInfoFixture : IProvideClassInfo
    {
        internal ITypeInfo Info;
        internal bool Fail;
        public void GetClassInfo(out ITypeInfo typeInfo)
        {
            typeInfo = Info;
            if (Fail) throw new COMException("Disposable class metadata failure");
        }
    }

    internal sealed class ComTypeInfoFixture : ITypeInfo, IDisposable
    {
        internal sealed class Function
        {
            internal string Name = "Caption";
            internal INVOKEKIND Kind = INVOKEKIND.INVOKE_FUNC;
            internal short Parameters;
            internal bool DescriptorFailure, NamesFailure;
            internal int NameCount = 1;
            internal int? MemberId;
        }
        internal sealed class Inherited
        {
            internal ITypeInfo Info;
            internal IMPLTYPEFLAGS Flags;
            internal bool FlagsFailure, ReferenceFailure, InfoFailure;
        }
        internal readonly List<Function> Functions = new List<Function>();
        internal readonly List<Inherited> Parents = new List<Inherited>();
        internal TYPEKIND Kind = TYPEKIND.TKIND_DISPATCH;
        internal Guid Identity = Guid.NewGuid();
        internal short? FunctionCount, ParentCount;
        internal bool AttributeFailure, AttributeFailureAfterAllocation, DocumentationFailure;
        internal int AttributeReleases, FunctionReleases;
        private readonly HashSet<IntPtr> allocations = new HashSet<IntPtr>();
        public void GetTypeAttr(out IntPtr pointer)
        {
            pointer = IntPtr.Zero;
            if (AttributeFailure) throw new COMException("Disposable attribute failure");
            pointer = Allocate(new TYPEATTR { guid = Identity, typekind = Kind,
                cFuncs = FunctionCount ?? (short)Functions.Count, cImplTypes = ParentCount ?? (short)Parents.Count });
            if (AttributeFailureAfterAllocation) throw new COMException("Disposable attribute failure after allocation");
        }
        public void GetFuncDesc(int index, out IntPtr pointer)
        {
            pointer = IntPtr.Zero;
            var function = Functions[index];
            if (function.DescriptorFailure) throw new COMException("Disposable function failure");
            pointer = Allocate(new FUNCDESC { memid = function.MemberId ?? index + 1, invkind = function.Kind, cParams = function.Parameters });
        }
        public void GetNames(int member, string[] names, int maxNames, out int count)
        {
            int mapped = Functions.FindIndex(f => f.MemberId == member);
            var function = Functions[mapped < 0 ? member - 1 : mapped];
            if (function.NamesFailure) throw new COMException("Disposable names failure");
            count = function.NameCount;
            if (maxNames > 0) names[0] = function.Name;
        }
        public void GetDocumentation(int member, out string name, out string description, out int helpContext, out string helpFile)
        {
            name = "Disposable interface"; description = "Metadata fixture"; helpContext = 0; helpFile = null;
            if (DocumentationFailure) throw new COMException("Disposable documentation failure");
        }
        public void GetImplTypeFlags(int index, out IMPLTYPEFLAGS flags)
        {
            flags = Parents[index].Flags;
            if (Parents[index].FlagsFailure) throw new COMException("Disposable flags failure");
        }
        public void GetRefTypeOfImplType(int index, out int reference)
        {
            reference = index;
            if (Parents[index].ReferenceFailure) throw new COMException("Disposable reference failure");
        }
        public void GetRefTypeInfo(int reference, out ITypeInfo info)
        {
            info = Parents[reference].Info;
            if (Parents[reference].InfoFailure) throw new COMException("Disposable inherited metadata failure");
        }
        public void ReleaseTypeAttr(IntPtr pointer) { Release(pointer); AttributeReleases++; }
        public void ReleaseFuncDesc(IntPtr pointer) { Release(pointer); FunctionReleases++; }
        public void ReleaseVarDesc(IntPtr pointer) { throw new NotSupportedException(); }
        private IntPtr Allocate<T>(T value)
        {
            var pointer = Marshal.AllocHGlobal(Marshal.SizeOf(typeof(T)));
            Marshal.StructureToPtr(value, pointer, false); allocations.Add(pointer); return pointer;
        }
        private void Release(IntPtr pointer)
        {
            if (!allocations.Remove(pointer)) throw new InvalidOperationException("Descriptor released twice or not owned.");
            Marshal.FreeHGlobal(pointer);
        }
        public void Dispose()
        {
            int count = allocations.Count;
            foreach (var pointer in allocations) Marshal.FreeHGlobal(pointer);
            allocations.Clear();
            if (count != 0) throw new InvalidOperationException("Metadata descriptors were not released: " + count);
        }
        public void GetTypeComp(out ITypeComp value) { throw new NotSupportedException(); }
        public void GetVarDesc(int index, out IntPtr value) { throw new NotSupportedException(); }
        public void GetIDsOfNames(string[] names, int count, int[] ids) { throw new NotSupportedException(); }
        public void Invoke(object instance, int member, short flags, ref DISPPARAMS arguments, IntPtr result, IntPtr exception, out int argumentError) { throw new NotSupportedException(); }
        public void GetDllEntry(int member, INVOKEKIND kind, IntPtr dll, IntPtr entry, IntPtr ordinal) { throw new NotSupportedException(); }
        public void AddressOfMember(int member, INVOKEKIND kind, out IntPtr address) { throw new NotSupportedException(); }
        public void CreateInstance(object outer, ref Guid iid, out object value) { throw new NotSupportedException(); }
        public void GetMops(int member, out string value) { throw new NotSupportedException(); }
        public void GetContainingTypeLib(out ITypeLib library, out int index) { throw new NotSupportedException(); }
    }
}

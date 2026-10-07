using System;
using System.Collections.Generic;
using System.Runtime.InteropServices;
using System.Runtime.InteropServices.ComTypes;
using System.Runtime.Remoting.Messaging;
using System.Runtime.Remoting.Proxies;
using ELEMDESC = System.Runtime.InteropServices.ComTypes.ELEMDESC;
using FUNCDESC = System.Runtime.InteropServices.ComTypes.FUNCDESC;
using INVOKEKIND = System.Runtime.InteropServices.ComTypes.INVOKEKIND;
using PARAMDESC = System.Runtime.InteropServices.ComTypes.PARAMDESC;
using PARAMFLAG = System.Runtime.InteropServices.ComTypes.PARAMFLAG;
using TYPEATTR = System.Runtime.InteropServices.ComTypes.TYPEATTR;
using TYPEDESC = System.Runtime.InteropServices.ComTypes.TYPEDESC;
using TYPEKIND = System.Runtime.InteropServices.ComTypes.TYPEKIND;
using VARDESC = System.Runtime.InteropServices.ComTypes.VARDESC;

namespace VBAi.Tests.Infrastructure
{
    internal sealed class OwnedReferenceMetadata : RealProxy, IDisposable
    {
        internal readonly ITypeInfo Info;
        internal TYPEKIND Kind = TYPEKIND.TKIND_DISPATCH;
        internal INVOKEKIND Invocation = INVOKEKIND.INVOKE_FUNC;
        internal short FunctionFlags;
        internal int MemberId = 1, NameCount = 2;
        internal VarEnum Return = VarEnum.VT_VOID;
        internal ELEMDESC[] Parameters = new ELEMDESC[0];
        internal VARDESC[] Variables = new VARDESC[0];
        internal bool Function = true, EmptyVariableName, FailDocumentation;
        internal int FunctionReleases, VariableReleases, TypeReleases;
        private readonly HashSet<IntPtr> allocations = new HashSet<IntPtr>();
        internal OwnedReferenceMetadata() : base(typeof(ITypeInfo)) { Info = (ITypeInfo)GetTransparentProxy(); }
        internal static ELEMDESC Parameter(VarEnum type, PARAMFLAG flags) => new ELEMDESC
        { tdesc = new TYPEDESC { vt = (short)type }, desc = new ELEMDESC.DESCUNION { paramdesc = new PARAMDESC { wParamFlags = flags } } };
        private IntPtr Allocate<T>(T value)
        { var p = Marshal.AllocHGlobal(Marshal.SizeOf(value)); Marshal.StructureToPtr(value, p, false); allocations.Add(p); return p; }
        private void Release(IntPtr p) { if (allocations.Remove(p)) Marshal.FreeHGlobal(p); }
        public override IMessage Invoke(IMessage message)
        {
            var call = (IMethodCallMessage)message; var args = call.Args;
            try
            {
                switch (call.MethodName)
                {
                    case "GetTypeAttr": args[0] = Allocate(new TYPEATTR { guid = Guid.NewGuid(), typekind = Kind, cFuncs = (short)(Function ? 1 : 0), cVars = (short)Variables.Length }); break;
                    case "GetFuncDesc":
                        IntPtr parameters = IntPtr.Zero;
                        if (Parameters.Length != 0)
                        {
                            int size = Marshal.SizeOf(typeof(ELEMDESC)); parameters = Marshal.AllocHGlobal(size * Parameters.Length); allocations.Add(parameters);
                            for (int i = 0; i < Parameters.Length; i++) Marshal.StructureToPtr(Parameters[i], IntPtr.Add(parameters, size * i), false);
                        }
                        args[1] = Allocate(new FUNCDESC
                        {
                            memid = MemberId,
                            invkind = Invocation,
                            wFuncFlags = FunctionFlags,
                            cParams = (short)Parameters.Length,
                            lprgelemdescParam = parameters,
                            elemdescFunc = Parameter(Return, 0)
                        }); break;
                    case "GetNames":
                        var names = (string[])args[1]; int found = Math.Min(NameCount, names.Length);
                        for (int i = 0; i < found; i++) names[i] = i == 0 ? "Run" : "value" + i;
                        args[3] = found; break;
                    case "GetVarDesc": args[1] = Allocate(Variables[(int)args[0]]); break;
                    case "GetDocumentation":
                        if (FailDocumentation) throw new COMException("Owned documentation failure.");
                        args[1] = EmptyVariableName && (int)args[0] >= 100 ? "" : "member" + args[0];
                        args[2] = "Native member documentation"; args[3] = 42; args[4] = "Owned.chm"; break;
                    case "ReleaseTypeAttr": Release((IntPtr)args[0]); TypeReleases++; break;
                    case "ReleaseFuncDesc":
                        var descriptor = (FUNCDESC)Marshal.PtrToStructure((IntPtr)args[0], typeof(FUNCDESC));
                        Release(descriptor.lprgelemdescParam); Release((IntPtr)args[0]); FunctionReleases++; break;
                    case "ReleaseVarDesc": Release((IntPtr)args[0]); VariableReleases++; break;
                    default: throw new InvalidOperationException("Unexpected type operation " + call.MethodName);
                }
                return new ReturnMessage(null, args, args.Length, call.LogicalCallContext, call);
            }
            catch (Exception error) { return new ReturnMessage(error, call); }
        }
        public void Dispose() { foreach (var p in allocations) Marshal.FreeHGlobal(p); allocations.Clear(); }
    }
}

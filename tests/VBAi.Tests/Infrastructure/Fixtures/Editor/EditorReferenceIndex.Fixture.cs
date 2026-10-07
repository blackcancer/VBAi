using Microsoft.VisualStudio.TestTools.UnitTesting;
using System;
using System.Runtime.InteropServices;
using System.Runtime.InteropServices.ComTypes;
using System.Runtime.Remoting.Messaging;
using System.Runtime.Remoting.Proxies;

namespace VBAi.Tests.Infrastructure
{
    internal sealed class OwnedNamelessTypeInfo : RealProxy
    {
        internal int NamesRead, FunctionsReleased, TypesReleased;
        internal ITypeInfo Info => (ITypeInfo)GetTransparentProxy();
        internal OwnedNamelessTypeInfo() : base(typeof(ITypeInfo)) { }
        public override IMessage Invoke(IMessage message)
        {
            var call = (IMethodCallMessage)message; var args = (object[])call.Args.Clone();
            switch (call.MethodName)
            {
                case "GetTypeAttr": args[0] = Allocate(new System.Runtime.InteropServices.ComTypes.TYPEATTR { guid = Guid.NewGuid(), cFuncs = 1 }); break;
                case "GetFuncDesc": Assert.AreEqual(0, args[0]); args[1] = Allocate(new System.Runtime.InteropServices.ComTypes.FUNCDESC { invkind = System.Runtime.InteropServices.ComTypes.INVOKEKIND.INVOKE_FUNC }); break;
                case "GetNames": NamesRead++; args[3] = 0; break;
                case "ReleaseTypeAttr": Marshal.FreeHGlobal((IntPtr)args[0]); TypesReleased++; break;
                case "ReleaseFuncDesc": Marshal.FreeHGlobal((IntPtr)args[0]); FunctionsReleased++; break;
                default: throw new AssertFailedException("Unexpected nameless type metadata operation: " + call.MethodName);
            }
            return new ReturnMessage(null, args, args.Length, call.LogicalCallContext, call);
        }
        private static IntPtr Allocate<T>(T value)
        { IntPtr pointer = Marshal.AllocHGlobal(Marshal.SizeOf(value)); Marshal.StructureToPtr(value, pointer, false); return pointer; }
    }
}
namespace CodexVBE.Tests.Unit
{
    using System;
    using System.Reflection;
    using System.Runtime.ExceptionServices;
    using CodexVBE;

    public sealed partial class ChatSessionStoreTests
    {
        private static object Call(ChatSessionStore store,string name,params object[] arguments)
        {
            try {return typeof(ChatSessionStore).GetMethod(name,BindingFlags.Instance|BindingFlags.NonPublic).Invoke(store,arguments);}
            catch(TargetInvocationException error) {ExceptionDispatchInfo.Capture(error.InnerException).Throw();throw;}
        }
        private static void Sql(ChatSessionStore store,string sql,params string[] values)
        {Call(store,"Execute",sql,values);}
    }
}

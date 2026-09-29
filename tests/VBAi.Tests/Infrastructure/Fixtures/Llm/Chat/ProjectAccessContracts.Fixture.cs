using System;
using System.ComponentModel;
using System.Reflection;
using Microsoft.VisualStudio.TestTools.UnitTesting;
namespace VBAi.Tests.Unit
{
    internal static class AccessContract
    {
        internal static object Call(object target,string method,params object[] arguments)
        {
            var type=target as Type ?? target.GetType();
            return type.GetMethod(method,BindingFlags.NonPublic|BindingFlags.Public|BindingFlags.Instance|BindingFlags.Static).Invoke(target is Type?null:target,arguments);
        }
        internal static void Set(object target,string field,object value)
        {
            target.GetType().GetField(field,BindingFlags.NonPublic|BindingFlags.Instance).SetValue(target,value);
        }
        internal static void Denied(object target,string method,params object[] arguments)
        {
            var error=Assert.ThrowsException<TargetInvocationException>(()=>Call(target,method,arguments));
            Assert.IsInstanceOfType(error.InnerException,typeof(InvalidOperationException));
        }
    }
    internal sealed class AccessDesignContext : LicenseContext
    {
        public override LicenseUsageMode UsageMode => LicenseUsageMode.Designtime;
    }
}

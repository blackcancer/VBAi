namespace VBAi.Tests.Unit
{
    using Microsoft.VisualStudio.TestTools.UnitTesting;
    using System.Reflection;
    using VBAi;

    public sealed partial class VbeDebugWindowsBoundaryTests
    {
        private static object CallPrivate(string name, params object[] arguments)
        {
            var method = typeof(VbeDebugWindows).GetMethod(name, BindingFlags.NonPublic | BindingFlags.Static);
            Assert.IsNotNull(method, "Missing private helper: " + name);
            return method.Invoke(null, arguments);
        }
    }
}

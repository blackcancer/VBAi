namespace VBAi.Tests.Unit
{
    using System;
    using System.Collections;
    using System.Collections.Generic;
    using System.Linq;
    using System.Reflection;
    using VBAi;
    using Microsoft.VisualStudio.TestTools.UnitTesting;

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

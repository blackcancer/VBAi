namespace CodexVBE.Tests.Unit
{
    using System;
    using System.Collections.Generic;
    using System.Reflection;
    using System.Runtime.InteropServices;
    using System.Runtime.Serialization;
    using System.Windows.Forms;
    using CodexVBE;
    using Microsoft.VisualStudio.TestTools.UnitTesting;

    public sealed partial class HostSettingsWindowTests
    {
        private const BindingFlags Private = BindingFlags.NonPublic | BindingFlags.Instance;
        private static T Field<T>(object owner, string name)
        {
            var field = owner.GetType().GetField(name, Private);
            Assert.IsNotNull(field, name);
            return (T)field.GetValue(owner);
        }

        public sealed class FakeMenus
        {
            public List<FakeBar> CommandBars { get; } = new List<FakeBar>();
        }

        public sealed class FakeBar
        {
            public int Type { get; set; }
            public List<FakeMenu> Controls { get; set; }
        }

        public sealed class FakeMenu
        {
            public string Caption { get; set; }
        }
    }
}

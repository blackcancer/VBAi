namespace CodexVBE.Tests.Unit
{
    using System;
    using System.Reflection;
    using System.Windows.Forms;
    using CodexVBE;
    using Microsoft.VisualStudio.TestTools.UnitTesting;

    public sealed partial class HostSettingsCoverageTests
    {
        private const BindingFlags PrivateInstance = BindingFlags.Instance | BindingFlags.NonPublic;
        private static T Field<T>(object instance, string name)
        {
            var field = instance.GetType().GetField(name, PrivateInstance);
            Assert.IsNotNull(field, name);
            return (T)field.GetValue(instance);
        }

        private static void Set(object instance, string name, object value)
        {
            var field = instance.GetType().GetField(name, PrivateInstance);
            Assert.IsNotNull(field, name);
            field.SetValue(instance, value);
        }

        private static object Call(object instance, string name, params object[] arguments)
        {
            var method = instance.GetType().GetMethod(name, PrivateInstance);
            Assert.IsNotNull(method, name);
            return method.Invoke(instance, arguments);
        }

        public sealed class FakeHost
        {
            public object[] CommandBars { get; set; }
        }

        public sealed class FakeBar
        {
            public int Type { get; set; }
            public object[] Controls { get; set; }
        }

        public sealed class InvalidBar
        {
            public int Type
            {
                get
                {
                    throw new InvalidOperationException("no type");
                }
            }
        }

        public sealed class FakeControl
        {
            public string Caption { get; set; }
        }

        public sealed class InvalidControl
        {
            public string Caption
            {
                get
                {
                    throw new InvalidOperationException("no caption");
                }
            }
        }

        public sealed class FakeNativeWindow
        {
            public int CloseCount { get; private set; }

            public void Close()
            {
                CloseCount++;
            }
        }

        public sealed class FakeOwnerHost
        {
            public FakeMainWindow MainWindow { get; set; }
        }

        public sealed class FakeMainWindow
        {
            public long HWnd { get; set; }
        }

        public sealed class RejectingNativeWindow
        {
            public void Close()
            {
                throw new InvalidOperationException("VBE closed first");
            }
        }
    }
}

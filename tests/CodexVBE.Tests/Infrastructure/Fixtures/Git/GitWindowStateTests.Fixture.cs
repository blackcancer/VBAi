namespace CodexVBE.Tests.Unit
{
    using System;
    using System.Drawing;
    using System.Reflection;
    using System.Runtime.Serialization;
    using System.Threading.Tasks;
    using System.Windows.Forms;
    using CodexVBE;
    using Microsoft.VisualStudio.TestTools.UnitTesting;

    public sealed partial class GitWindowStateTests
    {
        private const BindingFlags InstancePrivate = BindingFlags.Instance | BindingFlags.NonPublic;
        private static T Field<T>(GitWindow window, string name)
        {
            var field = typeof(GitWindow).GetField(name, InstancePrivate);
            Assert.IsNotNull(field, name);
            return (T)field.GetValue(window);
        }

        private static void Set(GitWindow window, string name, object value)
        {
            var field = typeof(GitWindow).GetField(name, InstancePrivate);
            Assert.IsNotNull(field, name);
            field.SetValue(window, value);
        }

        private static object Invoke(GitWindow window, string name, params object[] args)
        {
            var method = typeof(GitWindow).GetMethod(name, InstancePrivate);
            Assert.IsNotNull(method, name);
            if (name == "Perform" && args.Length == 1)
                args = new[]
                {
                    args[0],
                    (object)false
                };
            return method.Invoke(window, args);
        }

        private static DataGridView DiffGrid(CodeDiffView view)
        {
            var field = typeof(CodeDiffView).GetField("grid", InstancePrivate);
            Assert.IsNotNull(field);
            return (DataGridView)field.GetValue(view);
        }
    }
}

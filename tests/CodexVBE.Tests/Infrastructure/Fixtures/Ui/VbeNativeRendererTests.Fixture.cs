using System;
using System.Collections.Generic;
using System.Reflection;
using CodexVBE;

namespace CodexVBE.Tests.Unit
{
    /// <summary>Owns and restores a renderer snapshot; no native library or window is loaded.</summary>
    internal sealed class NativeRendererStopFixture : IDisposable
    {
        private static readonly BindingFlags Fields = BindingFlags.Static | BindingFlags.NonPublic;
        private readonly Dictionary<FieldInfo, object> previous = new Dictionary<FieldInfo, object>();
        private readonly Action<string, string> previousLog = LoadLog.AppendText;
        internal readonly List<string> Messages = new List<string>();
        internal uint Result;
        internal int Calls;
        internal Exception Failure;
        internal readonly IntPtr PinnedModule = new IntPtr(0x1234);
        internal readonly Delegate StopExport;

        internal NativeRendererStopFixture(bool active = true, bool available = true)
        {
            foreach (string name in new[] { "module", "active", "stop", "query" })
            {
                var field = typeof(VbeNativeRenderer).GetField(name, Fields);
                previous.Add(field, field.GetValue(null));
            }
            var stopField = typeof(VbeNativeRenderer).GetField("stop", Fields);
            StopExport = Delegate.CreateDelegate(stopField.FieldType, this,
                typeof(NativeRendererStopFixture).GetMethod(nameof(NativeStop), BindingFlags.Instance | BindingFlags.NonPublic));
            Set("module", PinnedModule); Set("active", active); Set("stop", available ? StopExport : null); Set("query", null);
            LoadLog.AppendText = (path, message) => Messages.Add(message);
        }

        private uint NativeStop()
        {
            Calls++;
            if (Failure != null) throw Failure;
            return Result;
        }
        private static void Set(string name, object value) => typeof(VbeNativeRenderer).GetField(name, Fields).SetValue(null, value);
        internal object Read(string name) => typeof(VbeNativeRenderer).GetField(name, Fields).GetValue(null);
        public void Dispose()
        {
            foreach (var pair in previous) pair.Key.SetValue(null, pair.Value);
            LoadLog.AppendText = previousLog;
        }
    }
}
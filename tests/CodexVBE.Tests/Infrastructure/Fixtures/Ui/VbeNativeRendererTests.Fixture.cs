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
            foreach (string name in new[] { "module", "active", "stop", "query", "start", "register", "refresh" })
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

        internal uint StartResult, RegisterResult, RefreshResult, QueryResult;
        internal int StartCalls, RegisterCalls, RefreshCalls, QueryCalls;
        internal IntPtr StartWindow, RegisteredWindow;
        internal Exception StartFailure;
        internal uint QuerySize;

        internal void ConfigureLifecycle(bool queryAvailable = true)
        {
            foreach (var pair in new[] { Tuple.Create("start", "NativeStart"), Tuple.Create("register", "NativeRegister"),
                Tuple.Create("refresh", "NativeRefresh"), Tuple.Create("query", "NativeQuery") })
            {
                var field = typeof(VbeNativeRenderer).GetField(pair.Item1, Fields);
                var method = GetType().GetMethod(pair.Item2, BindingFlags.Instance | BindingFlags.NonPublic);
                field.SetValue(null, pair.Item1 == "query" && !queryAvailable ? null : Delegate.CreateDelegate(field.FieldType, this, method));
            }
        }
        internal void SetActive(bool value) => Set("active", value);
        private uint NativeStart(IntPtr window)
        { StartCalls++; StartWindow = window; if (StartFailure != null) throw StartFailure; return StartResult; }
        private uint NativeRegister(IntPtr window) { RegisterCalls++; RegisteredWindow = window; return RegisterResult; }
        private uint NativeRefresh() { RefreshCalls++; return RefreshResult; }
        private uint NativeQuery(ref VbeNativeRenderer.RendererStatus status)
        {
            QueryCalls++; QuerySize = status.Size;
            status.Active = 1; status.Windows = 2; status.Imports = 3; status.RestoredImports = 4;
            status.Patterns = 5; status.Images = 6; status.Text = 7; status.Fills = 8; status.Unsupported = 9; status.Failures = 10;
            return QueryResult;
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
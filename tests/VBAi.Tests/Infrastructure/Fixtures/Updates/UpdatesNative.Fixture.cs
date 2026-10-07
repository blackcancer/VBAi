using Microsoft.VisualStudio.TestTools.UnitTesting;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Net;
using System.Net.Http;
using System.Reflection;
using System.Runtime.ExceptionServices;
using System.Threading;
using System.Threading.Tasks;
using VBAi.Tests.Infrastructure;

namespace VBAi.Tests.Unit
{
    internal sealed class UpdateOwnedRegistry : IDisposable
    {
        internal object Value;
        internal bool Disposed;
        public void Dispose() => Disposed = true;
    }
    internal sealed class UpdateRuntimeHttp : HttpMessageHandler
    {
        internal HttpStatusCode Status = HttpStatusCode.OK;
        internal string FinalUrl = WebViewRuntimePrerequisite.Bootstrapper;
        internal byte[] Bytes = new byte[] { 1, 2, 3 };
        internal int Calls;
        protected override Task<HttpResponseMessage> SendAsync(HttpRequestMessage request, CancellationToken token)
        {
            Calls++; Assert.AreEqual(WebViewRuntimePrerequisite.Bootstrapper, request.RequestUri.AbsoluteUri);
            return Task.FromResult(new HttpResponseMessage(Status) { RequestMessage = new HttpRequestMessage(HttpMethod.Get, FinalUrl), Content = new ByteArrayContent(Bytes) });
        }
    }
    internal sealed class UpdatesNativeFixture : IDisposable
    {
        internal readonly UpdateScope Scope = new UpdateScope();
        private readonly Dictionary<FieldInfo, object> saved = new Dictionary<FieldInfo, object>();
        private readonly Action<string, string> log = LoadLog.AppendText;
        internal readonly List<string> Logs = new List<string>();
        internal readonly List<ProcessStartInfo> Starts = new List<ProcessStartInfo>();
        internal readonly List<TimerCallback> Callbacks = new List<TimerCallback>();
        internal UpdatesNativeFixture()
        {
            foreach (var type in new[] { typeof(UpdateCoordinator), typeof(UpdateInstallerRunner), typeof(WebViewRuntimePrerequisite) })
                foreach (var field in type.GetFields(BindingFlags.Static | BindingFlags.NonPublic))
                    if (!field.IsLiteral && !field.IsInitOnly) saved.Add(field, field.GetValue(null));
            Set("timer", null); Set("lifetime", null);
            var create = UpdateCoordinator.CreateTimer;
            UpdateCoordinator.CreateTimer = (callback, state, due, period) =>
            {
                Assert.AreEqual(TimeSpan.Zero, due); Assert.AreEqual(TimeSpan.FromHours(1), period); Callbacks.Add(callback);
                return create(callback, state, Timeout.InfiniteTimeSpan, period);
            };
            UpdateCoordinator.StartProcess = start => { Starts.Add(start); return null; };
            LoadLog.AppendText = (path, text) => Logs.Add(text);
        }
        internal static object Read(string name) => typeof(UpdateCoordinator).GetField(name, BindingFlags.Static | BindingFlags.NonPublic).GetValue(null);
        internal static void Set(string name, object value) => typeof(UpdateCoordinator).GetField(name, BindingFlags.Static | BindingFlags.NonPublic).SetValue(null, value);
        internal static T Invoke<T>(Type type, string method, params object[] args)
        {
            try { return (T)type.GetMethod(method, BindingFlags.Static | BindingFlags.NonPublic).Invoke(null, args); }
            catch (TargetInvocationException error) { ExceptionDispatchInfo.Capture(error.InnerException).Throw(); throw; }
        }
        internal static Process ExitHelper(int code) => Process.Start(new ProcessStartInfo(Path.Combine(Environment.SystemDirectory, "cmd.exe"), "/d /c exit " + code)
        { UseShellExecute = false, CreateNoWindow = true, WindowStyle = ProcessWindowStyle.Hidden });
        public void Dispose()
        {
            UpdateCoordinator.Stop(); foreach (var field in saved) field.Key.SetValue(null, field.Value);
            LoadLog.AppendText = log; Scope.Dispose();
        }
    }
}

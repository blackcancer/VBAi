using System;
using System.Collections.Generic;
using System.Reflection;
using System.Runtime.ExceptionServices;
using System.Runtime.InteropServices;
using System.Threading;
using System.Threading.Tasks;
using CodexVBE;
using CodexVBE.Tests.Infrastructure;
using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace CodexVBE.Tests.Unit
{
    internal sealed class CrashOutlookProfile : IDisposable
    {
        internal int Count;
        internal bool Disposed;
        public void Dispose() => Disposed = true;
    }
    public sealed class CrashOutlookAccounts { public int Count { get; set; } = 1; }
    public sealed class CrashOutlookSession { public CrashOutlookAccounts Accounts { get; set; } = new CrashOutlookAccounts(); }
    public sealed class CrashOutlookMail
    {
        public string To { get; set; }
        public string Subject { get; set; }
        public string Body { get; set; }
        public Action Deliver { get; set; }
        public void Send() => Deliver();
    }
    public sealed class CrashOutlookApplication
    {
        public CrashOutlookSession Current { get; set; } = new CrashOutlookSession();
        public bool FailSession { get; set; }
        public CrashOutlookSession Session => FailSession ? throw new InvalidOperationException("owned Session failure") : Current;
        public CrashOutlookMail Mail { get; set; } = new CrashOutlookMail();
        public object CreateItem(int kind) { Assert.AreEqual(0, kind); return Mail; }
    }
    internal sealed class CrashReportDeliveryFixture : IDisposable
    {
        private readonly Dictionary<FieldInfo, object> saved = new Dictionary<FieldInfo, object>();
        private readonly Action<string, string> log = LoadLog.AppendText;
        internal readonly CrashOutlookApplication Application = new CrashOutlookApplication();
        internal readonly List<object> Released = new List<object>();
        internal readonly List<string> Profiles = new List<string>(), Logs = new List<string>();
        internal int Sends;
        internal CrashReportDeliveryFixture()
        {
            foreach (var field in typeof(CrashReportDelivery).GetFields(BindingFlags.Static | BindingFlags.NonPublic))
                if (!field.IsLiteral && !field.IsInitOnly) saved.Add(field, field.GetValue(null));
            CrashReportDelivery.LoadSettings = () => new LlmSettings { GitHubAccount = "owned-account" };
            CrashReportDelivery.ReadCredential = (account, token) => { Assert.AreEqual("owned-account", account); return Task.FromResult("owned-token"); };
            CrashReportDelivery.ProfileCount = version => { Profiles.Add(version); return null; };
            CrashReportDelivery.ActiveOutlook = progId => { Assert.AreEqual("Outlook.Application", progId); return Application; };
            CrashReportDelivery.OutlookType = progId => typeof(CrashOutlookApplication);
            CrashReportDelivery.CreateOutlook = type => { Assert.AreEqual(typeof(CrashOutlookApplication), type); return Application; };
            CrashReportDelivery.IsComReference = reference => true;
            CrashReportDelivery.ReleaseReference = reference => { Released.Add(reference); return 0; };
            Application.Mail.Deliver = () => Sends++;
            LoadLog.AppendText = (path, line) => Logs.Add(line);
        }
        internal LlmHttpFixture ConfigureHttp(string body)
        {
            var http = new LlmHttpFixture(body);
            CrashReportDelivery.CreateApi = (account, credential) => { Assert.AreEqual("owned-account", account); return new GitHubApi(account, http, credential); };
            return http;
        }
        internal static T Invoke<T>(string method, params object[] args)
        {
            try { return (T)typeof(CrashReportDelivery).GetMethod(method, BindingFlags.Static | BindingFlags.NonPublic).Invoke(null, args); }
            catch (TargetInvocationException error) { ExceptionDispatchInfo.Capture(error.InnerException).Throw(); throw; }
        }
        internal void UseNativeProfileReader()
        {
            foreach (var field in saved) if (field.Key.Name == "ProfileCount") CrashReportDelivery.ProfileCount = (Func<string, int?>)field.Value;
        }
        internal bool Send() => Invoke<bool>("SendOutlookNative", "owned title", "owned body");
        public void Dispose() { foreach (var field in saved) field.Key.SetValue(null, field.Value); LoadLog.AppendText = log; }
    }
}

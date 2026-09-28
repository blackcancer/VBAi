using System;
using System.Diagnostics;
using System.Threading;
using System.Windows.Forms;
using CodexVBE;
using CodexVBE.Tests.Infrastructure;
using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace CodexVBE.Tests.Unit
{
    internal static class UpdateUiPump
    {
        internal static void Until(Func<bool> done)
        {
            var watch = Stopwatch.StartNew();
            while (!done() && watch.ElapsedMilliseconds < 5000) { Application.DoEvents(); Thread.Sleep(1); }
            Assert.IsTrue(done(), "The owned update UI operation must settle.");
        }
        internal static void Call(object window, string method) => LlmBoundaryScope.Call(window, method, null, EventArgs.Empty);
    }
    internal sealed class UpdateProgressFixture : IDisposable
    {
        internal readonly UpdateScope Scope = new UpdateScope();
        internal readonly UpdateProgressWindow Window = new UpdateProgressWindow();
        internal readonly UpdateInstallJob Job;
        internal readonly UpdateInstallerRunner Runner;
        private readonly string culture = UpdateText.Culture;
        internal UpdateProgressFixture(bool background = false, string language = "en-US")
        {
            Job = UpdateInstallerRunnerTests.Job(Scope); Job.Culture = language;
            Window.Configure(Scope.Root, Job, background);
            UiInvoke.Field<System.Windows.Forms.Timer>(Window, "timer").Stop();
            Runner = UiInvoke.Field<UpdateInstallerRunner>(Window, "runner");
            Runner.VerifySignature = path => true;
            Runner.Install = path => 0;
            Runner.InstalledVersion = path => Job.TargetVersion;
            IntPtr handle = Window.Handle;
        }
        internal void Poll()
        {
            UpdateUiPump.Call(Window, "Poll");
            UpdateUiPump.Until(() => Window.IsDisposed || !UiInvoke.Field<bool>(Window, "polling"));
        }
        public void Dispose() { Window.Dispose(); UpdateText.Culture = culture; Scope.Dispose(); }
    }
    internal sealed class UpdateWindowFixture : IDisposable
    {
        internal readonly UpdateScope Scope = new UpdateScope();
        private readonly ThemeScope theme = new ThemeScope();
        internal readonly UpdateWindow Window = new UpdateWindow();
        internal readonly UpdatePreferences Preferences = new UpdatePreferences();
        internal bool Managed;
        internal int Stored;
        internal UpdateWindowFixture()
        {
            Window.ReadPreferences = () => Preferences;
            Window.StorePreferences = value => { Assert.AreSame(Preferences, value); Stored++; };
            Window.ManagedInstallation = () => Managed;
            Window.CreateFeed = () => new UpdateFeed(new UpdateFeedTests.Handler("[]"), token => throw new AssertFailedException("No real credential access"));
            IntPtr handle = Window.Handle;
        }
        internal UpdateRelease Release(bool installer = true, string body = "Owned release notes") => new UpdateRelease
        { tag_name = "999.0.0", body = body, assets = installer ? new[] { UpdateFeedTests.Asset("fixture installer") } : null };
        internal void Check(UpdateRelease release)
        {
            Window.CreateFeed = () => new UpdateFeed(new UpdateFeedTests.Handler(new System.Web.Script.Serialization.JavaScriptSerializer().Serialize(release == null ? new UpdateRelease[0] : new[] { release })), token => throw new AssertFailedException("No real credentials"));
            UpdateUiPump.Call(Window, "Check_Click"); UpdateUiPump.Until(() => !UiInvoke.Field<bool>(Window, "busy"));
        }
        internal string Status => UiInvoke.Field<Label>(Window, "status").Text;
        internal void Download()
        {
            Window.CreateFeed = () => new UpdateFeed(new UpdateFeedTests.Handler("fixture installer"), token => throw new AssertFailedException("No real credentials"));
            UpdateUiPump.Call(Window, "Download_Click"); UpdateUiPump.Until(() => !UiInvoke.Field<bool>(Window, "busy"));
        }
        public void Dispose() { Window.Dispose(); theme.Dispose(); Scope.Dispose(); }
    }}

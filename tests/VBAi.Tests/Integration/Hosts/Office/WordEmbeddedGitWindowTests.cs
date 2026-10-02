using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Runtime.ExceptionServices;
using System.Security.Cryptography;
using System.Text;
using System.Threading;
using System.Web.Script.Serialization;
using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace VBAi.Tests.Integration
{
    /// <summary>Qualifies Word's production VBE-thread capture through its actual installed GitWindow.</summary>
    [TestClass, TestCategory("Office"), TestCategory("NativeEmbeddedGitUi"), DoNotParallelize]
    public sealed class WordEmbeddedGitWindowTests
    {
        private static readonly List<Context> Retained = new List<Context>();

        [TestMethod]
        public void InstalledWordGitWindowCapturesSavedProjectAndCreatesExactLocalCheckpoint()
        {
            if (Environment.GetEnvironmentVariable("VBAi_RUN_WORD_EMBEDDED_GIT_TESTS") != "1" ||
                Environment.GetEnvironmentVariable("VBAi_RUN_WORD_GIT_TESTS") != "1" ||
                Environment.GetEnvironmentVariable("VBAi_RUN_OFFICE_TESTS") != "1")
                Assert.Inconclusive("Explicit owned Word, Git and embedded UI opt-ins are required.");
            string root = ExcelOwnedBootstrapPlan.RequireLocalAbsolutePath(Environment.GetEnvironmentVariable("VBAi_OFFICE_RESULTS"));
            string manifestPath = ExcelOwnedBootstrapPlan.RequireLocalAbsolutePath(Environment.GetEnvironmentVariable("VBAi_TEST_GITHUB_MANIFEST"));
            var manifest = new JavaScriptSerializer().Deserialize<Dictionary<string, object>>(File.ReadAllText(manifestPath));
            var plan = EmbeddedGitWindowTests.ValidateManifest(manifest);
            Guid expected = Guid.Parse(Environment.GetEnvironmentVariable("VBAi_TEST_EMBEDDED_GIT_MVID"));
            string hash = Environment.GetEnvironmentVariable("VBAi_TEST_EMBEDDED_GIT_SHA256");
            Assert.AreEqual(expected, typeof(VbeSession).Module.ModuleVersionId);
            Assert.AreEqual(hash, Sha(typeof(VbeSession).Assembly.Location), true);
            var context = new Context(Path.Combine(root, "word-owner-git-" + Guid.NewGuid().ToString("N")), plan);
            context.Record(new { Phase = "Preflight", ExpectedMvid = expected, ExpectedSha256 = hash,
                Scope = "Actual installed Word Git menu, owner-thread capture, local checkpoint and compare; no push/import/macro execution or chat-provider acceptance." });
            var owner = new Thread(() => Owner(context, expected, hash)) { IsBackground = true };
            owner.SetApartmentState(ApartmentState.STA); owner.Start();
            try
            {
                if (!context.Ready.Wait(TimeSpan.FromSeconds(90))) throw new TimeoutException("Word preparation has no terminal evidence.");
                if (context.OwnerError != null) ExceptionDispatchInfo.Capture(context.OwnerError).Throw();
                var ui = new Thread(() => Ui(context)) { IsBackground = true };
                ui.SetApartmentState(ApartmentState.MTA); ui.Start();
                if (!context.UiReady.Wait(TimeSpan.FromSeconds(10))) throw new TimeoutException("Word UIA was not armed.");
                context.StartMenu.Set();
                if (!context.UiDone.Wait(TimeSpan.FromSeconds(240))) throw new TimeoutException("Word UI operation is pending or uncertain.");
                if (!context.OwnerDone.Wait(TimeSpan.FromSeconds(30))) throw new TimeoutException("Word modal/cleanup remains pending.");
                var errors = new[] { context.UiError, context.OwnerError }.Where(error => error != null).ToArray();
                if (errors.Length > 1) throw new AggregateException("Word UI and owner failures.", errors);
                if (errors.Length == 1) ExceptionDispatchInfo.Capture(errors[0]).Throw();
                context.Record(new { Phase = "PASS", Scope = "WordOwnerCaptureLocalCheckpointCompareNormalExit" });
            }
            catch (Exception primary)
            {
                if (!context.MenuEmitted && !context.Retain && context.Ready.IsSet)
                { context.Stop = true; context.StartMenu.Set(); context.OwnerDone.Wait(TimeSpan.FromSeconds(30)); }
                if (!context.OwnerDone.IsSet)
                {
                    context.Stop = context.Retain = true;
                    if (context.Fixture != null) context.Fixture.NativeExecutionUnsettled = true;
                    lock (Retained) Retained.Add(context);
                }
                Exception preserved = EmbeddedGitUiProtocol.PreserveFailures(primary, context.UiError, context.OwnerError);
                context.Record(new { Phase = "FailedOrRetained", context.Retain, context.MenuEmitted, context.ModalClosed,
                    ProcessId = context.Fixture?.ProcessId, DocumentPath = context.Fixture?.DocumentPath,
                    FixtureRoot = context.Fixture?.Root,
                    OwnerTerminal = context.OwnerDone.IsSet, UiTerminal = context.UiDone.IsSet, Error = preserved.ToString(), ReplayAttempts = 0 });
                if (!ReferenceEquals(primary, preserved)) ExceptionDispatchInfo.Capture(preserved).Throw();
                throw;
            }
        }

        private static void Owner(Context context, Guid expected, string hash)
        {
            try
            {
                context.Fixture = OfficeVbeFixture.Start("Word");
                var status = context.Fixture.Data("status");
                Assert.AreEqual(expected.ToString("D"), status["AssemblyModuleVersionId"]);
                Assert.AreEqual(hash, Sha(Convert.ToString(status["AssemblyPath"])), true);
                context.Scope = context.Fixture.PrepareWordEmbeddedGit(context.Nonce, context.Record);
                context.DocumentHash = Sha(context.Scope.Path);
                context.Ready.Set();
                if (!context.StartMenu.Wait(TimeSpan.FromSeconds(30))) throw new TimeoutException("Word UI worker did not arm.");
                if (context.Stop) throw new InvalidOperationException("Stopped before Word menu invocation.");
                context.Fixture.ExecuteWordGitMenu(context.Scope, value => {
                    context.Record(value);
                    if (new JavaScriptSerializer().Serialize(value).Contains("MenuExecuteIntent"))
                    { context.MenuEmitted = true; context.MenuIntent.Set(); }
                });
                if (!context.UiDone.Wait(TimeSpan.FromSeconds(15)) || !context.ModalClosed)
                    throw new InvalidOperationException("Word modal returned without known exact closure.");
                context.Fixture.VerifyWordEmbeddedSource(context.Scope);
                Assert.AreEqual(context.DocumentHash, Sha(context.Scope.Path));
                context.Record(new { Phase = "WordSourceAndSavedBytesPreserved", context.DocumentHash });
            }
            catch (Exception error) { context.OwnerError = error; }
            finally
            {
                context.Ready.Set();
                if (context.MenuEmitted && !context.ModalClosed) context.Retain = true;
                if (context.Fixture != null)
                {
                    if (context.Fixture.WordGitMustRetain) context.Retain = true;
                    if (context.Retain)
                    { context.Fixture.NativeExecutionUnsettled = true; lock (Retained) Retained.Add(context); }
                    else
                    {
                        try
                        {
                            context.Record(new { Phase = "NormalCleanupIntent", ReplayAttempts = 0 });
                            context.Fixture.Dispose();
                            if (context.DocumentHash != null) Assert.AreEqual(context.DocumentHash, Sha(context.Scope.Path));
                            context.Record(new { Phase = "NormalCleanupReturned" });
                        }
                        catch (Exception cleanup) { context.OwnerError = context.OwnerError == null ? cleanup : new AggregateException(context.OwnerError, cleanup); }
                    }
                }
                context.OwnerDone.Set();
            }
        }

        private static void Ui(Context context)
        {
            var automation = new EmbeddedGitAutomation(context.Fixture.ProcessId,
                () => context.Fixture.RequireWordEmbeddedOwner(context.Scope), context.Scope, context.Record, () => context.Stop);
            try
            {
                context.Record(new { Phase = "MtaUiaArmed", Apartment = Thread.CurrentThread.GetApartmentState().ToString() });
                context.UiReady.Set();
                if (!context.MenuIntent.Wait(TimeSpan.FromSeconds(40))) throw new TimeoutException("Word menu intent was not emitted.");
                automation.OpenedWindow();
                automation.Link(context.Plan.Remote, context.Plan.Branch, context.Plan.Commit);
                automation.Checkpoint(context.Plan.Remote, context.Plan.Branch, context.Nonce, context.Plan.TabName);
                automation.CompareOnce();
            }
            catch (Exception error) { context.UiError = error; }
            finally
            {
                if (context.MenuEmitted)
                {
                    if (automation.HasObservedWindow && automation.Protocol.CanClose && !context.Stop)
                    {
                        try { automation.CloseKnownTerminal(); context.ModalClosed = true; }
                        catch (Exception cleanup) { context.Retain = true; context.UiError = context.UiError == null ? cleanup : new AggregateException(context.UiError, cleanup); }
                    }
                    else context.Retain = true;
                }
                context.UiDone.Set();
            }
        }

        private static string Sha(string path)
        {
            using (var file = new FileStream(path, FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
            using (var hash = SHA256.Create()) return BitConverter.ToString(hash.ComputeHash(file)).Replace("-", "");
        }

        private sealed class Context
        {
            internal readonly string Root, Nonce = "WORD_OWNER_" + Guid.NewGuid().ToString("N");
            internal readonly EmbeddedGitWindowTests.TestPlan Plan;
            internal readonly ManualResetEventSlim Ready = new ManualResetEventSlim(), StartMenu = new ManualResetEventSlim(),
                UiReady = new ManualResetEventSlim(), MenuIntent = new ManualResetEventSlim(), UiDone = new ManualResetEventSlim(), OwnerDone = new ManualResetEventSlim();
            internal OfficeVbeFixture Fixture;
            internal ExcelVbeFixture.EmbeddedGitScope Scope;
            internal string DocumentHash;
            internal Exception UiError, OwnerError;
            internal volatile bool Stop, Retain, MenuEmitted, ModalClosed;
            private readonly object sync = new object();
            private int sequence;
            internal Context(string root, EmbeddedGitWindowTests.TestPlan plan) { Root = root; Plan = plan; Directory.CreateDirectory(root); }
            internal void Record(object value)
            {
                lock (sync)
                {
                    if (++sequence > 160) throw new InvalidOperationException("Word evidence bound exhausted.");
                    using (var file = new FileStream(Path.Combine(Root, sequence.ToString("D3") + ".json"), FileMode.CreateNew, FileAccess.Write, FileShare.Read))
                    using (var writer = new StreamWriter(file, new UTF8Encoding(false))) writer.Write(new JavaScriptSerializer().Serialize(value));
                }
            }
        }
    }
}

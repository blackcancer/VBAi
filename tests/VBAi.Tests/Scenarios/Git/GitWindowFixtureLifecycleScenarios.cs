using Microsoft.VisualStudio.TestTools.UnitTesting;
using System;
using System.IO;
using System.Linq;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace VBAi.Tests.Unit
{
    /// <summary>Exercises failure containment across the real Git UI fixture and its owned scratch repository.</summary>
    public sealed partial class GitWindowCoverageTests
    {
        /// <summary>Restores ambient state and removes the UI handler even when actual window disposal throws.</summary>
        [WinFormsTestMethod]
        public void FixtureDisposalFailureRestoresCacheContextAndUiHandler()
        {
            var cache = GitWindow.CacheDirectory;
            var context = SynchronizationContext.Current;
            var failure = new InvalidOperationException("Fixture window disposal failed.");
            var fixture = new Fixture();
            EventHandler failDisposal = (sender, args) => { throw failure; };
            fixture.Window.Disposed += failDisposal;
            try
            {
                Assert.AreSame(failure, Assert.ThrowsException<InvalidOperationException>(() => fixture.Dispose()));
                Assert.AreSame(cache, GitWindow.CacheDirectory);
                Assert.AreSame(context, SynchronizationContext.Current);
                Assert.IsFalse(Directory.Exists(fixture.Git.Root), "Terminal Git work must still be cleaned after a window disposal failure.");
                Exception observed = null;
                ThreadExceptionEventHandler capture = (sender, args) => observed = args.Exception;
                var unrelatedUiFailure = new Exception("Subsequent UI failure.");
                Application.ThreadException += capture;
                try { Application.OnThreadException(unrelatedUiFailure); }
                finally { Application.ThreadException -= capture; }
                Assert.AreSame(unrelatedUiFailure, observed);
                Assert.IsNull(fixture.CapturedUiFailure, "A disposed fixture must no longer receive thread exceptions.");
                fixture.Dispose();
            }
            finally
            {
                fixture.Window.Disposed -= failDisposal;
                fixture.Window.Dispose();
                fixture.Git.Dispose();
                fixture.Dispose();
            }
        }

        /// <summary>Retains the original operation error alongside a real locked-file cleanup failure.</summary>
        [WinFormsTestMethod]
        public void FixtureCleanupFailurePreservesOriginalOperationErrorAndLockedScratch()
        {
            var cache = GitWindow.CacheDirectory;
            var context = SynchronizationContext.Current;
            var fixture = new Fixture();
            var original = new InvalidOperationException("Original disposable operation failure.");
            string lockedPath = Path.Combine(fixture.Git.Root, "cleanup-locked.txt");
            try
            {
                Assert.AreSame(original, Assert.ThrowsException<InvalidOperationException>(() => fixture.Pump(Task.FromException(original))));
                using (new FileStream(lockedPath, FileMode.Create, FileAccess.ReadWrite, FileShare.None))
                {
                    var combined = Assert.ThrowsException<AggregateException>(() => fixture.Dispose());
                    Assert.AreSame(original, combined.InnerExceptions[0]);
                    Assert.IsTrue(combined.InnerExceptions.Skip(1).Any(error => error is IOException));
                    StringAssert.Contains(combined.Message, fixture.Git.Root);
                    Assert.IsTrue(Directory.Exists(fixture.Git.Root));
                    Assert.IsTrue(File.Exists(lockedPath));
                }
                Assert.AreSame(cache, GitWindow.CacheDirectory);
                Assert.AreSame(context, SynchronizationContext.Current);
                fixture.Dispose();
            }
            finally { fixture.Window.Dispose(); fixture.Git.Dispose(); fixture.Dispose(); }
        }

        /// <summary>Preserves live window, transaction lock and scratch until the owned operation actually completes.</summary>
        [WinFormsTestMethod]
        public void FixtureDisposalRetainsActiveWorkUntilItsTrackedTaskIsTerminal()
        {
            var cache = GitWindow.CacheDirectory;
            var context = SynchronizationContext.Current;
            var fixture = new Fixture();
            var gate = new TaskCompletionSource<bool>(TaskCreationOptions.RunContinuationsAsynchronously);
            Task operation = null;
            try
            {
                operation = (Task)fixture.Call("Perform", new Func<Task>(() => gate.Task), false);
                Assert.IsFalse(operation.IsCompleted);
                Assert.IsTrue(fixture.Get<bool>("running"));
                var retained = Assert.ThrowsException<InvalidOperationException>(() => fixture.Dispose());
                StringAssert.Contains(retained.Message, fixture.Git.Root);
                Assert.IsTrue(Directory.Exists(fixture.Git.Root));
                Assert.IsFalse(fixture.Window.IsDisposed);
                Assert.IsFalse(operation.IsCompleted);
                using (File.Create(Path.Combine(fixture.Git.Root, "retained-marker.txt"))) { }
                Assert.ThrowsException<IOException>(() =>
                {
                    using (File.Open(Path.Combine(fixture.UiCache, "session.lock"), FileMode.Open, FileAccess.ReadWrite, FileShare.None)) { }
                });
                Assert.AreSame(cache, GitWindow.CacheDirectory);
                Assert.AreSame(context, SynchronizationContext.Current);
                fixture.Dispose();
                Assert.IsTrue(File.Exists(Path.Combine(fixture.Git.Root, "retained-marker.txt")), "A second Dispose must not retry retained cleanup.");
            }
            finally
            {
                gate.TrySetResult(true);
                if (operation != null) fixture.Pump(operation);
                fixture.Window.Dispose();
                fixture.Git.Dispose();
                fixture.Dispose();
            }
            Assert.IsTrue(operation.IsCompleted);
            Assert.IsFalse(Directory.Exists(fixture.Git.Root));
        }
    }
}

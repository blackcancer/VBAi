using Microsoft.VisualStudio.TestTools.UnitTesting;
using System;
using System.IO;
using System.Linq;
using System.Windows.Forms;
using VBAi.Tests.Infrastructure;

namespace VBAi.Tests.Unit
{
    public sealed partial class AddInShutdownObservationTests
    {
        /// <summary>Native close, updater and dispatcher failures all survive; none suppresses later attempts.</summary>
        [STATestMethod, DataRow(false, false), DataRow(false, true), DataRow(true, false), DataRow(true, true)]
        public void MultipleShutdownFailuresKeepTheirOriginalErrorsAfterEveryResourceIsAttempted(bool begin, bool diagnostics)
        {
            using (var scope = new ObservationScope())
            using (var dispatcher = new Control())
            {
                if (!diagnostics) AddIn.CreateShutdownDiagnostic = entry => null;
                var instance = new AddIn();
                var nativeError = new InvalidOperationException("native close failure");
                var stopError = new IOException("stop failure");
                var dispatcherError = new ApplicationException("dispatcher failure");
                var native = new ObservedNativeWindow { Error = nativeError }; int stops = 0;
                Set(instance, "nativeTestWindow", native); Set(instance, "dispatcher", dispatcher);
                AddIn.StopUpdateCheck = () => { stops++; throw stopError; };
                EventHandler failed = (sender, args) => { throw dispatcherError; };
                dispatcher.Disposed += failed;
                try
                {
                    var result = Assert.ThrowsException<AggregateException>(() => Shutdown(instance, begin));
                    CollectionAssert.AreEqual(new Exception[] { nativeError, stopError, dispatcherError }, result.InnerExceptions);
                    Assert.AreEqual(1, native.Calls); Assert.AreEqual(1, stops); Assert.IsTrue(dispatcher.IsDisposed);
                    foreach (string field in new[] { "nativeTestWindow", "testExplorerService", "chat", "server", "dispatcher", "vbe" })
                        Assert.IsNull(Get(instance, field), field);
                    if (diagnostics)
                        foreach (string stage in new[] { "nativeTestWindow.Close", "StopUpdateCheck", "dispatcher.Dispose" })
                            Assert.AreEqual(1, scope.Events.Count(row => Is(row, "Fault", stage)), stage);
                    else Assert.AreEqual(0, scope.Events.Count);
                    AddIn.StopUpdateCheck = () => { };
                    Shutdown(instance, !begin);
                    Assert.AreEqual(1, native.Calls, "Failed close must not be retried by the other shutdown callback.");
                }
                finally { dispatcher.Disposed -= failed; }
            }
        }

        /// <summary>A failing actual chat disposal cannot keep the bridge or owner dispatcher alive.</summary>
        [STATestMethod, DataRow(false), DataRow(true)]
        public void ActualChatDisposalFailureStillClosesBridgeAndDispatcher(bool begin)
        {
            using (var host = new HostUiScope())
            using (var scope = new ObservationScope())
            {
                var instance = host.Connected();
                var chat = (ChatWindow)Get(instance, "chat"); var dispatcher = (Control)Get(instance, "dispatcher");
                var primary = new IOException("chat disposal failed");
                EventHandler failed = (sender, args) => { throw primary; };
                chat.Disposed += failed;
                try
                {
                    Assert.AreSame(primary, Assert.ThrowsException<IOException>(() => Shutdown(instance, begin)));
                    Assert.IsTrue(chat.IsDisposed); Assert.IsTrue(dispatcher.IsDisposed);
                    Assert.IsNull(Get(instance, "chat")); Assert.IsNull(Get(instance, "server")); Assert.IsNull(Get(instance, "dispatcher"));
                    Assert.AreEqual(1, scope.Events.Count(row => Is(row, "Fault", "chat.Dispose")));
                    Before(scope, "Fault", "chat.Dispose", "Returned", "server.Dispose");
                }
                finally { chat.Disposed -= failed; AddIn.CreateShutdownDiagnostic = entry => null; host.Close(instance); }
            }
        }
    }
}

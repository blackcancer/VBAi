using Microsoft.VisualStudio.TestTools.UnitTesting;
using System;
using System.Diagnostics;
using System.Text;
using System.Threading;
using VBAi.Tests.Integration;

namespace VBAi.Tests.Unit
{
    /// <summary>Real kernel containment regressions using only disposable console workers.</summary>
    [TestClass, TestCategory("Unit")]
    public sealed class OllamaBackendJobTests
    {
        private static Process StartWorker(string body)
        {
            var start = new ProcessStartInfo(System.IO.Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.Windows),
                "System32\\WindowsPowerShell\\v1.0\\powershell.exe"),
                "-NoProfile -NonInteractive -EncodedCommand " + Convert.ToBase64String(Encoding.Unicode.GetBytes(body)))
            { UseShellExecute = false, CreateNoWindow = true };
            return Process.Start(start);
        }

        [DataTestMethod]
        [DataRow(false, true)]
        [DataRow(true, false)]
        [DataRow(false, false)]
        public void UnsettledOrRetainedNativeWorkRefusesBeforeStop(bool settled, bool officeAbsent)
        {
            using (var job = new OllamaBackendJob())
            {
                Assert.ThrowsException<InvalidOperationException>(() => job.Stop(settled, officeAbsent));
                Assert.AreEqual(0, job.ReadProcessIds().Length);
            }
        }

        [TestMethod]
        public void KernelJobStopsFutureChildAndPreservesUnrelatedOwnedControl()
        {
            using (var job = new OllamaBackendJob())
            using (var control = StartWorker("[Threading.Thread]::Sleep(60000)"))
            using (var backend = StartWorker("[Threading.Thread]::Sleep(3000); $p=Start-Process -FilePath ($env:WINDIR+'\\System32\\ping.exe') -ArgumentList '-n 60 127.0.0.1' -WindowStyle Hidden -PassThru; $p.WaitForExit()"))
            {
                try
                {
                    job.Attach(backend);
                    Assert.ThrowsException<InvalidOperationException>(() => job.Attach(control));
                    var watch = Stopwatch.StartNew();
                    while (job.ReadProcessIds().Length < 2 && watch.ElapsedMilliseconds < 15000) Thread.Sleep(50);
                    var members = job.ReadProcessIds();
                    Assert.AreEqual(2, members.Length, "The future disposable child must inherit containment.");
                    CollectionAssert.Contains(members, backend.Id);
                    CollectionAssert.DoesNotContain(members, control.Id);
                    CollectionAssert.AreEquivalent(members, job.Stop(true, true));
                    Assert.AreEqual(0, job.ReadProcessIds().Length);
                    Assert.IsTrue(backend.WaitForExit(5000), "Original held backend must exit.");
                    Assert.IsFalse(control.HasExited, "Unrelated process must stay alive.");
                    Assert.ThrowsException<InvalidOperationException>(() => job.Stop(true, true));
                }
                finally
                {
                    // Only our original console process handles; never Office or name-based cleanup.
                    if (!backend.HasExited) { job.Stop(true, true); backend.WaitForExit(5000); }
                    if (!control.HasExited) { control.Kill(); control.WaitForExit(5000); }
                }
            }
        }

        [TestMethod]
        public void DisposedHandleCannotOperateAndDisposeIsIdempotent()
        {
            var job = new OllamaBackendJob();
            job.Dispose(); job.Dispose();
            Assert.ThrowsException<ObjectDisposedException>(() => job.ReadProcessIds());
            Assert.ThrowsException<ObjectDisposedException>(() => job.Stop(true, true));
        }
    }
}

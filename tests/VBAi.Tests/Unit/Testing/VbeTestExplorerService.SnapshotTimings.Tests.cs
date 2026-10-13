using Microsoft.VisualStudio.TestTools.UnitTesting;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;

namespace VBAi.Tests.Unit
{
    public sealed partial class VbeTestExplorerServiceTests
    {
        [STATestMethod]
        [DataRow("none"), DataRow("source failure"), DataRow("sink failure")]
        public void SnapshotComponentTimingsPreserveSourceReadsAndFailureSemantics(string fault)
        {
            using (var fixture = new Fixture())
            {
                var events = new List<NativeExecutionTimings.StageEvent>();
                var timings = new NativeExecutionTimings(value =>
                {
                    if (fault == "sink failure") throw new InvalidOperationException("diagnostic refused");
                    events.Add(value);
                });
                typeof(VbeTestExplorerService).GetField("nativeTimings", BindingFlags.NonPublic | BindingFlags.Instance)
                    .SetValue(fixture.Service, timings);
                var method = typeof(VbeTestExplorerService).GetMethod("Snapshot", BindingFlags.NonPublic | BindingFlags.Instance);
                var sourceFailure = new InvalidOperationException("source refused");
                if (fault == "source failure") fixture.Project.VBComponents[0].CodeModule.BeforeReadLines = () => { throw sourceFailure; };
                VbaTestProjectSnapshot snapshot = null;
                Action read = () => snapshot = (VbaTestProjectSnapshot)method.Invoke(fixture.Service, new object[] { "owned", fixture.Project });
                if (fault == "source failure")
                {
                    var failure = Assert.ThrowsException<TargetInvocationException>(read);
                    Assert.AreSame(sourceFailure, failure.InnerException);
                    Assert.IsTrue(events.Any(value => value.Stage == "SnapshotModules" && value.State == "Failed"));
                    Assert.IsFalse(events.Any(value => value.Stage == "SnapshotReferences"));
                }
                else
                {
                    read();
                    Assert.AreEqual(fixture.Project.FileName, snapshot.HostPath);
                    Assert.AreEqual(fixture.Project.VBComponents.Count, snapshot.Modules.Length);
                    if (fault == "none")
                        CollectionAssert.AreEqual(new[] { "ProjectSnapshot", "SnapshotHostPath", "SnapshotModules", "SnapshotReferences" },
                            events.Where(value => value.State == "Started").Select(value => value.Stage).ToArray());
                    else { Assert.AreEqual(0, events.Count); Assert.IsNotNull(timings.DiagnosticFailure); }
                }
                Assert.AreEqual(0, fixture.Host.Invocations);
            }
        }

        [STATestMethod]
        public void SnapshotDiagnosticsNeverCacheRevisionAfterSourceChanges()
        {
            using (var fixture = new Fixture())
            {
                var catalog = fixture.Catalog();
                var events = new List<NativeExecutionTimings.StageEvent>();
                typeof(VbeTestExplorerService).GetField("nativeTimings", BindingFlags.NonPublic | BindingFlags.Instance)
                    .SetValue(fixture.Service, new NativeExecutionTimings(events.Add));
                fixture.Project.VBComponents[0].CodeModule.Source += "\r\n' changed after discovery";
                Assert.ThrowsException<VbaTestInvocationException>(() => fixture.Service.Validate(catalog));
                Assert.IsTrue(events.Any(value => value.Stage == "SnapshotModules" && value.State == "Returned"));
                Assert.IsTrue(events.Any(value => value.Stage == "SnapshotReferences" && value.State == "Returned"));
                Assert.AreEqual(0, fixture.Host.Invocations);
            }
        }
    }
}

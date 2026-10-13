using Microsoft.VisualStudio.TestTools.UnitTesting;
using System;
using System.Collections.Generic;
using System.IO;
using System.Threading;
using VBAi.Tests.Integration;

namespace VBAi.Tests.Unit
{
    /// <summary>Exercises the complete copy-consent policy with owned files and synthetic UI evidence; no Office is opened.</summary>
    [TestClass, TestCategory("Unit")]
    public sealed class OfficeCoverageCopyConsentTests
    {
        [DataTestMethod, DataRow("Word"), DataRow("PowerPoint")]
        public void ExactFreshSyntheticCopyReceivesOneFileSpecificConsentWithoutAnyReplay(string host)
        {
            using (var probe = new Probe(host))
            {
                var policy = probe.Create();
                policy.Poll(); policy.Poll();
                Assert.IsTrue(policy.Attempted); Assert.IsTrue(policy.Returned);
                Assert.AreEqual(probe.Copy, policy.ConsentedPath);
                Assert.AreEqual(1, probe.Invocations);
                CollectionAssert.AreEqual(new[] { "coverage-copy-consent-intent", "coverage-copy-consent-returned" }, probe.Phases);
                Assert.AreEqual(3, probe.Reads);
                Assert.AreEqual("synthetic source", File.ReadAllText(probe.Original));
                Assert.AreEqual("synthetic copy", File.ReadAllText(probe.Copy));
            }
        }

        [DataTestMethod, DataRow("pid"), DataRow("title"), DataRow("class"), DataRow("handle"), DataRow("runtime"),
            DataRow("buttonIdentity"), DataRow("disabled"), DataRow("buttons"), DataRow("invoke"), DataRow("paths"),
            DataRow("relative"), DataRow("root"), DataRow("name"), DataRow("guid"), DataRow("prior"),
            DataRow("missing"), DataRow("original"), DataRow("changedSource"), DataRow("extraCopy"), DataRow("owner")]
        public void UnverifiedScopeOrWarningNeverRecordsConsentOrInvokesAnAction(string fault)
        {
            using (var probe = new Probe("PowerPoint"))
            {
                var policy = probe.Create();
                switch (fault)
                {
                    case "pid": probe.Dialog.ProcessId++; break;
                    case "title": probe.Dialog.Title = "Other warning"; break;
                    case "class": probe.Dialog.ClassName = "PPTFrameClass"; break;
                    case "handle": probe.Dialog.Handle = 0; break;
                    case "runtime": probe.Dialog.RuntimeIdentity = null; break;
                    case "buttonIdentity": probe.Dialog.ButtonIdentity = null; break;
                    case "disabled": probe.Dialog.Enabled = false; break;
                    case "buttons": probe.Dialog.EnableButtons = 2; break;
                    case "invoke": probe.Dialog.Invoke = null; break;
                    case "paths": probe.Dialog.Paths = new[] { probe.Copy, probe.Original }; break;
                    case "relative": probe.Dialog.Paths = new[] { "coverage.pptm" }; break;
                    case "root": probe.Dialog.Paths = new[] { Path.Combine(probe.Folder, "Other", Guid.NewGuid().ToString("N"), "coverage.pptm") }; break;
                    case "name": probe.Dialog.Paths = new[] { Path.Combine(Path.GetDirectoryName(probe.Copy), "Other.pptm") }; break;
                    case "guid": probe.Dialog.Paths = new[] { Path.Combine(probe.CoverageRoot, "not-a-guid", "coverage.pptm") }; break;
                    case "prior": probe.Prior.Add(Path.GetDirectoryName(probe.Copy)); policy = probe.Create(); break;
                    case "missing": File.Delete(probe.Copy); break;
                    case "original": probe.Dialog.Paths = new[] { probe.Original }; break;
                    case "changedSource": File.WriteAllText(probe.Original, "source changed after dispatch"); break;
                    case "extraCopy": Directory.CreateDirectory(Path.Combine(probe.CoverageRoot, Guid.NewGuid().ToString("N"))); break;
                    case "owner": probe.OwnerAlive = false; break;
                    default: Assert.Fail("Unknown refusal scenario."); break;
                }
                Assert.ThrowsException<InvalidOperationException>(policy.Poll);
                Assert.IsFalse(policy.Attempted);
                Assert.AreEqual(0, probe.Phases.Count);
                Assert.AreEqual(0, probe.Invocations);
            }
        }

        [DataTestMethod, DataRow("null"), DataRow("ambiguous")]
        public void IncompleteOrAmbiguousEnumerationCannotActivateAnyWarning(string fault)
        {
            using (var probe = new Probe("Word"))
            {
                probe.Inspect = () => fault == "null" ? null : new[] { probe.Dialog, probe.Dialog };
                Assert.ThrowsException<InvalidOperationException>(probe.Create().Poll);
                Assert.AreEqual(0, probe.Invocations); Assert.AreEqual(0, probe.Phases.Count);
            }
        }

        [DataTestMethod, DataRow(2), DataRow(3)]
        public void DialogReplacementAtEitherFinalCheckIsRefusedWithoutInvocation(int changedRead)
        {
            using (var probe = new Probe("Word"))
            {
                probe.Inspect = () =>
                {
                    if (probe.Reads == changedRead) probe.Dialog.ButtonIdentity = "replacement button";
                    return new[] { probe.CopySnapshot() };
                };
                var policy = probe.Create();
                Assert.ThrowsException<InvalidOperationException>(policy.Poll);
                Assert.AreEqual(changedRead == 3, policy.Attempted);
                Assert.AreEqual(changedRead == 3 ? 1 : 0, probe.Phases.Count);
                Assert.AreEqual(0, probe.Invocations);
            }
        }

        [DataTestMethod, DataRow("intent"), DataRow("invoke"), DataRow("returned")]
        public void ReceiptOrNativeActionFailureLatchesTheAttemptAndNeverRetries(string failure)
        {
            using (var probe = new Probe("PowerPoint"))
            {
                probe.Failure = failure;
                var policy = probe.Create();
                Assert.ThrowsException<IOException>(policy.Poll);
                Assert.IsTrue(policy.Attempted);
                Assert.AreEqual(failure == "returned", policy.Returned);
                int reads = probe.Reads;
                policy.Poll();
                Assert.AreEqual(reads, probe.Reads);
                Assert.AreEqual(failure == "intent" ? 0 : 1, probe.Invocations);
            }
        }

        [TestMethod]
        public void NoWarningIsReadOnlyAndASecondaryThreadCannotEvenEnumerate()
        {
            using (var probe = new Probe("Word"))
            {
                probe.Inspect = () => new OfficeCoverageCopyConsent.Dialog[0];
                var policy = probe.Create();
                policy.Poll();
                Assert.IsFalse(policy.Attempted);
                Exception refusal = null;
                var thread = new Thread(() => { try { policy.Poll(); } catch (Exception error) { refusal = error; } });
                thread.Start(); thread.Join();
                Assert.IsInstanceOfType(refusal, typeof(InvalidOperationException));
                Assert.AreEqual(1, probe.Reads);
                Assert.AreEqual(0, probe.Invocations);
            }
        }

        private sealed class Probe : IDisposable
        {
            internal readonly string Folder = Path.Combine(Path.GetTempPath(), "VBAi-CopyConsent-" + Guid.NewGuid().ToString("N"));
            internal readonly string Original, CoverageRoot, Copy, Host;
            internal readonly List<string> Prior = new List<string>(), Phases = new List<string>();
            internal OfficeCoverageCopyConsent.Dialog Dialog;
            internal Func<OfficeCoverageCopyConsent.Dialog[]> Inspect;
            internal bool OwnerAlive = true;
            internal int Reads, Invocations;
            internal string Failure;
            internal Probe(string host)
            {
                Host = host;
                Original = Path.Combine(Folder, "Disposable" + (host == "Word" ? ".docm" : ".pptm"));
                CoverageRoot = Path.Combine(Folder, "CoverageRuns");
                Copy = Path.Combine(CoverageRoot, Guid.NewGuid().ToString("N"), "coverage" + Path.GetExtension(Original));
                Directory.CreateDirectory(Path.GetDirectoryName(Copy));
                File.WriteAllText(Original, "synthetic source"); File.WriteAllText(Copy, "synthetic copy");
                Dialog = new OfficeCoverageCopyConsent.Dialog
                {
                    Handle = 99, ProcessId = 123, Title = "Avis de sécurité pour Microsoft " + host,
                    ClassName = "NUIDialog", RuntimeIdentity = "owned dialog", ButtonIdentity = "owned button",
                    Paths = new[] { Copy }, EnableButtons = 1, Enabled = true,
                    Invoke = () => { Invocations++; if (Failure == "invoke") throw new IOException("Unknown UI completion"); }
                };
                Inspect = () => new[] { CopySnapshot() };
            }
            internal OfficeCoverageCopyConsent Create() => new OfficeCoverageCopyConsent(Host, 123, Original, CoverageRoot,
                Prior, () => { if (!OwnerAlive) throw new InvalidOperationException("Original process unavailable"); },
                () => { Reads++; return Inspect(); }, (phase, value) =>
                {
                    if ((Failure == "intent" && phase.EndsWith("intent", StringComparison.Ordinal))
                        || (Failure == "returned" && phase.EndsWith("returned", StringComparison.Ordinal))) throw new IOException("Receipt unavailable");
                    Phases.Add(phase);
                });
            internal OfficeCoverageCopyConsent.Dialog CopySnapshot() => new OfficeCoverageCopyConsent.Dialog
            {
                Handle = Dialog.Handle, ProcessId = Dialog.ProcessId, Title = Dialog.Title, ClassName = Dialog.ClassName,
                RuntimeIdentity = Dialog.RuntimeIdentity, ButtonIdentity = Dialog.ButtonIdentity,
                Paths = Dialog.Paths, EnableButtons = Dialog.EnableButtons, Enabled = Dialog.Enabled, Invoke = Dialog.Invoke
            };
            public void Dispose() { Directory.Delete(Folder, true); }
        }
    }
}

using Microsoft.VisualStudio.TestTools.UnitTesting;
using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading;
using System.Web.Script.Serialization;
using System.Windows.Automation;

namespace VBAi.Tests.Integration
{
    internal sealed partial class OfficeVbeFixture
    {
        private int coverageConsentSequence;

        /// <summary>Opt-in qualification of one reviewed synthetic copy on the main desktop, never a product trust-policy change.</summary>
        internal OfficeCoverageCopyConsent CreateCoverageCopyConsent(string coverageRoot, IEnumerable<string> priorFolders)
        {
            if (Environment.GetEnvironmentVariable("VBAi_RUN_OFFICE_COPY_CONSENT") != "1") return null;
            RequireCoverageConsentOwner();
            int sequence = checked(++coverageConsentSequence);
            return new OfficeCoverageCopyConsent(Kind, ProcessId, DocumentPath, coverageRoot, priorFolders,
                RequireCoverageConsentOwner, ReadCoverageWarnings, (phase, value) =>
                {
                    const string prefix = "coverage-copy-consent-";
                    if (phase == null || !phase.StartsWith(prefix, StringComparison.Ordinal))
                        throw new ArgumentException("Unknown coverage consent receipt phase.", nameof(phase));
                    RecordCoverageReceipt(Root, "coverage-copy-consent", sequence, phase.Substring(prefix.Length), value);
                });
        }

        /// <summary>Gives each independent copy scope its own immutable receipts while preserving the first scope's names.</summary>
        internal static string CoverageReceiptPath(string root, string prefix, int sequence, string phase)
        {
            if (string.IsNullOrWhiteSpace(root) || !Path.IsPathRooted(root))
                throw new ArgumentException("An absolute evidence root is required.", nameof(root));
            if (sequence <= 0) throw new ArgumentOutOfRangeException(nameof(sequence));
            bool known = prefix == "coverage-copy-consent" ? phase == "intent" || phase == "returned" :
                prefix == "word-coverage-events" && (phase == "enable-intent" || phase == "enable-returned" ||
                    phase == "enable-checked" || phase == "restore-intent" || phase == "restore-returned" || phase == "restore-released");
            if (!known) throw new ArgumentException("Unknown coverage receipt scope or phase.", nameof(phase));
            string scope = prefix + (sequence == 1 ? "" : "-" + sequence.ToString(CultureInfo.InvariantCulture));
            return Path.Combine(root, scope + "-" + phase + ".json");
        }

        /// <summary>Publishes one receipt without replacing evidence from an earlier phase or measurement.</summary>
        internal static void RecordCoverageReceipt(string root, string prefix, int sequence, string phase, object value)
        {
            string path = CoverageReceiptPath(root, prefix, sequence, phase);
            using (var stream = new FileStream(path, FileMode.CreateNew, FileAccess.Write, FileShare.Read))
            using (var writer = new StreamWriter(stream, new UTF8Encoding(false)))
                writer.Write(new JavaScriptSerializer().Serialize(value));
        }

        /// <summary>Checks only original-handle and desktop evidence; no COM getter is called while the host may be modal.</summary>
        private void RequireCoverageConsentOwner()
        {
            Assert.AreEqual(shutdownOwnerThread, Thread.CurrentThread.ManagedThreadId);
            Assert.IsTrue(Kind == "Word" || Kind == "PowerPoint");
            Assert.IsTrue(owned && ownedProcess != null && shutdownEvidence != null && !hostTeardownRefused);
            Assert.IsFalse(ownedProcess.HasExited);
            Assert.AreEqual(ProcessId, ownedProcess.Id);
            Assert.AreEqual(shutdownEvidence.Record["ProcessStartedUtc"], ownedProcess.StartTime.ToUniversalTime().ToString("o"));
            Assert.AreEqual(shutdownEvidence.Record["OriginalProcessHandle"], "0x" + unchecked((ulong)ownedProcess.Handle.ToInt64()).ToString("X16"));
            Assert.AreEqual(shutdownEvidence.Record["ProcessImage"], ExcelOwnedProcessImage.Read(ownedProcess.Handle));
            Assert.AreEqual(Kind == "Word" ? "WINWORD.EXE" : "POWERPNT.EXE",
                Path.GetFileName(Convert.ToString(shutdownEvidence.Record["ProcessImage"])).ToUpperInvariant());
            Assert.AreEqual(0, Convert.ToInt32(shutdownEvidence.Record["QuitEntries"]));
            IsolatedTestDesktop.RequireMainCurrent();
        }

        /// <summary>Reads only the owned application's addressed security windows, without changing focus.</summary>
        private OfficeCoverageCopyConsent.Dialog[] ReadCoverageWarnings()
        {
            RequireCoverageConsentOwner();
            string title = "Avis de sécurité pour Microsoft " + Kind;
            var windows = new List<IntPtr>();
            bool enumerated = EnumWindows((window, _) =>
            {
                uint pid;
                GetWindowThreadProcessId(window, out pid);
                if (pid == (uint)ProcessId && IsWindowVisible(window) && BoundedDialogText(window) == title)
                    windows.Add(window);
                return windows.Count < 16;
            }, IntPtr.Zero);
            if (!enumerated || windows.Count >= 16) throw new InvalidOperationException("Owned warning enumeration is incomplete.");
            return windows.Select(window =>
            {
                var root = AutomationElement.FromHandle(window);
                if (root.Current.ProcessId != ProcessId || root.Current.Name != title)
                    throw new InvalidOperationException("The owned warning changed during inspection.");
                var nativeClass = new StringBuilder(128);
                if (GetClassName(window, nativeClass, nativeClass.Capacity) == 0)
                    throw new InvalidOperationException("The native warning class is unreadable.");
                var edits = root.FindAll(TreeScope.Descendants,
                    new PropertyCondition(AutomationElement.ControlTypeProperty, ControlType.Edit));
                var buttons = root.FindAll(TreeScope.Descendants, new AndCondition(
                    new PropertyCondition(AutomationElement.ControlTypeProperty, ControlType.Button),
                    new PropertyCondition(AutomationElement.NameProperty, "Activer les macros")));
                var paths = edits.Cast<AutomationElement>().Select(edit =>
                {
                    if (edit.Current.ProcessId != ProcessId) throw new InvalidOperationException("Foreign warning path control.");
                    return ((ValuePattern)edit.GetCurrentPattern(ValuePattern.Pattern)).Current.Value;
                }).ToArray();
                AutomationElement button = buttons.Count == 1 ? buttons[0] : null;
                return new OfficeCoverageCopyConsent.Dialog
                {
                    Handle = window.ToInt64(), ProcessId = root.Current.ProcessId, Title = root.Current.Name,
                    ClassName = nativeClass.ToString(), Paths = paths, EnableButtons = buttons.Count,
                    RuntimeIdentity = string.Join("/", root.GetRuntimeId()),
                    ButtonIdentity = button == null ? null : string.Join("/", button.GetRuntimeId()),
                    Enabled = button != null && button.Current.ProcessId == ProcessId && button.Current.IsEnabled && !button.Current.IsOffscreen,
                    // UIA Invoke can wait for Word's synchronous DocumentOpen callback. Keep
                    // the owning STA pumping until the single background invocation returns.
                    Invoke = button == null ? null : (Action)(() => OfficeCoverageEventObserver.ReadPumping(() =>
                    {
                        ((InvokePattern)button.GetCurrentPattern(InvokePattern.Pattern)).Invoke();
                        return true;
                    }, System.Windows.Forms.Application.DoEvents, Thread.Sleep))
                };
            }).ToArray();
        }
    }
}

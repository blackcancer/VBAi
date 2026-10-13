using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Threading;
using System.Web.Script.Serialization;
using VBAi.Desktop.Helper;

namespace VBAi.Tests.Integration
{
    /// <summary>Fixed Word observer entry. Its process exits before the parent may close the owned document.</summary>
    public static class OllamaOfficeObserverWorker
    {
        public static int Run(string json)
        {
            var serializer = new JavaScriptSerializer { MaxJsonLength = 8 * 1024 * 1024 };
            var plan = serializer.Deserialize<OfficeChatObserverPlan>(json);
            var profile = OllamaQualificationProfile.Resolve();
            if (profile.Model != plan.Model || profile.Describe() != plan.Profile)
                throw new InvalidOperationException("Observer sampling profile differs from the parent manifest.");
            var observations = new List<object>();
            var receipt = new OfficeChatObserverReceipt
            {
                Nonce = plan.Nonce, ProcessId = Process.GetCurrentProcess().Id,
                HostPid = plan.HostPid, HostBirthUtc = plan.HostBirthUtc,
                TestsMvid = typeof(OllamaOfficeObserverWorker).Module.ModuleVersionId.ToString("D"),
                ProductMvid = typeof(VbeSession).Module.ModuleVersionId.ToString("D")
            };
            Action flush = () =>
            {
                receipt.Observations = observations.ToArray();
                File.WriteAllText(plan.Receipt, serializer.Serialize(receipt));
            };
            Action<object> record = observation => { observations.Add(observation); flush(); };
            flush(); // A partial receipt never authorizes host shutdown.
            using (var originalParent = Process.GetProcessById(plan.ParentPid))
            {
                // Pin this generation rather than interpreting a later PID lookup as liveness.
                var originalParentHandle = originalParent.Handle;
                string actualParentBirth = originalParent.StartTime.ToUniversalTime().ToString("o");
                Action requireOwner = () => RequireParentAlive(originalParent.WaitForExit(0), plan.ParentBirthUtc, actualParentBirth);
                requireOwner();
                record(new { Phase = "ObserverParentGenerationPinned", ProcessId = originalParent.Id,
                    OriginalQueryHandle = originalParentHandle.ToInt64(), BirthUtc = actualParentBirth });
                var worker = new Thread(() =>
                {
                    OllamaOfficeUi ui = null;
                    using (var culture = new VBAi.Tests.Infrastructure.LocalizationScope(plan.Culture))
                    {
                        try
                        {
                            ui = new OllamaOfficeUi(plan.HostPid, record, requireOwner);
                            OllamaOfficeChatScenario.Run(ui, plan.Label, plan.Project, plan.Module, plan.Marker, plan.Model,
                                stage => { SetPassed(receipt, stage); flush(); });
                        }
                        catch (Exception error)
                        {
                            receipt.Error = error.ToString();
                            try { if (ui != null) record(new { Phase = "FailureTerminalObservation", AlreadyIdle = ui.ObserveAlreadyIdle() }); }
                            catch (Exception observation) { record(new { Phase = "TerminalObservationFailed", Error = observation.ToString() }); }
                            try
                            {
                                if (ui != null) record(new { Phase = "FailureFinalTranscript", SyntheticQualificationOnly = true,
                                    Texts = ui.VisibleTranscript().Select(text => new { Characters = text.Length,
                                        Sha256 = EditorDocument.Hash(text), Preview = text.Substring(0, Math.Min(512, text.Length)) }).ToArray() });
                            }
                            catch (Exception observation) { record(new { Phase = "FinalTranscriptObservationFailed", Error = observation.ToString() }); }
                        }
                        finally
                        {
                            receipt.Settled = ui != null && !ui.SentUnsettled;
                            if (receipt.Settled)
                            {
                                try { ui.ReleaseSettledReferences(); }
                                catch (Exception cleanup)
                                {
                                    // Process exit, not successful RCW release, is the isolation boundary.
                                    receipt.Error = (receipt.Error ?? "") + "\nObserver cleanup: " + cleanup;
                                }
                            }
                            receipt.Completed = true; flush();
                        }
                    }
                });
                worker.SetApartmentState(ApartmentState.MTA); worker.Start(); worker.Join();
                return receipt.Settled && receipt.Error == null && receipt.StreamedWhileBusy &&
                    receipt.CancellationAcknowledged && receipt.NextSendCompleted && receipt.NativeMarkerVisible ? 0 : 1;
            }
        }

        /// <summary>Owner exit or generation replacement closes admission without another native action.</summary>
        internal static void RequireParentAlive(bool exited, string expectedBirth, string actualBirth)
        {
            if (exited || string.IsNullOrEmpty(expectedBirth) || expectedBirth != actualBirth)
                throw new InvalidOperationException("The original observer parent exited or changed; no further native action is admitted.");
        }

        internal static void SetPassed(OfficeChatObserverReceipt receipt, string stage)
        {
            switch (stage)
            {
                case "StreamedWhileBusy": receipt.StreamedWhileBusy = true; break;
                case "CancellationAcknowledged": receipt.CancellationAcknowledged = true; break;
                case "NextSendCompleted": receipt.NextSendCompleted = true; break;
                case "NativeMarkerVisible": receipt.NativeMarkerVisible = true; break;
                default: throw new ArgumentException("Unknown assistant acceptance oracle.", nameof(stage));
            }
        }
    }
}

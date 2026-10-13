using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Security.Principal;
using System.Web.Script.Serialization;
using VBAi.Desktop.Helper;

namespace VBAi.Tests.Integration
{
    /// <summary>Owns one observer creation handle. It neither retries UI actions nor closes an uncertain Word host.</summary>
    internal static class OllamaOfficeObserverProcess
    {
        private static readonly List<IsolatedTestDesktop.NativeChild> retainedObservers = new List<IsolatedTestDesktop.NativeChild>();

        internal static OfficeChatObserverReceipt Run(OllamaOfficeHost host, string root, string module, string marker,
            OllamaQualificationProfile profile, Action<object> record, Action settled)
        {
            if (!OllamaOfficeDesktop.MainEnabled)
                throw new InvalidOperationException("The isolated Word observer requires the explicit real-desktop opt-in.");
            OllamaOfficeDesktop.Require("Default");
            string folder = Path.Combine(root, "word-observer");
            if (Directory.Exists(folder)) throw new InvalidOperationException("An observer bank is one-shot.");
            Directory.CreateDirectory(folder);
            var plan = new OfficeChatObserverPlan
            {
                Version = 1, Nonce = Guid.NewGuid().ToString("N"), UserSid = WindowsIdentity.GetCurrent().User.Value,
                Helper = typeof(OfficeChatObserverEntry).Assembly.Location,
                Tests = typeof(OllamaOfficeObserverProcess).Assembly.Location, Product = typeof(VbeSession).Assembly.Location,
                TestsMvid = typeof(OllamaOfficeObserverProcess).Module.ModuleVersionId.ToString("D"),
                ProductMvid = typeof(VbeSession).Module.ModuleVersionId.ToString("D"),
                HostPid = host.ProcessId, Label = host.Label, Project = host.ToolProject, Module = module, Marker = marker,
                Model = profile.Model, Profile = profile.Describe(), Culture = UiText.Culture.Name,
                Receipt = Path.Combine(folder, "observer-receipt.json")
            };
            using (var parent = Process.GetCurrentProcess())
            using (var owned = Process.GetProcessById(host.ProcessId))
            {
                plan.ParentPid = parent.Id; plan.ParentBirthUtc = parent.StartTime.ToUniversalTime().ToString("o");
                plan.ParentImage = parent.MainModule.FileName;
                plan.HostBirthUtc = owned.StartTime.ToUniversalTime().ToString("o"); plan.HostImage = owned.MainModule.FileName;
            }
            plan.HelperSha256 = OfficeChatObserverEntry.Hash(plan.Helper);
            plan.TestsSha256 = OfficeChatObserverEntry.Hash(plan.Tests);
            plan.ProductSha256 = OfficeChatObserverEntry.Hash(plan.Product);
            string manifest = Path.Combine(folder, "observer-plan.json");
            OfficeChatObserverEntry.Validate(plan, manifest, plan.Helper, File.Exists, OfficeChatObserverEntry.Hash);
            var serializer = new JavaScriptSerializer { MaxJsonLength = 8 * 1024 * 1024 };
            File.WriteAllText(manifest, serializer.Serialize(plan));
            IsolatedTestDesktop.NativeChild child = null;
            bool exitObserved = false;
            try
            {
                child = IsolatedTestDesktop.LaunchMain(plan.Helper, new[] { "--observe-word-chat", manifest }, Path.GetDirectoryName(plan.Tests));
                record(new { Phase = "ObserverStartedOnce", child.ProcessId, OriginalCreationHandle = child.ProcessHandle.ToInt64(),
                    Desktop = "WinSta0\\Default", HelperSha256 = plan.HelperSha256, NativeReplayAllowed = false });
                var timer = Stopwatch.StartNew();
                while (!(exitObserved = child.Wait(20)))
                {
                    System.Windows.Forms.Application.DoEvents();
                    if (timer.Elapsed > TimeSpan.FromMinutes(20))
                        throw new InvalidOperationException("Original observer exit was not observed; retain the observer and host without force or replay.");
                }
                uint exitCode = child.ExitCode();
                var receipt = File.Exists(plan.Receipt) ? serializer.Deserialize<OfficeChatObserverReceipt>(File.ReadAllText(plan.Receipt)) : null;
                record(new { Phase = "OriginalObserverExitObserved", child.ProcessId,
                    OriginalCreationHandle = child.ProcessHandle.ToInt64(), ExitCode = exitCode, Receipt = receipt,
                    ForcedTermination = false });
                OfficeChatObserverEntry.RequireTerminal(plan, receipt, child.ProcessId, exitObserved, exitCode);
                settled(); // Only now can the parent perform its existing guarded teardown.
                return receipt;
            }
            finally
            {
                if (child != null)
                {
                    if (exitObserved) child.Dispose();
                    else lock (retainedObservers) retainedObservers.Add(child);
                }
            }
        }
    }
}

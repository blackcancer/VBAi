using Microsoft.VisualStudio.TestTools.UnitTesting;
using System;
using System.Threading;

namespace VBAi.Tests.Integration
{
    internal sealed partial class OfficeVbeFixture
    {
        /// <summary>Suppresses legacy auto macros before creating documents in the exact owned Word generation.</summary>
        private void PrepareOwnedWordMacroSafety()
        {
            try
            {
                SuppressWordAutoMacros(application, () =>
                {
                    RequireUsableOwnedHost();
                    Assert.AreEqual("Word", Kind);
                    Assert.AreEqual(shutdownOwnerThread, Thread.CurrentThread.ManagedThreadId);
                    Assert.AreEqual(ApartmentState.STA, Thread.CurrentThread.GetApartmentState());
                    Assert.IsTrue(owned && ownedProcess != null && !ownedProcess.HasExited);
                    Assert.AreEqual(ProcessId, ownedProcess.Id);
                    Assert.AreEqual(shutdownEvidence.Record["ProcessStartedUtc"], ownedProcess.StartTime.ToUniversalTime().ToString("o"));
                    Assert.AreEqual(shutdownEvidence.Record["OriginalProcessHandle"],
                        "0x" + unchecked((ulong)ownedProcess.Handle.ToInt64()).ToString("X16"));
                }, Release, value => retainedDiagnosticReferences.Add(value));
                steps.Add(new
                {
                    WordAutoMacroSuppressionReturned = true,
                    ProcessId,
                    AutomationSecurity = 3,
                    ReenabledBeforeQuit = false
                });
            }
            catch
            {
                NativeExecutionUnsettled = true;
                throw;
            }
        }

        /// <summary>Refuses document work after a failed suppression; preserves an uncertain WordBasic lease without replay.</summary>
        internal static void SuppressWordAutoMacros(object app, Action requireOwner,
            Action<object> release, Action<object> retain)
        {
            if (app == null) throw new ArgumentNullException(nameof(app));
            if (requireOwner == null) throw new ArgumentNullException(nameof(requireOwner));
            if (release == null) throw new ArgumentNullException(nameof(release));
            if (retain == null) throw new ArgumentNullException(nameof(retain));
            requireOwner();
            ((dynamic)app).AutomationSecurity = 3;
            object wordBasic = ((dynamic)app).WordBasic;
            if (wordBasic == null) throw new InvalidOperationException("The owned Word instance has no WordBasic automation object.");
            try
            {
                // Microsoft documents the parameterless command as suppressing Word auto macros.
                ((dynamic)wordBasic).DisableAutoMacros();
            }
            catch
            {
                retain(wordBasic);
                throw;
            }
            try { release(wordBasic); }
            catch { retain(wordBasic); throw; }
        }
    }
}

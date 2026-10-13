using Microsoft.VisualStudio.TestTools.UnitTesting;
using System;
using System.Collections.Generic;

namespace VBAi.Tests.Integration
{
    internal sealed partial class OfficeVbeFixture
    {
        private int wordCoverageEventScopeSequence;

        /// <summary>Enables native copy events only in the reviewed owned /a fixture; disabled by default.</summary>
        internal OfficeWordCoverageEventScope StartWordCoverageEventScope()
        {
            if (Environment.GetEnvironmentVariable("VBAi_RUN_WORD_COVERAGE_EVENT_SCOPE") != "1") return null;
            Assert.AreEqual("Word", Kind);
            RequireCoverageConsentOwner();
            Assert.IsFalse(NativeExecutionUnsettled);
            Assert.IsTrue(wordStartupMacroSafetyVerified, "The original startup suppression must be known.");
            int sequence = checked(++wordCoverageEventScopeSequence);
            var metadata = new List<object>();
            OfficeWordCoverageEventScope scope = null;
            scope = new OfficeWordCoverageEventScope(new OfficeWordCoverageEventScope.Boundary
            {
                RequireOwner = () => { RequireCoverageConsentOwner(); RequireMainWordOriginal(); RequireOwnedDocument(); },
                ReviewTemplates = () => RequireReviewedWordEventTemplates(application, metadata.Add),
                ReadSecurity = () => Convert.ToInt32(((dynamic)application).AutomationSecurity),
                AcquireCommands = () => (object)((dynamic)application).WordBasic,
                ApplySuppression = (commands, state) => ((dynamic)commands).DisableAutoMacros(state),
                Record = phase => RecordWordCoverageEventScope(phase, scope, metadata, sequence)
            });
            try { scope.Enter(); return scope; }
            catch { NativeExecutionUnsettled = scope.Uncertain; throw; }
        }

        /// <summary>Restores suppression only after the native copy has settled; otherwise retains the original lease.</summary>
        internal void RestoreWordCoverageEventScope(OfficeWordCoverageEventScope scope)
        {
            if (scope == null) return;
            if (NativeExecutionUnsettled) { scope.Retain(); return; }
            try { scope.Restore(); }
            catch { NativeExecutionUnsettled = scope.Uncertain; throw; }
        }

        private void RecordWordCoverageEventScope(string phase, OfficeWordCoverageEventScope scope, List<object> metadata, int sequence)
        {
            RecordCoverageReceipt(Root, "word-coverage-events", sequence, phase, new
                {
                    ProcessId, DocumentPath, Phase = phase, OriginalSuppression = 1,
                    scope.Entered, scope.Active, scope.RestoreEntered, scope.Restored,
                    scope.ReferencesReleased, scope.Uncertain, TemplateMetadata = metadata.ToArray(),
                    TrustSettingsChanged = false, CodeTextRead = false, DirectPersonalTemplateFileRead = false
                });
        }
    }
}

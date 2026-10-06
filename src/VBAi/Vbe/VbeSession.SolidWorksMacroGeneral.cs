using System;
using System.Threading.Tasks;

namespace VBAi
{

    /// <summary>Provides the narrow nested General-page read needed by an already-authorized SOLIDWORKS macro publication.</summary>
    internal sealed partial class VbeSession
    {
        // Only the already authorized publication may perform this nested read.
        // Public General dispatch remains excluded while a macro operation owns the STA.
        /// <summary>Reads the selected source project's General page only while the current session owns the in-flight publication on its VBE STA.</summary>
        /// <param name="source">Publication request; only a mode-2 read with exact project and expected revision is accepted.</param>
        /// <param name="authorize">Publication authorization callback repeated around native selection and General-page access.</param>
        /// <param name="context">Owner-thread check that must hold before every native step.</param>
        /// <returns>General snapshot result with closure/uncertainty evidence; a failed entered command quarantines the session.</returns>
        private async Task<object> ReadPublicationGeneralAsync(Request source, Action<bool> authorize, Action context)
        {
            context();
            if (!macroInFlight || !ReferenceEquals(macroOwner, this) || macroQuarantined || macroOwnerQuarantined ||
                generalInFlight || generalQuarantined || bridgeOperationsInFlight != 0 || source == null ||
                source.Command != "read_project_general" || source.ExpectedMode != 2 ||
                string.IsNullOrWhiteSpace(source.Project) || string.IsNullOrWhiteSpace(source.ExpectedProjectVersion))
                throw new InvalidOperationException("Only the original publication may inspect its exact source General page.");
            var request = new Request { Command = "read_project_general", Project = source.Project,
                ExpectedMode = 2, ExpectedProjectVersion = source.ExpectedProjectVersion };
            bool opened = false;
            Action<bool> scoped = live => {
                context();
                if (live) generalAuthorizationDepth++;
                try { authorize(live); }
                finally { if (live) generalAuthorizationDepth--; }
                context();
            };
            Action<VbeProjectGeneralOperation.Result> journal = result => {
                context();
                opened |= result.OpenAttempts != 0 && !result.Terminal;
                LoadLog.AppendText(LoadLog.PathName, DateTime.UtcNow.ToString("o") +
                    " Publication General claim Open=" + result.OpenAttempts + " Field=" + result.FieldAttempts +
                    " OK=" + result.OkAttempts + " Cancel=" + result.CancelAttempts +
                    " Terminal=" + result.Terminal + " Uncertain=" + result.Uncertain +
                    " Closed=" + result.DialogClosed + " ExecuteReturned=" + result.OriginalExecuteReturned + Environment.NewLine);
            };
            generalInFlight = true;
            try
            {
                scoped(true);
                var originalSelection = components.CapturePublicationGeneralSelection();
                components.SelectPublicationGeneralProject(request);
                scoped(true);
                request.ControlCaption = debugger.ReadGeneralCommandCaption();
                scoped(true);
                request.RevalidateProjectPropertyAuthorization = scoped;
                var result = (VbeProjectGeneralOperation.Result)await components.ProjectGeneralAsync(
                    request, false, debugger.CaptureGeneralCommand, journal, context);
                context();
                if (result.Uncertain || !result.Terminal || !result.OriginalExecuteReturned || !result.DialogClosed)
                    generalQuarantined = macroQuarantined = true;
                else
                {
                    scoped(true);
                    components.RestorePublicationGeneralSelection(originalSelection);
                    scoped(true);
                }
                return result;
            }
            catch
            {
                if (opened) generalQuarantined = macroQuarantined = true;
                throw;
            }
            finally { generalInFlight = false; }
        }
    }
}

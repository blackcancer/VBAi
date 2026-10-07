using System;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace VBAi
{

    /// <summary>Connects the Git dialog's UI state to its one-shot revision-bound modal import handoff.</summary>
    internal sealed partial class GitWindow
    {

        /// <summary>Modal runner proving that the actual owner-bound ShowDialog call returned before import admission.</summary>
        private GitModalSession modalSession;

        /// <summary>Current import request exposed in the read-only accessibility correlation string.</summary>
        private GitModalSession.Request modalRequest;

        /// <summary>Failure retained while an operation is being presented or its modal handoff completes.</summary>
        private Exception operationFailure;

        /// <summary>Set only while requesting the modal callback to return for a queued import.</summary>
        private bool leavingForImport;

        /// <summary>Releases operation-owned cache/cancellation state and hides progress controls, preserving cleanup failures.</summary>
        private void FinishOperation()
        {
            Exception cleanup = null;
            var held = cacheLock; cacheLock = null;
            try { held?.Dispose(); } catch (Exception error) { cleanup = error; }
            if (repository != null) { repository.Cancellation = System.Threading.CancellationToken.None; repository.Progress = null; }
            var cancellation = operationCancellation; operationCancellation = null;
            try { cancellation?.Dispose(); }
            catch (Exception error) { cleanup = cleanup == null ? error : new AggregateException(cleanup, error); }
            running = false;
            try { operationProgress.Visible = false; cancelOperation.Visible = false; cancelOperation.Enabled = false; UpdateButtons(); }
            catch (Exception error) { cleanup = cleanup == null ? error : new AggregateException(cleanup, error); }
            if (cleanup != null) System.Runtime.ExceptionServices.ExceptionDispatchInfo.Capture(cleanup).Throw();
        }

        /// <summary>Attaches one session and publishes its initial accessibility correlation state.</summary>
        /// <param name="session">Owner-bound session that controls modal return and import admission.</param>
        internal void AttachModalSession(GitModalSession session)
        {
            if (modalSession != null || session == null) throw new InvalidOperationException("One Git modal session only.");
            modalSession = session;
            try { PublishHandoffState(); }
            catch { modalSession = null; throw; }
        }

        /// <summary>Detaches the exact completed session only after no import remains active.</summary>
        /// <param name="session">Session being detached; identity must match the current attachment.</param>
        internal void DetachModalSession(GitModalSession session)
        {
            if (!ReferenceEquals(modalSession, session) || running) throw new InvalidOperationException("A pending Git operation cannot lose its modal owner.");
            modalSession = null;
        }

        /// <summary>Publishes non-persisted UIA correlation for the session, request, action, and request phase.</summary>
        private void PublishHandoffState()
        {
            // Read-only UIA correlation, not a native command or a persisted Git format.
            AccessibleDescription = "VBAi.GitSession/1/" + modalSession.Id + "/" +
                (modalRequest == null ? "none/none/Idle" : modalRequest.Id + "/" + modalRequest.Action + "/" + modalRequest.Phase);
        }

        /// <summary>Queues an import, closes the live modal, and waits for its returning stack to validate and admit the request.</summary>
        /// <param name="request">Prepared request bound to the current repository revision.</param>
        /// <returns>A task completed once admission succeeds or refusal is reported.</returns>
        private async Task AdmitImport(GitModalSession.Request request)
        {
            modalRequest = request;
            Exception failure = null;
            try
            {
                await modalSession.Queue(request, () =>
                {
                    leavingForImport = true; PublishHandoffState(); DialogResult = DialogResult.OK;
                });
            }
            catch (Exception error) { failure = error; throw; }
            finally
            {
                leavingForImport = false;
                try { PublishHandoffState(); }
                catch (Exception publication)
                {
                    if (failure != null) throw new AggregateException(failure, publication);
                    throw;
                }
            }
        }
    }
}

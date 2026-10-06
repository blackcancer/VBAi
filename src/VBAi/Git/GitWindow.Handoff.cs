using System;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace VBAi
{

    /// <summary>Owns the git window state and operations.</summary>
    internal sealed partial class GitWindow
    {

        /// <summary>Maintains the modal session state for git window.</summary>
        private GitModalSession modalSession;

        /// <summary>Maintains the modal request state for git window.</summary>
        private GitModalSession.Request modalRequest;

        /// <summary>Maintains the operation failure state for git window.</summary>
        private Exception operationFailure;

        /// <summary>Maintains the leaving for import state for git window.</summary>
        private bool leavingForImport;

        /// <summary>Handles finish operation for git window.</summary>
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

        /// <summary>Handles attach modal session for git window.</summary>
        /// <param name="session">git modal session that supplies the session for this operation.</param>
        internal void AttachModalSession(GitModalSession session)
        {
            if (modalSession != null || session == null) throw new InvalidOperationException("One Git modal session only.");
            modalSession = session;
            try { PublishHandoffState(); }
            catch { modalSession = null; throw; }
        }

        /// <summary>Handles detach modal session for git window.</summary>
        /// <param name="session">git modal session that supplies the session for this operation.</param>
        internal void DetachModalSession(GitModalSession session)
        {
            if (!ReferenceEquals(modalSession, session) || running) throw new InvalidOperationException("A pending Git operation cannot lose its modal owner.");
            modalSession = null;
        }

        /// <summary>Handles publish handoff state for git window.</summary>
        private void PublishHandoffState()
        {
            // Read-only UIA correlation, not a native command or a persisted Git format.
            AccessibleDescription = "VBAi.GitSession/1/" + modalSession.Id + "/" +
                (modalRequest == null ? "none/none/Idle" : modalRequest.Id + "/" + modalRequest.Action + "/" + modalRequest.Phase);
        }

        /// <summary>Handles admit import for git window.</summary>
        /// <param name="request">request that supplies the request for this operation.</param>
        /// <returns>task produced by the operation for admit import on git window.</returns>
        private async Task AdmitImport(GitModalSession.Request request)
        {
            modalRequest = request;
            Exception failure = null;
            try
            {
                await modalSession.Queue(request, () => {
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

using System;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace VBAi
{
    internal sealed partial class GitWindow
    {
        private GitModalSession modalSession;
        private GitModalSession.Request modalRequest;
        private Exception operationFailure;
        private bool leavingForImport;
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
        internal void AttachModalSession(GitModalSession session)
        {
            if (modalSession != null || session == null) throw new InvalidOperationException("One Git modal session only.");
            modalSession = session;
            try { PublishHandoffState(); }
            catch { modalSession = null; throw; }
        }
        internal void DetachModalSession(GitModalSession session)
        {
            if (!ReferenceEquals(modalSession, session) || running) throw new InvalidOperationException("A pending Git operation cannot lose its modal owner.");
            modalSession = null;
        }
        private void PublishHandoffState()
        {
            // Read-only UIA correlation, not a native command or a persisted Git format.
            AccessibleDescription = "VBAi.GitSession/1/" + modalSession.Id + "/" +
                (modalRequest == null ? "none/none/Idle" : modalRequest.Id + "/" + modalRequest.Action + "/" + modalRequest.Phase);
        }
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

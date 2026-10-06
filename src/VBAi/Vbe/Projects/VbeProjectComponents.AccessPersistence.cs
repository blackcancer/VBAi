using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Runtime.InteropServices;
using System.Threading;
using System.Threading.Tasks;

namespace VBAi
{

    /// <summary>Owns the vbe project components state and operations.</summary>
    internal sealed partial class VbeProjectComponents
    {

        /// <summary>Bounds read-only verification after Access returns from its single native Save command.</summary>
        internal TimeSpan AccessSaveVerificationTimeout = TimeSpan.FromSeconds(3);

        /// <summary>Injects the dialog boundary in managed tests; production binds the exact native VBE owner.</summary>
        internal Func<IntPtr, int, IEnumerable<AccessSaveApprovedComponent>, IAccessSaveConfirmation> AccessSaveConfirmationFactory = null;
        // Bridge, editor and chat share the same native VBE owner thread, across service instances.
        /// <summary>Maintains the access save pending state for vbe project components.</summary>
        [ThreadStatic] private static bool accessSavePending;

        /// <summary>Invokes Access Save once and observes delayed completion on its originating VBE STA.</summary>
        /// <param name="request">request that supplies the request for this operation.</param>
        /// <param name="native">i other host probe that supplies the native for this operation.</param>
        /// <returns>task&lt;object&gt; produced by the operation for save access document async on vbe project components.</returns>
        internal Task<object> SaveAccessDocumentAsync(Request request, IOtherHostProbe native)
        {
            return VbeUiTask.Run(() => SaveAccessDocumentCoreAsync(request, native));
        }

        /// <summary>Keeps every approved identity and revision guard while yielding only for read-only observations.</summary>
        /// <param name="request">request that supplies the request for this operation.</param>
        /// <param name="native">i other host probe that supplies the native for this operation.</param>
        /// <returns>task&lt;object&gt; produced by the operation for save access document core async on vbe project components.</returns>
        private async Task<object> SaveAccessDocumentCoreAsync(Request request, IOtherHostProbe native)
        {
            if (native == null || native.HostKind != "Access") throw new InvalidOperationException("Deferred Access saving requires the current Access host.");
            if (request == null || string.IsNullOrWhiteSpace(request.ExpectedProjectVersion)) throw new ArgumentException("ExpectedProjectVersion is required.");
            if (accessSavePending) throw new InvalidOperationException("An Access save is already awaiting verification; no second save was invoked.");
            if (AccessSaveVerificationTimeout <= TimeSpan.Zero || AccessSaveVerificationTimeout > TimeSpan.FromSeconds(30))
                throw new InvalidOperationException("Access saved-state verification requires a positive timeout of at most 30 seconds.");
            int ownerThread = Thread.CurrentThread.ManagedThreadId;
            object projectObject = GetDesignProject(request.Project);
            dynamic project = projectObject;
            AssertProjectVersion(request, project);
            object application = native.Application(), document = null;
            try
            {
                document = MatchOtherHost(projectObject, native, application);
                string path = RequireAbsolutePath(request.ExpectedHostPath);
                int format = OtherHostFormat("Access", path), processId = native.CurrentProcessId;
                object pane = vbe.ActiveCodePane;
                if (pane == null) throw new InvalidOperationException("Access Save requires an active code pane in the approved project.");
                object component = ((dynamic)pane).CodeModule.Parent;
                RequireAccessSaveSelection(projectObject, pane, component, native);
                if (native is NativeOtherHostProbe nativeProbe) nativeProbe.BindAccessSaveSelection(pane, component);
                string sourceSha = OtherHostSourceSha(projectObject);
                // This existing native-host revision digest excludes only the Saved notification.
                string revision = SolidWorksSaveRevision(request.Project);
                var approvedComponents = new List<AccessSaveApprovedComponent>();
                foreach (object current in project.VBComponents)
                {
                    int type = (int)((dynamic)current).Type;
                    if (type == 1 || type == 2)
                        approvedComponents.Add(new AccessSaveApprovedComponent((string)((dynamic)current).Name, type));
                }
                IAccessSaveConfirmation confirmation = null;
                if (AccessSaveConfirmationFactory != null || native is NativeOtherHostProbe)
                {
                    var window = new IntPtr(Convert.ToInt64(vbe.MainWindow.HWnd));
                    confirmation = AccessSaveConfirmationFactory != null
                        ? AccessSaveConfirmationFactory(window, processId, approvedComponents)
                        : CreateNativeAccessSaveConfirmation(window, processId, approvedComponents);
                    if (confirmation == null) throw new InvalidOperationException("Access Save confirmation guard is unavailable.");
                    confirmation.Prepare();
                    if (native is NativeOtherHostProbe guardedProbe) guardedProbe.AccessBeforeSave = () => {
                        if (Thread.CurrentThread.ManagedThreadId != ownerThread || native.HostKind != "Access" ||
                            native.CurrentProcessId != processId || native.ApplicationProcessId(application) != (uint)processId ||
                            (int)project.Mode != 2 || (int)project.Protection != 0 ||
                            !native.SameProject(projectObject, (object)GetDesignProject(request.Project)))
                            throw new InvalidOperationException("Access Save lost its approved owner or project before invocation.");
                        AssertProjectVersion(request, project);
                        RequireAccessSaveSelection(projectObject, pane, component, native);
                        if (OtherHostSourceSha(projectObject) != sourceSha || SolidWorksSaveRevision(request.Project) != revision)
                            throw new InvalidOperationException("Access Save source, metadata or references changed before invocation.");
                        var currentState = native.State(document);
                        if (currentState.ReadOnly || !OtherHostSamePath(currentState.Path, path) || currentState.Format != format ||
                            !OtherHostProjectPathMatches(projectObject, path, "Access"))
                            throw new InvalidOperationException("Access Save path or writable state changed before invocation.");
                        request.RevalidateSaveAuthorization?.Invoke();
                        confirmation.RequireBeforeSave();
                    };
                }
                accessSavePending = true;
                try
                {
                    // Reuse the exact existing preflight/invocation boundary. Its provisional Saved
                    // failure is internal here; errors for any other guard remain uncertain unchanged.
                    dynamic initial = SaveOtherHost(request, false, native, application);
                    if (!(bool)initial.Verified && (string)initial.Reason != "Save verification failed: ProjectSaved.") return initial;
                    var elapsed = Stopwatch.StartNew();
                    try
                    {
                        while (true)
                        {
                            await Task.Delay(25);
                            if (Thread.CurrentThread.ManagedThreadId != ownerThread)
                                throw new InvalidOperationException("Access Save verification left the owning VBE thread.");
                            if (native.HostKind != "Access" || native.CurrentProcessId != processId ||
                                native.ApplicationProcessId(application) != (uint)processId)
                                throw new InvalidOperationException("Access Save verification failed: host process changed.");
                            if ((int)project.Mode != 2 || (int)project.Protection != 0 ||
                                !native.SameProject(projectObject, (object)GetDesignProject(request.Project)))
                                throw new InvalidOperationException("Access Save verification failed: project identity, mode or protection changed.");
                            RequireAccessSaveSelection(projectObject, pane, component, native);
                            if (OtherHostSourceSha(projectObject) != sourceSha || SolidWorksSaveRevision(request.Project) != revision)
                                throw new InvalidOperationException("Access Save verification failed: source, metadata or references changed.");
                            object observedDocument = null;
                            try
                            {
                                // Access returns a new CurrentProject COM object for successive getters.
                                // Identity is the approved application/PID, full database path and unique
                                // retained VBProject; read the freshly matched wrapper rather than a stale one.
                                observedDocument = MatchOtherHost(projectObject, native, application);
                                var state = native.State(observedDocument);
                                if (state.ReadOnly || !OtherHostSamePath(state.Path, path) ||
                                    !OtherHostProjectPathMatches(projectObject, path, "Access") || state.Format != format)
                                    throw new InvalidOperationException("Access Save verification failed: path, format or writable state changed.");
                                if (!native.FileExists(path) || native.FileLength(path) < 1)
                                    throw new InvalidOperationException("Access Save verification failed: database file is absent or empty.");
                                var candidate = confirmation?.Observe();
                                if (candidate != null)
                                {
                                    if (elapsed.Elapsed >= AccessSaveVerificationTimeout)
                                        throw new InvalidOperationException("Access Save confirmation remains pending; no confirmation or Save is retried.");
                                    if (confirmation.ConfirmationAttempts == 0)
                                        confirmation.Confirm(candidate, () => {
                                            if (Thread.CurrentThread.ManagedThreadId != ownerThread ||
                                                native.HostKind != "Access" || native.CurrentProcessId != processId ||
                                                native.ApplicationProcessId(application) != (uint)processId ||
                                                (int)project.Mode != 2 || (int)project.Protection != 0 ||
                                                !native.SameProject(projectObject, (object)GetDesignProject(request.Project)))
                                                throw new InvalidOperationException("Access Save confirmation lost its approved owner or project.");
                                            RequireAccessSaveSelection(projectObject, pane, component, native);
                                            if (OtherHostSourceSha(projectObject) != sourceSha || SolidWorksSaveRevision(request.Project) != revision)
                                                throw new InvalidOperationException("Access Save confirmation source, metadata or references changed.");
                                            object confirmedDocument = null;
                                            try
                                            {
                                                confirmedDocument = MatchOtherHost(projectObject, native, application);
                                                var confirmedState = native.State(confirmedDocument);
                                                if (confirmedState.ReadOnly || !OtherHostSamePath(confirmedState.Path, path) ||
                                                    !OtherHostProjectPathMatches(projectObject, path, "Access") || confirmedState.Format != format ||
                                                    !native.FileExists(path) || native.FileLength(path) < 1 ||
                                                    elapsed.Elapsed >= AccessSaveVerificationTimeout)
                                                    throw new InvalidOperationException("Access Save confirmation path, writable state or deadline changed.");
                                                request.RevalidateSaveAuthorization?.Invoke();
                                            }
                                            finally { ReleaseAccessObservation(confirmedDocument); }
                                        }, () => {
                                            if (elapsed.Elapsed >= AccessSaveVerificationTimeout)
                                                throw new InvalidOperationException("Access Save confirmation delivery deadline expired; no native confirmation was queued.");
                                        });
                                    // Enqueue is not persistence. Yield for native processing and verify all guards again.
                                    continue;
                                }
                                if (!(bool)project.Saved)
                                {
                                    if (elapsed.Elapsed < AccessSaveVerificationTimeout) continue;
                                    throw new InvalidOperationException("Access Save verification timed out: ProjectSaved remains false.");
                                }
                                return new { Project = request.Project, Host = "Access", HostPath = path, SaveApi = "VBE.CommandBars.ID3",
                                    SaveInvoked = true, SaveAsInvoked = false, MutationInvoked = true, Verified = true, Uncertain = false,
                                    ProjectSaved = true, HostSaved = (bool?)null, Bytes = native.FileLength(path), OwnerProcessId = processId,
                                    SourceSha256 = sourceSha, CodePreserved = true, NativeFileFormatVerified = true,
                                    NativeQualification = "NOT_RUN", PersistenceReopenVerified = false,
                                    Verification = "DeferredOwnerThreadSavedReadback", VerificationMilliseconds = elapsed.ElapsedMilliseconds,
                                    ConfirmationAttempts = confirmation?.ConfirmationAttempts ?? 0,
                                    ConfirmationQueued = confirmation?.ConfirmationQueued ?? false,
                                    ConfirmationPending = confirmation?.ConfirmationPending ?? false,
                                    Limit = "Access saved-state, identity, path, metadata and unchanged live source were verified after one native Save. Access has no document Saved flag; reopen the database to verify persisted content." };
                            }
                            finally { ReleaseAccessObservation(observedDocument); }
                        }
                    }
                    catch (Exception error)
                    {
                        return new { Project = request.Project, Host = "Access", HostPath = path, SaveApi = "VBE.CommandBars.ID3",
                            SaveInvoked = true, SaveAsInvoked = false, MutationInvoked = true, Verified = false, Uncertain = true,
                            NativeQualification = "NOT_RUN", Reason = error.Message,
                            ConfirmationAttempts = confirmation?.ConfirmationAttempts ?? 0,
                            ConfirmationQueued = confirmation?.ConfirmationQueued ?? false,
                            ConfirmationPending = confirmation?.ConfirmationPending ?? false,
                            Next = "Inspect project_persistence_status and the database; do not retry automatically." };
                    }
                }
                finally
                {
                    accessSavePending = false;
                    if (native is NativeOtherHostProbe guardedProbe) guardedProbe.AccessBeforeSave = null;
                }
            }
            finally
            {
                ReleaseAccessObservation(document);
                ReleaseAccessObservation(application);
            }
        }

        /// <summary>Balances one native getter acquisition without invalidating a shared RCW through FinalRelease.</summary>
        /// <param name="value">object that supplies the value for this operation.</param>
        private static void ReleaseAccessObservation(object value)
        {
            if (value != null && Marshal.IsComObject(value)) Marshal.ReleaseComObject(value);
        }

        /// <summary>Reads exact selection and component ownership without changing focus or native selection.</summary>
        /// <param name="project">object that supplies the project for this operation.</param>
        /// <param name="pane">object that supplies the pane for this operation.</param>
        /// <param name="component">object that supplies the component for this operation.</param>
        /// <param name="native">i other host probe that supplies the native for this operation.</param>
        private void RequireAccessSaveSelection(object project, object pane, object component, IOtherHostProbe native)
        {
            object activeProject = vbe.ActiveVBProject, activePane = vbe.ActiveCodePane;
            if (activeProject == null || activePane == null || !native.SameProject(project, activeProject) ||
                !native.SameProject(pane, activePane) || !native.SameProject(component, (object)((dynamic)activePane).CodeModule.Parent))
                throw new InvalidOperationException("Access Save verification failed: native selection changed.");
            foreach (object current in ((dynamic)project).VBComponents)
                if (native.SameProject(component, current)) return;
            throw new InvalidOperationException("The active Access component no longer belongs to the approved project.");
        }
    }
}

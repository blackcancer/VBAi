using System;
using System.Diagnostics;
using System.Threading;
using System.Threading.Tasks;

namespace VBAi
{
    internal sealed partial class VbeProjectComponents
    {
        /// <summary>Bounds read-only verification after Access returns from its single native Save command.</summary>
        internal TimeSpan AccessSaveVerificationTimeout = TimeSpan.FromSeconds(3);
        // Bridge, editor and chat share the same native VBE owner thread, across service instances.
        [ThreadStatic] private static bool accessSavePending;

        /// <summary>Invokes Access Save once and observes delayed completion on its originating VBE STA.</summary>
        internal Task<object> SaveAccessDocumentAsync(Request request, IOtherHostProbe native)
        {
            return VbeUiTask.Run(() => SaveAccessDocumentCoreAsync(request, native));
        }

        /// <summary>Keeps every approved identity and revision guard while yielding only for read-only observations.</summary>
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
            object document = MatchOtherHost(projectObject, native);
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
            accessSavePending = true;
            try
            {
                // Reuse the exact existing preflight/invocation boundary. Its provisional Saved
                // failure is internal here; errors for any other guard remain uncertain unchanged.
                dynamic initial = SaveOtherHost(request, false, native);
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
                            native.ApplicationProcessId(native.Application()) != (uint)processId)
                            throw new InvalidOperationException("Access Save verification failed: host process changed.");
                        if ((int)project.Mode != 2 || (int)project.Protection != 0 ||
                            !native.SameProject(projectObject, (object)GetDesignProject(request.Project)))
                            throw new InvalidOperationException("Access Save verification failed: project identity, mode or protection changed.");
                        RequireAccessSaveSelection(projectObject, pane, component, native);
                        if (OtherHostSourceSha(projectObject) != sourceSha || SolidWorksSaveRevision(request.Project) != revision)
                            throw new InvalidOperationException("Access Save verification failed: source, metadata or references changed.");
                        if (!native.SameProject(document, MatchOtherHost(projectObject, native)))
                            throw new InvalidOperationException("Access Save verification failed: document identity changed.");
                        var state = native.State(document);
                        if (state.ReadOnly || !OtherHostSamePath(state.Path, path) ||
                            !OtherHostProjectPathMatches(projectObject, path, "Access") || state.Format != format)
                            throw new InvalidOperationException("Access Save verification failed: path, format or writable state changed.");
                        if (!native.FileExists(path) || native.FileLength(path) < 1)
                            throw new InvalidOperationException("Access Save verification failed: database file is absent or empty.");
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
                            Limit = "Access saved-state, identity, path, metadata and unchanged live source were verified after one native Save. Access has no document Saved flag; reopen the database to verify persisted content." };
                    }
                }
                catch (Exception error)
                {
                    return new { Project = request.Project, Host = "Access", HostPath = path, SaveApi = "VBE.CommandBars.ID3",
                        SaveInvoked = true, SaveAsInvoked = false, MutationInvoked = true, Verified = false, Uncertain = true,
                        NativeQualification = "NOT_RUN", Reason = error.Message,
                        Next = "Inspect project_persistence_status and the database; do not retry automatically." };
                }
            }
            finally { accessSavePending = false; }
        }

        /// <summary>Reads exact selection and component ownership without changing focus or native selection.</summary>
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

using System;
using System.IO;
using System.Linq;

namespace VBAi
{

    /// <summary>Creates a separate PPTM coverage presentation without saving or renaming its original.</summary>
    internal sealed class VbaTestPowerPointCoverageClone
    {

        /// <summary>PowerPoint SaveAs format identifier for a macro-enabled Open XML presentation.</summary>
        internal const int MacroEnabledPresentationFormat = 25; // ppSaveAsOpenXMLPresentationMacroEnabled

        /// <summary>Resolves and validates the owning PowerPoint process, presentation, and VBA project.</summary>
        internal VbaTestPowerPointValuesHost Host = new VbaTestPowerPointValuesHost();

        /// <summary>Saves a separate copy in PPTM format without renaming or saving the original presentation.</summary>
        internal Action<object, string> SaveCopy = (presentation, path) => ((dynamic)presentation).SaveCopyAs(path, MacroEnabledPresentationFormat, -2);

        /// <summary>Opens the copied presentation hidden and without a new window.</summary>
        internal Func<object, string, object> OpenCopy;

        /// <summary>Closes the verified owned copy after its instrumentation is marked discardable.</summary>
        internal Action<object> CloseCopy = presentation => ((dynamic)presentation).Close();

        /// <summary>Installs the native copy opener, which owns its Presentations getter acquisition.</summary>
        internal VbaTestPowerPointCoverageClone() { OpenCopy = OpenNativeCopy; }

        /// <summary>Opens the copy once and balances the collection getter without releasing the returned presentation.</summary>
        /// <param name="application">Borrowed, verified application belonging to the original process.</param>
        /// <param name="path">Exact disposable copy path.</param>
        /// <returns>New presentation acquisition transferred to the caller's scope.</returns>
        private object OpenNativeCopy(object application, string path)
        {
            var references = new VbaTestComReferences(Host.ReleaseAcquired);
            var returned = new VbaTestComReferences(Host.ReleaseAcquired);
            bool entered = false;
            try
            {
                dynamic presentations = references.Add((object)((dynamic)application).Presentations);
                entered = true;
                object copy = returned.Add((object)presentations.Open(path, 0, 0, 0));
                references.Dispose();
                if (copy != null) returned.Transfer(copy);
                return copy;
            }
            catch
            {
                if (entered) { references.Retain(); returned.Retain(); }
                throw;
            }
            finally
            {
                try { references.Dispose(); }
                finally { returned.Dispose(); }
            }
        }

        /// <summary>Creates a disposable PPTM coverage copy from a saved macro-capable PowerPoint source.</summary>
        /// <param name="project">Source VBProject to resolve to its owning presentation.</param>
        /// <param name="sourcePath">Expected saved source presentation path.</param>
        /// <param name="folder">Fresh directory for the PPTM copy.</param>
        /// <returns>Owned cloned project and a close action that discards only clone changes.</returns>
        internal static VbaTestCoverageClone CreatePowerPoint(object project, string sourcePath, string folder)
        { return new VbaTestPowerPointCoverageClone().Create(project, sourcePath, folder); }

        /// <summary>Creates and verifies a distinct macro-enabled presentation copy before returning its project for instrumentation.</summary>
        /// <param name="project">Source VBProject whose presentation identity must remain unchanged.</param>
        /// <param name="sourcePath">Saved source path with a supported macro-capable extension.</param>
        /// <param name="folder">Fresh directory for the retained coverage.pptm copy.</param>
        /// <returns>Owned copy with post-close absence verification.</returns>
        /// <exception cref="InvalidOperationException">The source, extension, destination, or copy identity fails validation.</exception>
        /// <exception cref="VbaTestInvocationException">Save/open/close completion is uncertain; the retained path is reported.</exception>
        internal VbaTestCoverageClone Create(object project, string sourcePath, string folder)
        {
            var source = (VbaTestPowerPointValuesHost.OwnedTarget)Host.ResolveTarget(project, sourcePath);
            var references = new VbaTestComReferences(Host.ReleaseAcquired);
            VbaTestPowerPointValuesHost.OwnedTarget ownedCopy = null;
            bool nativeEntered = false, transferred = false;
            void retain()
            {
                references.Retain();
                source.RetainOnUncertain();
                ownedCopy?.RetainOnUncertain();
            }
            try
            {
                string extension = Path.GetExtension(sourcePath);
                if (!new[] { ".pptm", ".ppsm", ".potm", ".ppt", ".pps", ".pot" }.Contains(extension, StringComparer.OrdinalIgnoreCase))
                    throw new InvalidOperationException("PowerPoint coverage requires a saved macro-capable presentation, show or template; add-ins are excluded.");
                VbaTestPowerPointValuesHost.RequireAbsolutePath(folder);
                string copyPath = Path.Combine(Path.GetFullPath(folder), "coverage.pptm");
                if (VbaTestPowerPointValuesHost.SamePath(copyPath, sourcePath) || File.Exists(copyPath))
                    throw new InvalidOperationException("The owned PPTM copy path is already occupied.");
                foreach (dynamic presentation in references.Items((object)((dynamic)source.Application).Presentations))
                    if (VbaTestPowerPointValuesHost.SamePath((string)presentation.FullName, copyPath))
                        throw new InvalidOperationException("The PPTM copy path already belongs to an open presentation.");
                Directory.CreateDirectory(folder);
                object copy;
                Host.ValidateTarget(source);
                nativeEntered = true;
                try { SaveCopy(source.Presentation, copyPath); }
                catch (Exception error) { throw Uncertain("Saving the PPTM copy", copyPath, error); }
                Host.ValidateTarget(source);
                try { copy = references.Add(OpenCopy(source.Application, copyPath)); }
                catch (Exception error) { throw Uncertain("Opening the PPTM copy", copyPath, error); }
                // No trust settings are changed. PresentationOpen/BeforeClose handlers need native qualification.
                try
                {
                    object copyProject = copy == null ? null : references.Add((object)((dynamic)copy).VBProject);
                    if (copy == null || Host.SameIdentity(copy, source.Presentation)
                        || Host.SameIdentity(copyProject, source.Project)
                        || !VbaTestPowerPointValuesHost.SamePath((string)((dynamic)copy).FullName, copyPath))
                        throw new InvalidOperationException("The opened coverage presentation is not a verified distinct owned copy. It was not closed.");
                    ownedCopy = (VbaTestPowerPointValuesHost.OwnedTarget)Host.ResolveTarget(copyProject, copyPath);
                    if (!Host.SameIdentity(ownedCopy.Presentation, copy))
                        throw new InvalidOperationException("The opened coverage copy identity could not be verified. It was not closed.");
                    Host.ValidateTarget(source);
                    var clone = new VbaTestCoverageClone
                    {
                        Project = copyProject,
                        Path = copyPath,
                        Close = () =>
                        {
                            bool closed = false;
                            try
                            {
                                CloseVerifiedCopy(source, ownedCopy, copyPath);
                                closed = true;
                            }
                            finally
                            {
                                if (!closed) retain();
                                else
                                {
                                    try { ownedCopy.Dispose(); source.Dispose(); references.Dispose(); }
                                    catch (Exception error)
                                    {
                                        retain();
                                        throw Uncertain("Releasing the verified PPTM copy references", copyPath, error);
                                    }
                                }
                            }
                        }
                    };
                    transferred = true;
                    return clone;
                }
                catch (Exception error)
                { throw Uncertain("Verifying ownership after opening the PPTM copy", copyPath, error); }
            }
            finally
            {
                if (!transferred)
                {
                    if (nativeEntered) retain();
                    else
                    {
                        try { references.Dispose(); }
                        finally { source.Dispose(); }
                    }
                }
            }
        }

        /// <summary>Revalidates, discards and closes only the copy, then verifies its absence.</summary>
        /// <param name="source">Original target, which must never be closed or marked saved.</param>
        /// <param name="copy">Verified disposable target whose acquisitions remain owned by the caller.</param>
        /// <param name="copyPath">Path retained if native close completion is uncertain.</param>
        private void CloseVerifiedCopy(VbaTestPowerPointValuesHost.OwnedTarget source,
            VbaTestPowerPointValuesHost.OwnedTarget copy, string copyPath)
        {
            var current = Host.ValidateTarget(copy);
            RequireDistinctCopy(source, current);
            try
            {
                // Saved=msoTrue discards only this copy's instrumentation without writing it to disk.
                ((dynamic)current.Presentation).Saved = -1;
                if (Convert.ToInt32(((dynamic)current.Presentation).Saved) != -1)
                    throw new InvalidOperationException("The owned copy did not retain the discard-changes state; Close was not attempted.");
                current = Host.ValidateTarget(copy);
                RequireDistinctCopy(source, current);
                CloseCopy(current.Presentation);
                using (var remainingReferences = new VbaTestComReferences(Host.ReleaseAcquired))
                    foreach (dynamic remaining in remainingReferences.Items((object)((dynamic)current.Application).Presentations))
                        if (Host.SameIdentity((object)remaining, current.Presentation))
                            throw new InvalidOperationException("The owned presentation remains open after Close returned.");
            }
            catch (Exception error) { throw Uncertain("Closing the owned PPTM copy", copyPath, error); }
        }

        /// <summary>Rejects an alias of either original identity before touching the copy's saved state or closing it.</summary>
        /// <param name="source">Borrowed original target.</param>
        /// <param name="copy">Borrowed copy target revalidated on the owner thread.</param>
        private void RequireDistinctCopy(VbaTestPowerPointValuesHost.OwnedTarget source, VbaTestPowerPointValuesHost.OwnedTarget copy)
        {
            if (Host.SameIdentity(copy.Project, source.Project) || Host.SameIdentity(copy.Presentation, source.Presentation))
                throw new InvalidOperationException("Closing the original presentation is forbidden.");
        }

        /// <summary>Wraps unverified native completion as uncertain, preserving the copy for inspection and forbidding replay.</summary>
        /// <param name="phase">Save, open, or close operation that did not return verified completion.</param>
        /// <param name="retainedPath">Disposable file path retained for recovery.</param>
        /// <param name="error">Underlying failure.</param>
        /// <returns>Invocation exception marked as outcome-unknown.</returns>
        private static VbaTestInvocationException Uncertain(string phase, string retainedPath, Exception error)
        { return new VbaTestInvocationException(phase + " did not return verified completion; no retry was attempted. Retained copy: " + retainedPath + ". " + error.Message, true, error); }
    }
}

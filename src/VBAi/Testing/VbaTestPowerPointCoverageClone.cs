using System;
using System.IO;
using System.Linq;

namespace VBAi
{

    /// <summary>Creates a separate PPTM coverage presentation without saving or renaming its original.</summary>
    internal sealed class VbaTestPowerPointCoverageClone
    {

        /// <summary>Maintains the macro enabled presentation format state for vba test power point coverage clone.</summary>
        internal const int MacroEnabledPresentationFormat = 25; // ppSaveAsOpenXMLPresentationMacroEnabled

        /// <summary>Maintains the host state for vba test power point coverage clone.</summary>
        internal VbaTestPowerPointValuesHost Host = new VbaTestPowerPointValuesHost();

        /// <summary>Maintains the save copy state for vba test power point coverage clone.</summary>
        internal Action<object, string> SaveCopy = (presentation, path) => ((dynamic)presentation).SaveCopyAs(path, MacroEnabledPresentationFormat, -2);

        /// <summary>Maintains the open copy state for vba test power point coverage clone.</summary>
        internal Func<object, string, object> OpenCopy = (application, path) => ((dynamic)application).Presentations.Open(path, 0, 0, 0);

        /// <summary>Maintains the close copy state for vba test power point coverage clone.</summary>
        internal Action<object> CloseCopy = presentation => ((dynamic)presentation).Close();

        /// <summary>Creates power point for vba test power point coverage clone.</summary>
        /// <param name="project">object that supplies the project for this operation.</param>
        /// <param name="sourcePath">Path used for the source path being processed.</param>
        /// <param name="folder">Text that supplies the folder value. Use the format required by the calling operation.</param>
        /// <returns>vba test coverage clone produced by the operation for create power point on vba test power point coverage clone.</returns>
        internal static VbaTestCoverageClone CreatePowerPoint(object project, string sourcePath, string folder)
        { return new VbaTestPowerPointCoverageClone().Create(project, sourcePath, folder); }

        /// <summary>Creates  for vba test power point coverage clone.</summary>
        /// <param name="project">object that supplies the project for this operation.</param>
        /// <param name="sourcePath">Path used for the source path being processed.</param>
        /// <param name="folder">Text that supplies the folder value. Use the format required by the calling operation.</param>
        /// <returns>vba test coverage clone produced by the operation for create on vba test power point coverage clone.</returns>
        internal VbaTestCoverageClone Create(object project, string sourcePath, string folder)
        {
            var source = (VbaTestPowerPointValuesHost.OwnedTarget)Host.ResolveTarget(project, sourcePath);
            string extension = Path.GetExtension(sourcePath);
            if (!new[] { ".pptm", ".ppsm", ".potm", ".ppt", ".pps", ".pot" }.Contains(extension, StringComparer.OrdinalIgnoreCase))
                throw new InvalidOperationException("PowerPoint coverage requires a saved macro-capable presentation, show or template; add-ins are excluded.");
            VbaTestPowerPointValuesHost.RequireAbsolutePath(folder);
            string copyPath = Path.Combine(Path.GetFullPath(folder), "coverage.pptm");
            if (VbaTestPowerPointValuesHost.SamePath(copyPath, sourcePath) || File.Exists(copyPath))
                throw new InvalidOperationException("The owned PPTM copy path is already occupied.");
            foreach (dynamic presentation in ((dynamic)source.Application).Presentations)
                if (VbaTestPowerPointValuesHost.SamePath((string)presentation.FullName, copyPath))
                    throw new InvalidOperationException("The PPTM copy path already belongs to an open presentation.");
            Directory.CreateDirectory(folder);
            object copy;
            Host.ValidateTarget(source);
            try { SaveCopy(source.Presentation, copyPath); }
            catch (Exception error) { throw Uncertain("Saving the PPTM copy", copyPath, error); }
            Host.ValidateTarget(source);
            try { copy = OpenCopy(source.Application, copyPath); }
            catch (Exception error) { throw Uncertain("Opening the PPTM copy", copyPath, error); }
            // No automation-security/trust settings are changed. PowerPoint has no EnableEvents
            // boundary: application-level PresentationOpen/BeforeClose handlers need native qualification.
            try
            {
                if (copy == null || Host.SameIdentity(copy, source.Presentation)
                    || Host.SameIdentity((object)((dynamic)copy).VBProject, source.Project)
                    || !VbaTestPowerPointValuesHost.SamePath((string)((dynamic)copy).FullName, copyPath))
                    throw new InvalidOperationException("The opened coverage presentation is not a verified distinct owned copy. It was not closed.");
                var ownedCopy = Host.ResolveTarget((object)((dynamic)copy).VBProject, copyPath);
                if (!Host.SameIdentity(((VbaTestPowerPointValuesHost.OwnedTarget)ownedCopy).Presentation, copy))
                    throw new InvalidOperationException("The opened coverage copy identity could not be verified. It was not closed.");
                Host.ValidateTarget(source);
                return new VbaTestCoverageClone { Project = ((dynamic)copy).VBProject, Path = copyPath, Close = () =>
                {
                    var current = Host.ValidateTarget(ownedCopy);
                    if (Host.SameIdentity(current.Project, source.Project) || Host.SameIdentity(current.Presentation, source.Presentation))
                        throw new InvalidOperationException("Closing the original presentation is forbidden.");
                    try {
                        // Saved=msoTrue explicitly discards only the verified copy's instrumentation
                        // without writing it to disk (Presentation.Saved/Close, Microsoft VBA API).
                        ((dynamic)current.Presentation).Saved = -1;
                        if (Convert.ToInt32(((dynamic)current.Presentation).Saved) != -1)
                            throw new InvalidOperationException("The owned copy did not retain the discard-changes state; Close was not attempted.");
                        current = Host.ValidateTarget(ownedCopy);
                        if (Host.SameIdentity(current.Project, source.Project) || Host.SameIdentity(current.Presentation, source.Presentation))
                            throw new InvalidOperationException("Closing the original presentation is forbidden.");
                        CloseCopy(current.Presentation);
                        foreach (dynamic remaining in ((dynamic)current.Application).Presentations)
                            if (Host.SameIdentity((object)remaining, current.Presentation))
                                throw new InvalidOperationException("The owned presentation remains open after Close returned.");
                    }
                    catch (Exception error) { throw Uncertain("Closing the owned PPTM copy", copyPath, error); }
                } };
            }
            catch (Exception error)
            { throw Uncertain("Verifying ownership after opening the PPTM copy", copyPath, error); }
        }

        /// <summary>Handles uncertain for vba test power point coverage clone.</summary>
        /// <param name="phase">Text that supplies the phase value. Use the format required by the calling operation.</param>
        /// <param name="retainedPath">Path used for the retained path being processed.</param>
        /// <param name="error">Exception describing the error failure.</param>
        /// <returns>vba test invocation exception produced by the operation for uncertain on vba test power point coverage clone.</returns>
        private static VbaTestInvocationException Uncertain(string phase, string retainedPath, Exception error)
        { return new VbaTestInvocationException(phase + " did not return verified completion; no retry was attempted. Retained copy: " + retainedPath + ". " + error.Message, true, error); }
    }
}

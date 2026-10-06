using System;

namespace VBAi
{

    /// <summary>Implements project component operations, including guarded publication workflows.</summary>
    internal sealed partial class VbeProjectComponents
    {

        /// <summary>Snapshot of the active VBE project identity and version needed for safe selection restoration.</summary>
        internal sealed class PublicationGeneralSelection
        {

            /// <summary>Retains the original active VBProject COM identity for selection restoration.</summary>
            internal object Canonical;

            /// <summary>Stores the original project's name and captured version for restoration checks.</summary>
            internal string Name, Version;
        }

        /// <summary>Captures the active project identity, name, and version before publication changes selection.</summary>
        /// <returns>State used to verify and restore the exact original active project.</returns>
        /// <exception cref="InvalidOperationException">No unambiguous active project is available.</exception>
        internal PublicationGeneralSelection CapturePublicationGeneralSelection()
        {
            object canonical = (object)vbe.ActiveVBProject;
            if (canonical == null) throw new InvalidOperationException("Original active project is unavailable.");
            string name = (string)((dynamic)canonical).Name;
            if (!GeneralProjectIdentity(canonical, (object)GetDesignProject(name)))
                throw new InvalidOperationException("Original active project identity is ambiguous.");
            return new PublicationGeneralSelection { Canonical = canonical, Name = name,
                Version = (string)((dynamic)ProjectProperties(name)).Version };
        }

        /// <summary>Restores the captured active project after confirming its identity and version are unchanged.</summary>
        /// <param name="selection">Original active project identity, name, and version returned by capture.</param>
        /// <exception cref="InvalidOperationException">The project changed or VBE did not restore the original selection.</exception>
        internal void RestorePublicationGeneralSelection(PublicationGeneralSelection selection)
        {
            if (selection == null || !GeneralProjectIdentity(selection.Canonical, (object)GetDesignProject(selection.Name)))
                throw new InvalidOperationException("Original active project changed before selection restoration.");
            var guard = new Request { Project = selection.Name, ExpectedMode = 2, ExpectedProjectVersion = selection.Version };
            AssertProjectVersion(guard, selection.Canonical);
            vbe.ActiveVBProject = (dynamic)selection.Canonical;
            if (!GeneralProjectIdentity(selection.Canonical, (object)vbe.ActiveVBProject) ||
                !GeneralProjectIdentity(selection.Canonical, (object)GetDesignProject(selection.Name)))
                throw new InvalidOperationException("Original active project selection was not restored.");
            AssertProjectVersion(guard, selection.Canonical);
        }

        /// <summary>Selects the exact unprotected standalone publication source in VBE design mode.</summary>
        /// <param name="request">Read-project-general request containing the expected project name and version.</param>
        /// <exception cref="ArgumentException">The command, mode, project name, or expected version is missing or incorrect.</exception>
        /// <exception cref="InvalidOperationException">The source is not the exact unprotected standalone project or changes during selection.</exception>
        internal void SelectPublicationGeneralProject(Request request)
        {
            if (request == null || request.Command != "read_project_general" || request.ExpectedMode != 2 ||
                string.IsNullOrWhiteSpace(request.Project) || string.IsNullOrWhiteSpace(request.ExpectedProjectVersion))
                throw new ArgumentException("An exact publication source and revision are required.");
            object original = (object)GetDesignProject(request.Project);
            dynamic project = original;
            if ((int)project.Type != 101 || (int)project.Protection != 0 || (string)project.Name != request.Project)
                throw new InvalidOperationException("Only the exact unprotected standalone publication source may be selected.");
            AssertProjectVersion(request, original);
            vbe.ActiveVBProject = project; // UI selection only; no project metadata setter.
            if (!GeneralProjectIdentity(original, (object)vbe.ActiveVBProject) ||
                !GeneralProjectIdentity(original, (object)GetDesignProject(request.Project)) ||
                (int)project.Mode != 2 || (int)project.Protection != 0 || (string)project.Name != request.Project)
                throw new InvalidOperationException("Original publication source changed during General selection.");
            AssertProjectVersion(request, original);
        }
    }
}

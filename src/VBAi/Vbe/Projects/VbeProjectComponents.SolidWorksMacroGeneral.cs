using System;

namespace VBAi
{

    /// <summary>Owns the vbe project components state and operations.</summary>
    internal sealed partial class VbeProjectComponents
    {

        /// <summary>Owns the publication general selection state and operations.</summary>
        internal sealed class PublicationGeneralSelection
        {

            /// <summary>Tracks the canonical state of publication general selection.</summary>
            internal object Canonical;

            /// <summary>Maintains the name and version state for publication general selection.</summary>
            internal string Name, Version;
        }

        /// <summary>Captures publication general selection for vbe project components.</summary>
        /// <returns>publication general selection produced by the operation for capture publication general selection on vbe project components.</returns>
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

        /// <summary>Handles restore publication general selection for vbe project components.</summary>
        /// <param name="selection">publication general selection that supplies the selection for this operation.</param>
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

        /// <summary>Handles select publication general project for vbe project components.</summary>
        /// <param name="request">request that supplies the request for this operation.</param>
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

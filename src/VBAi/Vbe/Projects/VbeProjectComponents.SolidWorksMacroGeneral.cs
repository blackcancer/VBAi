using System;

namespace VBAi
{
    internal sealed partial class VbeProjectComponents
    {
        internal sealed class PublicationGeneralSelection
        {
            internal object Canonical;
            internal string Name, Version;
        }

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

using System;
using System.Collections;
using System.Collections.Generic;
using System.ComponentModel;
using System.Linq;

namespace CodexVBE
{
    internal sealed partial class VbeForms
    {
        // Read-only preflight for a possible Frame-with-Labels copy. This does
        // not claim that creating the proposed hierarchy has been verified.
        public object FrameCopyPlan(Request request)
        {
            return BuildFrameCopyPlan(request, false);
        }

        public object FrameSimpleCopyPlan(Request request)
        {
            return BuildFrameCopyPlan(request, true);
        }

        private object BuildFrameCopyPlan(Request request, bool allowTextBox)
        {
            if (string.IsNullOrWhiteSpace(request.ControlPath) ||
                string.IsNullOrWhiteSpace(request.ExpectedTreeVersion))
                throw new ArgumentException("ControlPath and ExpectedTreeVersion are required.");
            ValidateName(request.NewName, "NewName");
            dynamic form = GetForm(GetProject(request.Project), request.Form);
            dynamic tree = Tree(request.Project, request.Form);
            if (!string.Equals((string)tree.TreeVersion, request.ExpectedTreeVersion,
                StringComparison.OrdinalIgnoreCase))
                throw new InvalidOperationException("The UserForm hierarchy changed since it was read.");
            if (!TreeContainsPath((IEnumerable)tree.Controls, request.ControlPath))
                throw new InvalidOperationException("ControlPath is not canonical in form_tree.");
            string[] parts = request.ControlPath.Split('/');
            if (parts.Length < 2 || parts.Length % 2 != 0 ||
                !string.Equals(parts[parts.Length - 2], "Controls", StringComparison.Ordinal))
                throw new ArgumentException("ControlPath must identify a Frame control.");

            object source = ResolveTreeItem(form.Designer, request.ControlPath);
            if (!string.Equals(TypeDescriptor.GetClassName(source), "Frame", StringComparison.OrdinalIgnoreCase))
                throw new InvalidOperationException("The selected control is not a native MSForms Frame.");
            dynamic original = source;
            var issues = new List<string>();
            var children = new List<object>();
            var names = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            foreach (dynamic existing in form.Designer.Controls)
                names.Add((string)existing.Name);
            if (names.Contains(request.NewName))
                issues.Add("New Frame name already exists in the UserForm: " + request.NewName);
            if (request.NewName.Length > 40)
                issues.Add("New Frame name exceeds 40 characters.");

            int directCount = 0;
            foreach (dynamic child in original.Controls)
            {
                if (!SameContainer((object)child.Parent, source, (string)form.Name)) continue;
                directCount++;
                if (directCount > 512)
                    throw new InvalidOperationException("Frame has more than 512 direct controls.");
                string childName = (string)child.Name;
                string type = TypeDescriptor.GetClassName((object)child);
                string proposedName = request.NewName + "_" + childName;
                string reason = null;
                bool label = string.Equals(type, "Label", StringComparison.OrdinalIgnoreCase);
                bool textBox = allowTextBox &&
                    string.Equals(type, "TextBox", StringComparison.OrdinalIgnoreCase);
                if (!label && !textBox)
                    reason = allowTextBox
                        ? "Only direct Label or TextBox children have tested positive copy profiles."
                        : "Only direct Label children have a tested positive copy profile.";
                else if (proposedName.Length > 40)
                    reason = "Proposed child name exceeds 40 characters.";
                else if (names.Contains(proposedName))
                    reason = "Proposed child name already exists in the UserForm.";
                if (reason != null) issues.Add(childName + ": " + reason);
                names.Add(proposedName);
                children.Add(new { SourcePath = request.ControlPath + "/Controls/" + childName,
                    SourceName = childName, Type = type, ProposedName = proposedName,
                    ProposedPath = request.ControlPath.Substring(0,
                        request.ControlPath.LastIndexOf('/')) + "/" + request.NewName +
                        "/Controls/" + proposedName,
                    Profile = label ? "Label positive profile" :
                        textBox ? "TextBox text-value positive profile" : null,
                    Eligible = reason == null, Issue = reason });
            }
            int collectionCount = Convert.ToInt32(original.Controls.Count);
            return new { Project = request.Project, Form = request.Form,
                SourcePath = request.ControlPath, ProposedFrameName = request.NewName,
                ExpectedTreeVersion = request.ExpectedTreeVersion,
                CollectionCount = collectionCount, DirectChildCount = directCount,
                Children = children, Issues = issues,
                EligibleForLimitedProbe = issues.Count == 0,
                MutationVerified = false, ReadOnly = true,
                Scope = allowTextBox
                    ? "Frame with direct Label/TextBox children only; all other properties and descendants are excluded."
                    : "Frame with direct Label children only; all other properties and descendants are excluded." };
        }
    }
}

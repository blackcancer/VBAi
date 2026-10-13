using System;
using System.Collections.Generic;
using System.Linq;

namespace VBAi
{
    /// <summary>Obtains a live project module inventory for the chat's navigation picker.</summary>
    internal sealed partial class VbeChatReferences
    {
        /// <summary>Reads modules at the time of opening the picker; stale cached entries never substitute for a refused read.</summary>
        /// <param name="project">A project-only reference.</param>
        /// <returns>Current module references in a stable display order.</returns>
        /// <exception cref="ArgumentException">The reference is not a project or its module inventory is malformed.</exception>
        /// <exception cref="InvalidOperationException">The live project selector is absent, ambiguous or inaccessible.</exception>
        internal VbeChatReference[] ModulesForProject(VbeChatReference project)
        {
            if (project == null || string.IsNullOrWhiteSpace(project.Project) || project.Module != null || project.Name != null)
                throw new ArgumentException("A project-only reference is required.", nameof(project));
            var response = Execute(new Request { Command = "list_modules", Project = project.Project });
            if (!response.Ok) throw new InvalidOperationException(response.Error);
            var raw = json.DeserializeObject(json.Serialize(response.Data)) as object[];
            if (raw == null || raw.Any(row => !(row is IDictionary<string, object>)))
                throw new ArgumentException("A complete live module inventory is required.");
            var rows = raw.Cast<IDictionary<string, object>>();
            string[] names = rows.Select(row => Field(row, "Name")).ToArray();
            if (names.Any(string.IsNullOrWhiteSpace) || names.Distinct(StringComparer.OrdinalIgnoreCase).Count() != names.Length)
                throw new ArgumentException("The live module inventory contains missing or ambiguous names.");
            return names.OrderBy(name => name, StringComparer.OrdinalIgnoreCase)
                .Select(name => new VbeChatReference { Project = project.Project, Module = name, Kind = "Module" }).ToArray();
        }
    }
}

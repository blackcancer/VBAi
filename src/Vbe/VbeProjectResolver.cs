using System;
using System.Collections.Generic;
using System.IO;

namespace CodexVBE
{
    internal static class VbeProjectResolver
    {
        // A saved project's FileName is a stable selector when several projects share a VBA name.
        public static dynamic Resolve(dynamic vbe, string selector)
        {
            if (string.IsNullOrWhiteSpace(selector)) throw new ArgumentException("Project is required.");
            bool byPath = Path.IsPathRooted(selector);
            string path = byPath ? Path.GetFullPath(selector) : null;
            var matches = new List<dynamic>();
            foreach (dynamic project in vbe.VBProjects)
            {
                if (!byPath)
                {
                    if (string.Equals((string)project.Name, selector, StringComparison.OrdinalIgnoreCase))
                        matches.Add(project);
                    continue;
                }
                string fileName;
                try { fileName = (string)project.FileName; }
                catch { continue; } // An unsaved VBProject may reject FileName.
                if (string.IsNullOrWhiteSpace(fileName)) continue;
                try
                {
                    if (string.Equals(Path.GetFullPath(fileName), path, StringComparison.OrdinalIgnoreCase))
                        matches.Add(project);
                }
                catch (ArgumentException) { }
                catch (NotSupportedException) { }
                catch (PathTooLongException) { }
            }
            if (matches.Count != 1)
                throw new InvalidOperationException("Project selector is absent or ambiguous: " + selector +
                    ". Use the exact FileName from list_projects when names collide.");
            return matches[0];
        }
    }
}

using System;
using System.Collections.Generic;
using System.IO;

namespace VBAi
{

    /// <summary>Sélectionne un projet VBA par son nom ou par le chemin de son fichier enregistré.</summary>
    internal static class VbeProjectResolver
    {
        // A saved project's FileName is a stable selector when several projects share a VBA name.
        /// <summary>Résout un unique projet et refuse les sélecteurs absents ou ambigus.</summary>
        /// <param name="vbe">Instance VBE dont les projets sont parcourus.</param>
        /// <param name="selector">Nom du projet, ou chemin absolu de son fichier.</param>
        /// <param name="readHostPath">Optional owning-thread path reader for isolated host contracts.</param>
        /// <param name="releaseAcquired">Optional owner-thread release boundary for collection getter acquisitions.</param>
        /// <returns>The unique project's acquired reference; ownership transfers to the caller.</returns>
        /// <exception cref="ArgumentException">Le sélecteur est vide ou blanc.</exception>
        /// <exception cref="InvalidOperationException">Aucun projet ou plusieurs projets correspondent.</exception>
        public static dynamic Resolve(dynamic vbe, string selector, Func<object, string> readHostPath = null, Action<object> releaseAcquired = null)
        {
            if (string.IsNullOrWhiteSpace(selector)) throw new ArgumentException("Project is required.");
            bool byPath = Path.IsPathRooted(selector);
            string path = byPath ? Path.GetFullPath(selector) : null;
            var matches = new List<dynamic>();
            using (var references = new VbaTestComReferences(releaseAcquired))
            {
                foreach (dynamic project in references.Items((object)vbe.VBProjects))
                {
                    if (!byPath)
                    {
                        if (string.Equals((string)project.Name, selector, StringComparison.OrdinalIgnoreCase))
                            matches.Add(project);
                        continue;
                    }
                    string fileName;
                    try { fileName = readHostPath == null ? VbeProjectHostPath.Read((object)project) : readHostPath((object)project); }
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
                        ". Use the exact HostPath from list_projects when names collide.");
                references.Transfer((object)matches[0]);
                return matches[0];
            }
        }
    }
}

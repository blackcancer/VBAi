using System;
using System.IO;
using System.Runtime.InteropServices;

namespace VBAi.Tests.Integration
{
    internal sealed partial class OfficeVbeFixture
    {
        /// <summary>Quits only the sole disposable presentation after final inventory and saved-state readback.</summary>
        private void QuitSoleOwnedPowerPoint()
        {
            PreparePowerPointQuit(RequireSoleOwnedPowerPoint,
                () => ((dynamic)document).Saved = -1,
                () => Convert.ToInt32(((dynamic)document).Saved) == -1,
                () => QuitOwnedOnce(() => ((dynamic)application).Quit()));
        }

        /// <summary>Reads the live document/project inventory before each admitted shutdown action.</summary>
        private void RequireSoleOwnedPowerPoint()
        {
            commandContainment.RequireTerminal();
            if (NativeExecutionUnsettled) throw new InvalidOperationException("PowerPoint execution remains unsettled; no shutdown action.");
            RequireOwnedDocument();
            object presentations = null, current = null, vbe = null, projects = null, project = null;
            try
            {
                presentations = ((dynamic)application).Presentations;
                int count = Convert.ToInt32(((dynamic)presentations).Count);
                if (count != 1) throw new InvalidOperationException("PowerPoint must contain exactly the sole owned presentation before Quit.");
                current = ((dynamic)presentations)[1];
                IntPtr first = IntPtr.Zero, second = IntPtr.Zero;
                bool same;
                try
                {
                    first = Marshal.GetIUnknownForObject(current);
                    second = Marshal.GetIUnknownForObject(document);
                    same = first == second;
                }
                finally { if (second != IntPtr.Zero) Marshal.Release(second); if (first != IntPtr.Zero) Marshal.Release(first); }
                vbe = ((dynamic)application).VBE;
                projects = ((dynamic)vbe).VBProjects;
                int projectCount = Convert.ToInt32(((dynamic)projects).Count);
                if (projectCount != 1) throw new InvalidOperationException("PowerPoint must contain exactly the owned VBA project before Quit.");
                project = ((dynamic)projects)[1];
                string currentPath = (string)((dynamic)current).FullName;
                string projectPath = (string)((dynamic)project).FileName;
                int mode = Convert.ToInt32(((dynamic)project).Mode);
                ValidatePowerPointQuitInventory(count, projectCount, same, Root, DocumentPath, currentPath, projectPath, mode);
                RequireOwnedDocument();
                steps.Add(new { PowerPointQuitInventoryVerified = true, Presentations = count, Projects = projectCount,
                    ProjectMode = mode, DocumentPath, CloseEntries = 0, NativeMutationReplayed = false });
            }
            finally
            {
                Release(project); Release(projects); Release(vbe);
                if (!ReferenceEquals(current, document)) Release(current);
                Release(presentations);
            }
        }

        /// <summary>Requires a single design-mode project belonging to the fixture's exact disposable file.</summary>
        internal static void ValidatePowerPointQuitInventory(int presentations, int projects, bool sameDocument,
            string root, string expectedPath, string actualPath, string projectPath, int mode)
        {
            if (presentations != 1 || projects != 1 || !sameDocument || mode != 2 /* vbext_vm_Design */ ||
                string.IsNullOrWhiteSpace(root) || string.IsNullOrWhiteSpace(expectedPath) ||
                string.IsNullOrWhiteSpace(actualPath) || string.IsNullOrWhiteSpace(projectPath))
                throw new InvalidOperationException("PowerPoint shutdown inventory is absent, changed or not in design mode.");
            string expected = Path.GetFullPath(expectedPath);
            string prefix = Path.GetFullPath(root).TrimEnd(Path.DirectorySeparatorChar, Path.AltDirectorySeparatorChar) + Path.DirectorySeparatorChar;
            if (!expected.StartsWith(prefix, StringComparison.OrdinalIgnoreCase) ||
                !string.Equals(expected, Path.GetFullPath(actualPath), StringComparison.OrdinalIgnoreCase) ||
                !string.Equals(expected, Path.GetFullPath(projectPath), StringComparison.OrdinalIgnoreCase))
                throw new InvalidOperationException("PowerPoint shutdown does not target the exact disposable presentation and project.");
        }

        /// <summary>Stages one saved flag, verifies it, revalidates inventory, then admits exactly one Quit.</summary>
        internal static void PreparePowerPointQuit(Action validate, Action markSaved, Func<bool> saved, Action quit)
        {
            if (validate == null || markSaved == null || saved == null || quit == null)
                throw new ArgumentNullException("PowerPoint shutdown boundary");
            validate();
            markSaved();
            if (!saved()) throw new InvalidOperationException("The owned PowerPoint saved flag did not retain its value; no Quit.");
            validate();
            quit();
        }
    }
}

using System;
using System.Linq;

namespace VBAi.Tests.Integration
{
    /// <summary>One macro-free Word control trial; an uncertain native phase is never repeated or automatically cleaned up.</summary>
    internal sealed class OfficeWordLifecycleControl
    {
        internal sealed class Boundary
        {
            internal Func<object> Acquire;
            internal Action<object> RequireOwner;
            internal Action RequireCopy;
            internal Action<object> RequireAutoMacroScope;
            internal Func<object, string[]> Paths;
            internal Func<object, int> ReadSecurity;
            internal Action<object, int> WriteSecurity;
            internal Func<object, object> WordBasic;
            internal Action<object, int> SuppressAutoMacros;
            internal Func<object, string, bool, object> Open;
            internal Func<object, string> DocumentPath;
            internal Func<object, bool> HasVba, ReadOnly;
            internal Action<object> Close;
            internal Action<object> Release = VbaTestComReferences.Release;
        }

        internal bool Uncertain { get; private set; }
        internal bool Entered { get; private set; }
        internal bool OpenEnteredWithForceDisable { get; private set; }
        internal bool OriginalSuppressionRestoreReturned { get; private set; }
        internal bool OriginalSecurityRestored { get; private set; }

        /// <summary>Runs once on the application's acquiring STA; returns only after close, absence and safety restoration.</summary>
        internal void Run(Boundary boundary, string original, string copy, bool visible, bool suppress)
        {
            if (Entered) throw new InvalidOperationException("The original control trial cannot be repeated.");
            Entered = true;
            if (boundary == null || boundary.Acquire == null || boundary.RequireOwner == null || boundary.RequireCopy == null || boundary.RequireAutoMacroScope == null ||
                boundary.Paths == null || boundary.ReadSecurity == null || boundary.WriteSecurity == null || boundary.WordBasic == null ||
                boundary.SuppressAutoMacros == null || boundary.Open == null || boundary.DocumentPath == null || boundary.HasVba == null ||
                boundary.ReadOnly == null || boundary.Close == null || boundary.Release == null)
                throw new ArgumentNullException(nameof(boundary));
            if (string.IsNullOrEmpty(original) || string.IsNullOrEmpty(copy) ||
                string.Equals(original, copy, StringComparison.OrdinalIgnoreCase))
                throw new ArgumentException("Distinct original and owned copy paths are required.");
            var references = new VbaTestComReferences(boundary.Release);
            try
            {
                boundary.RequireCopy();
                object application = references.Add(boundary.Acquire());
                if (application == null) throw new InvalidOperationException("The owned Word application is unavailable.");
                boundary.RequireOwner(application);
                RequireOnlyOriginal(boundary.Paths(application), original);
                int previous = boundary.ReadSecurity(application);
                if (previous != 2 && previous != 3) throw new InvalidOperationException("The original Word policy must be ByUI or ForceDisable.");
                if (!suppress) boundary.RequireAutoMacroScope(application);
                object commands = references.Add(boundary.WordBasic(application));
                if (commands == null) throw new InvalidOperationException("WordBasic is unavailable.");

                Uncertain = true;
                boundary.WriteSecurity(application, 3);
                if (boundary.ReadSecurity(application) != 3) throw new InvalidOperationException("The macro-free control requires ForceDisable.");
                // This flag is varied only while the application is ForceDisable, on reviewed
                // DOCX files. It does not enable VBA or change the user's Trust Center policy.
                boundary.SuppressAutoMacros(commands, suppress ? 1 : 0);
                boundary.RequireOwner(application);
                boundary.RequireCopy();
                if (boundary.ReadSecurity(application) != 3) throw new InvalidOperationException("Macro security changed before the controlled open.");
                if (!suppress) boundary.RequireAutoMacroScope(application);
                OpenEnteredWithForceDisable = true;
                object document = references.Add(boundary.Open(application, copy, visible));
                if (document == null || !string.Equals(copy, boundary.DocumentPath(document), StringComparison.OrdinalIgnoreCase) ||
                    boundary.HasVba(document) || !boundary.ReadOnly(document))
                    throw new InvalidOperationException("The opened control is not the exact read-only macro-free document.");
                var opened = boundary.Paths(application);
                if (opened == null || opened.Length != 2 ||
                    !opened.Any(path => string.Equals(path, original, StringComparison.OrdinalIgnoreCase)) ||
                    !opened.Any(path => string.Equals(path, copy, StringComparison.OrdinalIgnoreCase)))
                    throw new InvalidOperationException("A foreign or ambiguous document is present.");
                boundary.RequireOwner(application);
                boundary.Close(document);
                RequireOnlyOriginal(boundary.Paths(application), original);
                boundary.RequireCopy();
                boundary.SuppressAutoMacros(commands, 1); // Restore the fixture's known startup suppression before any macro run or Quit.
                OriginalSuppressionRestoreReturned = true;
                boundary.WriteSecurity(application, previous);
                if (boundary.ReadSecurity(application) != previous) throw new InvalidOperationException("The original owned security policy was not restored.");
                OriginalSecurityRestored = true;
                references.Dispose();
                Uncertain = false;
            }
            catch (Exception primary)
            {
                if (Uncertain) references.Retain();
                else
                {
                    try { references.Dispose(); }
                    catch (Exception cleanup)
                    {
                        Uncertain = true;
                        throw new AggregateException("Control refusal and reference recovery did not settle.", primary, cleanup);
                    }
                }
                throw;
            }
        }

        private static void RequireOnlyOriginal(string[] paths, string original)
        {
            if (paths == null || paths.Length != 1 || !string.Equals(paths[0], original, StringComparison.OrdinalIgnoreCase))
                throw new InvalidOperationException("Only the original owned document may remain open.");
        }
    }
}

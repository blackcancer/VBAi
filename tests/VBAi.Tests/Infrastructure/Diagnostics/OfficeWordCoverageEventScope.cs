using System;

namespace VBAi.Tests.Integration
{
    /// <summary>One owned Word copy event scope; uncertain writes and incomplete restoration are retained without replay.</summary>
    internal sealed class OfficeWordCoverageEventScope
    {
        internal sealed class Boundary
        {
            internal Action RequireOwner, ReviewTemplates;
            internal Func<int> ReadSecurity;
            internal Func<object> AcquireCommands;
            internal Action<object, int> ApplySuppression;
            internal Action<string> Record;
            internal Action<object> Release = VbaTestComReferences.Release;
        }
        private readonly Boundary boundary;
        private VbaTestComReferences references;
        private object commands;
        internal bool Entered { get; private set; }
        internal bool Active { get; private set; }
        internal bool Uncertain { get; private set; }
        internal bool RestoreEntered { get; private set; }
        internal bool Restored { get; private set; }
        internal bool ReferencesReleased { get; private set; }

        internal OfficeWordCoverageEventScope(Boundary boundary) { this.boundary = boundary; }

        internal void Enter()
        {
            if (Entered) throw new InvalidOperationException("The original event scope cannot be entered twice.");
            Entered = true;
            if (boundary == null || boundary.RequireOwner == null || boundary.ReviewTemplates == null || boundary.ReadSecurity == null ||
                boundary.AcquireCommands == null || boundary.ApplySuppression == null || boundary.Record == null || boundary.Release == null)
                throw new ArgumentNullException(nameof(boundary));
            references = new VbaTestComReferences(boundary.Release);
            try
            {
                boundary.RequireOwner();
                if (boundary.ReadSecurity() != 3) throw new InvalidOperationException("Event setup requires the owned ForceDisable policy.");
                boundary.ReviewTemplates();
                commands = references.Add(boundary.AcquireCommands());
                if (commands == null) throw new InvalidOperationException("The owned WordBasic object is unavailable.");
                boundary.RequireOwner();
                if (boundary.ReadSecurity() != 3) throw new InvalidOperationException("Security changed before event setup.");
                boundary.Record("enable-intent");
                Uncertain = true;
                boundary.ApplySuppression(commands, 0);
                Active = true;
                boundary.Record("enable-returned");
                boundary.ReviewTemplates();
                boundary.RequireOwner();
                if (boundary.ReadSecurity() != 3) throw new InvalidOperationException("Security changed after event setup.");
                boundary.Record("enable-checked");
                Uncertain = false;
            }
            catch (Exception primary)
            {
                if (Uncertain) references.Retain();
                else
                {
                    try { references.Dispose(); ReferencesReleased = true; }
                    catch (Exception cleanup)
                    {
                        Uncertain = true;
                        throw new AggregateException("Event setup and reference recovery did not settle.", primary, cleanup);
                    }
                }
                throw;
            }
        }

        internal void Restore()
        {
            if (!Active || Uncertain || RestoreEntered) throw new InvalidOperationException("Only the original settled event scope can restore once.");
            RestoreEntered = true;
            try
            {
                boundary.RequireOwner();
                int security = boundary.ReadSecurity();
                if (security != 2 && security != 3) throw new InvalidOperationException("The owned copy security policy changed.");
                boundary.Record("restore-intent");
                Uncertain = true;
                boundary.ApplySuppression(commands, 1);
                Restored = true;
                Active = false;
                boundary.Record("restore-returned");
                references.Dispose();
                ReferencesReleased = true;
                boundary.Record("restore-released");
                Uncertain = false;
            }
            catch { Retain(); throw; }
        }

        internal void Retain()
        {
            Uncertain = true;
            references?.Retain();
        }
    }
}

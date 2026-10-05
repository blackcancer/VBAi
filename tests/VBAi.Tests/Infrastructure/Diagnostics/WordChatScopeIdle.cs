using System;
using System.Windows.Automation;

namespace VBAi.Tests.Integration
{
    /// <summary>Observes an already selected Word scope and closes an expanded picker at most once.</summary>
    internal static class WordChatScopeIdle
    {
        /// <summary>One read-only sample of the exact owned scope picker after selection.</summary>
        internal sealed class Observation
        {
            internal bool Enabled;
            internal bool ExactSelection;
            internal ExpandCollapseState State;
        }

        /// <summary>Requires two enabled, exact, collapsed observations within the original fifteen-second bound.</summary>
        internal static void Wait(Action guard, Func<Observation> read, Action collapse,
            Func<long> elapsedMilliseconds, Action pause)
        {
            if (guard == null) throw new ArgumentNullException(nameof(guard));
            if (read == null) throw new ArgumentNullException(nameof(read));
            if (collapse == null) throw new ArgumentNullException(nameof(collapse));
            if (elapsedMilliseconds == null) throw new ArgumentNullException(nameof(elapsedMilliseconds));
            if (pause == null) throw new ArgumentNullException(nameof(pause));
            int stable = 0;
            bool collapseIssued = false;
            while (elapsedMilliseconds() < 15000)
            {
                guard();
                Observation current = read();
                if (current == null) throw new InvalidOperationException("The owned Word scope observation is absent.");
                if (current.State != ExpandCollapseState.Collapsed && current.State != ExpandCollapseState.Expanded)
                    throw new InvalidOperationException("The owned Word scope picker has an unexpected expansion state.");
                if (current.Enabled && current.ExactSelection && current.State == ExpandCollapseState.Collapsed)
                {
                    if (++stable == 2) return;
                }
                else
                {
                    stable = 0;
                    // Select can synchronously close and disable a WinForms picker.
                    // A disabled or changed scope only receives observations. If an
                    // exact enabled picker stays expanded, close once under guard;
                    // exceptions or an unchanged popup never cause another action.
                    if (current.Enabled && current.ExactSelection && current.State == ExpandCollapseState.Expanded && !collapseIssued)
                    {
                        guard();
                        collapseIssued = true;
                        collapse();
                    }
                }
                pause();
            }
            throw new TimeoutException("The selected saved Word chat scope did not become idle, exact and collapsed.");
        }
    }
}

using System;
using System.Linq;

namespace VBAi
{
    internal sealed partial class VbeDebug
    {
        internal string ReadGeneralCommandCaption()
        {
            string menuName = (string)vbe.CommandBars.ActiveMenuBar.Name;
            if (string.IsNullOrWhiteSpace(menuName)) throw new InvalidOperationException("Original VBE menu route is unavailable.");
            var matches = EnumerateCommands().Where(entry => entry.Id == 2578 && entry.Enabled &&
                entry.Path.StartsWith(menuName + " > ", StringComparison.Ordinal) &&
                (int)((dynamic)entry.Control).Type == 1).ToArray();
            if (matches.Length != 1 || string.IsNullOrWhiteSpace(matches[0].Caption))
                throw new InvalidOperationException("Unique observed active-menu General command required.");
            return matches[0].Caption;
        }

        internal Action<Action> CaptureGeneralCommand(Request request, Action requireApprovedTarget)
        {
            if (request == null || requireApprovedTarget == null || string.IsNullOrWhiteSpace(request.ControlCaption))
                throw new ArgumentException("An exact General command caption and approved project callback are required.");
            requireApprovedTarget();
            string menuName = (string)vbe.CommandBars.ActiveMenuBar.Name;
            if (string.IsNullOrWhiteSpace(menuName)) throw new InvalidOperationException("Original VBE menu route is unavailable.");
            var matches = EnumerateCommands().Where(entry => entry.Id == 2578 && entry.Enabled && entry.Caption == request.ControlCaption &&
                entry.Path.StartsWith(menuName + " > ", StringComparison.Ordinal)).ToArray();
            if (matches.Length != 1) throw new InvalidOperationException("The exact active-menu General command route is unavailable/ambiguous.");
            var captured = matches[0]; int originalType = (int)((dynamic)captured.Control).Type;
            if (originalType != 1) throw new InvalidOperationException("General command must be the exact native menu button.");
            bool consumed = false;
            return beforeExecute => {
                if (consumed) throw new InvalidOperationException("Original General command cannot be executed twice.");
                consumed = true;
                if ((string)vbe.CommandBars.ActiveMenuBar.Name != menuName) throw new InvalidOperationException("General active menu name changed.");
                var live = ResolvePostedNativeCommand(captured); // ID/path/caption/type; no proxy IUnknown equality.
                if ((int)((dynamic)live.Control).Type != originalType) throw new InvalidOperationException("General native command type changed.");
                requireApprovedTarget(); // All menu COM reads finish before final project/revision/pure policy guard.
                if (beforeExecute == null) throw new ArgumentException("Original Execute entry guard is required.");
                beforeExecute(); // No existing modal and original UI ownership, then mark only this operation entered.
                ((dynamic)live.Control).Execute(); // Exactly one supported Office native command invocation.
            };
        }
    }
}

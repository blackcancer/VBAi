using System;
using System.Collections.Generic;
using System.Linq;

namespace VBAi.Tests.Integration
{
    /// <summary>Pure selection of the exact newly visible native Options popup and localized Git item.</summary>
    internal static class WordChatGitMenuDiscovery
    {
        internal static bool IsExactGitItem(string name, string expectedLocalizedName, string controlType,
            int itemProcessId, int processId)
        {
            return !string.IsNullOrEmpty(expectedLocalizedName) &&
                string.Equals(name, expectedLocalizedName, StringComparison.Ordinal) &&
                controlType == "ControlType.MenuItem" && itemProcessId == processId;
        }

        internal sealed class Candidate
        {
            internal long PopupHandle, OwnerHandle;
            internal int NativeProcessId, UiProcessId, GitItemProcessId;
            internal uint NativeThreadId;
            internal bool Visible, NewlyVisible;
            internal string NativeClass, UiType;
            internal int MenuItemCount, GitLabelMatches, EnabledGitMatches;
        }

        internal static Candidate RequireUnique(IEnumerable<Candidate> inventory, int processId, uint threadId,
            long expectedOwner)
        {
            if (inventory == null) throw new ArgumentNullException(nameof(inventory));
            var rows = inventory.ToArray();
            if (rows.Length > 64 || rows.Any(row => row == null))
                throw new InvalidOperationException("Bounded Word chat popup inventory is invalid.");
            var matches = rows.Where(row => row.PopupHandle != 0 && row.Visible && row.NewlyVisible &&
                row.NativeProcessId == processId && row.UiProcessId == processId &&
                row.NativeThreadId == threadId && row.OwnerHandle == expectedOwner &&
                row.NativeClass != null && row.NativeClass.StartsWith("WindowsForms", StringComparison.Ordinal) &&
                row.UiType == "ControlType.Menu" && row.MenuItemCount > 0 && row.MenuItemCount <= 64 &&
                row.GitLabelMatches == 1 && row.EnabledGitMatches == 1 && row.GitItemProcessId == processId).ToArray();
            if (matches.Length != 1 || rows.Count(row => row.PopupHandle == matches[0].PopupHandle) != 1)
                throw new InvalidOperationException("The owned Word chat Options popup has no unique localized Git item.");
            return matches[0];
        }
    }
}

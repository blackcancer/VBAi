using System;
using System.Collections.Generic;
using System.Linq;
using VBAi.Desktop.Helper;

namespace VBAi.Tests.Integration
{
    /// <summary>Pure selection of the exact newly visible native Options popup and localized Git item.</summary>
    internal static class WordChatGitMenuDiscovery
    {
        private const uint WsChild = 0x40000000;
        private const uint RequiredDropDownOwnerExStyle = 0x00000180; // WS_EX_TOOLWINDOW | WS_EX_WINDOWEDGE

        internal static bool IsExactGitItem(string name, string expectedLocalizedName, string controlType,
            int itemProcessId, int processId)
        {
            return !string.IsNullOrEmpty(expectedLocalizedName) &&
                string.Equals(name, expectedLocalizedName, StringComparison.Ordinal) &&
                controlType == "ControlType.MenuItem" && itemProcessId == processId;
        }

        internal sealed class Candidate
        {
            internal long PopupHandle, OwnerHandle, GitItemNativeAncestor;
            internal int NativeProcessId, UiProcessId, GitItemProcessId;
            internal uint NativeThreadId;
            internal bool Visible, NewlyVisible;
            internal string NativeClass, UiType;
            internal int MenuItemCount, GitLabelMatches, EnabledGitMatches;
            internal OwnerShape OwnerShape;
        }

        internal sealed class OwnerShape
        {
            internal long Handle, Parent, Root, Owner;
            internal int ProcessId;
            internal uint ThreadId, Style, ExStyle;
            internal bool Live, Visible;
            internal string ClassName;
        }

        private static string AppDomainSuffix(string className)
        {
            int marker = className == null ? -1 : className.IndexOf(".app.", StringComparison.Ordinal);
            return marker < 0 ? null : className.Substring(marker);
        }

        internal static bool HasStrictPopupOwner(Candidate row, int processId, uint threadId, long expectedRoot,
            OwnerShape exactScopePicker = null)
        {
            if (row == null || row.OwnerHandle == 0 || row.OwnerShape == null) return false;
            OwnerShape owner = row.OwnerShape;
            if (!owner.Live || owner.Handle != row.OwnerHandle || owner.ProcessId != processId ||
                owner.ThreadId != threadId) return false;
            if (row.OwnerHandle == expectedRoot) return owner.Root == expectedRoot;
            if (exactScopePicker != null && owner.Handle == exactScopePicker.Handle)
                return owner.Root == expectedRoot && owner.Parent != 0 && owner.Owner == 0 &&
                    owner.Visible && (owner.Style & WsChild) != 0 &&
                    owner.ClassName != null && owner.ClassName.StartsWith("WindowsForms10.COMBOBOX.app.", StringComparison.Ordinal) &&
                    SameOwner(owner, exactScopePicker);
            string popupDomain = AppDomainSuffix(row.NativeClass);
            return row.OwnerHandle != row.PopupHandle && !owner.Visible && owner.Parent == 0 &&
                owner.Root == owner.Handle && owner.Owner == 0 && (owner.Style & WsChild) == 0 &&
                (owner.ExStyle & RequiredDropDownOwnerExStyle) == RequiredDropDownOwnerExStyle &&
                owner.ClassName != null && owner.ClassName.StartsWith("WindowsForms10.Window.0.app.", StringComparison.Ordinal) &&
                popupDomain != null && string.Equals(AppDomainSuffix(owner.ClassName), popupDomain, StringComparison.Ordinal);
        }

        private static bool SameOwner(OwnerShape before, OwnerShape now)
            => before != null && now != null && before.Handle == now.Handle && before.Parent == now.Parent &&
                before.Root == now.Root && before.Owner == now.Owner && before.ProcessId == now.ProcessId &&
                before.ThreadId == now.ThreadId && before.Style == now.Style && before.ExStyle == now.ExStyle &&
                before.Live == now.Live && before.Visible == now.Visible && before.ClassName == now.ClassName;

        internal static void RequireUnchangedPopupOwner(Candidate selected, string currentPopupClass,
            long currentPopupOwner, OwnerShape currentOwner, int processId, uint threadId, long expectedRoot,
            OwnerShape exactScopePicker = null)
        {
            if (selected == null || selected.OwnerShape == null || currentOwner == null ||
                selected.NativeClass != currentPopupClass || selected.OwnerHandle != currentPopupOwner)
                throw new InvalidOperationException("The exact Word chat Options popup owner changed before invocation.");
            var now = new Candidate { PopupHandle = selected.PopupHandle, NativeClass = currentPopupClass,
                OwnerHandle = currentPopupOwner, OwnerShape = currentOwner };
            OwnerShape before = selected.OwnerShape;
            if (!HasStrictPopupOwner(selected, processId, threadId, expectedRoot, exactScopePicker) ||
                !HasStrictPopupOwner(now, processId, threadId, expectedRoot, exactScopePicker) || !SameOwner(before, currentOwner))
                throw new InvalidOperationException("The exact Word chat Options popup owner changed before invocation.");
        }

        internal static Candidate RequireUnique(IEnumerable<Candidate> inventory, int processId, uint threadId,
            long expectedOwner, OwnerShape exactScopePicker = null)
        {
            if (inventory == null) throw new ArgumentNullException(nameof(inventory));
            var rows = inventory.ToArray();
            if (rows.Length > 64 || rows.Any(row => row == null))
                throw new InvalidOperationException("Bounded Word chat popup inventory is invalid.");
            var matches = rows.Where(row => row.PopupHandle != 0 && row.Visible && row.NewlyVisible &&
                row.NativeProcessId == processId && row.UiProcessId == processId &&
                row.NativeThreadId == threadId && HasStrictPopupOwner(row, processId, threadId, expectedOwner, exactScopePicker) &&
                NativeToolStripPopupIdentity.Matches(row.UiType, row.NativeClass) && row.MenuItemCount > 0 && row.MenuItemCount <= 64 &&
                row.GitLabelMatches == 1 && row.EnabledGitMatches == 1 && row.GitItemProcessId == processId &&
                row.GitItemNativeAncestor == row.PopupHandle).ToArray();
            if (matches.Length != 1 || rows.Count(row => row.PopupHandle == matches[0].PopupHandle) != 1)
                throw new InvalidOperationException("The owned Word chat Options popup has no unique localized Git item.");
            return matches[0];
        }
    }
}

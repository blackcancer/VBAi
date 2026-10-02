using System;
using System.Collections.Generic;
using System.Linq;

namespace VBAi.Tests.Integration
{
    /// <summary>Pure, content-free identity check for a native Word chat window observed under its owned VBE.</summary>
    internal static class WordChatWindowDiscovery
    {
        internal sealed class Candidate
        {
            internal long Handle;
            internal int NativeProcessId, UiProcessId;
            internal uint NativeThreadId;
            internal bool Visible, WithinOwnedVbe;
            internal string NativeClass, ControlType;
            internal int ScopePickerCount, OptionsCount;
            internal int ScopePickerProcessId, OptionsProcessId;
            internal string ScopePickerType, OptionsType;
            internal long ScopePickerHandle, OptionsHandle;
            internal uint ScopePickerThreadId, OptionsThreadId;
            internal bool ScopePickerWithinChat, OptionsWithinChat;
        }

        internal static Candidate RequireUnique(IEnumerable<Candidate> inventory, int processId, uint threadId)
        {
            if (inventory == null) throw new ArgumentNullException(nameof(inventory));
            var rows = inventory.ToArray();
            if (rows.Length > 64 || rows.Any(row => row == null))
                throw new InvalidOperationException("Bounded Word chat native-window inventory is invalid.");
            var matches = rows.Where(row => row.Handle != 0 && row.Visible && row.WithinOwnedVbe &&
                row.NativeProcessId == processId && row.UiProcessId == processId && row.NativeThreadId == threadId &&
                row.NativeClass != null && row.NativeClass.StartsWith("WindowsForms", StringComparison.Ordinal) &&
                row.ControlType == "ControlType.Window" && row.ScopePickerCount == 1 && row.OptionsCount == 1 &&
                row.ScopePickerProcessId == processId && row.OptionsProcessId == processId &&
                row.ScopePickerType == "ControlType.ComboBox" && row.OptionsType == "ControlType.Button" &&
                row.ScopePickerHandle != 0 && row.OptionsHandle != 0 &&
                row.ScopePickerThreadId == threadId && row.OptionsThreadId == threadId &&
                row.ScopePickerWithinChat && row.OptionsWithinChat).ToArray();
            if (matches.Length != 1 || rows.Count(row => row.Handle == matches[0].Handle) != 1)
                throw new InvalidOperationException("The owned Word chat window has no unique native scope-picker/options shape.");
            return matches[0];
        }
    }
}

using System;
using System.Collections.Generic;
using System.Linq;

namespace VBAi.Tests.Integration
{
    /// <summary>Pure, content-free identity check for a native Word chat window observed under its owned VBE.</summary>
    internal static class WordChatWindowDiscovery
    {
        internal delegate bool NativeVisitor(IntPtr window, IntPtr state);
        internal delegate bool NativeChildEnumerator(IntPtr parent, NativeVisitor visitor, IntPtr state);

        /// <summary>Bounds the callback inventory; EnumChildWindows has no usable return status.</summary>
        internal static IntPtr[] ReadNativeChildren(IntPtr parent, NativeChildEnumerator enumerate)
        {
            if (parent == IntPtr.Zero) throw new ArgumentException("A verified parent HWND is required.", nameof(parent));
            if (enumerate == null) throw new ArgumentNullException(nameof(enumerate));
            var found = new List<IntPtr>();
            NativeVisitor visitor = (window, unused) => { found.Add(window); return found.Count < 2048; };
            // Microsoft documents the BOOL return as unused. Only our callback bound is a status.
            enumerate(parent, visitor, IntPtr.Zero);
            if (found.Count >= 2048)
                throw new InvalidOperationException("Bounded Word VBE child-window inventory exceeded.");
            return found.ToArray();
        }

        internal sealed class Candidate
        {
            internal long Handle;
            internal int NativeProcessId, UiProcessId;
            internal uint NativeThreadId;
            internal bool Visible, WithinOwnedVbe;
            internal bool FixedChatCaption;
            internal string NativeClass, ControlType;
            internal int ScopePickerCount, OptionsCount;
            internal int ScopePickerProcessId, OptionsProcessId;
            internal string ScopePickerType, OptionsType;
            internal long ScopePickerHandle, OptionsHandle;
            internal uint ScopePickerThreadId, OptionsThreadId;
            internal bool ScopePickerWithinChat, OptionsWithinChat;
        }

        internal sealed class OwnerIdentity
        {
            internal long VbeHandle, VbeRoot, ChatHandle, ChatRoot, ChatOwner;
            internal bool ChatWithinVbe;
            internal int VbeRootProcessId, ChatRootProcessId;
            internal uint VbeRootThreadId, ChatRootThreadId;
        }

        /// <summary>A docked child resolves modal ownership to its top-level VBE root.</summary>
        internal static long RequireModalOwner(OwnerIdentity identity, int processId, uint threadId)
        {
            if (identity == null) throw new ArgumentNullException(nameof(identity));
            if (identity.VbeHandle == 0 || identity.VbeRoot == 0 || identity.ChatHandle == 0 || identity.ChatRoot == 0 ||
                identity.VbeRootProcessId != processId || identity.ChatRootProcessId != processId ||
                identity.VbeRootThreadId != threadId || identity.ChatRootThreadId != threadId)
                throw new InvalidOperationException("Word chat/VBE root identity is incomplete or foreign.");
            if (identity.ChatWithinVbe)
            {
                if (identity.ChatHandle == identity.ChatRoot || identity.ChatRoot != identity.VbeRoot)
                    throw new InvalidOperationException("Docked Word chat is not beneath the exact VBE root.");
            }
            else if (identity.ChatRoot != identity.ChatHandle || identity.ChatOwner != identity.VbeRoot)
                throw new InvalidOperationException("Floating Word chat is not owned by the exact VBE root.");
            return identity.ChatRoot;
        }

        internal static void RequireUnchangedModalOwner(OwnerIdentity before, OwnerIdentity now,
            long observedModalOwner, int processId, uint threadId)
        {
            long expected = RequireModalOwner(before, processId, threadId);
            if (now == null || expected != RequireModalOwner(now, processId, threadId) ||
                before.VbeHandle != now.VbeHandle || before.VbeRoot != now.VbeRoot ||
                before.ChatHandle != now.ChatHandle || before.ChatOwner != now.ChatOwner ||
                before.ChatWithinVbe != now.ChatWithinVbe || observedModalOwner != expected)
                throw new InvalidOperationException("The exact Word chat Git modal owner/root relationship changed.");
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
                (row.ControlType == "ControlType.Window" || row.ControlType == "ControlType.Pane") &&
                row.FixedChatCaption && row.ScopePickerCount == 1 && row.OptionsCount == 1 &&
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

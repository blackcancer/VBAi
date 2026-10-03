using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Globalization;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading;

namespace VBAi
{
    // General controls are the observed VBIDE dialog contract, never a host-specific Application API.
    internal sealed class VbeProjectGeneralNative : VbeProjectGeneralOperation.INative
    {
        private readonly IntPtr root;
        private readonly uint pid, thread;
        private readonly Action requireNativeContext;
        private string projectName;
        private bool fieldConsumed, closeConsumed;
        internal VbeProjectGeneralNative(IntPtr originalVbeRoot, Action requireNativeContext)
        {
            root = originalVbeRoot; pid = GetCurrentProcessId(); thread = GetCurrentThreadId();
            this.requireNativeContext = requireNativeContext ?? throw new ArgumentNullException(nameof(requireNativeContext));
            RequireOwner();
        }
        public void RequireOwner()
        {
            uint owner;
            Require(IntPtr.Size == 8 && Thread.CurrentThread.GetApartmentState() == ApartmentState.STA &&
                GetCurrentProcessId() == pid && GetCurrentThreadId() == thread && root != IntPtr.Zero &&
                GetWindowThreadProcessId(root, out owner) == thread && owner == pid && GetAncestor(root, 2) == root,
                "General must remain on its original current-process x64 VBE UI STA/root.");
            requireNativeContext(); // Caller can additionally bind private/input/sentinel proof for qualification.
        }
        public void Prepare()
        {
            RequireOwner(); Require(IsWindowEnabled(root) && Dialogs().Count == 0, "Existing owned modal/disabled VBE refuses a new General operation.");
        }
        private List<IntPtr> Dialogs()
        {
            RequireOwner(); var found = new List<IntPtr>(); var watch = Stopwatch.StartNew(); int visited = 0; Exception failure = null;
            bool complete = EnumWindows((window, state) => {
                try
                {
                    if (++visited > 8192 || watch.ElapsedMilliseconds > 5000) return false;
                    uint owner; uint tid = GetWindowThreadProcessId(window, out owner);
                    if (owner == pid && IsWindowVisible(window))
                    {
                        string kind = Class(window);
                        if (kind == "#32770" || kind.StartsWith("bosa_sdm", StringComparison.OrdinalIgnoreCase))
                        {
                            Require(tid == thread && kind == "#32770", "Unknown owned modal/thread refuses General automation."); found.Add(window);
                        }
                    }
                    return true;
                }
                catch (Exception error) { failure = error; return false; }
            }, IntPtr.Zero);
            if (failure != null) throw failure;
            Require(complete && visited <= 8192 && watch.ElapsedMilliseconds <= 5000, "Owned modal inventory is incomplete."); RequireOwner(); return found;
        }
        private void RequireWindow(IntPtr window)
        {
            RequireOwner(); uint owner;
            Require(window != IntPtr.Zero && GetWindowThreadProcessId(window, out owner) == thread && owner == pid,
                "General native target changed PID/UI thread.");
        }
        private void Budget(Stopwatch watch) { RequireOwner(); Require(watch.ElapsedMilliseconds <= 5000, "General snapshot exceeded its cooperative five-second budget."); }
        private string Text(IntPtr window, Stopwatch watch)
        {
            Budget(watch); RequireWindow(window);
            var text = new StringBuilder(4096); UIntPtr result;
            // This is an owning-thread trusted native getter; SendMessageTimeout is not a hard wall timeout on that same thread.
            Require(SendMessageTimeoutText(window, 13, new UIntPtr(4096), text, 0x23, 150, out result) != IntPtr.Zero && result.ToUInt64() < 4095,
                "General text getter is unavailable/truncated.");
            Budget(watch); RequireWindow(window); return text.ToString();
        }
        private void RequireDialog(IntPtr dialog, string name, Stopwatch watch)
        {
            Budget(watch); RequireWindow(dialog);
            Require(Class(dialog) == "#32770" && IsWindowVisible(dialog) && Dialogs().SequenceEqual(new[] { dialog }), "General modal is ambiguous or changed.");
            IntPtr owner = GetWindow(dialog, 4);
            Require(owner == root || GetAncestor(owner, 2) == root, "General is not owned by the original VBE root.");
            string caption = Text(dialog, watch);
            Require(caption == name + " - Propriétés du projet" || caption == name + " - Project Properties", "General exact project dialog title changed.");
            Budget(watch);
        }
        public VbeProjectGeneralOperation.Snapshot Capture(string exactProjectName)
        {
            RequireOwner(); var watch = Stopwatch.StartNew(); var dialogs = Dialogs(); Budget(watch);
            if (dialogs.Count == 0) return null;
            Require(dialogs.Count == 1, "Multiple owned modals refuse General."); IntPtr dialog = dialogs[0]; RequireDialog(dialog, exactProjectName, watch);
            var children = new List<IntPtr>(); bool exceeded = false; Exception childFailure = null;
            EnumChildWindows(dialog, (child, state) => {
                try { Budget(watch); if (children.Count >= 1024) { exceeded = true; return false; } RequireWindow(child); children.Add(child); return true; }
                catch (Exception error) { childFailure = error; return false; }
            }, IntPtr.Zero);
            if (childFailure != null) throw childFailure; Require(!exceeded && children.Count != 0, "General child inventory is incomplete.");
            Func<int, IntPtr> edit = id => {
                Budget(watch); var matches = children.Where(child => GetDlgCtrlID(child) == id && Class(child) == "Edit" && IsWindowVisible(child)).ToArray();
                Require(matches.Length == 1, "General expected visible edit is missing or ambiguous.");
                RequireWindow(matches[0]); Require(MatchesEditableContract(GetDlgCtrlID(matches[0]), id, Class(matches[0]), GetWindowLongPtr(matches[0], -16).ToInt64(), IsWindowVisible(matches[0]), IsWindowEnabled(matches[0])), "General edit is disabled/password/readonly."); return matches[0];
            };
            var tabs = children.Where(child => Class(child) == "SysTabControl32" && GetParent(child) == dialog && IsWindowVisible(child) && IsWindowEnabled(child)).ToArray();
            Require(tabs.Length == 1, "General exact native tab is unavailable."); RequireWindow(tabs[0]); UIntPtr selected;
            Require(SendMessageTimeoutScalar(tabs[0], 0x130B, UIntPtr.Zero, IntPtr.Zero, 0x23, 150, out selected) != IntPtr.Zero && selected.ToUInt64() == 0,
                "General tab0 must already be selected; tab changes are forbidden.");
            var snapshot = new VbeProjectGeneralOperation.Snapshot { Dialog = dialog, Tab = tabs[0], NameEdit = edit(4941), DescriptionEdit = edit(4940), HelpFileEdit = edit(4948), Context = edit(4949), CompilationEdit = edit(4958) };
            snapshot.Page = GetParent(snapshot.Context); RequireWindow(snapshot.Page);
            Require(Class(snapshot.Page) == "#32770" && GetDlgCtrlID(snapshot.Page) == 0 && GetParent(snapshot.Page) == dialog && IsWindowVisible(snapshot.Page), "General child page identity changed.");
            foreach (IntPtr window in new[] { snapshot.NameEdit, snapshot.DescriptionEdit, snapshot.HelpFileEdit, snapshot.Context, snapshot.CompilationEdit })
                Require(GetParent(window) == snapshot.Page && GetAncestor(window, 2) == dialog, "General edits belong to another native page.");
            snapshot.Name = Text(snapshot.NameEdit, watch); snapshot.Description = Text(snapshot.DescriptionEdit, watch);
            snapshot.HelpFile = Text(snapshot.HelpFileEdit, watch); snapshot.ContextText = Text(snapshot.Context, watch); snapshot.Compilation = Text(snapshot.CompilationEdit, watch);
            Require(snapshot.Name == exactProjectName, "General native project name differs from the canonical approved project.");
            RequireDialog(dialog, exactProjectName, watch); Budget(watch); projectName = exactProjectName; return snapshot;
        }
        public void RequireSame(VbeProjectGeneralOperation.Snapshot expected, bool compareValues)
        {
            Require(expected != null && projectName != null, "General snapshot was not captured.");
            var current = Capture(projectName);
            Require(expected.SameNative(current) && (!compareValues || expected.OptionsVersion == current.OptionsVersion), "Original General native handles/values changed.");
        }
        public void WriteContext(VbeProjectGeneralOperation.Snapshot expected, int value, Action beforeEntry)
        { WriteField(expected, expected.Context, value.ToString(CultureInfo.InvariantCulture), beforeEntry); }
        public void RequireHelpFileRepresentable(VbeProjectGeneralOperation.Snapshot expected, string value)
        {
            RequireSame(expected, true); RequireTextRepresentable(expected.HelpFileEdit, value);
        }
        private void RequireTextRepresentable(IntPtr target, string value)
        {
            RequireWindow(target); bool unicode = IsWindowUnicode(target); uint codePage = GetACP();
            RequireWindow(target);
            RequireExactTextRepresentation(value, checked((int)codePage), unicode);
            RequireWindow(target);
        }
        internal static void RequireExactTextRepresentation(string value, int codePage, bool unicode)
        {
            if (value == null) throw new ArgumentNullException(nameof(value));
            Require(codePage > 0, "Native General text code page is unavailable.");
            try
            {
                // Exception fallbacks disable replacement and best-fit mappings. Unicode controls
                // still require valid UTF-16; no normalization or path substitution is performed.
                var encoding = Encoding.GetEncoding(unicode ? 1200 : codePage,
                    EncoderFallback.ExceptionFallback, DecoderFallback.ExceptionFallback);
                string roundTrip = encoding.GetString(encoding.GetBytes(value));
                Require(string.Equals(value, roundTrip, StringComparison.Ordinal), "Native General text cannot preserve the exact requested value.");
            }
            catch (ArgumentException error)
            {
                throw new InvalidOperationException("Native General text cannot preserve the exact requested value in its code page; no field write was entered.", error);
            }
        }
        public void WriteHelpFile(VbeProjectGeneralOperation.Snapshot expected, string value, Action beforeEntry)
        { WriteField(expected, expected.HelpFileEdit, value ?? throw new ArgumentNullException(nameof(value)), beforeEntry); }
        private void WriteField(VbeProjectGeneralOperation.Snapshot expected, IntPtr target, string value, Action beforeEntry)
        {
            Require(!fieldConsumed, "Original General field write is already consumed."); fieldConsumed = true;
            RequireSame(expected, true); RequireOwner(); UIntPtr result;
            RequireTextRepresentable(target, value);
            Require(beforeEntry != null, "Final General field authorization is required."); beforeEntry(); RequireOwner();
            Require(SendMessageTimeoutWrite(target, 12, UIntPtr.Zero, value, 0x23, 150, out result) != IntPtr.Zero && result.ToUInt64() != 0,
                "The single General field write is uncertain; no alternative write or retry.");
        }
        public void Close(VbeProjectGeneralOperation.Snapshot expected, int buttonId, Action beforeEnqueue)
        {
            Require((buttonId == 1 || buttonId == 2) && !closeConsumed, "Only one original General OK/Cancel is allowed."); closeConsumed = true;
            RequireSame(expected, true); var watch = Stopwatch.StartNew(); IntPtr button = GetDlgItem(expected.Dialog, buttonId); RequireWindow(button);
            Require(Class(button) == "Button" && GetParent(button) == expected.Dialog && GetDlgCtrlID(button) == buttonId && IsWindowVisible(button) && IsWindowEnabled(button), "General exact close button changed.");
            string caption = Text(button, watch);
            Require(MatchesCloseCaption(buttonId, caption), "General close caption is unsupported.");
            RequireDialog(expected.Dialog, projectName, watch); RequireWindow(button);
            Require(GetDlgItem(expected.Dialog, buttonId) == button && GetParent(button) == expected.Dialog && GetDlgCtrlID(button) == buttonId && Class(button) == "Button" && IsWindowVisible(button) && IsWindowEnabled(button), "General final close identity changed.");
            Budget(watch); Require(beforeEnqueue != null, "Final General close authorization is required."); beforeEnqueue(); RequireOwner();
            Require(PostMessage(expected.Dialog, 0x111, new UIntPtr((uint)buttonId), button), "General close enqueue failed; no retry.");
        }
        public bool Closed(VbeProjectGeneralOperation.Snapshot expected)
        {
            RequireOwner(); if (IsWindow(expected.Dialog)) { RequireSame(expected, false); return false; }
            Require(Dialogs().Count == 0 && IsWindowEnabled(root), "General closure/root enabled state is unproved."); return true;
        }
        private static void Require(bool value, string reason) { if (!value) throw new InvalidOperationException(reason); }
        internal static bool MatchesEditableContract(int id, int expectedId, string kind, long style, bool visible, bool enabled)
            => id == expectedId && kind == "Edit" && visible && enabled && (style & (0x20 | 0x800)) == 0;
        internal static bool MatchesCloseCaption(int buttonId, string caption)
            => buttonId == 1 ? caption == "OK" || caption == "&OK" : buttonId == 2 && (caption == "Annuler" || caption == "&Annuler" || caption == "Cancel" || caption == "&Cancel");
        private static string Class(IntPtr window) { var text = new StringBuilder(256); Require(GetClassName(window, text, text.Capacity) > 0, "Native class unavailable."); return text.ToString(); }
        private delegate bool Visitor(IntPtr window, IntPtr state);
        [DllImport("kernel32.dll")] private static extern uint GetCurrentProcessId();
        [DllImport("kernel32.dll")] private static extern uint GetCurrentThreadId();
        [DllImport("kernel32.dll")] private static extern uint GetACP();
        [DllImport("user32.dll")] private static extern uint GetWindowThreadProcessId(IntPtr window, out uint pid);
        [DllImport("user32.dll")] private static extern IntPtr GetAncestor(IntPtr window, uint flags);
        [DllImport("user32.dll")] private static extern IntPtr GetWindow(IntPtr window, uint command);
        [DllImport("user32.dll")] private static extern IntPtr GetParent(IntPtr window);
        [DllImport("user32.dll")] private static extern IntPtr GetDlgItem(IntPtr window, int id);
        [DllImport("user32.dll")] private static extern int GetDlgCtrlID(IntPtr window);
        [DllImport("user32.dll")] private static extern bool IsWindow(IntPtr window);
        [DllImport("user32.dll")] private static extern bool IsWindowVisible(IntPtr window);
        [DllImport("user32.dll")] private static extern bool IsWindowEnabled(IntPtr window);
        [DllImport("user32.dll")] private static extern bool IsWindowUnicode(IntPtr window);
        [DllImport("user32.dll", EntryPoint = "GetWindowLongPtrW")] private static extern IntPtr GetWindowLongPtr(IntPtr window, int index);
        [DllImport("user32.dll", EntryPoint = "GetClassNameW", CharSet = CharSet.Unicode)] private static extern int GetClassName(IntPtr window, StringBuilder text, int capacity);
        [DllImport("user32.dll")] private static extern bool EnumWindows(Visitor visitor, IntPtr state);
        [DllImport("user32.dll")] private static extern bool EnumChildWindows(IntPtr window, Visitor visitor, IntPtr state);
        [DllImport("user32.dll", EntryPoint = "PostMessageW", SetLastError = true)] private static extern bool PostMessage(IntPtr window, uint message, UIntPtr parameter, IntPtr data);
        [DllImport("user32.dll", EntryPoint = "SendMessageTimeoutW", CharSet = CharSet.Unicode, SetLastError = true)] private static extern IntPtr SendMessageTimeoutText(IntPtr window, uint message, UIntPtr parameter, StringBuilder text, uint flags, uint timeout, out UIntPtr result);
        [DllImport("user32.dll", EntryPoint = "SendMessageTimeoutW", CharSet = CharSet.Unicode, SetLastError = true)] private static extern IntPtr SendMessageTimeoutWrite(IntPtr window, uint message, UIntPtr parameter, string text, uint flags, uint timeout, out UIntPtr result);
        [DllImport("user32.dll", EntryPoint = "SendMessageTimeoutW", SetLastError = true)] private static extern IntPtr SendMessageTimeoutScalar(IntPtr window, uint message, UIntPtr parameter, IntPtr data, uint flags, uint timeout, out UIntPtr result);
    }
}

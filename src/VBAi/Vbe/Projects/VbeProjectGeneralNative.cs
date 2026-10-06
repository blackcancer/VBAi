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
    /// <summary>Owns the vbe project general native state and operations.</summary>
    internal sealed class VbeProjectGeneralNative : VbeProjectGeneralOperation.INative
    {

        /// <summary>Maintains the root state for vbe project general native.</summary>
        private readonly IntPtr root;

        /// <summary>Identifies the pid and thread associated with vbe project general native.</summary>
        private readonly uint pid, thread;

        /// <summary>Maintains the require native context state for vbe project general native.</summary>
        private readonly Action requireNativeContext;

        /// <summary>Maintains the project name state for vbe project general native.</summary>
        private string projectName;

        /// <summary>Maintains the field consumed and close consumed state for vbe project general native.</summary>
        private bool fieldConsumed, closeConsumed;

        /// <summary>Initializes a VbeProjectGeneralNative instance with the supplied state.</summary>
        /// <param name="originalVbeRoot">Native handle that supplies the original vbe root for this operation.</param>
        /// <param name="requireNativeContext">action that supplies the require native context for this operation.</param>
        internal VbeProjectGeneralNative(IntPtr originalVbeRoot, Action requireNativeContext)
        {
            root = originalVbeRoot; pid = GetCurrentProcessId(); thread = GetCurrentThreadId();
            this.requireNativeContext = requireNativeContext ?? throw new ArgumentNullException(nameof(requireNativeContext));
            RequireOwner();
        }

        /// <summary>Requires owner for vbe project general native.</summary>
        public void RequireOwner()
        {
            uint owner;
            Require(IntPtr.Size == 8 && Thread.CurrentThread.GetApartmentState() == ApartmentState.STA &&
                GetCurrentProcessId() == pid && GetCurrentThreadId() == thread && root != IntPtr.Zero &&
                GetWindowThreadProcessId(root, out owner) == thread && owner == pid && GetAncestor(root, 2) == root,
                "General must remain on its original current-process x64 VBE UI STA/root.");
            requireNativeContext(); // Caller can additionally bind private/input/sentinel proof for qualification.
        }

        /// <summary>Handles prepare for vbe project general native.</summary>
        public void Prepare()
        {
            RequireOwner(); Require(IsWindowEnabled(root) && Dialogs().Count == 0, "Existing owned modal/disabled VBE refuses a new General operation.");
        }

        /// <summary>Handles dialogs for vbe project general native.</summary>
        /// <returns>list&lt;int ptr&gt; produced by the operation for dialogs on vbe project general native.</returns>
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

        /// <summary>Requires window for vbe project general native.</summary>
        /// <param name="window">Native handle that supplies the window for this operation.</param>
        private void RequireWindow(IntPtr window)
        {
            RequireOwner(); uint owner;
            Require(window != IntPtr.Zero && GetWindowThreadProcessId(window, out owner) == thread && owner == pid,
                "General native target changed PID/UI thread.");
        }

        /// <summary>Handles budget for vbe project general native.</summary>
        /// <param name="watch">stopwatch that supplies the watch for this operation.</param>
        private void Budget(Stopwatch watch) { RequireOwner(); Require(watch.ElapsedMilliseconds <= 5000, "General snapshot exceeded its cooperative five-second budget."); }

        /// <summary>Handles text for vbe project general native.</summary>
        /// <param name="window">Native handle that supplies the window for this operation.</param>
        /// <param name="watch">stopwatch that supplies the watch for this operation.</param>
        /// <returns>Text produced by the operation for text on vbe project general native.</returns>
        private string Text(IntPtr window, Stopwatch watch)
        {
            Budget(watch); RequireWindow(window);
            var text = new StringBuilder(4096); UIntPtr result;
            // This is an owning-thread trusted native getter; SendMessageTimeout is not a hard wall timeout on that same thread.
            Require(SendMessageTimeoutText(window, 13, new UIntPtr(4096), text, 0x23, 150, out result) != IntPtr.Zero && result.ToUInt64() < 4095,
                "General text getter is unavailable/truncated.");
            Budget(watch); RequireWindow(window); return text.ToString();
        }

        /// <summary>Requires dialog for vbe project general native.</summary>
        /// <param name="dialog">Native handle that supplies the dialog for this operation.</param>
        /// <param name="name">Text that supplies the name value. Use the format required by the calling operation.</param>
        /// <param name="watch">stopwatch that supplies the watch for this operation.</param>
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

        /// <summary>Captures the existing General dialog and its visible project fields without switching tabs or opening a dialog.</summary>
        /// <param name="exactProjectName">Canonical project name used to match the unique owned Project Properties dialog and verify its Name field.</param>
        /// <returns>A snapshot of the exact page, tab, edit controls, and their current text, or <see langword="null"/> when no owned modal is open.</returns>
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

        /// <summary>Re-captures the dialog and requires its native control identities and, optionally, all captured text to match the expected snapshot.</summary>
        /// <param name="expected">Previously captured snapshot used as the concurrency token.</param>
        /// <param name="compareValues">When true, also compares the options version derived from the captured field values.</param>
        public void RequireSame(VbeProjectGeneralOperation.Snapshot expected, bool compareValues)
        {
            Require(expected != null && projectName != null, "General snapshot was not captured.");
            var current = Capture(projectName);
            Require(expected.SameNative(current) && (!compareValues || expected.OptionsVersion == current.OptionsVersion), "Original General native handles/values changed.");
        }

        /// <summary>Writes one HelpContextID value through the captured General edit after the final authorization callback.</summary>
        /// <param name="expected">Current snapshot whose dialog and field identity must still match.</param>
        /// <param name="value">HelpContextID value formatted with invariant decimal digits.</param>
        /// <param name="beforeEntry">Final authorization check run immediately before the single native text mutation.</param>
        public void WriteContext(VbeProjectGeneralOperation.Snapshot expected, int value, Action beforeEntry)
        { WriteField(expected, expected.Context, value.ToString(CultureInfo.InvariantCulture), beforeEntry); }

        /// <summary>Verifies the current snapshot and proves the requested HelpFile round-trips through the target edit's text encoding.</summary>
        /// <param name="expected">Snapshot whose native controls and current option values must remain unchanged.</param>
        /// <param name="value">HelpFile text to encode and decode without best-fit substitution or normalization.</param>
        public void RequireHelpFileRepresentable(VbeProjectGeneralOperation.Snapshot expected, string value)
        {
            RequireSame(expected, true); RequireTextRepresentable(expected.HelpFileEdit, value);
        }

        /// <summary>Checks that text can be represented exactly by the target edit's Unicode or active ANSI code-page contract.</summary>
        /// <param name="target">Original-process edit control whose encoding mode is inspected.</param>
        /// <param name="value">Text to round-trip without best-fit replacement.</param>
        private void RequireTextRepresentable(IntPtr target, string value)
        {
            RequireWindow(target); bool unicode = IsWindowUnicode(target); uint codePage = GetACP();
            RequireWindow(target);
            RequireExactTextRepresentation(value, checked((int)codePage), unicode);
            RequireWindow(target);
        }

        /// <summary>Refuses unrepresentable text rather than changing it through replacement, best-fit conversion, or normalization.</summary>
        /// <param name="value">Text whose exact round-trip representation is required.</param>
        /// <param name="codePage">Active Windows ANSI code page; must be positive even when the control is Unicode.</param>
        /// <param name="unicode">Whether the edit control uses UTF-16 instead of the supplied ANSI code page.</param>
        internal static void RequireExactTextRepresentation(string value, int codePage, bool unicode)
        {
            if (value == null) throw new ArgumentNullException(nameof(value));
            Require(codePage > 0, "Native General text code page is unavailable.");
            // Unsupported/unknown code pages are infrastructure failures, not a
            // known representation refusal eligible for a guarded Cancel.
            var encoding = Encoding.GetEncoding(unicode ? 1200 : codePage,
                EncoderFallback.ExceptionFallback, DecoderFallback.ExceptionFallback);
            try
            {
                // Exception fallbacks disable replacement and best-fit mappings. Unicode controls
                // still require valid UTF-16; no normalization or path substitution is performed.
                string roundTrip = encoding.GetString(encoding.GetBytes(value));
                if (!string.Equals(value, roundTrip, StringComparison.Ordinal))
                    throw new VbeProjectGeneralOperation.TextRepresentationRefusedException();
            }
            catch (EncoderFallbackException error) { throw new VbeProjectGeneralOperation.TextRepresentationRefusedException(error); }
            catch (DecoderFallbackException error) { throw new VbeProjectGeneralOperation.TextRepresentationRefusedException(error); }
        }

        /// <summary>Writes the HelpFile field once after checking the captured snapshot and final authorization.</summary>
        /// <param name="expected">Snapshot whose target dialog and option values must remain current.</param>
        /// <param name="value">New HelpFile text; null is refused before native dispatch.</param>
        /// <param name="beforeEntry">Final authorization callback run directly before the native edit message.</param>
        public void WriteHelpFile(VbeProjectGeneralOperation.Snapshot expected, string value, Action beforeEntry)
        { WriteField(expected, expected.HelpFileEdit, value ?? throw new ArgumentNullException(nameof(value)), beforeEntry); }

        /// <summary>Consumes the one permitted field-write attempt, revalidates the dialog and text encoding, then sends one native setter message.</summary>
        /// <param name="expected">Snapshot used to reject stale native handles or changed project options.</param>
        /// <param name="target">Exact captured edit control for HelpFile or HelpContextID.</param>
        /// <param name="value">Exact text to set after representation checks.</param>
        /// <param name="beforeEntry">Required final authorization callback; a timeout or uncertain setter result is never retried.</param>
        private void WriteField(VbeProjectGeneralOperation.Snapshot expected, IntPtr target, string value, Action beforeEntry)
        {
            Require(!fieldConsumed, "Original General field write is already consumed."); fieldConsumed = true;
            RequireSame(expected, true); RequireOwner(); UIntPtr result;
            RequireTextRepresentable(target, value);
            Require(beforeEntry != null, "Final General field authorization is required."); beforeEntry(); RequireOwner();
            Require(SendMessageTimeoutWrite(target, 12, UIntPtr.Zero, value, 0x23, 150, out result) != IntPtr.Zero && result.ToUInt64() != 0,
                "The single General field write is uncertain; no alternative write or retry.");
        }

        /// <summary>Posts exactly one verified OK or Cancel command to the captured dialog after a final identity and authorization check.</summary>
        /// <param name="expected">Snapshot identifying the original dialog and controls.</param>
        /// <param name="buttonId">Native button ID 1 for OK or 2 for Cancel.</param>
        /// <param name="beforeEnqueue">Final authorization callback invoked immediately before posting the close command.</param>
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

        /// <summary>Verifies that the original dialog is gone, no owned modal remains, and the VBE root is enabled.</summary>
        /// <param name="expected">Snapshot identifying the dialog whose closure is being checked.</param>
        /// <returns>True when closure is proved; false while the original dialog remains open.</returns>
        public bool Closed(VbeProjectGeneralOperation.Snapshot expected)
        {
            RequireOwner(); if (IsWindow(expected.Dialog)) { RequireSame(expected, false); return false; }
            Require(Dialogs().Count == 0 && IsWindowEnabled(root), "General closure/root enabled state is unproved."); return true;
        }

        /// <summary>Requires  for vbe project general native.</summary>
        /// <param name="value">Indicates whether value is enabled.</param>
        /// <param name="reason">Text that supplies the reason value. Use the format required by the calling operation.</param>
        private static void Require(bool value, string reason) { if (!value) throw new InvalidOperationException(reason); }

        /// <summary>Handles matches editable contract for vbe project general native.</summary>
        /// <param name="id">int that supplies the id for this operation.</param>
        /// <param name="expectedId">int that supplies the expected id for this operation.</param>
        /// <param name="kind">Text that supplies the kind value. Use the format required by the calling operation.</param>
        /// <param name="style">long that supplies the style for this operation.</param>
        /// <param name="visible">Indicates whether visible is enabled.</param>
        /// <param name="enabled">Indicates whether enabled is enabled.</param>
        /// <returns>Boolean indicating the result of the check for matches editable contract on vbe project general native.</returns>
        internal static bool MatchesEditableContract(int id, int expectedId, string kind, long style, bool visible, bool enabled)
            => id == expectedId && kind == "Edit" && visible && enabled && (style & (0x20 | 0x800)) == 0;

        /// <summary>Handles matches close caption for vbe project general native.</summary>
        /// <param name="buttonId">int that supplies the button id for this operation.</param>
        /// <param name="caption">Text that supplies the caption value. Use the format required by the calling operation.</param>
        /// <returns>Boolean indicating the result of the check for matches close caption on vbe project general native.</returns>
        internal static bool MatchesCloseCaption(int buttonId, string caption)
            => buttonId == 1 ? caption == "OK" || caption == "&OK" : buttonId == 2 && (caption == "Annuler" || caption == "&Annuler" || caption == "Cancel" || caption == "&Cancel");

        /// <summary>Handles class for vbe project general native.</summary>
        /// <param name="window">Native handle that supplies the window for this operation.</param>
        /// <returns>Text produced by the operation for class on vbe project general native.</returns>
        private static string Class(IntPtr window) { var text = new StringBuilder(256); Require(GetClassName(window, text, text.Capacity) > 0, "Native class unavailable."); return text.ToString(); }

        /// <summary>Defines the visitor callback.</summary>
        /// <param name="window">Native handle that supplies the window for this operation.</param>
        /// <param name="state">Native handle that supplies the state for this operation.</param>
        /// <returns>Boolean indicating the result of the check for operation on vbe project general native.</returns>
        private delegate bool Visitor(IntPtr window, IntPtr state);

        /// <summary>Returns current process id for vbe project general native.</summary>
        /// <returns>uint produced by the operation for get current process id on vbe project general native.</returns>
        [DllImport("kernel32.dll")] private static extern uint GetCurrentProcessId();

        /// <summary>Returns current thread id for vbe project general native.</summary>
        /// <returns>uint produced by the operation for get current thread id on vbe project general native.</returns>
        [DllImport("kernel32.dll")] private static extern uint GetCurrentThreadId();

        /// <summary>Returns acp for vbe project general native.</summary>
        /// <returns>uint produced by the operation for get acp on vbe project general native.</returns>
        [DllImport("kernel32.dll")] private static extern uint GetACP();

        /// <summary>Returns window thread process id for vbe project general native.</summary>
        /// <param name="window">Native handle that supplies the window for this operation.</param>
        /// <param name="pid">uint that supplies the pid for this operation.</param>
        /// <returns>uint produced by the operation for get window thread process id on vbe project general native.</returns>
        [DllImport("user32.dll")] private static extern uint GetWindowThreadProcessId(IntPtr window, out uint pid);

        /// <summary>Returns ancestor for vbe project general native.</summary>
        /// <param name="window">Native handle that supplies the window for this operation.</param>
        /// <param name="flags">uint that supplies the flags for this operation.</param>
        /// <returns>int ptr produced by the operation for get ancestor on vbe project general native.</returns>
        [DllImport("user32.dll")] private static extern IntPtr GetAncestor(IntPtr window, uint flags);

        /// <summary>Returns window for vbe project general native.</summary>
        /// <param name="window">Native handle that supplies the window for this operation.</param>
        /// <param name="command">uint that supplies the command for this operation.</param>
        /// <returns>int ptr produced by the operation for get window on vbe project general native.</returns>
        [DllImport("user32.dll")] private static extern IntPtr GetWindow(IntPtr window, uint command);

        /// <summary>Returns parent for vbe project general native.</summary>
        /// <param name="window">Native handle that supplies the window for this operation.</param>
        /// <returns>int ptr produced by the operation for get parent on vbe project general native.</returns>
        [DllImport("user32.dll")] private static extern IntPtr GetParent(IntPtr window);

        /// <summary>Returns dlg item for vbe project general native.</summary>
        /// <param name="window">Native handle that supplies the window for this operation.</param>
        /// <param name="id">int that supplies the id for this operation.</param>
        /// <returns>int ptr produced by the operation for get dlg item on vbe project general native.</returns>
        [DllImport("user32.dll")] private static extern IntPtr GetDlgItem(IntPtr window, int id);

        /// <summary>Returns dlg ctrl id for vbe project general native.</summary>
        /// <param name="window">Native handle that supplies the window for this operation.</param>
        /// <returns>int produced by the operation for get dlg ctrl id on vbe project general native.</returns>
        [DllImport("user32.dll")] private static extern int GetDlgCtrlID(IntPtr window);

        /// <summary>Determines whether window for vbe project general native.</summary>
        /// <param name="window">Native handle that supplies the window for this operation.</param>
        /// <returns>Boolean indicating the result of the check for is window on vbe project general native.</returns>
        [DllImport("user32.dll")] private static extern bool IsWindow(IntPtr window);

        /// <summary>Determines whether window visible for vbe project general native.</summary>
        /// <param name="window">Native handle that supplies the window for this operation.</param>
        /// <returns>Boolean indicating the result of the check for is window visible on vbe project general native.</returns>
        [DllImport("user32.dll")] private static extern bool IsWindowVisible(IntPtr window);

        /// <summary>Determines whether window enabled for vbe project general native.</summary>
        /// <param name="window">Native handle that supplies the window for this operation.</param>
        /// <returns>Boolean indicating the result of the check for is window enabled on vbe project general native.</returns>
        [DllImport("user32.dll")] private static extern bool IsWindowEnabled(IntPtr window);

        /// <summary>Determines whether window unicode for vbe project general native.</summary>
        /// <param name="window">Native handle that supplies the window for this operation.</param>
        /// <returns>Boolean indicating the result of the check for is window unicode on vbe project general native.</returns>
        [DllImport("user32.dll")] private static extern bool IsWindowUnicode(IntPtr window);

        /// <summary>Returns window long ptr for vbe project general native.</summary>
        /// <param name="window">Native handle that supplies the window for this operation.</param>
        /// <param name="index">int that supplies the index for this operation.</param>
        /// <returns>int ptr produced by the operation for get window long ptr on vbe project general native.</returns>
        [DllImport("user32.dll", EntryPoint = "GetWindowLongPtrW")] private static extern IntPtr GetWindowLongPtr(IntPtr window, int index);

        /// <summary>Returns class name for vbe project general native.</summary>
        /// <param name="window">Native handle that supplies the window for this operation.</param>
        /// <param name="text">string builder that supplies the text for this operation.</param>
        /// <param name="capacity">int that supplies the capacity for this operation.</param>
        /// <returns>int produced by the operation for get class name on vbe project general native.</returns>
        [DllImport("user32.dll", EntryPoint = "GetClassNameW", CharSet = CharSet.Unicode)] private static extern int GetClassName(IntPtr window, StringBuilder text, int capacity);

        /// <summary>Handles enum windows for vbe project general native.</summary>
        /// <param name="visitor">visitor that supplies the visitor for this operation.</param>
        /// <param name="state">Native handle that supplies the state for this operation.</param>
        /// <returns>Boolean indicating the result of the check for enum windows on vbe project general native.</returns>
        [DllImport("user32.dll")] private static extern bool EnumWindows(Visitor visitor, IntPtr state);

        /// <summary>Handles enum child windows for vbe project general native.</summary>
        /// <param name="window">Native handle that supplies the window for this operation.</param>
        /// <param name="visitor">visitor that supplies the visitor for this operation.</param>
        /// <param name="state">Native handle that supplies the state for this operation.</param>
        /// <returns>Boolean indicating the result of the check for enum child windows on vbe project general native.</returns>
        [DllImport("user32.dll")] private static extern bool EnumChildWindows(IntPtr window, Visitor visitor, IntPtr state);

        /// <summary>Handles post message for vbe project general native.</summary>
        /// <param name="window">Native handle that supplies the window for this operation.</param>
        /// <param name="message">uint that supplies the message for this operation.</param>
        /// <param name="parameter">Native handle that supplies the parameter for this operation.</param>
        /// <param name="data">Native handle that supplies the data for this operation.</param>
        /// <returns>Boolean indicating the result of the check for post message on vbe project general native.</returns>
        [DllImport("user32.dll", EntryPoint = "PostMessageW", SetLastError = true)] private static extern bool PostMessage(IntPtr window, uint message, UIntPtr parameter, IntPtr data);

        /// <summary>Handles send message timeout text for vbe project general native.</summary>
        /// <param name="window">Native handle that supplies the window for this operation.</param>
        /// <param name="message">uint that supplies the message for this operation.</param>
        /// <param name="parameter">Native handle that supplies the parameter for this operation.</param>
        /// <param name="text">string builder that supplies the text for this operation.</param>
        /// <param name="flags">uint that supplies the flags for this operation.</param>
        /// <param name="timeout">uint that supplies the timeout for this operation.</param>
        /// <param name="result">Native handle that supplies the result for this operation.</param>
        /// <returns>int ptr produced by the operation for send message timeout text on vbe project general native.</returns>
        [DllImport("user32.dll", EntryPoint = "SendMessageTimeoutW", CharSet = CharSet.Unicode, SetLastError = true)] private static extern IntPtr SendMessageTimeoutText(IntPtr window, uint message, UIntPtr parameter, StringBuilder text, uint flags, uint timeout, out UIntPtr result);

        /// <summary>Handles send message timeout write for vbe project general native.</summary>
        /// <param name="window">Native handle that supplies the window for this operation.</param>
        /// <param name="message">uint that supplies the message for this operation.</param>
        /// <param name="parameter">Native handle that supplies the parameter for this operation.</param>
        /// <param name="text">Text that supplies the text value. Use the format required by the calling operation.</param>
        /// <param name="flags">uint that supplies the flags for this operation.</param>
        /// <param name="timeout">uint that supplies the timeout for this operation.</param>
        /// <param name="result">Native handle that supplies the result for this operation.</param>
        /// <returns>int ptr produced by the operation for send message timeout write on vbe project general native.</returns>
        [DllImport("user32.dll", EntryPoint = "SendMessageTimeoutW", CharSet = CharSet.Unicode, SetLastError = true)] private static extern IntPtr SendMessageTimeoutWrite(IntPtr window, uint message, UIntPtr parameter, string text, uint flags, uint timeout, out UIntPtr result);

        /// <summary>Handles send message timeout scalar for vbe project general native.</summary>
        /// <param name="window">Native handle that supplies the window for this operation.</param>
        /// <param name="message">uint that supplies the message for this operation.</param>
        /// <param name="parameter">Native handle that supplies the parameter for this operation.</param>
        /// <param name="data">Native handle that supplies the data for this operation.</param>
        /// <param name="flags">uint that supplies the flags for this operation.</param>
        /// <param name="timeout">uint that supplies the timeout for this operation.</param>
        /// <param name="result">Native handle that supplies the result for this operation.</param>
        /// <returns>int ptr produced by the operation for send message timeout scalar on vbe project general native.</returns>
        [DllImport("user32.dll", EntryPoint = "SendMessageTimeoutW", SetLastError = true)] private static extern IntPtr SendMessageTimeoutScalar(IntPtr window, uint message, UIntPtr parameter, IntPtr data, uint flags, uint timeout, out UIntPtr result);
    }
}

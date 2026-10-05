using System;
using System.Collections.Generic;
using System.Runtime.InteropServices;
using System.Text;
using System.Text.RegularExpressions;

// Qualification teardown only: one No for the one named disposable dirty standalone source.
// No project mutation API, focus/input, macro execution, or catch-all alert dismissal.
public static class Q020NativeTeardown
{
    public sealed class UiContextUnavailableException : InvalidOperationException
    { public UiContextUnavailableException(string message) : base(message) {} }
    // Called only after ExitApp entry: cached generation plus native identities, never a process/module reopen.
    public static void RequireExitingContext(int pid, long originalRoot, uint originalThread, long dialog)
    {
        var root = new IntPtr(originalRoot); RequireWindow(root, pid, originalThread);
        if (GetAncestor(root, 2) != root || Class(root) != "wndclass_desked_gsk")
            throw new UiContextUnavailableException("Original exiting VBE root context unavailable.");
        bool foundRoot = false, foundDialog = dialog == 0; int count = 0; Exception failure = null;
        bool complete = EnumDesktopWindows(GetThreadDesktop(GetCurrentThreadId()), (w, s) => {
            try {
                if (++count > 8192) throw new InvalidOperationException("Desktop inventory bound exceeded.");
                if (w == root) { RequireWindow(w, pid, originalThread); foundRoot = true; }
                if (dialog != 0 && w.ToInt64() == dialog) { RequireWindow(w, pid, originalThread); foundDialog = true; }
                return true;
            } catch (Exception e) { failure = e; return false; }
        }, IntPtr.Zero);
        if (failure != null) throw failure;
        if (!complete || !foundRoot || !foundDialog) throw new UiContextUnavailableException("Exiting UI desktop membership unavailable; no discard.");
    }
    public sealed class Candidate
    {
        public long Dialog, NoButton, OriginalFrame;
        public uint Pid, Thread;
        public string ProjectName, Prompt, NoCaption;
    }
    private delegate bool Visitor(IntPtr hwnd, IntPtr state);
    public static Candidate Observe(int pid, long frame, string projectName)
    {
        if (!Regex.IsMatch(projectName ?? "", @"^[A-Za-z_][A-Za-z0-9_]{0,30}$")) throw new InvalidOperationException("Exact disposable source name unavailable.");
        var original = new IntPtr(frame);
        var dialogs = new List<IntPtr>(); Exception error = null; int count = 0;
        bool complete = EnumDesktopWindows(GetThreadDesktop(GetCurrentThreadId()), (w, s) => {
            try {
                if (++count > 8192) throw new InvalidOperationException("Desktop inventory bound exceeded.");
                uint p; uint t = GetWindowThreadProcessId(w, out p);
                if (p != (uint)pid) return true;
                if (t == 0) throw new InvalidOperationException("Owned desktop window identity unavailable.");
                if (p == pid && IsWindowVisible(w) && Class(w) == "#32770") dialogs.Add(w);
                return true;
            } catch (Exception e) { error = e; return false; }
        }, IntPtr.Zero);
        if (error != null) throw new InvalidOperationException("Private modal inventory incomplete.", error);
        if (!complete) throw new UiContextUnavailableException("Exiting modal inventory unavailable; no discard.");
        if (dialogs.Count == 0) return null;
        if (dialogs.Count != 1) throw new InvalidOperationException("Unknown/ambiguous owned modal; do not discard.");
        RequireWindow(original, pid, null);
        IntPtr dialog = dialogs[0]; uint dialogPid; uint tid = GetWindowThreadProcessId(dialog, out dialogPid);
        IntPtr rootOwner = GetAncestor(dialog, 3);
        RequireWindow(rootOwner, pid, tid);
        // The VBE survives native SW frame teardown. The supplied root is the original VBE root, not the disappeared SW frame.
        RequireWindow(original, pid, tid);
        var buttons = new Dictionary<int, IntPtr>(); var text = new List<string>(); int icons = 0; count = 0; error = null;
        EnumChildWindows(dialog, (w, s) => {
            try {
                if (++count > 32) throw new InvalidOperationException("Prompt control inventory bound exceeded.");
                RequireWindow(w, pid, tid); if (!IsWindowVisible(w)) return true;
                string cls = Class(w); int id = GetDlgCtrlID(w);
                if (cls == "Static") { string value = Text(w); if (!string.IsNullOrWhiteSpace(value)) text.Add(value); else if (id == 20) icons++; else throw new InvalidOperationException("Unknown empty static control."); }
                else if (cls == "Button" && (id == 6 || id == 7 || id == 2 || id == 9) && IsWindowEnabled(w)) {
                    if (buttons.ContainsKey(id)) throw new InvalidOperationException("Prompt button ID is ambiguous."); buttons.Add(id, w);
                } else throw new InvalidOperationException("Unknown prompt control shape.");
                return true;
            } catch (Exception e) { error = e; return false; }
        }, IntPtr.Zero);
        if (error != null) throw new InvalidOperationException("Prompt control inventory incomplete.", error);
        if (text.Count != 1 || buttons.Count != 3 || icons > 1 || !buttons.ContainsKey(6) || !buttons.ContainsKey(7) || (!buttons.ContainsKey(2) && !buttons.ContainsKey(9)))
            throw new InvalidOperationException("Only the complete named Yes/No/Cancel save confirmation is supported.");
        string message = text[0]; string escaped = Regex.Escape(projectName);
        string french = @"^(?:Voulez-vous enregistrer les modifications apportées (?:à|au projet) |Enregistrer les modifications (?:de|apportées à) )['""«]?(?:" + escaped + @")['""»]?\s*\?$";
        string english = @"^(?:Save changes to |Do you want to save (?:the )?changes to )['""]?(?:" + escaped + @")['""]?\s*\?$";
        bool fr = Regex.IsMatch(message, french, RegexOptions.CultureInvariant);
        bool en = Regex.IsMatch(message, english, RegexOptions.CultureInvariant);
        string yes = Text(buttons[6]).Replace("&", ""), no = Text(buttons[7]).Replace("&", ""), cancel = buttons.ContainsKey(2) ? Text(buttons[2]).Replace("&", "") : null;
        bool nativeFrench = buttons.ContainsKey(9) && icons == 1 && rootOwner == original && Class(original) == "wndclass_desked_gsk" && message == "Le Projet " + projectName + " a été modifié. Voulez-vous enregistrer les modifications maintenant ?" && yes == "Oui" && no == "Non" && Text(buttons[9]).Replace("&", "") == "Aide";
        if (!nativeFrench && ((!fr || yes != "Oui" || no != "Non" || cancel != "Annuler") && (!en || yes != "Yes" || no != "No" || cancel != "Cancel")))
            throw new InvalidOperationException("Owned modal is not the predeclared disposable-source save prompt.");
        return new Candidate { Dialog = dialog.ToInt64(), NoButton = buttons[7].ToInt64(), OriginalFrame = frame,
            Pid = (uint)pid, Thread = tid, ProjectName = projectName, Prompt = message, NoCaption = Text(buttons[7]) };
    }
    public static void Discard(Candidate approved, Action requireOriginalPrivateContext)
    {
        if (approved == null || requireOriginalPrivateContext == null) throw new ArgumentNullException();
        requireOriginalPrivateContext();
        Candidate current = Observe((int)approved.Pid, approved.OriginalFrame, approved.ProjectName);
        if (current == null) throw new UiContextUnavailableException("Original disposable-source prompt disappeared before delivery.");
        if (current.Dialog != approved.Dialog || current.NoButton != approved.NoButton ||
            current.Thread != approved.Thread || current.Prompt != approved.Prompt || current.NoCaption != approved.NoCaption)
            throw new InvalidOperationException("Original disposable-source prompt changed before delivery.");
        requireOriginalPrivateContext(); RequireWindow(new IntPtr(approved.NoButton), (int)approved.Pid, approved.Thread);
        if (!IsWindowVisible(new IntPtr(approved.Dialog)) || !IsWindowEnabled(new IntPtr(approved.NoButton)) ||
            GetDlgCtrlID(new IntPtr(approved.NoButton)) != 7 || Class(new IntPtr(approved.NoButton)) != "Button" ||
            GetAncestor(new IntPtr(approved.NoButton), 2).ToInt64() != approved.Dialog)
            throw new InvalidOperationException("Exact No button changed before delivery.");
        if (!PostMessage(new IntPtr(approved.NoButton), 0xF5, UIntPtr.Zero, IntPtr.Zero))
            throw new InvalidOperationException("One discard enqueue failed; do not retry.");
    }
    private static void RequireWindow(IntPtr w, int pid, uint? thread)
    { uint p; uint t = GetWindowThreadProcessId(w, out p); if (!IsWindow(w) || p != pid || t == 0 || thread.HasValue && t != thread.Value) throw new UiContextUnavailableException("Owned teardown window identity changed."); }
    private static string Class(IntPtr w) { var text = new StringBuilder(256); if (GetClassName(w, text, text.Capacity) == 0) throw new UiContextUnavailableException("Window class unavailable."); return text.ToString(); }
    private static string Text(IntPtr w) { var text = new StringBuilder(1024); UIntPtr length; if (SendMessageTimeout(w, 13, new UIntPtr((uint)text.Capacity), text, 0x23, 100, out length) == IntPtr.Zero || length.ToUInt64() >= 1023) throw new InvalidOperationException("Prompt text read incomplete."); return text.ToString(); }
    [DllImport("user32.dll", SetLastError=true)] private static extern bool EnumDesktopWindows(IntPtr desk, Visitor callback, IntPtr state);
    [DllImport("user32.dll")] private static extern bool EnumChildWindows(IntPtr w, Visitor callback, IntPtr state);
    [DllImport("user32.dll")] private static extern IntPtr GetThreadDesktop(uint thread);
    [DllImport("kernel32.dll")] private static extern uint GetCurrentThreadId();
    [DllImport("user32.dll")] private static extern uint GetWindowThreadProcessId(IntPtr w, out uint pid);
    [DllImport("user32.dll", CharSet=CharSet.Unicode)] private static extern int GetClassName(IntPtr w, StringBuilder text, int count);
    [DllImport("user32.dll")] private static extern bool IsWindow(IntPtr w);
    [DllImport("user32.dll")] private static extern bool IsWindowVisible(IntPtr w);
    [DllImport("user32.dll")] private static extern bool IsWindowEnabled(IntPtr w);
    [DllImport("user32.dll")] private static extern int GetDlgCtrlID(IntPtr w);
    [DllImport("user32.dll")] private static extern IntPtr GetAncestor(IntPtr w, uint flag);
    [DllImport("user32.dll", CharSet=CharSet.Unicode, SetLastError=true)] private static extern IntPtr SendMessageTimeout(IntPtr w,uint message,UIntPtr first,StringBuilder text,uint flags,uint timeout,out UIntPtr result);
    [DllImport("user32.dll", CharSet=CharSet.Unicode, SetLastError=true)] private static extern bool PostMessage(IntPtr w,uint message,UIntPtr first,IntPtr second);
}

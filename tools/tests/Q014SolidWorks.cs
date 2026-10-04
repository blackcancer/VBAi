using System;
using System.Runtime.InteropServices;
using System.Runtime.InteropServices.ComTypes;
using System.Threading;
using SolidWorks.Interop.sldworks;

// Keep native SOLIDWORKS RCWs inside typed C#: PowerShell's dynamic IDispatch
// type-info scan is unavailable in the observed 2019 ROT context.
public static class Q014SolidWorks
{
    private static object Exact(int pid)
    {
        IRunningObjectTable table = null; IBindCtx context = null; IEnumMoniker iterator = null;
        try {
            Marshal.ThrowExceptionForHR(GetRunningObjectTable(0, out table));
            Marshal.ThrowExceptionForHR(CreateBindCtx(0, out context));
            table.EnumRunning(out iterator); var m = new IMoniker[1];
            while (iterator.Next(1, m, IntPtr.Zero) == 0) {
                try {
                    string name; m[0].GetDisplayName(context, null, out name);
                    if (!string.Equals(name, "SolidWorks_PID_" + pid, StringComparison.OrdinalIgnoreCase)) continue;
                    object value; table.GetObject(m[0], out value); return value;
                } catch (COMException) { } finally { if (m[0] != null) Marshal.ReleaseComObject(m[0]); }
            }
            throw new InvalidOperationException("Exact SOLIDWORKS PID ROT unavailable; no activation attempted.");
        } finally { if (iterator != null) Marshal.ReleaseComObject(iterator); if (context != null) Marshal.ReleaseComObject(context); if (table != null) Marshal.ReleaseComObject(table); }
    }
    [DllImport("ole32.dll")] private static extern int GetRunningObjectTable(int reserved, out IRunningObjectTable table);
    [DllImport("ole32.dll")] private static extern int CreateBindCtx(int reserved, out IBindCtx context);
    public static string Verify(int pid)
    {
        object raw = Exact(pid);
        try {
            var sw = (ISldWorks)raw;
            if (sw.GetProcessID() != pid) throw new InvalidOperationException("Typed SOLIDWORKS PID differs.");
            return sw.RevisionNumber();
        } finally { Marshal.ReleaseComObject(raw); }
    }

    public sealed class CommandAttempt
    {
        private readonly object gate = new object();
        private bool completed, returned;
        private string error;
        public bool Completed { get { lock (gate) return completed; } }
        public bool Returned { get { lock (gate) return returned; } }
        public string Error { get { lock (gate) return error; } }
        internal void Finish(bool value, Exception failure)
        { lock (gate) { returned = value; error = failure == null ? null : failure.ToString(); completed = true; } }
    }

    public static CommandAttempt BeginNormalExit(int pid, string revision)
    {
        var attempt = new CommandAttempt();
        var worker = new Thread(() => {
            object raw = null;
            try {
                raw = Exact(pid); var sw = (ISldWorks)raw;
                if (sw.GetProcessID() != pid || sw.RevisionNumber() != revision || sw.GetDocumentCount() != 0)
                    throw new InvalidOperationException("Owned host changed or a CAD document is open.");
                sw.ExitApp(); attempt.Finish(true, null);
            } catch (Exception error) { attempt.Finish(false, error); }
            finally { if (raw != null) Marshal.ReleaseComObject(raw); }
        });
        worker.SetApartmentState(ApartmentState.STA); worker.Start(); return attempt;
    }

    // A file command can remain modal. Retain its original STA/RCW while the
    // campaign addresses only the verified dialog on the private desktop.
    public static CommandAttempt BeginFileCommand(int pid, string revision, int command)
    {
        if (command != 573 && command != 84) throw new ArgumentOutOfRangeException("command");
        var attempt = new CommandAttempt();
        var worker = new Thread(() => {
            object raw = null;
            try {
                raw = Exact(pid); var sw = (ISldWorks)raw;
                if (sw.GetProcessID() != pid || sw.RevisionNumber() != revision)
                    throw new InvalidOperationException("Exact native command host changed.");
                attempt.Finish(sw.RunCommand(command, ""), null);
            } catch (Exception error) { attempt.Finish(false, error); }
            finally { if (raw != null) Marshal.ReleaseComObject(raw); }
        });
        worker.SetApartmentState(ApartmentState.STA);
        worker.IsBackground = false;
        worker.Start();
        return attempt;
    }
}

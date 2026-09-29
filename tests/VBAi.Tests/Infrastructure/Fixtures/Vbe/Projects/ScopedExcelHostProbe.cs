namespace VBAi.Tests.Infrastructure
{
    using System;
    using System.Runtime.InteropServices;
    using VBAi;

    /// <summary>Sonde d’hôte pour qualifier le service de production sur une instance Excel jetable explicitement possédée.</summary>
    public sealed class ScopedExcelHostProbe : VbeProjectComponents.IExcelHostProbe
    {
        /// <summary>Application COM de l’instance jetable.</summary>
        public object Application { get; set; }
        /// <summary>PID vérifié de l’instance jetable.</summary>
        public int ProcessId { get; set; }
        public bool IsExcel => true;
        public int CurrentProcessId => ProcessId;
        public object ExcelApplication() => Application;
        public uint WindowProcessId(IntPtr window) { GetWindowThreadProcessId(window,out uint pid);return pid; }
        [DllImport("user32.dll")] private static extern uint GetWindowThreadProcessId(IntPtr window,out uint pid);
    }
}

using System;
using System.Collections.Generic;
using System.Runtime.InteropServices;
using System.Threading;

namespace VBAi.Tests.Integration
{
    internal sealed partial class ExcelVbeFixture
    {
        internal IDictionary<string, object> ReadOwnedProjectIdentity()
        {
            string desktop = Environment.GetEnvironmentVariable("VBAi_TEST_DESKTOP_NAME");
            RequireProjectReadinessOwner(owned, ProcessId, ownedProcess?.Id ?? 0,
                privateDesktopChild?.ProcessId ?? 0, ownedProcess == null || ownedProcess.HasExited,
                Thread.CurrentThread.GetApartmentState() == ApartmentState.STA, OllamaOfficeDesktop.MainEnabled);
            OllamaOfficeDesktop.Require(desktop);
            uint pid;
            IntPtr hwnd = new IntPtr(Convert.ToInt64(((dynamic)application).Hwnd));
            if (GetWindowThreadProcessId(hwnd, out pid) == 0 || pid != ProcessId)
                throw new InvalidOperationException("Owned Excel native application changed.");
            OllamaOfficeDesktop.RequireWindow(desktop, (uint)ProcessId, true, hwnd);
            if (Convert.ToInt32(((dynamic)workbooks).Count) != 1 ||
                !string.IsNullOrEmpty(Convert.ToString(((dynamic)workbook).Path)))
                throw new InvalidOperationException("Exact sole unsaved workbook required.");
            object currentWorkbook = null; IntPtr retainedWorkbookIdentity = IntPtr.Zero, currentWorkbookIdentity = IntPtr.Zero;
            try {
                currentWorkbook = ((dynamic)workbooks).Item(1);
                retainedWorkbookIdentity = Marshal.GetIUnknownForObject(workbook);
                currentWorkbookIdentity = Marshal.GetIUnknownForObject(currentWorkbook);
                if (retainedWorkbookIdentity != currentWorkbookIdentity)
                    throw new InvalidOperationException("The retained workbook is not the sole native document.");
            } finally {
                if (currentWorkbookIdentity != IntPtr.Zero) Marshal.Release(currentWorkbookIdentity);
                if (retainedWorkbookIdentity != IntPtr.Zero) Marshal.Release(retainedWorkbookIdentity);
                if (!ReferenceEquals(currentWorkbook, workbook)) Release(currentWorkbook);
            }
            object vbe = null, projects = null, ownedProject = null;
            IntPtr ownedIdentity = IntPtr.Zero;
            try
            {
                ownedProject = ((dynamic)workbook).VBProject;
                ownedIdentity = Marshal.GetIUnknownForObject(ownedProject);
                string name = Convert.ToString(((dynamic)ownedProject).Name);
                vbe = ((dynamic)application).VBE; projects = ((dynamic)vbe).VBProjects;
                int count = Convert.ToInt32(((dynamic)projects).Count), matches = 0;
                if (count < 1 || count > 64 || string.IsNullOrWhiteSpace(name))
                    throw new InvalidOperationException("Bounded native project inventory required.");
                for (int index = 1; index <= count; index++)
                {
                    object project = null; IntPtr identity = IntPtr.Zero;
                    try {
                        project = ((dynamic)projects).Item(index);
                        identity = Marshal.GetIUnknownForObject(project);
                        if (identity == ownedIdentity) matches++;
                    } finally {
                        if (identity != IntPtr.Zero) Marshal.Release(identity);
                        if (!ReferenceEquals(project, ownedProject)) Release(project);
                    }
                }
                return new Dictionary<string, object> {
                    ["ProjectName"] = name, ["NativeProjectCount"] = count, ["OwnedIdentityMatches"] = matches,
                    ["WorkbookName"] = Convert.ToString(((dynamic)workbook).Name), ["WorkbookPath"] = "",
                    ["ProcessId"] = ProcessId, ["Desktop"] = OllamaOfficeDesktop.MainEnabled ? "Default" : desktop,
                    ["ReadOnly"] = true, ["OwnedWorkbookIdentityMatches"] = true
                };
            }
            finally {
                if (ownedIdentity != IntPtr.Zero) Marshal.Release(ownedIdentity);
                Release(ownedProject); Release(projects); Release(vbe);
            }
        }

        internal void WriteQualificationEvidence(string name, object data) => WriteEvidence(name, data);

        /// <summary>Allows explicit Main ownership without replacing the original process or private-child identity.</summary>
        internal static void RequireProjectReadinessOwner(bool owned, int expectedPid, int originalPid,
            int privatePid, bool exited, bool owningSta, bool main)
        {
            if (!owned || expectedPid <= 0 || originalPid != expectedPid || exited || !owningSta ||
                (!main && privatePid != expectedPid) || (main && privatePid != 0))
                throw new InvalidOperationException("Original owned Excel generation and STA in the selected desktop scope required.");
        }
    }
}

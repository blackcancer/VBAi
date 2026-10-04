using System;
using System.Collections.Generic;

// Loaded only by the qualification STA with Visual Studio's own public interop.
public static class Q014VisualStudio
{
    public static object Read(object raw)
    {
        var dte = (EnvDTE.DTE)raw;
        var targets = new List<string>();
        foreach (EnvDTE.Process process in dte.Debugger.DebuggedProcesses)
            targets.Add(process.ProcessID + " " + process.Name);
        return new { Solution = dte.Solution.FullName, Mode = (int)dte.Debugger.CurrentMode, Targets = targets };
    }

    private static EnvDTE.DTE CheckedReady(object raw, string solution)
    {
        var dte = (EnvDTE.DTE)raw;
        if (dte.Debugger.CurrentMode != EnvDTE.dbgDebugMode.dbgDesignMode ||
            dte.Debugger.DebuggedProcesses.Count != 0 ||
            !string.Equals(dte.Solution.FullName, solution, StringComparison.OrdinalIgnoreCase))
            throw new InvalidOperationException("Only the exact idle qualification solution may launch.");
        int projects = 0;
        foreach (EnvDTE.Project project in dte.Solution.Projects)
        {
            if (string.Equals(project.Name, "Q014SolidWorks", StringComparison.Ordinal)) projects++;
            else if (!string.Equals(project.Kind, EnvDTE.Constants.vsProjectKindMisc, StringComparison.OrdinalIgnoreCase))
                throw new InvalidOperationException("Another project is present in the utility solution.");
        }
        if (projects != 1) throw new InvalidOperationException("Exact utility project unavailable.");
        foreach (EnvDTE.Document document in dte.Documents)
            if (!document.Saved) throw new InvalidOperationException("An unsaved document prevents launch.");
        return dte;
    }

    public static bool IsReady(object raw, string solution)
    {
        return CheckedReady(raw, solution) != null;
    }

    public static void Start(object raw, string solution)
    {
        var dte = CheckedReady(raw, solution);
        dte.Solution.SolutionBuild.StartupProjects = new object[] { "Q014SolidWorks.vcxproj" };
        dte.ExecuteCommand("Debug.Start", "");
    }

    public static void Quit(object raw, string solution)
    {
        var dte = (EnvDTE.DTE)raw;
        if (dte.Debugger.CurrentMode != EnvDTE.dbgDebugMode.dbgDesignMode || dte.Debugger.DebuggedProcesses.Count != 0 ||
            !string.Equals(dte.Solution.FullName, solution, StringComparison.OrdinalIgnoreCase))
            throw new InvalidOperationException("Exact utility debugger must be idle before normal quit.");
        foreach (EnvDTE.Document document in dte.Documents)
            if (!document.Saved) throw new InvalidOperationException("Unsaved IDE document prevents normal quit.");
        dte.Quit();
    }
}

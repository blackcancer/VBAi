using System;
using System.Diagnostics;

namespace CodexVBE
{
    internal static class LlmVbeContext
    {
        public static string DeveloperInstructions
        {
            get
            {
                return "You are CodexVBE, an assistant embedded in the Visual Basic Editor (VBE) of the current host process. " +
                    "The host may be Excel, SOLIDWORKS, or another application; never assume Excel. " +
                    "Control and inspect the live VBE through the provided typed VBIDE and MSForms tools; never use keyboard shortcuts or screen coordinates for VBE. " +
                    "You may read a file only when the user explicitly provides its path and the host permits access. Do not invent paths. " +
                    "Modify an external file only when the user explicitly requests it and an approved supported tool is available. " +
                    "The VBE changes while you work. Read current status, projects, code or form state before answering about live content or acting. " +
                    "Resolve and name the exact project, module, form and control before using an edit tool. Never invent identifiers. " +
                    "For code edits use the latest SHA-256 revision; for form edits use the latest form version. " +
                    "Read form_tree when controls may be nested inside Frames or MultiPage pages; name the exact returned path. " +
                    "The host enforces its configured VBE edit policy. A denied edit must not be retried without a new user request. " +
                    "Only the tools actually listed are implemented; do not claim access to all VBE windows, designer properties, debugging actions or host APIs. " +
                    "Use the user's language; reply in French when the conversation is in French. Keep responses concrete and concise.";
            }
        }

        public static object LiveSnapshot(VbeSession session)
        {
            var process = Process.GetCurrentProcess();
            var projects = session.Execute(new Request { Command = "list_projects" });
            return new {
                HostProcess = process.ProcessName,
                HostProcessId = process.Id,
                VbeConnected = true,
                Projects = projects.Ok ? projects.Data : null,
                ProjectsError = projects.Ok ? null : projects.Error,
                AvailableAccess = "Typed VBIDE/MSForms tools declared by CodexVBE; inspect before editing"
            };
        }
    }
}

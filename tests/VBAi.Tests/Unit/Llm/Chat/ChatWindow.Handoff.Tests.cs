using Microsoft.VisualStudio.TestTools.UnitTesting;
using System;
using System.IO;
using System.Windows.Forms;

namespace VBAi.Tests.Unit
{
    public sealed partial class ChatWindowStateTests
    {
        [STATestMethod, DataRow(false), DataRow(true)]
        public void GitChatWithoutAmbientContextKeepsDisposalAndFailureStatusOnItsOriginalSta(bool fail)
        {
            using (var runtime = new RuntimeScope())
            {
                string document = Path.Combine(runtime.Root, "Handoff.docm"); runtime.Vbe.VBProjects[0].FileName = document;
                runtime.Host = request => Response.Success(request.Command == "list_projects"
                    ? (object)new[] { new { Name = "P", FileName = document, Path = runtime.Root } }
                    : new { SelectedProject = "P", SelectedProjectPath = document });
                var previous = GitWindow.CacheDirectory;
                try
                {
                    GitWindow.CacheDirectory = key => Path.Combine(runtime.Root, "handoff-cache");
                    using (var chat = LoadedWindow(runtime.Session))
                    {
                        var probe = new GitHandoffCallerProbe(fail); ChatWindow.ShowModal = probe.Show;
                        var label = Get<Label>(chat, "status");
                        label.TextChanged += (sender, args) => { if (label.Text.Contains("post-handoff presentation failed")) probe.Report(); };
                        probe.Start(() => Call(chat, "GitHub_Click", null, EventArgs.Empty));
                        probe.Finish(); Assert.IsFalse(GitModalSession.IsActive(chat.GitModalOwner()));
                    }
                }
                finally { GitWindow.CacheDirectory = previous; }
            }
        }
    }
}

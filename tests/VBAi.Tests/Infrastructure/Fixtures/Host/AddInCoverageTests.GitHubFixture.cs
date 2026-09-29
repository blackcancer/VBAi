using System;
namespace VBAi.Tests.Unit
{
    public sealed partial class AddInCoverageTests
    {
        public sealed class GitHubLaunchHost
        {
            public Exception ProjectError { get; set; }
            public object Project { get; set; }
            public object ActiveVBProject { get { if (ProjectError != null) throw ProjectError; return Project; } }
        }
        public sealed class GitHubUnavailableProject
        {
            public Exception PathError { get; set; }
            public string FileName { get { throw PathError; } }
        }
    }
}

using Microsoft.VisualStudio.TestTools.UnitTesting;
using System;

namespace VBAi.Tests.Integration
{
    internal sealed partial class ExcelVbeFixture
    {
        /// <summary>Binds the immutable diagnostic to this original Excel generation and native VBE owner.</summary>
        internal OwnerGitQualificationManifest OwnerGitIdentity(string workbookPath, string evidenceRoot,
            string repositoryChild, string branch, OwnerGitQualificationStep[] steps, string remote, string remoteCommit)
        {
            Assert.IsTrue(owned && ownedProcess != null && ownedProcess.Id == ProcessId && !ownedProcess.HasExited);
            object project = null, editor = null, window = null;
            try
            {
                project = OwnGitProjectRcw();
                Assert.AreEqual(workbookPath, VbeProjectHostPath.Read(project), true);
                editor = ((dynamic)project).VBE; window = ((dynamic)editor).MainWindow;
                long handle = Convert.ToInt64(((dynamic)window).HWnd);
                uint pid; uint tid = GetWindowThreadProcessId(new IntPtr(handle), out pid);
                Assert.AreEqual((uint)ProcessId, pid); Assert.IsTrue(tid != 0 && handle != 0);
                return new OwnerGitQualificationManifest
                {
                    Version = 1,
                    OwnerPid = ProcessId,
                    OwnerBirthUtcTicks = ownedProcess.StartTime.ToUniversalTime().Ticks,
                    OwnerNativeTid = tid,
                    VbeHandle = handle,
                    AssemblyMvid = typeof(VbeSession).Module.ModuleVersionId.ToString("D"),
                    AssemblySha256 = EmbeddedRawHash(typeof(VbeSession).Assembly.Location),
                    FixtureRoot = Root,
                    WorkbookPath = workbookPath,
                    EvidenceRoot = evidenceRoot,
                    RepoRelativePath = repositoryChild,
                    Project = workbookPath,
                    Branch = branch,
                    RemoteUrl = remote,
                    RemoteCommit = remoteCommit,
                    Steps = steps
                };
            }
            finally { ReleaseGitLayoutAlias(window); ReleaseGitLayoutAlias(editor); Release(project); }
        }
    }
}

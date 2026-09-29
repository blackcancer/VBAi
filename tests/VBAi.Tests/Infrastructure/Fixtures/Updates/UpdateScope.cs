using System;
using System.IO;
using VBAi;
namespace VBAi.Tests.Infrastructure
{
    internal sealed class UpdateScope : IDisposable
    {
        internal readonly string Root = Path.Combine(Path.GetTempPath(), "vbai-update-" + Guid.NewGuid().ToString("N"));
        private readonly string originalRoot = UpdatePaths.Root, originalInstallation = UpdateState.InstallationDirectoryOverride;
        private readonly Func<UpdateFeed> feed = UpdateCoordinator.CreateFeed;
        private readonly Action<bool> worker = UpdateCoordinator.LaunchWorker;
        internal UpdateScope()
        {
            Directory.CreateDirectory(Root);
            UpdatePaths.Root = Root; UpdateState.InstallationDirectoryOverride = Path.Combine(Root, "installed");
            Directory.CreateDirectory(UpdateState.InstallationDirectory);
            UpdateCoordinator.LaunchWorker = background => throw new InvalidOperationException("Native updater must never start in a unit test.");
        }
        internal void Managed()
        {
            File.WriteAllText(Path.Combine(UpdateState.InstallationDirectory, "vbai-installation.json"),
                "{\"Product\":\"VBAi\",\"Architecture\":\"win-x64\",\"UpdateProtocol\":1,\"InstallationId\":\"" + Guid.NewGuid() + "\"}");
        }
        public void Dispose()
        {
            UpdatePaths.Root = originalRoot; UpdateState.InstallationDirectoryOverride = originalInstallation;
            UpdateCoordinator.CreateFeed = feed; UpdateCoordinator.LaunchWorker = worker;
            if (Directory.Exists(Root)) Directory.Delete(Root, true);
        }
    }
}

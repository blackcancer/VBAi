namespace VBAi.Tests.Unit
{
    using System;
    using System.Reflection;
    using System.Text;
    using VBAi;

    public sealed partial class MacroGitRepositoryTests
    {
        internal static MacroGitRepository.Result Reply(string text, int exitCode = 0)
            => new MacroGitRepository.Result { Bytes = Encoding.UTF8.GetBytes(text), ExitCode = exitCode };
        internal static MacroGitRepository.Result Run(MacroGitRepository repository, string[] args, byte[] input = null, bool useRepository = true, bool allowFailure = false)
        {
            try { return (MacroGitRepository.Result)typeof(MacroGitRepository).GetMethod("Run", BindingFlags.Instance | BindingFlags.NonPublic)
                    .Invoke(repository, new object[] { args, input, useRepository, allowFailure }); }
            catch (TargetInvocationException ex) { System.Runtime.ExceptionServices.ExceptionDispatchInfo.Capture(ex.InnerException).Throw(); throw; }
        }
        internal static object Call(MacroGitRepository repository, string method, params object[] args)
        {
            try { return typeof(MacroGitRepository).GetMethod(method, BindingFlags.Instance | BindingFlags.NonPublic).Invoke(repository, args); }
            catch (TargetInvocationException ex) { System.Runtime.ExceptionServices.ExceptionDispatchInfo.Capture(ex.InnerException).Throw(); throw; }
        }
        internal static MacroGitRepository.Result Native(MacroGitRepository repository, string[] args, byte[] input, bool use, bool allow)
        {
            var previous = repository.CommandOverride;
            repository.CommandOverride = null;
            try { return Run(repository, args, input, use, allow); }
            finally { repository.CommandOverride = previous; }
        }
    }
}

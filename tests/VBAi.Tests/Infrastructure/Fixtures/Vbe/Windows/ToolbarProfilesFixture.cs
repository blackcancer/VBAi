namespace VBAi.Tests.Unit
{
    using System;
    using System.Collections.Generic;
    using System.IO;
    using System.Reflection;
    using System.Runtime.ExceptionServices;
    using VBAi;

    public sealed partial class ToolbarProfilesTests
    {
        private const string PersistentTag = "VBAi.ToolbarCommand.Persistent.";
        private sealed class ProfileScope : IDisposable
        {
            internal readonly string Root = System.IO.Path.Combine(System.IO.Path.GetTempPath(), "VBAi-profile-tests-" + Guid.NewGuid().ToString("N"));
            internal string Path => System.IO.Path.Combine(Root, "profiles.sqlite");
            internal VbeToolbarProfiles Profiles => new VbeToolbarProfiles(Path);
            public void Dispose() { if (Directory.Exists(Root)) Directory.Delete(Root, true); }
        }
        private static VbeToolbarProfiles.Command ProfileCommand() => new VbeToolbarProfiles.Command { Id = 42, Caption = "Native command", Tag = PersistentTag + "owned" };
        private static VbeToolbarProfiles.Bar ProfileBar() => new VbeToolbarProfiles.Bar { Name = "VBAi - Profile", Visible = true, Position = 1, Commands = new[] { ProfileCommand() } };
        private static void ValidateProfile(string method, object value)
        {
            try { typeof(VbeToolbarProfiles).GetMethod(method, BindingFlags.Static | BindingFlags.NonPublic).Invoke(null, new[] { value }); }
            catch (TargetInvocationException error) { ExceptionDispatchInfo.Capture(error.InnerException).Throw(); throw; }
        }
        private static Request ProfileRequest(VbeEditorWindows service, string name)
        {
            return new Request
            {
                ObjectName = name,
                ExpectedToolbarCollectionVersion = (string)Data(service.Toolbars())["ToolbarCollectionVersion"],
                ExpectedToolbarControlsVersion = (string)Data(service.ToolbarControls(new Request { ObjectName = name }))["ToolbarControlsVersion"]
            };
        }
        private static Request LayoutProfileRequest(VbeEditorWindows service, string name, string action)
        {
            var bars = (System.Collections.IEnumerable)Data(service.Toolbars())["Toolbars"];
            foreach (Dictionary<string, object> bar in bars)
                if ((string)((Dictionary<string, object>)bar["Properties"])["Name"] == name)
                    return new Request { ObjectName = name, Action = action, ExpectedWindowVersion = (string)bar["WindowVersion"], ExpectedToolbarLayoutVersion = (string)bar["ToolbarLayoutVersion"] };
            throw new InvalidOperationException("Missing fixture toolbar.");
        }
        private static void SaveProfile(VbeEditorWindows service, ToolbarCustomizationTests.Bar bar, bool create)
        {
            try { typeof(VbeEditorWindows).GetMethod("SaveToolbarProfile", BindingFlags.Instance | BindingFlags.NonPublic).Invoke(service, new object[] { bar, create }); }
            catch (TargetInvocationException error) { ExceptionDispatchInfo.Capture(error.InnerException).Throw(); throw; }
        }
    }
}

using System;
using System.Diagnostics;
namespace CodexVBE
{
    internal static class SafeLinks
    {
        internal static bool Allowed(string text) { return Uri.TryCreate(text, UriKind.Absolute, out var uri) && (uri.Scheme == "https" || uri.Scheme == "http") && string.IsNullOrEmpty(uri.UserInfo); }
        internal static void Open(string text) { Open(text, Process.Start); }
        internal static void Open(string text, Func<ProcessStartInfo, Process> start)
        {
            if (!Allowed(text)) throw new ArgumentException(UiText.Get("Only HTTP and HTTPS links can be opened."));
            start(new ProcessStartInfo(text) { UseShellExecute = true });
        }
    }
}

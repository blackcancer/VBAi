using System;
using System.ComponentModel;
using System.Runtime.InteropServices;

namespace VBAi.Tests.Integration
{
    /// <summary>Issues one close to an already observed modal after its exact native ownership guard.</summary>
    internal static class WordChatModalClose
    {
        /// <summary>Checks ownership before posting; a failed or uncertain delivery is never retried.</summary>
        internal static void PostOnce(IntPtr modal, Action guard, Func<IntPtr, bool> post)
        {
            if (modal == IntPtr.Zero) throw new ArgumentException("The exact modal handle is required.", nameof(modal));
            if (guard == null) throw new ArgumentNullException(nameof(guard));
            if (post == null) throw new ArgumentNullException(nameof(post));
            guard();
            if (!post(modal)) throw new Win32Exception(Marshal.GetLastWin32Error(), "The one owned modal close was not posted.");
        }
    }
}

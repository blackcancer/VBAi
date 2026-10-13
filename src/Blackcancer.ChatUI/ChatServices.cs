// SPDX-License-Identifier: MPL-2.0
using System;
using System.Diagnostics;
using System.Drawing;
using System.Windows.Forms;

namespace Blackcancer.ChatUI
{
    /// <summary>Host callbacks belonging to one transcript, never to a global host.</summary>
    public sealed class ChatServices
    {
        public Func<string, string> Translate { get; set; } = text => text;
        public Action<string> CopyText { get; set; } = text => { if (!string.IsNullOrEmpty(text)) Clipboard.SetText(text); };
        public Action<string> OpenLink { get; set; } = OpenHttpsLink;
        /// <summary>Optional host colorizer. Rendering and navigation work without this callback.</summary>
        public Action<RichTextBox, string, string> ColorizeCode { get; set; }
        /// <summary>Optional borrowed host icon for independent reading windows. The window clones it.</summary>
        public Func<Icon> GetWindowIcon { get; set; }
        /// <summary>The host can explicitly admit its own additional link schemes for this transcript.</summary>
        public Func<string, bool> IsAllowedLink { get; set; } = IsSafeLink;
        public string Text(string value) => Translate?.Invoke(value) ?? value;
        public static bool IsSafeLink(string value)
        {
            Uri uri;
            return Uri.TryCreate(value, UriKind.Absolute, out uri)
                && uri.Scheme == Uri.UriSchemeHttps && !string.IsNullOrEmpty(uri.Host);
        }
        private static void OpenHttpsLink(string value)
        {
            if (!IsSafeLink(value)) throw new ArgumentException("Only absolute HTTPS links are supported.", nameof(value));
            Process.Start(new ProcessStartInfo(value) { UseShellExecute = true });
        }
    }

    /// <summary>Palette and fonts supplied by a host. Supplied fonts remain owned by that host.</summary>
    public sealed class SharedChatStyle
    {
        public Color Background { get; set; } = SystemColors.Window;
        public Color Foreground { get; set; } = SystemColors.WindowText;
        public Color CodeBackground { get; set; } = Color.FromArgb(245, 245, 245);
        public Color Border { get; set; } = SystemColors.ControlDark;
        public Color Link { get; set; } = Color.RoyalBlue;
        public Font BodyFont { get; set; } = SystemFonts.MessageBoxFont;
        public Font CodeFont { get; set; }
        public bool RightToLeft { get; set; }
    }
}

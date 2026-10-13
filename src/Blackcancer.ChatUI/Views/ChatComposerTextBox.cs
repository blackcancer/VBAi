// SPDX-License-Identifier: MPL-2.0
// Ported from SolidAi native chat and VBAi native Designer-backed composer; source-vendored host-independent UI.
using System;
using System.ComponentModel;
using System.Windows.Forms;

namespace Blackcancer.ChatUI.Views
{
    /// <summary>Native composer which owns Return even in a COM host's dialog message loop.</summary>
    [ToolboxItem(true)]
    public sealed class ChatComposerTextBox : TextBox
    {
        private bool returnHandled;
        protected override void WndProc(ref Message message)
        {
            if (message.Msg == 0x0087 /* WM_GETDLGCODE */)
            {
                base.WndProc(ref message);
                // A native dialog manager must dispatch Return to this editor.
                // Keep Tab and other navigation under the host's normal processing.
                if (message.WParam == IntPtr.Zero || message.WParam.ToInt32() == (int)Keys.Enter)
                    message.Result = new IntPtr(message.Result.ToInt64() | 0x0004 /* DLGC_WANTALLKEYS */);
                return;
            }
            if (message.Msg == 0x0100 /* WM_KEYDOWN */ && message.WParam.ToInt32() == (int)Keys.Enter)
            {
                var args = new KeyEventArgs(Keys.Enter | ModifierKeys);
                OnKeyDown(args);
                returnHandled = args.Handled || args.SuppressKeyPress;
                if (returnHandled) return;
                if ((args.Modifiers & Keys.Shift) != 0 && !ReadOnly)
                {
                    SelectedText = Environment.NewLine;
                    returnHandled = true;
                    return;
                }
            }
            if (message.Msg == 0x0102 /* WM_CHAR */ && message.WParam.ToInt32() == '\r')
            {
                // TranslateMessage may already have queued WM_CHAR before WM_KEYDOWN
                // was consumed. Do not submit twice or insert an unwanted newline.
                if (!returnHandled)
                {
                    var args = new KeyEventArgs(Keys.Enter | ModifierKeys);
                    OnKeyDown(args);
                    if (!args.Handled && !args.SuppressKeyPress && (args.Modifiers & Keys.Shift) != 0 && !ReadOnly)
                        SelectedText = Environment.NewLine;
                }
                returnHandled = false;
                return;
            }
            if (message.Msg == 0x0101 /* WM_KEYUP */ && message.WParam.ToInt32() == (int)Keys.Enter)
                returnHandled = false;
            base.WndProc(ref message);
        }
    }
}

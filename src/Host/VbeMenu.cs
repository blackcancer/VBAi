using System;
using System.Collections.Generic;
using System.Reflection;
using System.Runtime.InteropServices;

namespace CodexVBE
{
    // Uses the host VBE's own CommandBars and Office CommandBarButton COM event.
    internal sealed class VbeMenu : IDisposable
    {
        private static readonly Guid ClickInterface = new Guid("000C0351-0000-0000-C000-000000000046");
        private readonly object viewButton;
        private readonly object settingsButton;
        private readonly ClickHandler viewHandler;
        private readonly ClickHandler settingsHandler;
        private readonly List<Tuple<object, ClickHandler>> editorButtons = new List<Tuple<object, ClickHandler>>();
        private readonly List<System.Drawing.Bitmap> menuImages = new List<System.Drawing.Bitmap>();

        private delegate void ClickHandler(object control, ref bool cancelDefault);

        public VbeMenu(object vbe, Action showAssistant, Action showSettings, Action showGitHub, Action<string> editorAction = null)
        {
            dynamic view = FindMenu(vbe, "affichage", "view");
            dynamic tools = FindMenu(vbe, "outils", "tools");
            viewButton = view.Controls.Add(1, Missing.Value, Missing.Value, Missing.Value, true);
            settingsButton = tools.Controls.Add(1, Missing.Value, Missing.Value, Missing.Value, true);
            ((dynamic)viewButton).Caption = UiText.Get("VBAi assistant");
            ((dynamic)viewButton).Tag = "CodexVBE.Assistant";
            ((dynamic)settingsButton).Caption = UiText.Get("VBAi settings…");
            ((dynamic)settingsButton).Tag = "CodexVBE.Settings";
            ((dynamic)viewButton).TooltipText = UiText.Get("Open VBAi — Your AI agent for VBA");
            ((dynamic)settingsButton).TooltipText = UiText.Get("Configure providers and the GitHub account");
            viewHandler = (object control, ref bool cancel) => { cancel = true; showAssistant(); };
            settingsHandler = (object control, ref bool cancel) => { cancel = true; showSettings(); };
            try
            {
                ComEventsHelper.Combine(viewButton, ClickInterface, 1, viewHandler);
                ComEventsHelper.Combine(settingsButton, ClickInterface, 1, settingsHandler);
                SetIcon(viewButton, typeof(ChatWindow));
                SetIcon(settingsButton, typeof(LlmSettingsWindow));
                object gitButton = view.Controls.Add(1, Missing.Value, Missing.Value, Missing.Value, true);
                ((dynamic)gitButton).Caption = "GitHub VBAi…";
                ((dynamic)gitButton).Tag = "CodexVBE.GitHub";
                ((dynamic)gitButton).TooltipText = UiText.Get("Synchronize the active VBA project and manage its branches and checkpoints");
                ClickHandler gitHandler = (object control, ref bool cancel) => { cancel = true; showGitHub(); };
                editorButtons.Add(Tuple.Create(gitButton, gitHandler));
                ComEventsHelper.Combine(gitButton, ClickInterface, 1, gitHandler);
                SetIcon(gitButton, typeof(GitWindow));
            }
            catch
            {
                Dispose();
                throw;
            }
            try
            {
                if (editorAction != null)
                {
                    foreach (dynamic bar in ((dynamic)vbe).CommandBars)
                    {
                        string name = Normalize((string)bar.Name);
                        if (name != "code window" && name != "code window (break)" && name != "fenêtre code") continue;
                        foreach (var action in new[] { "/expliquer", "/corriger", "/refactoriser" })
                        {
                            string command = action;
                            object button = bar.Controls.Add(1, Missing.Value, Missing.Value, Missing.Value, true);
                            ((dynamic)button).Caption = "VBAi · " + UiText.Get(action == "/expliquer" ? "Explain" : action == "/corriger" ? "Fix" : "Refactor");
                            ((dynamic)button).Tag = "CodexVBE." + action.Substring(1);
                            ClickHandler handler = (object control, ref bool cancel) => { cancel = true; editorAction(command); };
                            editorButtons.Add(Tuple.Create(button, handler));
                            ComEventsHelper.Combine(button, ClickInterface, 1, handler);
                            SetIcon(button, typeof(ChatWindow));
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                LoadLog.Write("VBE editor context menus unavailable: " + ex);
            }
        }

        private void SetIcon(object button, Type windowType)
        {
            try
            {
                var resources = new System.ComponentModel.ComponentResourceManager(windowType);
                using (var icon = (System.Drawing.Icon)resources.GetObject("$this.Icon"))
                using (var small = new System.Drawing.Icon(icon, 16, 16))
                using (var source = small.ToBitmap())
                {
                    // Office CommandBars use an OLE picture and a separate monochrome mask.
                    var picture = new System.Drawing.Bitmap(16, 16, System.Drawing.Imaging.PixelFormat.Format24bppRgb);
                    var mask = new System.Drawing.Bitmap(16, 16, System.Drawing.Imaging.PixelFormat.Format24bppRgb);
                    menuImages.Add(picture);
                    menuImages.Add(mask);
                    for (int y = 0; y < 16; y++)
                        for (int x = 0; x < 16; x++)
                        {
                            var color = source.GetPixel(x, y);
                            bool opaque = color.A >= 128;
                            picture.SetPixel(x, y, opaque ? System.Drawing.Color.FromArgb(color.R, color.G, color.B) : System.Drawing.Color.White);
                            mask.SetPixel(x, y, opaque ? System.Drawing.Color.Black : System.Drawing.Color.White);
                        }
                    ((dynamic)button).Picture = MenuPicture.ToOle(picture);
                    ((dynamic)button).Mask = MenuPicture.ToOle(mask);
                    ((dynamic)button).Style = 3; // msoButtonIconAndCaption
                }
            }
            catch (Exception ex) { LoadLog.Write("VBE menu icon unavailable: " + ex.Message); }
        }

        private sealed class MenuPicture : System.Windows.Forms.AxHost
        {
            private MenuPicture() : base("") { }
            internal static object ToOle(System.Drawing.Image image) { return GetIPictureDispFromPicture(image); }
        }

        private static dynamic FindMenu(object application, params string[] captions)
        {
            foreach (dynamic bar in ((dynamic)application).CommandBars)
            {
                int type;
                try { type = (int)bar.Type; } catch { continue; }
                if (type != 1) continue; // msoBarTypeMenuBar
                foreach (dynamic item in bar.Controls)
                {
                    string name;
                    try { name = Normalize((string)item.Caption); } catch { continue; }
                    foreach (string target in captions)
                        if (name == target) return item;
                }
            }
            throw new InvalidOperationException("VBE menu not found: " + string.Join("/", captions));
        }

        private static string Normalize(string caption)
        {
            return (caption ?? "").Replace("&", "").Trim().ToLowerInvariant();
        }

        public void Dispose()
        {
            foreach (var item in editorButtons)
            {
                try { ComEventsHelper.Remove(item.Item1, ClickInterface, 1, item.Item2); } catch { }
                try { ((dynamic)item.Item1).Delete(); } catch { }
            }
            editorButtons.Clear();
            if (viewButton != null)
            {
                try { if (viewHandler != null) ComEventsHelper.Remove(viewButton, ClickInterface, 1, viewHandler); } catch { }
                try { ((dynamic)viewButton).Delete(); } catch { }
            }
            if (settingsButton != null)
            {
                try { if (settingsHandler != null) ComEventsHelper.Remove(settingsButton, ClickInterface, 1, settingsHandler); } catch { }
                try { ((dynamic)settingsButton).Delete(); } catch { }
            }
            foreach (var image in menuImages) image.Dispose();
            menuImages.Clear();
        }
    }
}

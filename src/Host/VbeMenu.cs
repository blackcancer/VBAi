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

        private delegate void ClickHandler(object control, ref bool cancelDefault);

        public VbeMenu(object vbe, Action showAssistant, Action showSettings, Action<string> editorAction = null)
        {
            dynamic view = FindMenu(vbe, "affichage", "view");
            dynamic tools = FindMenu(vbe, "outils", "tools");
            viewButton = view.Controls.Add(1, Missing.Value, Missing.Value, Missing.Value, true);
            settingsButton = tools.Controls.Add(1, Missing.Value, Missing.Value, Missing.Value, true);
            ((dynamic)viewButton).Caption = "Assistant CodexVBE";
            ((dynamic)viewButton).Tag = "CodexVBE.Assistant";
            ((dynamic)settingsButton).Caption = "Configuration CodexVBE…";
            ((dynamic)settingsButton).Tag = "CodexVBE.Settings";
            viewHandler = (object control, ref bool cancel) => { cancel = true; showAssistant(); };
            settingsHandler = (object control, ref bool cancel) => { cancel = true; showSettings(); };
            try
            {
                ComEventsHelper.Combine(viewButton, ClickInterface, 1, viewHandler);
                ComEventsHelper.Combine(settingsButton, ClickInterface, 1, settingsHandler);
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
                            ((dynamic)button).Caption = "CodexVBE · " + action.Substring(1);
                            ((dynamic)button).Tag = "CodexVBE." + action.Substring(1);
                            ClickHandler handler = (object control, ref bool cancel) => { cancel = true; editorAction(command); };
                            editorButtons.Add(Tuple.Create(button, handler));
                            ComEventsHelper.Combine(button, ClickInterface, 1, handler);
                        }
                    }
                }
            }
            catch
            {
                Dispose();
                throw;
            }
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
        }
    }
}

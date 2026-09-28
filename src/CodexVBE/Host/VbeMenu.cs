using System;
using System.Collections.Generic;
using System.Reflection;
using System.Runtime.InteropServices;

namespace CodexVBE
{
    // Uses the host VBE's own CommandBars and Office CommandBarButton COM event.
        /// <summary>Ajoute les boutons principaux et, si demandé, les commandes de l’éditeur.</summary>
    internal sealed class VbeMenu : IDisposable
    {
        /// <summary>IID de l’interface Office utilisée pour recevoir les clics de CommandBarButton.</summary>
        private static readonly Guid ClickInterface = new Guid("000C0351-0000-0000-C000-000000000046");
        /// <summary>Bouton VBAi ajouté au menu View du VBE.</summary>
        private readonly object viewButton;
        /// <summary>Bouton VBAi ajouté au menu Tools du VBE.</summary>
        private readonly object settingsButton;
        /// <summary>Gestionnaire COM du bouton assistant.</summary>
        private readonly ClickHandler viewHandler;
        /// <summary>Gestionnaire COM du bouton de paramètres.</summary>
        private readonly ClickHandler settingsHandler;
        /// <summary>Boutons GitHub et commandes ajoutés aux menus contextuels de l’éditeur.</summary>
        private readonly List<Tuple<object, ClickHandler>> editorButtons = new List<Tuple<object, ClickHandler>>();
        /// <summary>Images OLE et masques détenus pendant la durée de vie du menu.</summary>
        private readonly List<System.Drawing.Bitmap> menuImages = new List<System.Drawing.Bitmap>();
        /// <summary>Fonction d’abonnement aux événements COM des boutons.</summary>
        private readonly Action<object, Guid, int, Delegate> subscribe;
        /// <summary>Fonction d’abonnement aux événements COM des boutons.</summary>
        private readonly Action<object, Guid, int, Delegate> unsubscribe;
        /// <summary>Fonction appliquant l’icône associée à la fenêtre.</summary>
        private readonly Action<object, Type> applyIcon;
        /// <summary>Empêche la suppression répétée des commandes et images.</summary>
        private bool disposed;

        /// <summary>Signature du gestionnaire de clic Office avec indicateur d’annulation par défaut.</summary>
        /// <param name="control">Bouton Office qui a déclenché l’événement.</param>
        /// <param name="cancelDefault">Indique si l’action Office standard doit être annulée.</param>
        private delegate void ClickHandler(object control, ref bool cancelDefault);

        /// <summary>Ajoute les boutons principaux et, si demandé, les commandes de l’éditeur.</summary>
        /// <param name="vbe">Instance VBE dont les barres de commande sont modifiées.</param>
        /// <param name="showAssistant">Action qui affiche l’assistant.</param>
        /// <param name="showSettings">Action qui ouvre les paramètres.</param>
        /// <param name="showGitHub">Action qui ouvre l’interface GitHub.</param>
        /// <param name="editorAction">Action facultative appelée depuis un menu contextuel de l’éditeur.</param>
        public VbeMenu(object vbe, Action showAssistant, Action showSettings, Action showGitHub, Action<string> editorAction = null)
            : this(vbe, showAssistant, showSettings, showGitHub, editorAction, null, null, null)
        {
        }

        /// <summary>Crée les commandes et utilise les fonctions COM injectées si elles sont fournies.</summary>
        /// <param name="vbe">Instance VBE dont les barres de commande sont modifiées.</param>
        /// <param name="showAssistant">Action qui affiche l’assistant.</param>
        /// <param name="showSettings">Action qui ouvre les paramètres.</param>
        /// <param name="showGitHub">Action qui ouvre l’interface GitHub.</param>
        /// <param name="editorAction">Action facultative appelée depuis un menu d’éditeur.</param>
        /// <param name="subscribe">Fonction d’abonnement COM, ou valeur par défaut si null.</param>
        /// <param name="unsubscribe">Fonction de désabonnement COM, ou valeur par défaut si null.</param>
        /// <param name="applyIcon">Fonction d’application d’icône, ou valeur par défaut si null.</param>
        internal VbeMenu(object vbe, Action showAssistant, Action showSettings, Action showGitHub,
        /// <summary>Fonction d’abonnement aux événements COM des boutons.</summary>
            Action<string> editorAction, Action<object, Guid, int, Delegate> subscribe,
        /// <summary>Fonction d’abonnement aux événements COM des boutons.</summary>
            Action<object, Guid, int, Delegate> unsubscribe, Action<object, Type> applyIcon)
        {
        /// <summary>Fonction d’abonnement aux événements COM des boutons.</summary>
            this.subscribe = subscribe ?? new Action<object, Guid, int, Delegate>(ComEventsHelper.Combine);
            this.unsubscribe = unsubscribe ?? ((button, iid, dispid, handler) =>
                ComEventsHelper.Remove(button, iid, dispid, handler));
            this.applyIcon = applyIcon ?? SetIcon;
            dynamic view = FindMenu(vbe, true);
            dynamic tools = FindMenu(vbe, false);
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
                this.subscribe(viewButton, ClickInterface, 1, viewHandler);
                this.subscribe(settingsButton, ClickInterface, 1, settingsHandler);
                this.applyIcon(viewButton, typeof(ChatWindow));
                this.applyIcon(settingsButton, typeof(LlmSettingsWindow));
                object gitButton = view.Controls.Add(1, Missing.Value, Missing.Value, Missing.Value, true);
                ((dynamic)gitButton).Caption = "GitHub VBAi…";
                ((dynamic)gitButton).Tag = "CodexVBE.GitHub";
                ((dynamic)gitButton).TooltipText = UiText.Get("Synchronize the active VBA project and manage its branches and checkpoints");
                ClickHandler gitHandler = (object control, ref bool cancel) => { cancel = true; showGitHub(); };
                editorButtons.Add(Tuple.Create(gitButton, gitHandler));
                this.subscribe(gitButton, ClickInterface, 1, gitHandler);
                this.applyIcon(gitButton, typeof(GitWindow));
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
                            this.subscribe(button, ClickInterface, 1, handler);
                            this.applyIcon(button, typeof(ChatWindow));
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                LoadLog.Write("VBE editor context menus unavailable: " + ex);
            }
        }

        /// <summary>Convertit l’icône de la fenêtre en image et masque OLE pour le bouton Office.</summary>
        /// <param name="button">Bouton de barre de commande auquel appliquer l’image.</param>
        /// <param name="windowType">Type de formulaire fournissant la ressource d’icône.</param>
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

        /// <summary>Expose la conversion d’une image WinForms vers IPictureDisp.</summary>
        private sealed class MenuPicture : System.Windows.Forms.AxHost
        {
        /// <summary>Expose la conversion d’une image WinForms vers IPictureDisp.</summary>
            private MenuPicture() : base("") { }
        /// <summary>Convertit une image .NET en représentation OLE IPictureDisp.</summary>
        /// <param name="image">Image .NET à convertir en image OLE.</param>
        /// <returns>Objet IPictureDisp utilisable par CommandBarButton.</returns>
internal static object ToOle(System.Drawing.Image image) { return GetIPictureDispFromPicture(image); }
        }

        /// <summary>Recherche le menu VBE View ou Tools en tenant compte de la langue de l’hôte.</summary>
        /// <param name="application">Objet dont les barres de commande sont recherchées.</param>
        /// <param name="view">Indique si le menu recherché est View plutôt que Tools.</param>
        /// <returns>Menu VBE correspondant à la langue et au type demandés.</returns>
private static dynamic FindMenu(object application, bool view)
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
                    if (UiLanguages.IsMenu(name, view)) return item;
                }
            }
            throw new InvalidOperationException("VBE menu not found: " + (view ? "View" : "Tools"));
        }

        /// <summary>Normalise une légende de menu pour comparaison sans casse ni esperluette.</summary>
        /// <param name="caption">Légende de menu à normaliser.</param>
        /// <returns>Légende sans esperluette, espaces périphériques ni différence de casse.</returns>
private static string Normalize(string caption)
        {
            return (caption ?? "").Replace("&", "").Trim().ToLowerInvariant();
        }

        /// <summary>Désabonne et supprime les boutons, puis libère les images associées.</summary>
        public void Dispose()
        {
            if (disposed) return;
            disposed = true;
            foreach (var item in editorButtons)
            {
                try { unsubscribe(item.Item1, ClickInterface, 1, item.Item2); } catch { }
                try { ((dynamic)item.Item1).Delete(); } catch { }
            }
            editorButtons.Clear();
            if (viewButton != null)
            {
                try { if (viewHandler != null) unsubscribe(viewButton, ClickInterface, 1, viewHandler); } catch { }
                try { ((dynamic)viewButton).Delete(); } catch { }
            }
            if (settingsButton != null)
            {
                try { if (settingsHandler != null) unsubscribe(settingsButton, ClickInterface, 1, settingsHandler); } catch { }
                try { ((dynamic)settingsButton).Delete(); } catch { }
            }
            foreach (var image in menuImages) image.Dispose();
            menuImages.Clear();
        }
    }
}

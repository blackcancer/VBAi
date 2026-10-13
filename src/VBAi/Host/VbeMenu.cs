using System;
using System.Collections.Generic;
using System.Reflection;
using System.Runtime.InteropServices;

namespace VBAi
{
    // Uses the host VBE's own CommandBars and Office CommandBarButton COM event.
    /// <summary>Ajoute les boutons principaux et, si demandé, les commandes de l’éditeur.</summary>
    internal sealed class VbeMenu : IDisposable
    {

        /// <summary>Abonne un délégué à un événement COM d’une source Office.</summary>
        internal static Action<object, Guid, int, Delegate> SubscribeDefault = ComEventsHelper.Combine;

        /// <summary>Lit l’icône embarquée dans les ressources du type de fenêtre.</summary>
        internal Func<Type, System.Drawing.Icon> ReadIcon = ReadIconNative;

        /// <summary>Récupère l’icône <c>$this.Icon</c> des ressources associées à une fenêtre.</summary>
        /// <param name="windowType">Type de fenêtre qui possède les ressources.</param>
        /// <returns>Icône extraite des ressources du type.</returns>
        private static System.Drawing.Icon ReadIconNative(Type windowType)
        {
            if (windowType == typeof(LlmSettingsWindow)) return VbeWindowIcons.Icon("settings");
            if (windowType == typeof(GitWindow)) return VbeWindowIcons.Icon("github");
            if (windowType == typeof(ChatWindow) || windowType == typeof(ModernEditorWindow) ||
                windowType == typeof(AboutWindow) || windowType == typeof(CrashReportWindow) || windowType == typeof(UpdateWindow))
                return VbeWindowIcons.Icon("assistant");
            return (System.Drawing.Icon)new System.ComponentModel.ComponentResourceManager(windowType).GetObject("$this.Icon");
        }

        /// <summary>IID de l’interface Office utilisée pour recevoir les clics de CommandBarButton.</summary>
        private static readonly Guid ClickInterface = new Guid("000C0351-0000-0000-C000-000000000046");

        /// <summary>Bouton VBAi ajouté au menu View du VBE.</summary>
        private object viewButton;

        /// <summary>Bouton VBAi ajouté au menu Tools du VBE.</summary>
        private object settingsButton;

        /// <summary>Gestionnaire COM du bouton assistant.</summary>
        private ClickHandler viewHandler;

        /// <summary>Gestionnaire COM du bouton de paramètres.</summary>
        private ClickHandler settingsHandler;

        /// <summary>Boutons GitHub et commandes ajoutés aux menus contextuels de l’éditeur.</summary>
        private readonly List<Tuple<object, ClickHandler>> editorButtons = new List<Tuple<object, ClickHandler>>();

        /// <summary>Images OLE et masques détenus pendant la durée de vie du menu.</summary>
        private readonly List<System.Drawing.Bitmap> menuImages = new List<System.Drawing.Bitmap>();

        /// <summary>Fonction d’abonnement aux événements COM des boutons.</summary>
        private readonly Action<object, Guid, int, Delegate> subscribe;

        /// <summary>Fonction de désabonnement des événements COM des boutons.</summary>
        private readonly Action<object, Guid, int, Delegate> unsubscribe;

        /// <summary>Fonction appliquant l’icône associée à la fenêtre.</summary>
        private readonly Action<object, Type> applyIcon;

        /// <summary>Empêche la suppression répétée des commandes et images.</summary>
        private bool disposed;

        /// <summary>Owns each native getter and button acquisition until menu shutdown.</summary>
        private readonly VbaTestComReferences nativeReferences;

        /// <summary>Tracks buttons before metadata or event registration can fail.</summary>
        private readonly List<object> createdButtons = new List<object>();

        /// <summary>Owning STA identity; no COM cleanup is dispatched from another thread.</summary>
        private readonly int ownerThread;


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
        /// <param name="showAbout">Action facultative ouvrant À propos sans dépendre du chat.</param>
        /// <param name="showCrashReport">Action facultative ouvrant le rapport de problème.</param>
        /// <param name="showUpdates">Action facultative ouvrant les mises à jour.</param>
        /// <param name="showHelp">Opens local help; null uses the packaged guide.</param>
        public VbeMenu(object vbe, Action showAssistant, Action showSettings, Action showGitHub, Action<string> editorAction = null, Action showAbout = null, Action showCrashReport = null, Action showUpdates = null, Action showHelp = null)
            : this(vbe, showAssistant, showSettings, showGitHub, editorAction, null, null, null, showAbout, showCrashReport, showUpdates, showHelp)
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
        /// <param name="showAbout">Action facultative ouvrant À propos sans dépendre du chat.</param>
        /// <param name="showCrashReport">Action facultative ouvrant le rapport de problème.</param>
        /// <param name="showUpdates">Action facultative ouvrant les mises à jour.</param>
        /// <param name="showHelp">Opens local help; null uses the packaged guide.</param>
        /// <param name="releaseAcquired">Balances one getter acquisition without finalizing borrowed owners.</param>
        /// <param name="indexedCollection">Identifies native one-based collections; defaults to COM collection detection.</param>
        internal VbeMenu(object vbe, Action showAssistant, Action showSettings, Action showGitHub,
            Action<string> editorAction, Action<object, Guid, int, Delegate> subscribe,
            Action<object, Guid, int, Delegate> unsubscribe, Action<object, Type> applyIcon, Action showAbout = null, Action showCrashReport = null, Action showUpdates = null, Action showHelp = null,
            Action<object> releaseAcquired = null, Func<object, bool> indexedCollection = null)
        {
            ownerThread = System.Threading.Thread.CurrentThread.ManagedThreadId;
            nativeReferences = new VbaTestComReferences(releaseAcquired,
                indexedCollection ?? (Func<object, bool>)Marshal.IsComObject);
            try
            {
                this.subscribe = subscribe ?? SubscribeDefault;
                this.unsubscribe = unsubscribe ?? ((button, iid, dispid, handler) =>
                    ComEventsHelper.Remove(button, iid, dispid, handler));
                this.applyIcon = applyIcon ?? SetIcon;
                dynamic view = FindPopupInScope(vbe, true, false, nativeReferences);
                dynamic tools = FindPopupInScope(vbe, false, false, nativeReferences);
                viewButton = CreateNativeButton((object)view);
                settingsButton = CreateNativeButton((object)tools);
                ((dynamic)viewButton).Caption = UiText.Get("VBAi assistant");
                ((dynamic)viewButton).Tag = "VBAi.Assistant";
                ((dynamic)settingsButton).Caption = UiText.Get("VBAi settings…");
                ((dynamic)settingsButton).Tag = "VBAi.Settings";
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
                    if (editorAction != null)
                    {
                        object modernButton = CreateNativeButton((object)view);
                        ((dynamic)modernButton).Caption = UiText.Get("VBAi editor");
                        ((dynamic)modernButton).Tag = "VBAi.ModernEditor";
                        ((dynamic)modernButton).TooltipText = UiText.Get("Open active module");
                        ClickHandler modernHandler = (object control, ref bool cancel) => { cancel = true; editorAction("/editor"); };
                        editorButtons.Add(Tuple.Create(modernButton, modernHandler));
                        dynamic testsButton = CreateNativeButton((object)view);
                        testsButton.Caption = UiText.Get("VBAi test explorer…");
                        testsButton.Tag = "VBAi.Tests";
                        ClickHandler testsHandler = (object control, ref bool cancel) => { cancel = true; editorAction("/tests"); };
                        editorButtons.Add(Tuple.Create((object)testsButton, testsHandler));
                        this.subscribe((object)testsButton, ClickInterface, 1, testsHandler);
                        this.applyIcon((object)testsButton, typeof(AboutWindow));
                        this.subscribe(modernButton, ClickInterface, 1, modernHandler);
                        this.applyIcon(modernButton, typeof(AboutWindow));
                    }
                    object gitButton = CreateNativeButton((object)view);
                    ((dynamic)gitButton).Caption = "GitHub VBAi…";
                    ((dynamic)gitButton).Tag = "VBAi.GitHub";
                    ((dynamic)gitButton).TooltipText = UiText.Get("Synchronize the active VBA project and manage its branches and checkpoints");
                    ClickHandler gitHandler = (object control, ref bool cancel) => { cancel = true; showGitHub(); };
                    editorButtons.Add(Tuple.Create(gitButton, gitHandler));
                    this.subscribe(gitButton, ClickInterface, 1, gitHandler);
                    this.applyIcon(gitButton, typeof(GitWindow));
                    if (showUpdates != null)
                    {
                        object updateButton = CreateNativeButton((object)tools);
                        ((dynamic)updateButton).Caption = UiText.Get("VBAi updates");
                        ((dynamic)updateButton).Tag = "VBAi.Updates";
                        ((dynamic)updateButton).TooltipText = UiText.Get("Check for updates");
                        ClickHandler updateHandler = (object control, ref bool cancel) => { cancel = true; showUpdates(); };
                        editorButtons.Add(Tuple.Create(updateButton, updateHandler));
                        this.subscribe(updateButton, ClickInterface, 1, updateHandler);
                        this.applyIcon(updateButton, typeof(UpdateWindow));
                    }
                    if (showCrashReport != null)
                    {
                        object reportButton = CreateNativeButton((object)tools);
                        ((dynamic)reportButton).Caption = UiText.Get("Report an issue");
                        ((dynamic)reportButton).Tag = "VBAi.CrashReport";
                        ((dynamic)reportButton).TooltipText = UiText.Get("GitHub first; Outlook if unavailable.");
                        ClickHandler reportHandler = (object control, ref bool cancel) => { cancel = true; showCrashReport(); };
                        editorButtons.Add(Tuple.Create(reportButton, reportHandler));
                        this.subscribe(reportButton, ClickInterface, 1, reportHandler);
                        this.applyIcon(reportButton, typeof(CrashReportWindow));
                    }
                    dynamic help = FindPopupInScope(vbe, false, true, nativeReferences);
                    if (help != null)
                    {
                        object helpButton = CreateNativeButton((object)help);
                        ((dynamic)helpButton).Caption = "VBAi · " + UiText.Get("User guide");
                        ((dynamic)helpButton).Tag = "VBAi.Help";
                        ((dynamic)helpButton).TooltipText = UiText.Get("Open the local VBAi user guide in the interface language");
                        ClickHandler helpHandler = (object control, ref bool cancel) => { cancel = true; (showHelp ?? (() => UiHelp.Open(null)))(); };
                        editorButtons.Add(Tuple.Create(helpButton, helpHandler));
                        this.subscribe(helpButton, ClickInterface, 1, helpHandler);
                        this.applyIcon(helpButton, typeof(AboutWindow));
                    }
                    if (showAbout != null)
                    {
                        object aboutButton = CreateNativeButton((object)tools);
                        ((dynamic)aboutButton).Caption = UiText.Get("About VBAi");
                        ((dynamic)aboutButton).Tag = "VBAi.About";
                        ((dynamic)aboutButton).TooltipText = UiText.Get("Version, technical details and project resources");
                        ClickHandler aboutHandler = (object control, ref bool cancel) => { cancel = true; showAbout(); };
                        editorButtons.Add(Tuple.Create(aboutButton, aboutHandler));
                        this.subscribe(aboutButton, ClickInterface, 1, aboutHandler);
                        this.applyIcon(aboutButton, typeof(AboutWindow));
                    }
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
                        foreach (dynamic bar in nativeReferences.Items((object)((dynamic)vbe).CommandBars))
                        {
                            string name = Normalize((string)bar.Name);
                            if (name != "code window" && name != "code window (break)" && name != "fenêtre code") continue;
                            foreach (var action in new[] { "/expliquer", "/corriger", "/refactoriser" })
                            {
                                string command = action;
                                object button = CreateNativeButton((object)bar);
                                ((dynamic)button).Caption = "VBAi · " + UiText.Get(action == "/expliquer" ? "Explain" : action == "/corriger" ? "Fix" : "Refactor");
                                ((dynamic)button).Tag = "VBAi." + action.Substring(1);
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
                    LoadLog.Write("VBE editor context menus unavailable: " + ex.ToString());
                }
            }
            catch
            {
                Dispose();
                throw;
            }
        }

        /// <summary>Convertit l’icône de la fenêtre en image et masque OLE pour le bouton Office.</summary>
        /// <param name="button">Bouton de barre de commande auquel appliquer l’image.</param>
        /// <param name="windowType">Type de formulaire fournissant la ressource d’icône.</param>
        private void SetIcon(object button, Type windowType)
        {
            try
            {
                using (var icon = ReadIcon(windowType))
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
            return AcquirePopup(application, view, false);
        }

        /// <summary>Returns one acquired popup; its caller owns the returned acquisition.</summary>
        private static object FindHelpMenu(object application)
        {
            return AcquirePopup(application, false, true);
        }

        /// <summary>Balances discovery even if releasing an unrelated getter fails.</summary>
        private static object AcquirePopup(object application, bool view, bool help)
        {
            using (var resultOwner = new VbaTestComReferences())
            {
                object result;
                using (var discovery = new VbaTestComReferences())
                {
                    result = FindPopupInScope(application, view, help, discovery);
                    if (result != null)
                    {
                        discovery.Transfer(result);
                        resultOwner.Add(result);
                    }
                }
                if (result != null) resultOwner.Transfer(result);
                return result;
            }
        }

        /// <summary>Uses one-based native collections; the borrowed VBE is never acquired.</summary>
        internal static object FindPopupInScope(object application, bool view, bool help, VbaTestComReferences references)
        {
            var names = new HashSet<string>(StringComparer.OrdinalIgnoreCase) {
                "?", "help", "aide", "ayuda", "hilfe", "ajuda", "guida", "ヘルプ", "도움말",
                "帮助", "說明", "说明", "справка", "تعليمات", "مساعدة", "सहायता"
            };
            foreach (dynamic bar in references.Items((object)((dynamic)application).CommandBars))
            {
                try { if ((int)bar.Type != 1) continue; } catch { continue; }
                foreach (dynamic item in references.Items((object)bar.Controls))
                {
                    string name;
                    try { name = Normalize((string)item.Caption); } catch { continue; }
                    if (help ? names.Contains(UiLanguages.NormalizeMenu(name)) : UiLanguages.IsMenu(name, view))
                        return item;
                }
            }
            if (help) return null;
            throw new InvalidOperationException("VBE menu not found: " + (view ? "View" : "Tools"));
        }

        /// <summary>Records both Controls and Add acquisitions before configuring the new button.</summary>
        private object CreateNativeButton(object popup)
        {
            dynamic controls = nativeReferences.Add((object)((dynamic)popup).Controls);
            object button = nativeReferences.Add((object)controls.Add(1, Missing.Value, Missing.Value, Missing.Value, true));
            createdButtons.Add(button);
            return button;
        }

        /// <summary>Attempts one deletion per distinct created button, including partial construction.</summary>
        private static void DeleteButtonOnce(object button, List<object> visited)
        {
            if (button == null || visited.Exists(item => ReferenceEquals(item, button))) return;
            visited.Add(button);
            ((dynamic)button).Delete();
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
            if (ownerThread != 0 && ownerThread != System.Threading.Thread.CurrentThread.ManagedThreadId)
                throw new InvalidOperationException("VBE menu cleanup must remain on its owning thread.");
            disposed = true;
            var deletedButtons = new List<object>();
            foreach (var item in editorButtons)
            {
                try { unsubscribe(item.Item1, ClickInterface, 1, item.Item2); } catch { }
                try { DeleteButtonOnce(item.Item1, deletedButtons); } catch { }
            }
            editorButtons.Clear();
            if (viewButton != null)
            {
                try { if (viewHandler != null) unsubscribe(viewButton, ClickInterface, 1, viewHandler); } catch { }
                try { DeleteButtonOnce(viewButton, deletedButtons); } catch { }
            }
            if (settingsButton != null)
            {
                try { if (settingsHandler != null) unsubscribe(settingsButton, ClickInterface, 1, settingsHandler); } catch { }
                try { DeleteButtonOnce(settingsButton, deletedButtons); } catch { }
            }
            foreach (var button in createdButtons)
                try { DeleteButtonOnce(button, deletedButtons); } catch { }
            createdButtons.Clear();
            viewButton = null;
            settingsButton = null;
            viewHandler = null;
            settingsHandler = null;
            foreach (var image in menuImages) image.Dispose();
            menuImages.Clear();
            nativeReferences?.Dispose();
        }
    }
}

using System;
using System.Collections.Generic;
using System.Linq;
using System.Security.Cryptography;
using System.Text;
using System.Web.Script.Serialization;

namespace CodexVBE
{
    /// <summary>État géométrique et identité d’une fenêtre de document VBE visible.</summary>
internal sealed class EditorWindowBounds
    {
        /// <summary>Légende native de la fenêtre.</summary>
        /// <value>Caption lue sur la fenêtre VBE.</value>
public string Caption { get; set; }
        /// <summary>Identité d’empreinte de la fenêtre et de son document.</summary>
        /// <value>SHA-256 de l’identité du composant ou de l’Explorateur d’objets.</value>
public string Identity { get; set; }
        /// <summary>Type natif de fenêtre VBE.</summary>
        /// <value>Valeur Type lue depuis VBIDE.</value>
public int Type { get; set; }
        /// <summary>État natif de la fenêtre.</summary>
        /// <value>Valeur WindowState observée.</value>
public int State { get; set; }
        /// <summary>Coordonnée gauche de la fenêtre.</summary>
        /// <value>Position native Left.</value>
public int Left { get; set; }
        /// <summary>Coordonnée supérieure de la fenêtre.</summary>
        /// <value>Position native Top.</value>
public int Top { get; set; }
        /// <summary>Largeur native de la fenêtre.</summary>
        /// <value>Valeur Width en pixels de l’hôte.</value>
public int Width { get; set; }
        /// <summary>Hauteur native de la fenêtre.</summary>
        /// <value>Valeur Height en pixels de l’hôte.</value>
public int Height { get; set; }
    }
    /// <summary>Inventorie et organise les fenêtres de documents ouvertes dans l’éditeur VBE.</summary>
internal sealed partial class VbeDebug
    {
        /// <summary>Associe une fenêtre visible à une identité stable de projet et de composant.</summary>
        /// <param name="window">Fenêtre de document à identifier.</param>
        /// <param name="type">Type natif de la fenêtre.</param>
        /// <returns>Empreinte de l’identité trouvée sans ambiguïté.</returns>
private string DocumentIdentity(dynamic window, int type)
        {
            var matches = new HashSet<string>(StringComparer.Ordinal);
            if (type == 0)
            {
                foreach (dynamic pane in vbe.CodePanes)
                    if (string.Equals((string)pane.Window.Caption, (string)window.Caption, StringComparison.Ordinal))
                    {
                        dynamic component = pane.CodeModule.Parent;
                        matches.Add(ComponentIdentity(component.Collection.Parent, component));
                    }
            }
            else if (type == 1)
            {
                foreach (dynamic project in vbe.VBProjects)
                    foreach (dynamic component in project.VBComponents)
                    {
                        if (!(bool)component.HasOpenDesigner) continue;
                        if (string.Equals((string)component.DesignerWindow().Caption, (string)window.Caption, StringComparison.Ordinal))
                            matches.Add(ComponentIdentity(project, component));
                    }
            }
            else matches.Add("ObjectBrowser");
            if (matches.Count != 1) throw new InvalidOperationException("The document window cannot be mapped to one project/component identity.");
            using (var sha = SHA256.Create()) return BitConverter.ToString(sha.ComputeHash(Encoding.UTF8.GetBytes(type + ":" + matches.Single()))).Replace("-", "").ToLowerInvariant();
        }
        /// <summary>Forme l’identité d’un composant depuis le chemin ou le nom du projet et le nom du composant.</summary>
        /// <param name="project">Projet qui contient le composant.</param>
        /// <param name="component">Composant à identifier.</param>
        /// <returns>Représentation sérialisée de l’identité du composant.</returns>
private static string ComponentIdentity(dynamic project, dynamic component)
        {
            string path = null; try { path = (string)project.FileName; } catch { }
            return new JavaScriptSerializer().Serialize(new { Project = string.IsNullOrWhiteSpace(path) ? (string)project.Name : path, Module = (string)component.Name });
        }
        /// <summary>Retourne les fenêtres de document visibles et la version de leur disposition.</summary>
        /// <returns>Instantanés géométriques triés et empreinte globale de disposition.</returns>
public object EditorLayout()
        {
            var windows = ReadEditorBounds();
            return new { Windows = windows, WindowVersion = LayoutHash(windows), Scope = "All visible VBE document windows, across projects" };
        }
        /// <summary>Exécute la commande native de cascade ou de mosaïque après contrôle de la disposition.</summary>
        /// <param name="request">Action, légende exacte de commande et version attendue des fenêtres.</param>
        /// <returns>Résultat de l’appel natif, fenêtres avant/après et état de vérification géométrique.</returns>
public object ArrangeEditorWindows(Request request)
        {
            int id;
            switch (request.Action)
            {
                case "cascade": id = 1826; break;
                case "tile_vertical": id = 2561; break;
                case "tile_horizontal": id = 2562; break;
                default: throw new ArgumentException("Use cascade, tile_vertical or tile_horizontal.");
            }
            var before = ReadEditorBounds();
            if (before.Length < 2 || before.Length > 64) throw new InvalidOperationException("Arrangement requires 2-64 visible document windows.");
            if (string.IsNullOrWhiteSpace(request.ExpectedWindowVersion) || !string.Equals(request.ExpectedWindowVersion, LayoutHash(before), StringComparison.OrdinalIgnoreCase))
                throw new InvalidOperationException("The document layout changed. Read editor_layout again.");
            var command = EnumerateCommands().FirstOrDefault(x => x.Id == id && x.Enabled && string.Equals(x.Caption, request.ControlCaption, StringComparison.Ordinal));
            if (command == null) throw new InvalidOperationException("The exact native arrangement command is absent or disabled. Read list_commands again.");
            string error = null;
            try { ((dynamic)command.Control).Execute(); } catch (Exception ex) { error = ex.Message; }
            EditorWindowBounds[] after = null;
            try { after = ReadEditorBounds(); } catch (Exception ex) { error = error ?? ex.Message; }
            bool sameWindows = after != null && before.Select(WindowKey).SequenceEqual(after.Select(WindowKey));
            bool verified = error == null && sameWindows && VerifyArrangement(request.Action, after);
            return new { request.Action, CommandId = id, Applied = error == null ? (bool?)true : null,
                Verified = verified, VerificationPending = !verified, Before = before, After = after,
                WindowVersion = after == null ? null : LayoutHash(after), NativeError = error,
                Scope = "All visible VBE document windows, across projects", PersistenceVerified = false };
        }
        /// <summary>Lit les fenêtres de document visibles en excluant les cadres liés et les types hors périmètre.</summary>
        /// <returns>Fenêtres ordonnées par identité; échoue si leur identité n’est pas unique.</returns>
private EditorWindowBounds[] ReadEditorBounds()
        {
            var rows = new List<EditorWindowBounds>();
            foreach (dynamic window in vbe.Windows)
            {
                int type = (int)window.Type;
                if (type < 0 || type > 2 || !(bool)window.Visible || window.LinkedWindowFrame != null) continue;
                rows.Add(new EditorWindowBounds { Caption = (string)window.Caption, Identity = DocumentIdentity(window, type), Type = type, State = (int)window.WindowState,
                    Left = (int)window.Left, Top = (int)window.Top, Width = (int)window.Width, Height = (int)window.Height });
            }
            if (rows.Select(WindowKey).Distinct(StringComparer.Ordinal).Count() != rows.Count)
                throw new InvalidOperationException("Visible document identities are ambiguous.");
            return rows.OrderBy(WindowKey, StringComparer.Ordinal).ToArray();
        }
        /// <summary>Construit la clé stable d’une fenêtre à partir de son type et de son identité.</summary>
        /// <param name="window">Fenêtre à identifier.</param>
        /// <returns>Clé textuelle déterministe.</returns>
private static string WindowKey(EditorWindowBounds window) { return window.Type + ":" + window.Identity; }
        /// <summary>Calcule l’empreinte d’un ensemble trié d’états de fenêtres.</summary>
        /// <param name="rows">Géométries observées.</param>
        /// <returns>Empreinte SHA-256 de leur représentation JSON.</returns>
private static string LayoutHash(EditorWindowBounds[] rows)
        {
            using (var sha = SHA256.Create()) return BitConverter.ToString(sha.ComputeHash(Encoding.UTF8.GetBytes(new JavaScriptSerializer().Serialize(rows)))).Replace("-", "").ToLowerInvariant();
        }
        /// <summary>Vérifie la géométrie résultante d’une cascade ou d’une mosaïque sans modifier les fenêtres.</summary>
        /// <param name="action">Commande attendue : cascade, tile_vertical ou tile_horizontal.</param>
        /// <param name="rows">Géométries relues après la commande native.</param>
        /// <returns><see langword="true"/> si les tailles et relations géométriques correspondent à l’action.</returns>
internal static bool VerifyArrangement(string action, EditorWindowBounds[] rows)
        {
            if (rows == null || rows.Length < 2 || rows.Any(x => x.State != 0 || x.Width <= 0 || x.Height <= 0)) return false;
            if (rows.Max(x => x.Width) - rows.Min(x => x.Width) > 2 || rows.Max(x => x.Height) - rows.Min(x => x.Height) > 2) return false;
            if (action == "cascade")
            {
                var ordered = rows.OrderBy(x => x.Left).ToArray();
                return ordered.Zip(ordered.Skip(1), (a, b) => b.Left > a.Left && b.Top > a.Top && b.Left < (long)a.Left + a.Width && b.Top < (long)a.Top + a.Height).All(x => x);
            }
            if (action != "tile_vertical" && action != "tile_horizontal") return false;
            for (int i = 0; i < rows.Length; i++)
                for (int j = i + 1; j < rows.Length; j++)
                {
                    var a = rows[i]; var b = rows[j];
                    long overlapWidth = Math.Min((long)a.Left + a.Width, (long)b.Left + b.Width) - Math.Max(a.Left, b.Left);
                    long overlapHeight = Math.Min((long)a.Top + a.Height, (long)b.Top + b.Height) - Math.Max(a.Top, b.Top);
                    if (overlapWidth > 2 && overlapHeight > 2) return false;
                }
            long width = rows.Max(x => (long)x.Left + x.Width) - rows.Min(x => x.Left);
            long height = rows.Max(x => (long)x.Top + x.Height) - rows.Min(x => x.Top);
            double area = rows.Sum(x => (double)x.Width * x.Height);
            if (area < 0.95 * width * height) return false;
            int columns = rows.Select(x => x.Left).Distinct().Count(), lines = rows.Select(x => x.Top).Distinct().Count();
            return action == "tile_vertical" ? columns >= lines : lines >= columns;
        }
    }
}

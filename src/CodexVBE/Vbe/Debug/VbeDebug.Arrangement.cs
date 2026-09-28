using System;
using System.Collections.Generic;
using System.Linq;
using System.Security.Cryptography;
using System.Text;
using System.Web.Script.Serialization;

namespace CodexVBE
{
    internal sealed class EditorWindowBounds
    {
        public string Caption { get; set; }
        public string Identity { get; set; }
        public int Type { get; set; }
        public int State { get; set; }
        public int Left { get; set; }
        public int Top { get; set; }
        public int Width { get; set; }
        public int Height { get; set; }
    }
    internal sealed partial class VbeDebug
    {
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
        private static string ComponentIdentity(dynamic project, dynamic component)
        {
            string path = null; try { path = (string)project.FileName; } catch { }
            return new JavaScriptSerializer().Serialize(new { Project = string.IsNullOrWhiteSpace(path) ? (string)project.Name : path, Module = (string)component.Name });
        }
        public object EditorLayout()
        {
            var windows = ReadEditorBounds();
            return new { Windows = windows, WindowVersion = LayoutHash(windows), Scope = "All visible VBE document windows, across projects" };
        }
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
        private static string WindowKey(EditorWindowBounds window) { return window.Type + ":" + window.Identity; }
        private static string LayoutHash(EditorWindowBounds[] rows)
        {
            using (var sha = SHA256.Create()) return BitConverter.ToString(sha.ComputeHash(Encoding.UTF8.GetBytes(new JavaScriptSerializer().Serialize(rows)))).Replace("-", "").ToLowerInvariant();
        }
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

using System;
using System.Collections.Generic;
using System.Drawing;
using System.Drawing.Drawing2D;
using System.Globalization;
using System.IO;
using System.Text.RegularExpressions;
using System.Xml;

namespace VBAi
{

    /// <summary>Draws the bundled original SVG paths at the current DPI and foreground color.</summary>
    internal static class UiCommandIcons
    {

        /// <summary>Caches parsed SVG path geometry by symbol; access is protected by the dictionary lock.</summary>
        private static readonly Dictionary<UiSymbol, GraphicsPath> Paths = new Dictionary<UiSymbol, GraphicsPath>();

        /// <summary>Renders an icon without bitmap scaling or theme-specific asset copies.</summary>
        /// <param name="graphics">Drawing context; ownership remains with the caller.</param>
        /// <param name="symbol">Bundled icon resource to draw.</param>
        /// <param name="bounds">Available drawing rectangle.</param>
        /// <param name="color">Foreground stroke color.</param>
        /// <param name="dpi">Display density used to scale logical dimensions.</param>
        /// <returns><see langword="true"/> when an icon path was cached or drawn; false when the resource has no path.</returns>
        internal static bool Draw(Graphics graphics, UiSymbol symbol, Rectangle bounds, Color color, int dpi)
        {
            lock (Paths)
            {
                if (!Paths.TryGetValue(symbol, out var path))
                {
                    using (var stream = typeof(UiCommandIcons).Assembly.GetManifestResourceStream("VBAi.CommandIcons." + symbol + ".svg"))
                    {
                        path = ReadPath(stream);
                    }
                    if (path == null) return false;
                    Paths.Add(symbol, path);
                }
                float size = Math.Min(Math.Min(bounds.Width, bounds.Height) - 4, 18F * dpi / 96F);
                if (size <= 0) return true;
                var state = graphics.Save();
                try
                {
                    graphics.TranslateTransform(bounds.X + (bounds.Width - size) / 2, bounds.Y + (bounds.Height - size) / 2);
                    graphics.ScaleTransform(size / 24, size / 24);
                    using (var pen = new Pen(color, 1.7F) { StartCap = LineCap.Round, EndCap = LineCap.Round, LineJoin = LineJoin.Round }) graphics.DrawPath(pen, path);
                }
                finally { graphics.Restore(state); }
                return true;
            }
        }

        /// <summary>Reads the first supported path without resolving external XML resources.</summary>
        /// <param name="stream">Icon resource stream, or null when the resource is absent.</param>
        /// <returns>The parsed path, or null when the resource contains no path.</returns>
        private static GraphicsPath ReadPath(Stream stream)
        {
            if (stream == null) return null;
            using (var reader = XmlReader.Create(stream, new XmlReaderSettings { DtdProcessing = DtdProcessing.Prohibit, XmlResolver = null }))
                while (reader.Read()) if (reader.NodeType == XmlNodeType.Element && reader.LocalName == "path") return Parse(reader.GetAttribute("d"));
            return null;
        }
        // The bundled SVGs intentionally use absolute M/L/C/Z commands only.
        /// <summary>Parses the absolute M, L, C and Z commands used by bundled SVG paths.</summary>
        /// <param name="data">SVG path data containing only absolute M, L, C, and Z commands.</param>
        /// <returns>A disposable path containing the parsed figures.</returns>
        /// <exception cref="FormatException">The path contains an unsupported command or malformed numeric data.</exception>
        private static GraphicsPath Parse(string data)
        {
            var tokens = Regex.Matches(data, @"[A-Za-z]|-?\d+(?:\.\d+)?");
            var path = new GraphicsPath(); int index = 0; PointF current = PointF.Empty;
            Func<float> number = () => float.Parse(tokens[index++].Value, CultureInfo.InvariantCulture);
            Func<PointF> point = () => new PointF(number(), number());
            try
            {
                while (index < tokens.Count)
                {
                    string command = tokens[index++].Value;
                    if (command == "M") { path.StartFigure(); current = point(); }
                    else if (command == "L") { var next = point(); path.AddLine(current, next); current = next; }
                    else if (command == "C") { var a = point(); var b = point(); var end = point(); path.AddBezier(current, a, b, end); current = end; }
                    else if (command == "Z") path.CloseFigure();
                    else throw new FormatException("Unsupported command in bundled icon: " + command);
                }
                return path;
            }
            catch { path.Dispose(); throw; }
        }
    }
}

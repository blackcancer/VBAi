using System;
using System.Linq;

namespace CodexVBE
{
    internal sealed class FormLayoutBox
    {
        public string Path { get; set; }
        public double Left { get; set; }
        public double Top { get; set; }
        public double Width { get; set; }
        public double Height { get; set; }
    }
    internal static class FormLayoutPlan
    {
        internal static FormLayoutBox[] Create(FormLayoutBox[] input, string action, double step, double parentWidth, double parentHeight)
        {
            if (input == null || input.Length < 2 || input.Length > 64) throw new ArgumentException("Select 2 to 64 controls in the same container.");
            if (double.IsNaN(parentWidth) || double.IsInfinity(parentWidth) || parentWidth <= 0 ||
                double.IsNaN(parentHeight) || double.IsInfinity(parentHeight) || parentHeight <= 0)
                throw new ArgumentException("Container dimensions must be finite and positive.");
            if (input.Any(x => x == null || new[] { x.Left, x.Top, x.Width, x.Height }.Any(n => double.IsNaN(n) || double.IsInfinity(n)) || x.Width <= 0 || x.Height <= 0))
                throw new ArgumentException("Invalid source control geometry.");
            var result = input.Select(x => new FormLayoutBox { Path = x.Path, Left = x.Left, Top = x.Top, Width = x.Width, Height = x.Height }).ToArray();
            var anchor = input[0];
            if (action == "space_horizontal" || action == "space_vertical" ||
                action == "increase_horizontal_spacing" || action == "increase_vertical_spacing" ||
                action == "decrease_horizontal_spacing" || action == "decrease_vertical_spacing")
            {
                if (double.IsNaN(step) || double.IsInfinity(step) || step < 0 || step > 1000)
                    throw new ArgumentException("Spacing must be a finite number in [0, 1000] points.");
                bool horizontal = action.Contains("horizontal");
                bool fixedGap = action.StartsWith("space_", StringComparison.Ordinal);
                double delta = action.StartsWith("decrease_", StringComparison.Ordinal) ? -step : step;
                var ordered = result.OrderBy(x => horizontal ? x.Left : x.Top).ThenBy(x => x.Path, StringComparer.Ordinal).ToArray();
                double previousOriginalEnd = horizontal ? ordered[0].Left + ordered[0].Width : ordered[0].Top + ordered[0].Height;
                for (int i = 1; i < ordered.Length; i++)
                {
                    var item = ordered[i]; var previous = ordered[i - 1];
                    double originalStart = horizontal ? item.Left : item.Top;
                    double gap = fixedGap ? step : originalStart - previousOriginalEnd + delta;
                    previousOriginalEnd = originalStart + (horizontal ? item.Width : item.Height);
                    if (gap < 0) throw new InvalidOperationException("The requested spacing would overlap selected controls.");
                    if (horizontal) item.Left = previous.Left + previous.Width + gap;
                    else item.Top = previous.Top + previous.Height + gap;
                }
            }
            else if (action == "distribute_horizontal" || action == "distribute_vertical")
            {
                if (input.Length < 3) throw new ArgumentException("Distribution requires at least three controls.");
                bool horizontal = action == "distribute_horizontal";
                var ordered = result.OrderBy(x => horizontal ? x.Left : x.Top).ThenBy(x => x.Path, StringComparer.Ordinal).ToArray();
                double start = horizontal ? ordered[0].Left : ordered[0].Top;
                var last = ordered[ordered.Length - 1];
                double end = horizontal ? last.Left + last.Width : last.Top + last.Height;
                double gap = (end - start - ordered.Sum(x => horizontal ? x.Width : x.Height)) / (ordered.Length - 1);
                if (gap < 0) throw new InvalidOperationException("The selected bounds cannot contain these controls without overlap.");
                foreach (var item in ordered) { if (horizontal) item.Left = start; else item.Top = start; start += (horizontal ? item.Width : item.Height) + gap; }
            }
            else
            {
                double left = input.Min(x => x.Left), top = input.Min(x => x.Top);
                double right = input.Max(x => x.Left + x.Width), bottom = input.Max(x => x.Top + x.Height);
                foreach (var item in result)
                    switch (action)
                    {
                        case "align_left": item.Left = anchor.Left; break;
                        case "align_right": item.Left = anchor.Left + anchor.Width - item.Width; break;
                        case "align_centers": item.Left = anchor.Left + (anchor.Width - item.Width) / 2; break;
                        case "align_middles": item.Top = anchor.Top + (anchor.Height - item.Height) / 2; break;
                        case "align_top": item.Top = anchor.Top; break;
                        case "align_bottom": item.Top = anchor.Top + anchor.Height - item.Height; break;
                        case "same_width": item.Width = anchor.Width; break;
                        case "same_height": item.Height = anchor.Height; break;
                        case "same_size": item.Width = anchor.Width; item.Height = anchor.Height; break;
                        case "center_horizontal": item.Left += (parentWidth - (right - left)) / 2 - left; break;
                        case "center_vertical": item.Top += (parentHeight - (bottom - top)) / 2 - top; break;
                        case "snap_grid":
                            if (step <= 0 || step > 100 || double.IsNaN(step)) throw new ArgumentException("Grid spacing must be in (0, 100] points.");
                            item.Left = Math.Round(item.Left / step, MidpointRounding.AwayFromZero) * step;
                            item.Top = Math.Round(item.Top / step, MidpointRounding.AwayFromZero) * step; break;
                        default: throw new ArgumentException("Unknown layout action.");
                    }
            }
            foreach (var item in result)
                // Sizes are copied from validated input or its validated anchor; only positions can leave the container.
                if (new[] { item.Left, item.Top, item.Width, item.Height }.Any(x => double.IsNaN(x) || double.IsInfinity(x)) || item.Left < 0 || item.Top < 0 || item.Left + item.Width > parentWidth + 0.1 || item.Top + item.Height > parentHeight + 0.1)
                    throw new InvalidOperationException("The layout would place a control outside its container.");
            return result;
        }
    }
}

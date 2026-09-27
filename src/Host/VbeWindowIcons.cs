using System;
using System.Drawing;

namespace CodexVBE
{
    internal static class VbeWindowIcons
    {
        public static Icon Icon(string name)
        {
            using (var stream = typeof(VbeWindowIcons).Assembly.GetManifestResourceStream("CodexVBE.Icons." + name + ".ico"))
            {
                if (stream == null) return null;
                using (var source = new Icon(stream)) return (Icon)source.Clone();
            }
        }

        public static Image Image(string name)
        {
            using (var stream = typeof(VbeWindowIcons).Assembly.GetManifestResourceStream("CodexVBE.Icons." + name + ".png"))
            {
                if (stream == null) return null;
                using (var source = System.Drawing.Image.FromStream(stream)) return new Bitmap(source);
            }
        }
    }
}

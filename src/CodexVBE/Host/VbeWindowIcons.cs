using System;
using System.Drawing;

namespace CodexVBE
{
    /// <summary>Charge les icônes et images du complément depuis ses ressources embarquées.</summary>
    internal static class VbeWindowIcons
    {
        /// <summary>Charge une icône ICO par nom de ressource et retourne une copie indépendante.</summary>
        /// <param name="name">Nom logique de la ressource, sans extension.</param>
        /// <returns>Icône clonée, ou null si la ressource est absente.</returns>
        public static Icon Icon(string name)
        {
            using (var stream = typeof(VbeWindowIcons).Assembly.GetManifestResourceStream("CodexVBE.Icons." + name + ".ico"))
            {
                if (stream == null) return null;
                using (var source = new Icon(stream)) return (Icon)source.Clone();
            }
        }

        /// <summary>Charge une image PNG par nom de ressource et retourne une bitmap indépendante.</summary>
        /// <param name="name">Nom logique de la ressource, sans extension.</param>
        /// <returns>Bitmap copiée, ou null si la ressource est absente.</returns>
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

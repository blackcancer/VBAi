using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Windows.Forms;

namespace CodexVBE
{
    /// <summary>Copie en mémoire les formats de presse-papiers MSForms qui peuvent être restaurés sans perte vérifiable.</summary>
    internal sealed class DesignerClipboardBackup
    {
        /// <summary>Taille maximale totale acceptée pour les données sauvegardées.</summary>
        internal const int MaximumBytes = 8 * 1024 * 1024;
        /// <summary>Données copiées par format, avec les flux mémorisés sous forme d’octets.</summary>
        private readonly Dictionary<string, object> values = new Dictionary<string, object>(StringComparer.Ordinal);
        /// <summary>Formats source omis car leur valeur n’était pas disponible.</summary>
        /// <value>Noms des formats ignorés pendant la capture.</value>
        public string[] OmittedFormats { get; private set; }
        /// <summary>Taille totale des données reconnues et copiées.</summary>
        /// <value>Nombre d’octets estimé ou lu pendant la capture.</value>
        public int ByteCount { get; private set; }
        /// <summary>Capture les formats disponibles en vérifiant la limite et la présence du flux MS Forms Bag.</summary>
        /// <param name="source">Données du presse-papiers avant mutation du Designer.</param>
        /// <returns>Copie pouvant recréer les formats pris en charge.</returns>
        /// <exception cref="InvalidOperationException">Le presse-papiers est vide, dépasse la limite ou contient un format non pris en charge.</exception>
        internal static DesignerClipboardBackup Capture(IDataObject source)
        {
            if (source == null) throw new InvalidOperationException("No Designer clipboard data; cut was not attempted.");
            var result = new DesignerClipboardBackup(); var omitted = new List<string>();
            foreach (string format in source.GetFormats(false))
            {
                object value = source.GetData(format, false);
                if (value == null) { omitted.Add(format); continue; }
                long size;
                if (value is MemoryStream stream) size = stream.Length;
                else if (value is string text) size = (long)text.Length * 2;
                else throw new InvalidOperationException("Unsupported Designer clipboard format: " + format + "; cut was not attempted.");
                if (size > MaximumBytes - result.ByteCount) throw new InvalidOperationException("Designer clipboard exceeds 8 MiB; cut was not attempted.");
                result.ByteCount += (int)size;
                result.values.Add(format, value is MemoryStream bytes ? (object)bytes.ToArray() : value);
            }
            if (!result.values.TryGetValue("MS Forms Bag", out object bag) || !(bag is byte[] data) || data.Length == 0)
                throw new InvalidOperationException("MS Forms Bag unavailable; cut was not attempted.");
            result.OmittedFormats = omitted.ToArray();
            return result;
        }
        /// <summary>Recrée un objet WinForms IDataObject avec les formats copiés.</summary>
        /// <returns>Objet prêt à être écrit dans le presse-papiers.</returns>
        internal DataObject CreateDataObject()
        {
            var data = new DataObject();
            foreach (var entry in values)
                data.SetData(entry.Key, false, entry.Value is byte[] bytes ? (object)new MemoryStream(bytes, false) : entry.Value);
            return data;
        }
        /// <summary>Compare les formats sauvegardés aux données relues, octet par octet pour les flux.</summary>
        /// <param name="source">Objet du presse-papiers après écriture de la sauvegarde.</param>
        /// <returns><see langword="true"/> si chaque valeur capturée correspond à la relecture.</returns>
        internal bool Matches(IDataObject source)
        {
            if (source == null) return false;
            foreach (var entry in values)
            {
                object actual = source.GetData(entry.Key, false);
                if (entry.Value is byte[] expected)
                {
                    if (!(actual is MemoryStream stream) || !expected.SequenceEqual(stream.ToArray())) return false;
                }
                else if (!Equals(entry.Value, actual)) return false;
            }
            return true;
        }
    }
}

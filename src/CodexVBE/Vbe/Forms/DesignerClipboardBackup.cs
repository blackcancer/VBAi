using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Windows.Forms;

namespace CodexVBE
{
    internal sealed class DesignerClipboardBackup
    {
        internal const int MaximumBytes = 8 * 1024 * 1024;
        private readonly Dictionary<string, object> values = new Dictionary<string, object>(StringComparer.Ordinal);
        public string[] OmittedFormats { get; private set; }
        public int ByteCount { get; private set; }
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
        internal DataObject CreateDataObject()
        {
            var data = new DataObject();
            foreach (var entry in values)
                data.SetData(entry.Key, false, entry.Value is byte[] bytes ? (object)new MemoryStream(bytes, false) : entry.Value);
            return data;
        }
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

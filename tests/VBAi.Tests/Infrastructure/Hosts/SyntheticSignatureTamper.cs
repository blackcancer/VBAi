using System;
using System.Collections;
using System.Collections.Generic;
using System.IO;
using System.IO.Compression;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Text.RegularExpressions;

namespace VBAi.Tests.Integration
{
    /// <summary>Creates an owned negative fixture by changing one literal in saved VBA source, without asking Office to remove the signature.</summary>
    internal static class SyntheticSignatureTamper
    {
        private const int Limit = 1048576;
        private const string ProjectEntry = "xl/vbaProject.bin";
        private const BindingFlags PrivateInstance = BindingFlags.Instance | BindingFlags.NonPublic;

        /// <summary>Copies a synthetic signed workbook, retaining every ZIP entry payload except its one mutated VBA source byte.</summary>
        internal static string CreateCopy(string original, string resultDirectory)
        {
            original = Path.GetFullPath(original);
            resultDirectory = Path.GetFullPath(resultDirectory);
            if (Path.GetFileName(original) != "SyntheticSignature.xlsm" ||
                !original.StartsWith(resultDirectory.TrimEnd(Path.DirectorySeparatorChar) + Path.DirectorySeparatorChar, StringComparison.OrdinalIgnoreCase))
                throw new InvalidOperationException("Only the campaign's synthetic signature workbook is admitted.");
            string output = Path.Combine(resultDirectory, "TamperedSyntheticSignature.xlsm");
            using (var input = File.OpenRead(original))
            using (var archive = new ZipArchive(input, ZipArchiveMode.Read))
            {
                var entries = archive.Entries.ToArray();
                if (entries.Length > 256 || entries.Select(e => e.FullName).Distinct(StringComparer.Ordinal).Count() != entries.Length)
                    throw new InvalidDataException("Bounded unique ZIP entry names required.");
                var project = entries.Single(e => e.FullName == ProjectEntry);
                byte[] projectBytes = ReadEntry(project);
                byte[] patched = PatchProject(projectBytes);
                // Validate the source mutation before creating any output file.
                using (var destination = new FileStream(output, FileMode.CreateNew, FileAccess.ReadWrite, FileShare.None))
                using (var target = new ZipArchive(destination, ZipArchiveMode.Create))
                {
                    foreach (var entry in entries)
                    {
                        byte[] data = entry.FullName == ProjectEntry ? patched : ReadEntry(entry);
                        var copied = target.CreateEntry(entry.FullName, CompressionLevel.Optimal);
                        copied.LastWriteTime = entry.LastWriteTime;
                        copied.ExternalAttributes = entry.ExternalAttributes;
                        using (var stream = copied.Open()) stream.Write(data, 0, data.Length);
                    }
                }
            }
            // Strongest-present signature payloads must stay byte-identical; changing source is the only payload mutation.
            VerifyCopy(original, output);
            return output;
        }

        private static byte[] ReadEntry(ZipArchiveEntry entry)
        {
            if (entry.Length > Limit) throw new InvalidDataException("Only bounded synthetic entry payloads are admitted.");
            using (var input = entry.Open())
            using (var output = new MemoryStream())
            {
                byte[] buffer = new byte[4096]; int count;
                while ((count = input.Read(buffer, 0, buffer.Length)) != 0)
                {
                    if (output.Length + count > Limit) throw new InvalidDataException("Entry exceeds its admission bound.");
                    output.Write(buffer, 0, count);
                }
                if (output.Length != entry.Length) throw new InvalidDataException("Entry extent changed.");
                return output.ToArray();
            }
        }

        private static Dictionary<string, byte[]> Streams(byte[] bytes)
        {
            if (bytes == null || bytes.Length > Limit) throw new InvalidDataException("Bounded CFB required.");
            Type type = typeof(FormResourcePreflight).GetNestedType("CompoundFile", BindingFlags.NonPublic);
            object file = type.GetConstructor(PrivateInstance, null, new[] { typeof(byte[]), typeof(int), typeof(int) }, null)
                .Invoke(new object[] { bytes, 0, bytes.Length });
            type.GetMethod("Validate", PrivateInstance).Invoke(file, null);
            var result = new Dictionary<string, byte[]>(StringComparer.Ordinal);
            foreach (object entry in (IEnumerable)type.GetField("entries", PrivateInstance).GetValue(file))
            {
                Type entryType = entry.GetType();
                if ((byte)entryType.GetField("Kind", PrivateInstance).GetValue(entry) != 2) continue;
                string path = (string)entryType.GetField("Path", PrivateInstance).GetValue(entry);
                result.Add(path, (byte[])entryType.GetField("Data", PrivateInstance).GetValue(entry));
            }
            return result;
        }

        /// <summary>Produces a single source literal mutation and proves that all other CFB streams remain unchanged.</summary>
        internal static byte[] PatchProject(byte[] bytes)
        {
            var streams = Streams(bytes);
            byte[] directory = Decompress(streams["/VBA/dir"], out _);
            byte[] module = streams["/VBA/SignatureFixture"];
            int start = ModuleOffset(directory);
            if (start < 0 || start >= module.Length) throw new InvalidDataException("Source offset is outside the module.");
            byte[] compressed = module.Skip(start).ToArray();
            int[] origins;
            byte[] source = Decompress(compressed, out origins);
            if (source.Any(b => b > 127)) throw new InvalidDataException("Only the fixture's ASCII source is admitted.");
            string text = Encoding.ASCII.GetString(source);
            if (!text.StartsWith("Attribute VB_Name = \"SignatureFixture\"\r\n", StringComparison.Ordinal))
                throw new InvalidDataException("Unexpected module identity.");
            var matches = Regex.Matches(text, "Private Const Marker As String = \"SIGNATURE_([a-f0-9]{32})\"");
            if (matches.Count != 1) throw new InvalidDataException("Exactly one synthetic source marker required.");
            Group nonce = matches[0].Groups[1];
            for (int index = nonce.Index; index < nonce.Index + nonce.Length; index++)
            {
                int literal = origins[index];
                byte replacement = source[index] == (byte)'a' ? (byte)'b' : (byte)'a';
                byte[] changedContainer = (byte[])compressed.Clone(); changedContainer[literal] = replacement;
                byte[] changedSource = Decompress(changedContainer, out _);
                byte[] expected = (byte[])source.Clone(); expected[index] = replacement;
                if (!changedSource.SequenceEqual(expected)) continue; // A copied literal may appear in more than one source position.
                byte[] changedModule = (byte[])module.Clone(); changedModule[start + literal] = replacement;
                int contextStart = Math.Max(0, start + literal - 8);
                byte[] context = module.Skip(contextStart).Take(17).ToArray();
                int physical = UniqueIndex(bytes, context);
                if (physical < 0) continue; // Never guess a sector layout or patch an ambiguous byte.
                byte[] changedFile = (byte[])bytes.Clone();
                changedFile[physical + start + literal - contextStart] = replacement;
                var after = Streams(changedFile);
                if (after.Count != streams.Count || !after.Keys.OrderBy(k => k).SequenceEqual(streams.Keys.OrderBy(k => k)))
                    throw new InvalidDataException("CFB stream inventory changed.");
                foreach (var entry in streams)
                    if (!after[entry.Key].SequenceEqual(entry.Key == "/VBA/SignatureFixture" ? changedModule : entry.Value))
                        throw new InvalidDataException("A stream outside the declared source literal changed.");
                return changedFile;
            }
            throw new InvalidDataException("No unique physical source literal can be changed without collateral source changes.");
        }

        private static int UniqueIndex(byte[] data, byte[] pattern)
        {
            int result = -1;
            for (int offset = 0; offset <= data.Length - pattern.Length; offset++)
            {
                bool match = true;
                for (int i = 0; i < pattern.Length; i++) if (data[offset + i] != pattern[i]) { match = false; break; }
                if (!match) continue;
                if (result != -1) return -1;
                result = offset;
            }
            return result;
        }

        private static int ModuleOffset(byte[] directory)
        {
            byte[] name = Encoding.ASCII.GetBytes("SignatureFixture");
            byte[] marker = new byte[6 + name.Length]; marker[0] = 0x19;
            Buffer.BlockCopy(BitConverter.GetBytes(name.Length), 0, marker, 2, 4);
            Buffer.BlockCopy(name, 0, marker, 6, name.Length);
            int position = UniqueIndex(directory, marker);
            if (position < 0) throw new InvalidDataException("Unique MODULE record required.");
            var records = new Dictionary<ushort, byte[]>();
            var allowed = new HashSet<ushort> { 0x19, 0x47, 0x1a, 0x32, 0x1c, 0x48, 0x31, 0x1e, 0x2c, 0x21, 0x22, 0x25, 0x28, 0x2b };
            for (int count = 0; count < 32; count++)
            {
                if (position + 6 > directory.Length) break;
                ushort tag = BitConverter.ToUInt16(directory, position);
                uint size = BitConverter.ToUInt32(directory, position + 2); position += 6;
                if (!allowed.Contains(tag) || records.ContainsKey(tag) || size > Limit || position + size > directory.Length)
                    throw new InvalidDataException("Unexpected module record.");
                records.Add(tag, directory.Skip(position).Take((int)size).ToArray()); position += (int)size;
                if (tag != 0x2b) continue;
                if (size != 0 || !records[0x19].SequenceEqual(name) || !records[0x1a].SequenceEqual(name) ||
                    records[0x31].Length != 4 || records.ContainsKey(0x21) == records.ContainsKey(0x22))
                    throw new InvalidDataException("Module identity, type or source offset differs.");
                return checked((int)BitConverter.ToUInt32(records[0x31], 0));
            }
            throw new InvalidDataException("Complete module record unavailable.");
        }

        /// <summary>Decodes bounded MS-OVBA chunks, retaining the originating literal byte for every source character.</summary>
        internal static byte[] Decompress(byte[] data, out int[] origins)
        {
            if (data == null || data.Length == 0 || data.Length > Limit || data[0] != 1)
                throw new InvalidDataException("Bounded compressed container required.");
            var output = new List<byte>(); var map = new List<int>(); int position = 1;
            while (position < data.Length)
            {
                if (position + 2 > data.Length) throw new InvalidDataException("Incomplete chunk.");
                ushort header = BitConverter.ToUInt16(data, position);
                int size = (header & 0xfff) + 3, end = position + size, origin = output.Count;
                if ((header & 0x7000) != 0x3000 || end > data.Length) throw new InvalidDataException("Invalid chunk extent.");
                position += 2;
                if ((header & 0x8000) == 0)
                {
                    if (size != 4098) throw new InvalidDataException("Raw chunk must contain 4096 bytes.");
                    while (position < end) { map.Add(position); output.Add(data[position++]); }
                }
                else while (position < end)
                {
                    byte flags = data[position++];
                    for (int bit = 0; bit < 8 && position < end; bit++)
                    {
                        int count = output.Count - origin;
                        if ((flags & (1 << bit)) == 0)
                        {
                            if (count >= 4096) throw new InvalidDataException("Literal exceeds chunk.");
                            map.Add(position); output.Add(data[position++]); continue;
                        }
                        if (position + 2 > end || count == 0) throw new InvalidDataException("Incomplete or premature copy token.");
                        ushort token = BitConverter.ToUInt16(data, position); position += 2;
                        int width = 4; while ((1 << width) < count) width++;
                        int mask = 0xffff >> width;
                        int length = (token & mask) + 3, distance = (token >> (16 - width)) + 1;
                        if (distance > count || count + length > 4096) throw new InvalidDataException("Copy exceeds chunk.");
                        for (int i = 0; i < length; i++) { map.Add(map[map.Count - distance]); output.Add(output[output.Count - distance]); }
                    }
                }
                if (output.Count > Limit) throw new InvalidDataException("Decompressed source exceeds bound.");
            }
            origins = map.ToArray(); return output.ToArray();
        }

        private static void VerifyCopy(string original, string copy)
        {
            using (var left = ZipFile.OpenRead(original))
            using (var right = ZipFile.OpenRead(copy))
            {
                if (!left.Entries.Select(e => e.FullName).OrderBy(n => n).SequenceEqual(right.Entries.Select(e => e.FullName).OrderBy(n => n)))
                    throw new InvalidDataException("ZIP inventory changed.");
                foreach (var entry in left.Entries)
                {
                    byte[] before = ReadEntry(entry), after = ReadEntry(right.GetEntry(entry.FullName));
                    int differences = before.Length == after.Length ? before.Where((b, i) => b != after[i]).Count() : -1;
                    if (differences != (entry.FullName == ProjectEntry ? 1 : 0)) throw new InvalidDataException("Unexpected ZIP payload mutation.");
                }
            }
        }
    }
}

using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;

namespace VBAi
{
    /// <summary>Checks exported UserForm OLE containers before any native import.</summary>
    internal static class FormResourcePreflight
    {
        /// <summary>
        /// Validates the bounded LB/08 envelope used by native VBA UserForm exports
        /// and its MS-CFB allocation graph. Does not activate OLE objects, rewrite
        /// resources, or interpret the embedded MS-OFORMS control properties.
        /// </summary>
        internal static void ValidateOleObjectBlob(byte[] resources, int offset)
        {
            Read(resources, offset).Validate();
        }

        // Comparison only. Transport, checkpoints and imported files keep their original bytes.
        internal static byte[] ComparisonBytes(byte[] resources, IEnumerable<int> offsets)
        {
            using (var buffer = new MemoryStream())
            using (var writer = new BinaryWriter(buffer, Encoding.UTF8, true))
            {
                int cursor = 0;
                foreach (int offset in offsets.Distinct().OrderBy(x => x))
                {
                    if (offset < cursor) throw Invalid();
                    var compound = Read(resources, offset);
                    compound.Validate();
                    writer.Write(offset - cursor);
                    writer.Write(resources, cursor, offset - cursor);
                    // LB/08 has been observed on native exports. Preserve every envelope
                    // byte except the physical CFB length, which changes with allocation.
                    writer.Write(resources, offset, 4);
                    writer.Write(resources, offset + 8, 16);
                    byte[] content = compound.ComparisonBytes();
                    writer.Write(content.Length); writer.Write(content);
                    cursor = offset + 24 + compound.length;
                }
                writer.Write(resources.Length - cursor);
                writer.Write(resources, cursor, resources.Length - cursor);
                return buffer.ToArray();
            }
        }

        private static CompoundFile Read(byte[] resources, int offset)
        {
            if (resources == null || resources.Length > VbaGitSnapshot.MaxBytes || offset < 0 ||
                offset > resources.Length - 24 || resources[offset] != 0x4c || resources[offset + 1] != 0x42 ||
                resources[offset + 2] != 8 || resources[offset + 3] != 0)
                throw Invalid();
            uint length = BitConverter.ToUInt32(resources, offset + 4);
            if (length > resources.Length - offset - 24 || length < 512)
                throw Invalid();
            return new CompoundFile(resources, offset + 24, (int)length);
        }

        private static InvalidOperationException Invalid()
        {
            return new InvalidOperationException("Invalid or truncated UserForm OLE resource container.");
        }

        /// <summary>Reads only bounded MS-CFB metadata; never invokes an OLE decoder.</summary>
        private sealed class CompoundFile
        {
            private static readonly byte[] FormClassId = new Guid("C62A69F0-16DC-11CE-9E98-00AA00574A4F").ToByteArray();
            private static readonly Encoding DirectoryNameEncoding = new UnicodeEncoding(false, false, true);
            private const uint End = 0xfffffffe, Free = 0xffffffff, FatSector = 0xfffffffd, DifatSector = 0xfffffffc;
            private readonly byte[] bytes;
            private readonly int origin;
            internal readonly int length;
            private int sectorSize, sectorCount, major;
            private uint[] fat;
            private bool[] claimed;
            private List<Entry> entries;

            internal CompoundFile(byte[] bytes, int origin, int length)
            {
                this.bytes = bytes; this.origin = origin; this.length = length;
            }

            internal void Validate()
            {
                byte[] signature = { 0xd0, 0xcf, 0x11, 0xe0, 0xa1, 0xb1, 0x1a, 0xe1 };
                for (int i = 0; i < signature.Length; i++) if (bytes[origin + i] != signature[i]) throw Invalid();
                major = U16(26);
                int shift = U16(30);
                if (U16(28) != 0xfffe || U16(32) != 6 || U32(56) != 4096 ||
                    (major != 3 && major != 4) || shift != (major == 3 ? 9 : 12)) throw Invalid();
                sectorSize = 1 << shift;
                if (length % sectorSize != 0 || length < sectorSize * 3 || (major == 3 && U32(40) != 0)) throw Invalid();
                sectorCount = length / sectorSize - 1;
                claimed = new bool[sectorCount];
                fat = new uint[sectorCount];
                for (int i = 0; i < fat.Length; i++) fat[i] = Free;

                var fatLocations = new List<uint>();
                for (int i = 0; i < 109; i++)
                {
                    uint value = U32(76 + i * 4);
                    if (value != Free) { Claim(value); fatLocations.Add(value); }
                }
                var difatLocations = new List<uint>();
                uint difat = U32(68), difatCount = U32(72);
                if (difatCount > sectorCount) throw Invalid();
                for (uint i = 0; i < difatCount; i++)
                {
                    Claim(difat); difatLocations.Add(difat);
                    int position = Sector(difat);
                    for (int entry = 0; entry < sectorSize / 4 - 1; entry++)
                    {
                        uint value = U32(position + entry * 4);
                        if (value != Free) { Claim(value); fatLocations.Add(value); }
                    }
                    difat = U32(position + sectorSize - 4);
                }
                if (difat != End || fatLocations.Count != U32(44) || fatLocations.Count == 0 ||
                    (long)fatLocations.Count * (sectorSize / 4) < sectorCount) throw Invalid();
                int fatIndex = 0;
                foreach (uint sector in fatLocations)
                {
                    int position = Sector(sector);
                    for (int i = 0; i < sectorSize / 4 && fatIndex < fat.Length; i++) fat[fatIndex++] = U32(position + i * 4);
                }
                foreach (uint sector in fatLocations) if (fat[sector] != FatSector) throw Invalid();
                foreach (uint sector in difatLocations) if (fat[sector] != DifatSector) throw Invalid();

                var directory = Chain(U32(48), null);
                if (directory.Count == 0 || (major == 4 && directory.Count != U32(40))) throw Invalid();
                entries = new List<Entry>();
                foreach (uint sector in directory)
                {
                    int position = Sector(sector);
                    for (int i = 0; i < sectorSize; i += 128) entries.Add(ReadEntry(position + i));
                }
                if (entries[0].Kind != 5 || entries[0].Name != "Root Entry" || entries[0].Left != Free || entries[0].Right != Free)
                    throw Invalid();

                // Walk iteratively: hostile nesting and sibling loops cannot exhaust the stack.
                var visited = new bool[entries.Count]; visited[0] = true;
                entries[0].Path = "";
                var pending = new Stack<Tuple<uint, string>>(); pending.Push(Tuple.Create(entries[0].Child, ""));
                var paths = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
                long pathCharacters = 0;
                while (pending.Count > 0)
                {
                    var pendingEntry = pending.Pop();
                    uint id = pendingEntry.Item1;
                    if (id == Free) continue;
                    if (id >= entries.Count || visited[id] || (entries[(int)id].Kind != 1 && entries[(int)id].Kind != 2)) throw Invalid();
                    visited[id] = true;
                    var entry = entries[(int)id];
                    pathCharacters += (long)pendingEntry.Item2.Length + 1 + entry.Name.Length;
                    if (pathCharacters > VbaGitSnapshot.MaxBytes) throw Invalid();
                    entry.Path = pendingEntry.Item2 + "/" + entry.Name;
                    if (!paths.Add(entry.Path)) throw Invalid();
                    pending.Push(Tuple.Create(entry.Left, pendingEntry.Item2)); pending.Push(Tuple.Create(entry.Right, pendingEntry.Item2));
                    if (entry.Kind == 1) pending.Push(Tuple.Create(entry.Child, entry.Path));
                    else if (entry.Child != Free) throw Invalid();
                }
                for (int i = 1; i < entries.Count; i++) if ((entries[i].Kind != 0) != visited[i]) throw Invalid();

                uint miniFatCount = U32(64);
                if (miniFatCount > sectorCount) throw Invalid();
                var miniFatChain = Chain(U32(60), (long)miniFatCount * sectorSize);
                var miniFat = new uint[miniFatChain.Count * (sectorSize / 4)];
                int index = 0;
                foreach (uint sector in miniFatChain)
                    for (int i = 0; i < sectorSize; i += 4) miniFat[index++] = U32(Sector(sector) + i);
                byte[] miniStream = ReadChain(Chain(entries[0].Start, entries[0].Size), entries[0].Size);
                long miniCount = (entries[0].Size + 63) / 64;
                if (miniCount > miniFat.Length) throw Invalid();
                var miniClaimed = new bool[(int)miniCount];
                foreach (var entry in entries)
                {
                    if (entry.Kind != 2) continue;
                    if (entry.Size >= 4096) { entry.Data = ReadChain(Chain(entry.Start, entry.Size), entry.Size); continue; }
                    entry.Data = new byte[(int)entry.Size];
                    long count = (entry.Size + 63) / 64;
                    uint next = entry.Start;
                    for (long i = 0; i < count; i++)
                    {
                        if (next >= miniCount || miniClaimed[next] || (long)next * 64 + Math.Min(64, entry.Size - i * 64) > entries[0].Size) throw Invalid();
                        miniClaimed[next] = true;
                        Buffer.BlockCopy(miniStream, (int)next * 64, entry.Data, (int)i * 64, (int)Math.Min(64, entry.Size - i * 64));
                        next = miniFat[next];
                    }
                    if (next != End) throw Invalid();
                }
            }

            internal byte[] ComparisonBytes()
            {
                // MS-CFB sector allocation, slack bytes, directory tree ordering and
                // creation/modification timestamps do not describe control properties.
                // Keep every logical stream, name, storage CLSID and state bit.
                var streams = entries.Where(x => x.Kind == 2).ToDictionary(x => x.Path, StringComparer.Ordinal);
                foreach (var storage in entries.Where(x => x.Kind == 1 || x.Kind == 5))
                {
                    // An unrelated OLE object's f/o streams can coincidentally match
                    // this grammar. Its class identity must also identify a UserForm.
                    if (!storage.Metadata.Take(16).SequenceEqual(FormClassId)) continue;
                    Entry form, objects;
                    if (!streams.TryGetValue(storage.Path + "/f", out form) ||
                        !streams.TryGetValue(storage.Path + "/o", out objects)) continue;
                    var normalized = FormStreamPadding.Normalize(form.Data, objects.Data);
                    form.Data = normalized[0]; objects.Data = normalized[1];
                }
                using (var buffer = new MemoryStream())
                using (var writer = new BinaryWriter(buffer, Encoding.UTF8, true))
                {
                    writer.Write(bytes, origin + 8, 16); // Header CLSID, even for unknown extensions.
                    foreach (var entry in entries.Where(x => x.Kind != 0).OrderBy(x => x.Path, StringComparer.Ordinal))
                    {
                        writer.Write(entry.Kind); writer.Write(entry.Path); writer.Write(entry.Name);
                        writer.Write(entry.Metadata);
                        writer.Write(entry.Data?.Length ?? 0);
                        if (entry.Data != null) writer.Write(entry.Data);
                    }
                    return buffer.ToArray();
                }
            }

            private byte[] ReadChain(List<uint> chain, long size)
            {
                var result = new byte[(int)size];
                int cursor = 0;
                foreach (uint sector in chain)
                {
                    int count = Math.Min(sectorSize, result.Length - cursor);
                    Buffer.BlockCopy(bytes, origin + Sector(sector), result, cursor, count);
                    cursor += count;
                }
                return result;
            }

            private Entry ReadEntry(int position)
            {
                byte kind = bytes[origin + position + 66];
                if (kind == 0) return new Entry();
                int nameLength = U16(position + 64);
                if ((kind != 1 && kind != 2 && kind != 5) || nameLength < 2 || nameLength > 64 || nameLength % 2 != 0 ||
                    U16(position + nameLength - 2) != 0) throw Invalid();
                string name;
                try { name = DirectoryNameEncoding.GetString(bytes, origin + position, nameLength - 2); }
                catch (DecoderFallbackException) { throw Invalid(); }
                if (name.Length == 0 || name.IndexOfAny(new[] { '\0', '/', '\\', ':', '!' }) >= 0) throw Invalid();
                ulong size = major == 3 ? U32(position + 120) : U64(position + 120);
                if ((kind == 2 || kind == 5) && size > (ulong)length) throw Invalid();
                var metadata = new byte[20]; Buffer.BlockCopy(bytes, origin + position + 80, metadata, 0, metadata.Length);
                return new Entry { Kind = kind, Name = name, Metadata = metadata, Left = U32(position + 68), Right = U32(position + 72),
                    Child = U32(position + 76), Start = U32(position + 116), Size = (long)size };
            }

            private List<uint> Chain(uint start, long? size)
            {
                if (size > length || size < 0) throw Invalid();
                long? expected = size.HasValue ? (size.Value + sectorSize - 1) / sectorSize : (long?)null;
                var result = new List<uint>();
                uint next = start;
                while (next != End)
                {
                    if (expected.HasValue && result.Count >= expected.Value) throw Invalid();
                    Claim(next); result.Add(next); next = fat[next];
                }
                if (expected.HasValue && result.Count != expected.Value) throw Invalid();
                return result;
            }

            private void Claim(uint sector)
            {
                if (sector >= sectorCount || claimed[sector]) throw Invalid();
                claimed[sector] = true;
            }

            private int Sector(uint sector)
            {
                if (sector >= sectorCount) throw Invalid();
                return ((int)sector + 1) * sectorSize;
            }

            private ushort U16(int position) { Bounds(position, 2); return BitConverter.ToUInt16(bytes, origin + position); }
            private uint U32(int position) { Bounds(position, 4); return BitConverter.ToUInt32(bytes, origin + position); }
            private ulong U64(int position) { Bounds(position, 8); return BitConverter.ToUInt64(bytes, origin + position); }
            private void Bounds(int position, int count) { if (position < 0 || position > length - count) throw Invalid(); }

            private sealed class Entry
            {
                internal byte Kind;
                internal string Name;
                internal string Path;
                internal byte[] Data, Metadata;
                internal uint Left, Right, Child, Start;
                internal long Size;
            }
        }
    }
}

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
        /// <param name="resources">byte[] that supplies the resources for this operation.</param>
        /// <param name="offset">int that supplies the offset for this operation.</param>
        internal static void ValidateOleObjectBlob(byte[] resources, int offset)
        {
            Read(resources, offset).Validate();
        }

        /// <summary>Reads font restoration bindings without activating or changing the exported resources.</summary>
        /// <param name="resources">byte[] that supplies the resources for this operation.</param>
        /// <param name="offset">int that supplies the offset for this operation.</param>
        /// <returns>form font binding[] produced by the operation for read font bindings on form resource preflight.</returns>
        internal static FormStreamPadding.FormFontBinding[] ReadFontBindings(byte[] resources, int offset)
        {
            var compound = Read(resources, offset); compound.Validate();
            return compound.ReadFontBindings();
        }

        // Comparison only. Transport, checkpoints and imported files keep their original bytes.
        /// <summary>Handles comparison bytes for form resource preflight.</summary>
        /// <param name="resources">byte[] that supplies the resources for this operation.</param>
        /// <param name="offsets">i enumerable&lt;int&gt; that supplies the offsets for this operation.</param>
        /// <returns>byte[] produced by the operation for comparison bytes on form resource preflight.</returns>
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

        /// <summary>Reads  for form resource preflight.</summary>
        /// <param name="resources">byte[] that supplies the resources for this operation.</param>
        /// <param name="offset">int that supplies the offset for this operation.</param>
        /// <returns>compound file produced by the operation for read on form resource preflight.</returns>
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

        /// <summary>Handles invalid for form resource preflight.</summary>
        /// <returns>invalid operation exception produced by the operation for invalid on form resource preflight.</returns>
        private static InvalidOperationException Invalid()
        {
            return new InvalidOperationException("Invalid or truncated UserForm OLE resource container.");
        }

        /// <summary>Reads only bounded MS-CFB metadata; never invokes an OLE decoder.</summary>
        private sealed class CompoundFile
        {

            /// <summary>Identifies the form class id associated with compound file.</summary>
            private static readonly byte[] FormClassId = new Guid("C62A69F0-16DC-11CE-9E98-00AA00574A4F").ToByteArray();

            /// <summary>Maintains the directory name encoding state for compound file.</summary>
            private static readonly Encoding DirectoryNameEncoding = new UnicodeEncoding(false, false, true);

            /// <summary>Maintains the end and free and fat sector and difat sector state for compound file.</summary>
            private const uint End = 0xfffffffe, Free = 0xffffffff, FatSector = 0xfffffffd, DifatSector = 0xfffffffc;

            /// <summary>Maintains the bytes state for compound file.</summary>
            private readonly byte[] bytes;

            /// <summary>Maintains the origin state for compound file.</summary>
            private readonly int origin;

            /// <summary>Maintains the length state for compound file.</summary>
            internal readonly int length;

            /// <summary>Counts the sector size and sector count and major maintained by compound file.</summary>
            private int sectorSize, sectorCount, major;

            /// <summary>Maintains the fat state for compound file.</summary>
            private uint[] fat;

            /// <summary>Maintains the claimed state for compound file.</summary>
            private bool[] claimed;

            /// <summary>Maintains the entries state for compound file.</summary>
            private List<Entry> entries;

            /// <summary>Initializes a CompoundFile instance with the supplied state.</summary>
            /// <param name="bytes">byte[] that supplies the bytes for this operation.</param>
            /// <param name="origin">int that supplies the origin for this operation.</param>
            /// <param name="length">int that supplies the length for this operation.</param>
            internal CompoundFile(byte[] bytes, int origin, int length)
            {
                this.bytes = bytes; this.origin = origin; this.length = length;
            }

            /// <summary>Validates  for compound file.</summary>
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

            /// <summary>Handles comparison bytes for compound file.</summary>
            /// <returns>byte[] produced by the operation for comparison bytes on compound file.</returns>
            internal byte[] ComparisonBytes()
            {
                // MS-CFB sector allocation, slack bytes, directory tree ordering and
                // creation/modification timestamps do not describe control properties.
                // Keep every logical stream, name, storage CLSID and state bit.
                var streams = entries.Where(x => x.Kind == 2).ToDictionary(x => x.Path, x => x.Data, StringComparer.Ordinal);
                var storageMetadata = entries.Where(x => x.Kind == 1 || x.Kind == 5).ToDictionary(x => x.Path, x => x.Metadata, StringComparer.Ordinal);
                // Normalize exactly one complete UserForm graph. An unsupported parent prevents descendant normalization.
                if (entries[0].Metadata.Take(16).SequenceEqual(FormClassId))
                {
                    var normalized = FormStreamPadding.NormalizeGraph(streams, storageMetadata);
                    foreach (var entry in entries.Where(x => x.Kind == 2)) entry.Data = normalized[entry.Path];
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

            /// <summary>Reads font bindings for compound file.</summary>
            /// <returns>form font binding[] produced by the operation for read font bindings on compound file.</returns>
            internal FormStreamPadding.FormFontBinding[] ReadFontBindings()
            {
                if (!entries[0].Metadata.Take(16).SequenceEqual(FormClassId)) return null;
                var streams = entries.Where(x => x.Kind == 2).ToDictionary(x => x.Path, x => x.Data, StringComparer.Ordinal);
                var metadata = entries.Where(x => x.Kind == 1 || x.Kind == 5).ToDictionary(x => x.Path, x => x.Metadata, StringComparer.Ordinal);
                return FormStreamPadding.ReadFontBindings(streams, metadata);
            }

            /// <summary>Reads chain for compound file.</summary>
            /// <param name="chain">list&lt;uint&gt; that supplies the chain for this operation.</param>
            /// <param name="size">long that supplies the size for this operation.</param>
            /// <returns>byte[] produced by the operation for read chain on compound file.</returns>
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

            /// <summary>Reads entry for compound file.</summary>
            /// <param name="position">int that supplies the position for this operation.</param>
            /// <returns>entry produced by the operation for read entry on compound file.</returns>
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

            /// <summary>Handles chain for compound file.</summary>
            /// <param name="start">uint that supplies the start for this operation.</param>
            /// <param name="size">long that supplies the size for this operation.</param>
            /// <returns>list&lt;uint&gt; produced by the operation for chain on compound file.</returns>
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

            /// <summary>Handles claim for compound file.</summary>
            /// <param name="sector">uint that supplies the sector for this operation.</param>
            private void Claim(uint sector)
            {
                if (sector >= sectorCount || claimed[sector]) throw Invalid();
                claimed[sector] = true;
            }

            /// <summary>Handles sector for compound file.</summary>
            /// <param name="sector">uint that supplies the sector for this operation.</param>
            /// <returns>int produced by the operation for sector on compound file.</returns>
            private int Sector(uint sector)
            {
                if (sector >= sectorCount) throw Invalid();
                return ((int)sector + 1) * sectorSize;
            }

            /// <summary>Handles u16 for compound file.</summary>
            /// <param name="position">int that supplies the position for this operation.</param>
            /// <returns>ushort produced by the operation for u16 on compound file.</returns>
            private ushort U16(int position) { Bounds(position, 2); return BitConverter.ToUInt16(bytes, origin + position); }

            /// <summary>Handles u32 for compound file.</summary>
            /// <param name="position">int that supplies the position for this operation.</param>
            /// <returns>uint produced by the operation for u32 on compound file.</returns>
            private uint U32(int position) { Bounds(position, 4); return BitConverter.ToUInt32(bytes, origin + position); }

            /// <summary>Handles u64 for compound file.</summary>
            /// <param name="position">int that supplies the position for this operation.</param>
            /// <returns>ulong produced by the operation for u64 on compound file.</returns>
            private ulong U64(int position) { Bounds(position, 8); return BitConverter.ToUInt64(bytes, origin + position); }

            /// <summary>Handles bounds for compound file.</summary>
            /// <param name="position">int that supplies the position for this operation.</param>
            /// <param name="count">int that supplies the count for this operation.</param>
            private void Bounds(int position, int count) { if (position < 0 || position > length - count) throw Invalid(); }

            /// <summary>Owns the entry state and operations.</summary>
            private sealed class Entry
            {

                /// <summary>Maintains the kind state for entry.</summary>
                internal byte Kind;

                /// <summary>Maintains the name state for entry.</summary>
                internal string Name;

                /// <summary>Keeps the path path available to entry.</summary>
                internal string Path;

                /// <summary>Maintains the data and metadata state for entry.</summary>
                internal byte[] Data, Metadata;

                /// <summary>Maintains the left and right and child and start state for entry.</summary>
                internal uint Left, Right, Child, Start;

                /// <summary>Maintains the size state for entry.</summary>
                internal long Size;
            }
        }
    }
}

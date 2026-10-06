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
        /// <param name="resources">Complete exported FRX byte array.</param>
        /// <param name="offset">Byte offset of an LB/08 OLE-object envelope.</param>
        /// <exception cref="InvalidOperationException">The envelope or its bounded CFB allocation graph is malformed or unsupported.</exception>
        internal static void ValidateOleObjectBlob(byte[] resources, int offset)
        {
            Read(resources, offset).Validate();
        }

        /// <summary>Reads font restoration bindings without activating or changing the exported resources.</summary>
        /// <param name="resources">Complete exported FRX byte array.</param>
        /// <param name="offset">Byte offset of the target UserForm OLE-object envelope.</param>
        /// <returns>Root and nested StdFont descriptor bindings found without activating OLE objects.</returns>
        internal static FormStreamPadding.FormFontBinding[] ReadFontBindings(byte[] resources, int offset)
        {
            var compound = Read(resources, offset); compound.Validate();
            return compound.ReadFontBindings();
        }

        // Comparison only. Transport, checkpoints and imported files keep their original bytes.
        /// <summary>Builds a comparison-only form of FRX bytes, canonicalizing recognized CFB allocation while preserving envelope data.</summary>
        /// <param name="resources">Original resource stream; returned bytes are never written back as import data.</param>
        /// <param name="offsets">Declared OLE envelope offsets, processed in sorted distinct order.</param>
        /// <returns>Comparison representation; overlapping or malformed ranges throw instead of being normalized.</returns>
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

        /// <summary>Validates an LB/08 envelope header and returns its bounded compound-file payload.</summary>
        /// <param name="resources">Complete resource stream.</param>
        /// <param name="offset">Envelope start offset.</param>
        /// <returns>Compound-file reader spanning only the declared payload.</returns>
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

        /// <summary>Creates the common rejection for malformed or unsupported UserForm OLE resource data.</summary>
        /// <returns>An exception describing invalid or truncated resource content.</returns>
        private static InvalidOperationException Invalid()
        {
            return new InvalidOperationException("Invalid or truncated UserForm OLE resource container.");
        }

        /// <summary>Reads only bounded MS-CFB metadata; never invokes an OLE decoder.</summary>
        private sealed class CompoundFile
        {

            /// <summary>Expected MSForms UserForm class identifier embedded in the compound file directory.</summary>
            private static readonly byte[] FormClassId = new Guid("C62A69F0-16DC-11CE-9E98-00AA00574A4F").ToByteArray();

            /// <summary>Strict UTF-16LE decoder for MS-CFB directory entry names.</summary>
            private static readonly Encoding DirectoryNameEncoding = new UnicodeEncoding(false, false, true);

            /// <summary>Reserved MS-CFB markers for chain end, free sectors, FAT sectors, and DIFAT sectors.</summary>
            private const uint End = 0xfffffffe, Free = 0xffffffff, FatSector = 0xfffffffd, DifatSector = 0xfffffffc;

            /// <summary>Original resource bytes containing the bounded compound-file payload.</summary>
            private readonly byte[] bytes;

            /// <summary>Absolute byte offset at which this payload begins.</summary>
            private readonly int origin;

            /// <summary>Declared CFB payload extent in bytes.</summary>
            internal readonly int length;

            /// <summary>Validated major CFB version, sector byte width, and sector count.</summary>
            private int sectorSize, sectorCount, major;

            /// <summary>Expanded FAT entries mapping each regular sector to its next chain sector.</summary>
            private uint[] fat;

            /// <summary>Tracks sectors already assigned to a metadata or stream chain to reject overlap.</summary>
            private bool[] claimed;

            /// <summary>Parsed directory entries whose stream chains can be traversed within the validated FAT.</summary>
            private List<Entry> entries;

            /// <summary>Initializes a CompoundFile instance with the supplied state.</summary>
            /// <param name="bytes">FRX backing bytes.</param>
            /// <param name="origin">Absolute start of the CFB payload.</param>
            /// <param name="length">Payload length declared by its LB/08 envelope.</param>
            internal CompoundFile(byte[] bytes, int origin, int length)
            {
                this.bytes = bytes; this.origin = origin; this.length = length;
            }

            /// <summary>Validates CFB header, FAT/DIFAT, directory, allocation chains, and the expected UserForm root storage.</summary>
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

            /// <summary>Serializes logical CFB entries for comparison while omitting sector allocation and timestamps.</summary>
            /// <returns>Canonical comparison bytes; original resource transport bytes remain untouched.</returns>
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

            /// <summary>Extracts supported root and nested StdFont descriptors from a validated UserForm storage graph.</summary>
            /// <returns>Font bindings, or null when the root storage is not the expected UserForm class.</returns>
            internal FormStreamPadding.FormFontBinding[] ReadFontBindings()
            {
                if (!entries[0].Metadata.Take(16).SequenceEqual(FormClassId)) return null;
                var streams = entries.Where(x => x.Kind == 2).ToDictionary(x => x.Path, x => x.Data, StringComparer.Ordinal);
                var metadata = entries.Where(x => x.Kind == 1 || x.Kind == 5).ToDictionary(x => x.Path, x => x.Metadata, StringComparer.Ordinal);
                return FormStreamPadding.ReadFontBindings(streams, metadata);
            }

            /// <summary>Copies the requested logical stream bytes from a previously validated sector chain.</summary>
            /// <param name="chain">Ordered CFB sector identifiers.</param>
            /// <param name="size">Declared logical stream length in bytes.</param>
            /// <returns>Stream content without sector slack bytes.</returns>
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

            /// <summary>Decodes one fixed-size CFB directory record, validating its name, type, and stream size.</summary>
            /// <param name="position">Byte offset of the directory record within the payload.</param>
            /// <returns>Parsed storage, stream, root, or empty entry.</returns>
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

            /// <summary>Follows a FAT chain, claiming each sector and enforcing its expected length when known.</summary>
            /// <param name="start">First sector identifier.</param>
            /// <param name="size">Expected stream byte length, or null for a chain with no declared exact length.</param>
            /// <returns>Ordered sector identifiers through the end marker.</returns>
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

            /// <summary>Marks one sector as owned by a single metadata or stream chain.</summary>
            /// <param name="sector">Sector index to claim.</param>
            private void Claim(uint sector)
            {
                if (sector >= sectorCount || claimed[sector]) throw Invalid();
                claimed[sector] = true;
            }

            /// <summary>Converts a zero-based sector index to an absolute payload-relative byte offset.</summary>
            /// <param name="sector">Sector index.</param>
            /// <returns>Byte offset following the CFB header sector.</returns>
            private int Sector(uint sector)
            {
                if (sector >= sectorCount) throw Invalid();
                return ((int)sector + 1) * sectorSize;
            }

            /// <summary>Reads a bounded little-endian 16-bit value from the payload.</summary>
            /// <param name="position">Payload-relative byte offset.</param>
            /// <returns>Decoded unsigned value.</returns>
            private ushort U16(int position) { Bounds(position, 2); return BitConverter.ToUInt16(bytes, origin + position); }

            /// <summary>Reads a bounded little-endian 32-bit value from the payload.</summary>
            /// <param name="position">Payload-relative byte offset.</param>
            /// <returns>Decoded unsigned value.</returns>
            private uint U32(int position) { Bounds(position, 4); return BitConverter.ToUInt32(bytes, origin + position); }

            /// <summary>Reads a bounded little-endian 64-bit value from the payload.</summary>
            /// <param name="position">Payload-relative byte offset.</param>
            /// <returns>Decoded unsigned value.</returns>
            private ulong U64(int position) { Bounds(position, 8); return BitConverter.ToUInt64(bytes, origin + position); }

            /// <summary>Rejects a requested byte range outside the declared CFB payload.</summary>
            /// <param name="position">Payload-relative range start.</param>
            /// <param name="count">Number of bytes to read.</param>
            private void Bounds(int position, int count) { if (position < 0 || position > length - count) throw Invalid(); }

            /// <summary>Parsed directory entry and its logical path, metadata, allocation chain, and content.</summary>
            private sealed class Entry
            {

                /// <summary>CFB directory kind: empty, storage, stream, or root storage.</summary>
                internal byte Kind;

                /// <summary>Decoded directory name.</summary>
                internal string Name;

                /// <summary>Full logical path constructed from the parent storage hierarchy.</summary>
                internal string Path;

                /// <summary>Logical stream bytes and the 20-byte persisted storage metadata compared by the canonicalizer.</summary>
                internal byte[] Data, Metadata;

                /// <summary>Directory tree links plus the first sector of this entry's allocation chain.</summary>
                internal uint Left, Right, Child, Start;

                /// <summary>Declared logical stream size in bytes.</summary>
                internal long Size;
            }
        }
    }
}

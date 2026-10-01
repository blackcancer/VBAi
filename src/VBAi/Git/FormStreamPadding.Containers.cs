using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;

namespace VBAi
{
    internal static partial class FormStreamPadding
    {
        private static readonly byte[] GraphFormClass = new Guid("C62A69F0-16DC-11CE-9E98-00AA00574A4F").ToByteArray();
        private static readonly byte[] GraphFrameClass = new Guid("6E182020-F460-11CE-9BCD-00AA00608E01").ToByteArray();
        private static readonly byte[] GraphMultiPageClass = new Guid("46E31370-3F7A-11CE-BED6-00AA00611080").ToByteArray();

        /// <summary>Normalizes comparison clones only after the complete bounded UserForm storage graph is validated.</summary>
        /// <remarks>Any refusal returns the exact original dictionary and arrays. Transport/import bytes are never modified.</remarks>
        internal static IReadOnlyDictionary<string, byte[]> NormalizeGraph(IReadOnlyDictionary<string, byte[]> streams,
            IReadOnlyDictionary<string, byte[]> storageMetadata, string rootPath = "")
        {
            if (streams == null || storageMetadata == null || rootPath == null) return streams;
            try
            {
                Require(streams.Count <= 16384 && storageMetadata.Count <= 4096);
                long size = 0;
                foreach (var stream in streams)
                {
                    Require(stream.Key != null && stream.Value != null && stream.Value.Length <= MaxStreamBytes);
                    size += stream.Value.Length; Require(size <= MaxStreamBytes);
                }
                var graph = new StorageGraph(streams, storageMetadata);
                var pending = new Stack<StorageNode>();
                pending.Push(new StorageNode(rootPath, 7, 0, 0));
                while (pending.Count != 0)
                {
                    StorageNode node = pending.Pop();
                    graph.Parse(node);
                    foreach (var site in node.Sites.Where(x => !x.Streamed))
                    {
                        Require(node.Depth < 64);
                        string name = "i" + site.Identity.ToString("D2", CultureInfo.InvariantCulture);
                        pending.Push(new StorageNode(node.Path + "/" + name, site.Type, site.Identity, node.Depth + 1));
                    }
                }
                Require(graph.StorageClaims.Count == storageMetadata.Count && graph.StreamClaims.Count == streams.Count);
                return graph.Comparison;
            }
            catch (UnsupportedLayoutException) { return streams; }
        }

        /// <summary>Reads design-surface metadata while retaining all property bytes and validating its optional extent.</summary>
        private static void ParseDesignExtender(Reader form)
        {
            // MS-OFORMS 2.2.10.9, 2.2.10.11.1-.3 and 2.5.5.1.
            // https://learn.microsoft.com/en-us/openspecs/office_file_formats/ms-oforms/8448f23a-86b7-40d9-839b-3a1aa14e501a
            Reader block = form.Block(0x0200);
            uint mask = block.UInt32(); Require((mask & ~0x1fu) == 0);
            uint flags = block.Field(mask, 0, 4, 0x15f55); Require((flags & ~0x3ffffu) == 0);
            block.Field(mask, 1, 4); block.Field(mask, 2, 4);
            uint click = block.Field(mask, 3, 1), twice = block.Field(mask, 4, 1);
            Require(click <= 1 || click == 0xfe || click == 0xff);
            Require(twice <= 2 || twice == 0xfe);
            block.Align(4); block.Finish();
        }

        private sealed class TabLinks
        {
            internal readonly List<string> Items = new List<string>(), Names = new List<string>();
        }

        private sealed class StorageSite
        {
            internal uint Identity, Type;
            internal bool Streamed;
            internal string Name;
        }

        private sealed class StorageNode
        {
            internal readonly string Path;
            internal readonly uint Type, Identity;
            internal readonly int Depth;
            internal readonly List<StorageSite> Sites = new List<StorageSite>();
            private readonly HashSet<uint> identities = new HashSet<uint>();
            internal TabLinks Tabs;
            internal StorageNode(string path, uint type, uint identity, int depth)
            { Path = path; Type = type; Identity = identity; Depth = depth; }

            internal void AddSite(uint identity, uint type, uint flags, bool hasObjectSize, string name)
            {
                // MS-OFORMS 2.5.4.1: Streamed=false means own ID-named storage; PromoteControls applies only to parents.
                // https://learn.microsoft.com/en-us/openspecs/office_file_formats/ms-oforms/ed58f23c-ec1f-43f8-a593-df2626191d27
                bool streamed = (flags & 0x10) != 0, parent = type == 14 || type == 57 || type == 7;
                Require(identity <= int.MaxValue && identities.Add(identity) && Sites.Count < 16384);
                Require((flags & ~0x4233fu) == 0 && streamed != parent && hasObjectSize == streamed);
                Require(((flags & 0x40000) != 0) == parent);
                Require(Type == 57 ? (type == 18 || type == 7) : type != 7);
                Sites.Add(new StorageSite { Identity = identity, Type = type, Streamed = streamed, Name = name });
            }
        }

        private sealed class StorageGraph
        {
            private readonly IReadOnlyDictionary<string, byte[]> streams, metadata;
            internal readonly Dictionary<string, byte[]> Comparison;
            internal readonly HashSet<string> StorageClaims = new HashSet<string>(StringComparer.Ordinal);
            internal readonly HashSet<string> StreamClaims = new HashSet<string>(StringComparer.Ordinal);
            private int siteCount;

            internal StorageGraph(IReadOnlyDictionary<string, byte[]> streams, IReadOnlyDictionary<string, byte[]> metadata)
            {
                this.streams = streams; this.metadata = metadata;
                Comparison = streams.ToDictionary(x => x.Key, x => x.Value, StringComparer.Ordinal);
            }

            internal void Parse(StorageNode node)
            {
                // Cached parent identities match Microsoft Forms coclasses, with Page represented by a FormControl storage.
                // https://learn.microsoft.com/en-us/dotnet/api/microsoft.vbe.interop.forms.frameclass
                // https://learn.microsoft.com/en-us/dotnet/api/microsoft.vbe.interop.forms.multipageclass
                // https://learn.microsoft.com/en-us/dotnet/api/microsoft.vbe.interop.forms.userformclass
                Require(StorageClaims.Add(node.Path));
                Require(metadata.TryGetValue(node.Path, out var identity));
                Require(identity != null && identity.Length == 20);
                byte[] expected = node.Type == 14 ? GraphFrameClass : node.Type == 57 ? GraphMultiPageClass : GraphFormClass;
                for (int i = 0; i < 16; i++) Require(identity[i] == expected[i]);
                byte[] form = Claim(node.Path + "/f", true), objects = Claim(node.Path + "/o", true);
                Claim(node.Path + "/\u0001CompObj", false);
                byte[] formClone = (byte[])form.Clone(), objectClone = (byte[])objects.Clone();
                ParseForm(new Reader(formClone, 0, formClone.Length, 0), new Reader(objectClone, 0, objectClone.Length, 0), node);
                siteCount += node.Sites.Count; Require(siteCount <= 16384);
                Comparison[node.Path + "/f"] = formClone; Comparison[node.Path + "/o"] = objectClone;
                if (node.Type == 57) ValidateMultiPage(node, Claim(node.Path + "/x", true));
            }

            private byte[] Claim(string path, bool required)
            {
                if (!streams.TryGetValue(path, out var bytes)) { Require(!required); return null; }
                Require(StreamClaims.Add(path)); return bytes;
            }

            private static void ValidateMultiPage(StorageNode node, byte[] bytes)
            {
                // MS-OFORMS 2.1.2.3 and 2.2.6. The x stream remains exact even for its ignored first PageProperties.
                // https://learn.microsoft.com/en-us/openspecs/office_file_formats/ms-oforms/aa2a2c6b-05e2-461f-a366-c10f15ef7731
                // https://officeprotocoldoc.z19.web.core.windows.net/files/MS-OFORMS/%5BMS-OFORMS%5D.pdf
                var pages = node.Sites.Where(x => x.Type == 7).ToArray();
                Require(node.Tabs != null && node.Sites.Count == pages.Length + 1 && node.Tabs.Items.Count == pages.Length && node.Tabs.Names.Count == pages.Length);
                // Parse a clone to validate x without changing its bytes, including hypothetical padding.
                var reader = new Reader((byte[])bytes.Clone(), 0, bytes.Length, 0);
                for (int i = 0; i <= pages.Length; i++)
                {
                    Reader page = reader.Block(0x0200); uint mask = page.UInt32(); Require((mask & ~6u) == 0);
                    page.Field(mask, 1, 4); page.Field(mask, 2, 4); page.Finish();
                }
                Reader block = reader.Block(0x0200); uint properties = block.UInt32(); Require((properties & ~0x0eu) == 0);
                uint count = block.Field(properties, 1, 4);
                block.Field(properties, 2, 4); // Optional ID retained; the specification does not equate it to outer siteID.
                // fFlags is a value flag, with no scalar payload in MultiPagePropertiesDataBlock.
                block.Finish(); Require(count == (uint)pages.Length && count <= (uint)(reader.Remaining / 4));
                var order = new HashSet<uint>();
                for (int i = 0; i < pages.Length; i++)
                {
                    uint pageID = reader.UInt32(); Require(order.Add(pageID));
                    StorageSite page = pages.FirstOrDefault(x => x.Identity == pageID);
                    // PageIDs bind each tab index to an exact Page storage. TabNames and site names are independent:
                    // the format specifies their ordering, not equality of the independently persisted names.
                    Require(page != null);
                }
                reader.Finish();
            }
        }
    }
}

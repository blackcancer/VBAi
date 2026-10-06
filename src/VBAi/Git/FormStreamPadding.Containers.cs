using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;

namespace VBAi
{

    /// <summary>Owns the form stream padding state and operations.</summary>
    internal static partial class FormStreamPadding
    {

        /// <summary>Maintains the graph form class state for form stream padding.</summary>
        private static readonly byte[] GraphFormClass = new Guid("C62A69F0-16DC-11CE-9E98-00AA00574A4F").ToByteArray();

        /// <summary>Maintains the graph frame class state for form stream padding.</summary>
        private static readonly byte[] GraphFrameClass = new Guid("6E182020-F460-11CE-9BCD-00AA00608E01").ToByteArray();

        /// <summary>Maintains the graph multi page class state for form stream padding.</summary>
        private static readonly byte[] GraphMultiPageClass = new Guid("46E31370-3F7A-11CE-BED6-00AA00611080").ToByteArray();

        /// <summary>Normalizes comparison clones only after the complete bounded UserForm storage graph is validated.</summary>
        /// <remarks>Any refusal returns the exact original dictionary and arrays. Transport/import bytes are never modified.</remarks>
        /// <param name="streams">i read only dictionary&lt;string, byte[]&gt; that supplies the streams for this operation.</param>
        /// <param name="storageMetadata">i read only dictionary&lt;string, byte[]&gt; that supplies the storage metadata for this operation.</param>
        /// <param name="rootPath">Path used for the root path being processed.</param>
        /// <returns>i read only dictionary&lt;string, byte[]&gt; produced by the operation for normalize graph on form stream padding.</returns>
        internal static IReadOnlyDictionary<string, byte[]> NormalizeGraph(IReadOnlyDictionary<string, byte[]> streams,
            IReadOnlyDictionary<string, byte[]> storageMetadata, string rootPath = "")
        {
            if (streams == null || storageMetadata == null || rootPath == null) return streams;
            try
            {
                List<StorageNode> nodes;
                return ReadGraph(streams, storageMetadata, rootPath, out nodes).Comparison;
            }
            catch (UnsupportedLayoutException) { return streams; }
        }

        /// <summary>Extracts declared standard fonts only when every storage and control in the graph validates.</summary>
        /// <param name="streams">i read only dictionary&lt;string, byte[]&gt; that supplies the streams for this operation.</param>
        /// <param name="storageMetadata">i read only dictionary&lt;string, byte[]&gt; that supplies the storage metadata for this operation.</param>
        /// <returns>form font binding[] produced by the operation for read font bindings on form stream padding.</returns>
        internal static FormFontBinding[] ReadFontBindings(IReadOnlyDictionary<string, byte[]> streams,
            IReadOnlyDictionary<string, byte[]> storageMetadata)
        {
            if (streams == null || storageMetadata == null) return null;
            try
            {
                List<StorageNode> nodes; ReadGraph(streams, storageMetadata, "", out nodes, true);
                Require(!nodes.Any(node => node.UnsupportedFont));
                return nodes.Where(node => node.Font != null)
                    .Select(node => new FormFontBinding(node.OwnerPath, node.Font, node.Type)).ToArray();
            }
            catch (UnsupportedLayoutException) { return null; }
        }

        /// <summary>Associates an exact persisted font payload with its validated container hierarchy.</summary>
        internal sealed class FormFontBinding
        {

            /// <summary>Keeps the owner path path available to form font binding.</summary>
            internal readonly string OwnerPath;

            /// <summary>Maintains the descriptor state for form font binding.</summary>
            internal readonly byte[] Descriptor;

            /// <summary>Maintains the type state for form font binding.</summary>
            internal readonly uint Type;

            /// <summary>Initializes a FormFontBinding instance with the supplied state.</summary>
            /// <param name="ownerPath">Path used for the owner path being processed.</param>
            /// <param name="descriptor">byte[] that supplies the descriptor for this operation.</param>
            /// <param name="type">uint that supplies the type for this operation.</param>
            internal FormFontBinding(string ownerPath, byte[] descriptor, uint type)
            { OwnerPath = ownerPath; Descriptor = (byte[])descriptor.Clone(); Type = type; }
        }

        /// <summary>Reads graph for form stream padding.</summary>
        /// <param name="streams">i read only dictionary&lt;string, byte[]&gt; that supplies the streams for this operation.</param>
        /// <param name="storageMetadata">i read only dictionary&lt;string, byte[]&gt; that supplies the storage metadata for this operation.</param>
        /// <param name="rootPath">Path used for the root path being processed.</param>
        /// <param name="nodes">list&lt;storage node&gt; that supplies the nodes for this operation.</param>
        /// <param name="collectOwnerPaths">Indicates whether collect owner paths is enabled.</param>
        /// <returns>storage graph produced by the operation for read graph on form stream padding.</returns>
        private static StorageGraph ReadGraph(IReadOnlyDictionary<string, byte[]> streams,
            IReadOnlyDictionary<string, byte[]> storageMetadata, string rootPath, out List<StorageNode> nodes,
            bool collectOwnerPaths = false)
        {
            nodes = new List<StorageNode>();
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
                nodes.Add(node);
                foreach (var site in node.Sites.Where(x => !x.Streamed))
                {
                    Require(node.Depth < 64);
                    string name = "i" + site.Identity.ToString("D2", CultureInfo.InvariantCulture);
                    string owner = "";
                    if (collectOwnerPaths)
                    {
                        // Bound and validate each name before concatenating paths.
                        // Comparison alone does not need native owner identities.
                        Require(site.Name != null && site.Name.Length >= 1 && site.Name.Length <= 40 &&
                            System.Text.RegularExpressions.Regex.IsMatch(site.Name, @"^[A-Za-z_][A-Za-z0-9_]*$"));
                        owner = node.OwnerPath + (node.OwnerPath.Length == 0 ? "" : "/") +
                            (site.Type == 7 ? "Pages/" : "Controls/") + site.Name;
                    }
                    pending.Push(new StorageNode(node.Path + "/" + name, site.Type, site.Identity, node.Depth + 1, owner));
                }
            }
            Require(graph.StorageClaims.Count == storageMetadata.Count && graph.StreamClaims.Count == streams.Count);
            return graph;
        }

        /// <summary>Reads design-surface metadata while retaining all property bytes and validating its optional extent.</summary>
        /// <param name="form">reader that supplies the form for this operation.</param>
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

        /// <summary>Owns the tab links state and operations.</summary>
        private sealed class TabLinks
        {

            /// <summary>Maintains the items and names state for tab links.</summary>
            internal readonly List<string> Items = new List<string>(), Names = new List<string>();
        }

        /// <summary>Owns the storage site state and operations.</summary>
        private sealed class StorageSite
        {

            /// <summary>Maintains the identity and type state for storage site.</summary>
            internal uint Identity, Type;

            /// <summary>Maintains the streamed state for storage site.</summary>
            internal bool Streamed;

            /// <summary>Maintains the name state for storage site.</summary>
            internal string Name;
        }

        /// <summary>Owns the storage node state and operations.</summary>
        private sealed class StorageNode
        {

            /// <summary>Keeps the path path available to storage node.</summary>
            internal readonly string Path;

            /// <summary>Keeps the owner path path available to storage node.</summary>
            internal readonly string OwnerPath;

            /// <summary>Maintains the font state for storage node.</summary>
            internal byte[] Font;

            /// <summary>Maintains the unsupported font state for storage node.</summary>
            internal bool UnsupportedFont;

            /// <summary>Maintains the type and identity state for storage node.</summary>
            internal readonly uint Type, Identity;

            /// <summary>Maintains the depth state for storage node.</summary>
            internal readonly int Depth;

            /// <summary>Maintains the sites state for storage node.</summary>
            internal readonly List<StorageSite> Sites = new List<StorageSite>();

            /// <summary>Maintains the identities state for storage node.</summary>
            private readonly HashSet<uint> identities = new HashSet<uint>();

            /// <summary>Maintains the tabs state for storage node.</summary>
            internal TabLinks Tabs;

            /// <summary>Initializes a StorageNode instance with the supplied state.</summary>
            /// <param name="path">Path used for the path being processed.</param>
            /// <param name="type">uint that supplies the type for this operation.</param>
            /// <param name="identity">uint that supplies the identity for this operation.</param>
            /// <param name="depth">int that supplies the depth for this operation.</param>
            /// <param name="ownerPath">Path used for the owner path being processed.</param>
            internal StorageNode(string path, uint type, uint identity, int depth, string ownerPath = "")
            { Path = path; Type = type; Identity = identity; Depth = depth; OwnerPath = ownerPath; }

            /// <summary>Adds site for storage node.</summary>
            /// <param name="identity">uint that supplies the identity for this operation.</param>
            /// <param name="type">uint that supplies the type for this operation.</param>
            /// <param name="flags">uint that supplies the flags for this operation.</param>
            /// <param name="hasObjectSize">Indicates whether has object size is enabled.</param>
            /// <param name="name">Text that supplies the name value. Use the format required by the calling operation.</param>
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

        /// <summary>Owns the storage graph state and operations.</summary>
        private sealed class StorageGraph
        {

            /// <summary>Maintains the streams and metadata state for storage graph.</summary>
            private readonly IReadOnlyDictionary<string, byte[]> streams, metadata;

            /// <summary>Maintains the comparison state for storage graph.</summary>
            internal readonly Dictionary<string, byte[]> Comparison;

            /// <summary>Maintains the storage claims state for storage graph.</summary>
            internal readonly HashSet<string> StorageClaims = new HashSet<string>(StringComparer.Ordinal);

            /// <summary>Maintains the stream claims state for storage graph.</summary>
            internal readonly HashSet<string> StreamClaims = new HashSet<string>(StringComparer.Ordinal);

            /// <summary>Counts the site count maintained by storage graph.</summary>
            private int siteCount;

            /// <summary>Initializes a StorageGraph instance with the supplied state.</summary>
            /// <param name="streams">i read only dictionary&lt;string, byte[]&gt; that supplies the streams for this operation.</param>
            /// <param name="metadata">i read only dictionary&lt;string, byte[]&gt; that supplies the metadata for this operation.</param>
            internal StorageGraph(IReadOnlyDictionary<string, byte[]> streams, IReadOnlyDictionary<string, byte[]> metadata)
            {
                this.streams = streams; this.metadata = metadata;
                Comparison = streams.ToDictionary(x => x.Key, x => x.Value, StringComparer.Ordinal);
            }

            /// <summary>Parses  for storage graph.</summary>
            /// <param name="node">storage node that supplies the node for this operation.</param>
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

            /// <summary>Handles claim for storage graph.</summary>
            /// <param name="path">Path used for the path being processed.</param>
            /// <param name="required">Indicates whether required is enabled.</param>
            /// <returns>byte[] produced by the operation for claim on storage graph.</returns>
            private byte[] Claim(string path, bool required)
            {
                if (!streams.TryGetValue(path, out var bytes)) { Require(!required); return null; }
                Require(StreamClaims.Add(path)); return bytes;
            }

            /// <summary>Validates multi page for storage graph.</summary>
            /// <param name="node">storage node that supplies the node for this operation.</param>
            /// <param name="bytes">byte[] that supplies the bytes for this operation.</param>
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

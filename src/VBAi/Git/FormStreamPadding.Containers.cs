using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;

namespace VBAi
{

    /// <summary>Validates and canonicalizes a complete nested MSForms UserForm storage graph.</summary>
    internal static partial class FormStreamPadding
    {

        /// <summary>CFB CLSID identifying a UserForm storage.</summary>
        private static readonly byte[] GraphFormClass = new Guid("C62A69F0-16DC-11CE-9E98-00AA00574A4F").ToByteArray();

        /// <summary>CFB CLSID identifying a Frame storage.</summary>
        private static readonly byte[] GraphFrameClass = new Guid("6E182020-F460-11CE-9BCD-00AA00608E01").ToByteArray();

        /// <summary>CFB CLSID identifying a MultiPage storage.</summary>
        private static readonly byte[] GraphMultiPageClass = new Guid("46E31370-3F7A-11CE-BED6-00AA00611080").ToByteArray();

        /// <summary>Normalizes comparison clones only after the complete bounded UserForm storage graph is validated.</summary>
        /// <remarks>Any refusal returns the exact original dictionary and arrays. Transport/import bytes are never modified.</remarks>
        /// <param name="streams">Logical CFB streams keyed by nested storage path.</param>
        /// <param name="storageMetadata">Twenty-byte persisted metadata records keyed by storage path.</param>
        /// <param name="rootPath">Root storage path, empty for the top-level form.</param>
        /// <returns>Comparison-only cloned streams after full graph validation, or the original dictionary on unsupported layout.</returns>
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
        /// <param name="streams">Logical streams indexed by CFB storage path.</param>
        /// <param name="storageMetadata">Validated storage class metadata.</param>
        /// <returns>Exact StdFont bindings with owner paths, or null when the graph cannot be fully validated.</returns>
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

            /// <summary>Logical path of the form or control that owns this font property.</summary>
            internal readonly string OwnerPath;

            /// <summary>Cloned exact serialized font descriptor bytes.</summary>
            internal readonly byte[] Descriptor;

            /// <summary>Persisted control type associated with the font owner.</summary>
            internal readonly uint Type;

            /// <summary>Initializes a FormFontBinding instance with the supplied state.</summary>
            /// <param name="ownerPath">Validated logical storage/control path.</param>
            /// <param name="descriptor">Persisted font descriptor; copied to prevent aliasing.</param>
            /// <param name="type">VB control type owning the descriptor.</param>
            internal FormFontBinding(string ownerPath, byte[] descriptor, uint type)
            { OwnerPath = ownerPath; Descriptor = (byte[])descriptor.Clone(); Type = type; }
        }

        /// <summary>Walks bounded Form, Frame, and MultiPage storages, validating unique child identities and complete stream ownership.</summary>
        /// <param name="streams">Logical CFB stream map.</param>
        /// <param name="storageMetadata">Storage CLSID and metadata map.</param>
        /// <param name="rootPath">Storage path to parse first.</param>
        /// <param name="nodes">Receives parsed storage nodes in traversal order.</param>
        /// <param name="collectOwnerPaths">Whether child names are required to build font-binding owner paths.</param>
        /// <returns>Graph containing cloned normalized streams and exact claims.</returns>
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
        /// <param name="form">Reader positioned at the optional DesignExtender block.</param>
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

        /// <summary>Pairs persisted MultiPage tab labels with page-storage order.</summary>
        private sealed class TabLinks
        {

            /// <summary>Tab item identifiers and display names in matching ordinal positions.</summary>
            internal readonly List<string> Items = new List<string>(), Names = new List<string>();
        }

        /// <summary>Validated child-control identity, type, stream mode, and optional name.</summary>
        private sealed class StorageSite
        {

            /// <summary>Persisted site identity and VB control type.</summary>
            internal uint Identity, Type;

            /// <summary>Whether the control payload is in the object stream rather than its own storage.</summary>
            internal bool Streamed;

            /// <summary>Optional persisted child name used for owner-path mapping.</summary>
            internal string Name;
        }

        /// <summary>Represents one nested Form/Frame/MultiPage storage and the validated children it owns.</summary>
        private sealed class StorageNode
        {

            /// <summary>CFB storage path used to claim this node's streams.</summary>
            internal readonly string Path;

            /// <summary>Logical form/control owner path used in font restoration bindings.</summary>
            internal readonly string OwnerPath;

            /// <summary>Exact supported standard-font descriptor found in this storage.</summary>
            internal byte[] Font;

            /// <summary>Marks an encountered font representation that cannot be safely transferred.</summary>
            internal bool UnsupportedFont;

            /// <summary>Persisted control type and parent-site identity for this storage.</summary>
            internal readonly uint Type, Identity;

            /// <summary>Depth from the root, bounded to prevent excessive nesting.</summary>
            internal readonly int Depth;

            /// <summary>Child sites declared by this storage.</summary>
            internal readonly List<StorageSite> Sites = new List<StorageSite>();

            /// <summary>Site identities already seen in this node, used to reject duplicates.</summary>
            private readonly HashSet<uint> identities = new HashSet<uint>();

            /// <summary>Optional tab labels collected for MultiPage child validation.</summary>
            internal TabLinks Tabs;

            /// <summary>Initializes a StorageNode instance with the supplied state.</summary>
            /// <param name="path">Storage path within the CFB tree.</param>
            /// <param name="type">Control type represented by this storage.</param>
            /// <param name="identity">Identity assigned by its parent site.</param>
            /// <param name="depth">Nesting depth from the root storage.</param>
            /// <param name="ownerPath">Optional logical path used to associate fonts with controls.</param>
            internal StorageNode(string path, uint type, uint identity, int depth, string ownerPath = "")
            { Path = path; Type = type; Identity = identity; Depth = depth; OwnerPath = ownerPath; }

            /// <summary>Validates and records one child site, enforcing unique identity and legal streamed/storage combinations.</summary>
            /// <param name="identity">Unique site ID within this storage.</param>
            /// <param name="type">Persisted control type.</param>
            /// <param name="flags">Site flags that encode stream and parent behavior.</param>
            /// <param name="hasObjectSize">Whether the site declares an object-stream extent.</param>
            /// <param name="name">Optional child name used for owner-path reconstruction.</param>
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

        /// <summary>Validates stream/storage claims and builds the comparison-only stream clones.</summary>
        private sealed class StorageGraph
        {

            /// <summary>Input logical stream data and storage metadata maps.</summary>
            private readonly IReadOnlyDictionary<string, byte[]> streams, metadata;

            /// <summary>Cloned streams with only proven padding changes applied.</summary>
            internal readonly Dictionary<string, byte[]> Comparison;

            /// <summary>Storage paths parsed exactly once.</summary>
            internal readonly HashSet<string> StorageClaims = new HashSet<string>(StringComparer.Ordinal);

            /// <summary>Stream paths claimed exactly once.</summary>
            internal readonly HashSet<string> StreamClaims = new HashSet<string>(StringComparer.Ordinal);

            /// <summary>Total sites parsed across the graph, bounded to 16,384.</summary>
            private int siteCount;

            /// <summary>Initializes a StorageGraph instance with the supplied state.</summary>
            /// <param name="streams">Input stream map whose arrays are copied for comparison.</param>
            /// <param name="metadata">Persisted storage metadata map.</param>
            internal StorageGraph(IReadOnlyDictionary<string, byte[]> streams, IReadOnlyDictionary<string, byte[]> metadata)
            {
                this.streams = streams; this.metadata = metadata;
                Comparison = streams.ToDictionary(x => x.Key, x => x.Value, StringComparer.Ordinal);
            }

            /// <summary>Validates one storage CLSID, claims its required streams, parses children, and updates comparison clones.</summary>
            /// <param name="node">Storage node being validated.</param>
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

            /// <summary>Claims a stream path once, rejecting duplicates and missing required streams.</summary>
            /// <param name="path">Logical CFB stream path.</param>
            /// <param name="required">Whether absence invalidates the graph.</param>
            /// <returns>Stream bytes, or null when an optional stream is absent.</returns>
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

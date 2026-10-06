using System;

namespace VBAi
{

    /// <summary>Canonicalizes only documented MS-OFORMS padding in supported logical f/o streams.</summary>
    internal static partial class FormStreamPadding
    {
        // Primary grammar: MS-OFORMS 2.1.1.2.4, 2.2.1, 2.2.4, 2.2.10 and 2.3.
        // https://learn.microsoft.com/en-us/openspecs/office_file_formats/ms-oforms/622ed335-0723-4491-b271-e4767d7453e3
        // Property descriptors, lengths, masks, strings and extension bytes remain significant.
        /// <summary>Maintains the max stream bytes state for form stream padding.</summary>
        private const int MaxStreamBytes = 32 * 1024 * 1024;

        /// <summary>Identifies the text font guid associated with form stream padding.</summary>
        private static readonly byte[] TextFontGuid = new Guid("AFC20920-DA4E-11CE-B943-00AA006887B4").ToByteArray();

        /// <summary>Identifies the std font guid associated with form stream padding.</summary>
        private static readonly byte[] StdFontGuid = new Guid("0BE35203-8F91-11CE-9DE3-00AA004BB851").ToByteArray();

        /// <summary>
        /// Returns cloned streams with padding zeroed, or the original pair when any layout is unsupported.
        /// This is a comparison representation, never a resource to import. No input bytes are modified.
        /// </summary>
        /// <param name="form">byte[] that supplies the form for this operation.</param>
        /// <param name="objects">byte[] that supplies the objects for this operation.</param>
        /// <returns>byte[][] produced by the operation for normalize on form stream padding.</returns>
        internal static byte[][] Normalize(byte[] form, byte[] objects)
        {
            var original = new[] { form, objects };
            if (form == null || objects == null || form.Length > MaxStreamBytes || objects.Length > MaxStreamBytes)
                return original;
            var normalizedForm = (byte[])form.Clone();
            var normalizedObjects = (byte[])objects.Clone();
            try
            {
                ParseForm(new Reader(normalizedForm, 0, normalizedForm.Length, 0),
                    new Reader(normalizedObjects, 0, normalizedObjects.Length, 0));
                return new[] { normalizedForm, normalizedObjects };
            }
            catch (UnsupportedLayoutException)
            {
                // All-or-nothing: an unknown child or trailing extension invalidates every cleared byte.
                return original;
            }
        }

        /// <summary>Determines whether it has  for form stream padding.</summary>
        /// <param name="mask">uint that supplies the mask for this operation.</param>
        /// <param name="bit">int that supplies the bit for this operation.</param>
        /// <returns>Boolean indicating the result of the check for has on form stream padding.</returns>
        private static bool Has(uint mask, int bit) { return (mask & (1u << bit)) != 0; }

        /// <summary>Requires  for form stream padding.</summary>
        /// <param name="condition">Indicates whether condition is enabled.</param>
        private static void Require(bool condition) { if (!condition) throw new UnsupportedLayoutException(); }

        /// <summary>Parses form for form stream padding.</summary>
        /// <param name="form">reader that supplies the form for this operation.</param>
        /// <param name="objects">reader that supplies the objects for this operation.</param>
        /// <param name="node">storage node that supplies the node for this operation.</param>
        private static void ParseForm(Reader form, Reader objects, StorageNode node = null)
        {
            Reader block = form.Block(0x0400);
            uint mask = block.UInt32();
            Require((mask & ~0x0fffbfceu) == 0 && Has(mask, 27));
            // Pictures and mouse icons require an additional codec. Do not normalize partial forms.
            Require(!Has(mask, 15) && !Has(mask, 21));
            block.Field(mask, 1, 4); block.Field(mask, 2, 4); block.Field(mask, 3, 4);
            uint flags = block.Field(mask, 6, 4, 4);
            Require((flags & ~0x0000c004u) == 0 && (node != null || (flags & 0x4000) == 0));
            block.Field(mask, 7, 1); block.Field(mask, 8, 1); block.Field(mask, 9, 1);
            block.Field(mask, 13, 4); block.Field(mask, 16, 1); block.Field(mask, 17, 1);
            block.Field(mask, 18, 4);
            uint caption = block.Field(mask, 19, 4);
            if (Has(mask, 20)) Require(block.Field(mask, 20, 2) == 0xffff);
            block.Field(mask, 22, 4); block.Field(mask, 23, 1); block.Field(mask, 25, 1);
            block.Field(mask, 26, 4); block.Field(mask, 27, 4);
            block.Align(4);
            if (Has(mask, 10)) block.Skip(8);
            if (Has(mask, 11)) block.Skip(8);
            if (Has(mask, 12)) block.Skip(8);
            if (Has(mask, 19)) block.String(caption);
            block.Finish();
            if (Has(mask, 20))
            {
                ParseFormFont(form, node);
            }

            // MS-OFORMS 2.2.10.6: only the empty class table is supported.
            if ((flags & 0x8000) == 0) Require(form.UInt16() == 0);
            uint count = form.UInt32();
            uint length = form.UInt32();
            Require(count <= (uint)(form.Remaining / 8));
            Reader sites = form.Section(length);
            int depthStart = sites.Position;
            uint represented = 0;
            while (represented < count)
            {
                Require(sites.Byte() == 0); // Direct children within this storage; nested parents have separate storages.
                byte typeOrCount = sites.Byte();
                if ((typeOrCount & 0x80) != 0)
                {
                    uint run = (uint)(typeOrCount & 0x7f);
                    Require(run != 0 && run <= count - represented && sites.Byte() == 1);
                    represented += run;
                }
                else
                {
                    Require(typeOrCount == 1);
                    represented++;
                }
            }
            sites.Padding((4 - ((sites.Position - depthStart) & 3)) & 3);
            for (uint i = 0; i < count; i++) ParseSite(sites, objects, node);
            sites.Finish();
            if ((flags & 0x4000) != 0) ParseDesignExtender(form);
            form.Finish();
            objects.Finish();
        }

        // MS-OFORMS 2.4.6 / 2.4.12. Font bytes remain significant; only their exact
        // documented extent is consumed, so subsequent site padding can be parsed.
        /// <summary>Parses form font for form stream padding.</summary>
        /// <param name="form">reader that supplies the form for this operation.</param>
        /// <param name="node">storage node that supplies the node for this operation.</param>
        private static void ParseFormFont(Reader form, StorageNode node = null)
        {
            int start = form.Position;
            bool text = true, standard = true;
            for (int i = 0; i < TextFontGuid.Length; i++)
            {
                byte value = form.Byte();
                text &= value == TextFontGuid[i];
                standard &= value == StdFontGuid[i];
            }
            if (text) { ParseText(form); if (node != null) node.UnsupportedFont = true; return; }
            Require(standard && form.Byte() == 1);
            form.UInt16(); // Signed charset, retained without interpreting its value.
            Require((form.Byte() & ~0x0e) == 0); // Bold and unused FONTFLAGS must be zero.
            Require(form.UInt16() <= 1000);
            uint height = form.UInt32();
            Require(height > 0 && height <= 655350000);
            int length = form.Byte();
            Require(length < 32);
            for (int i = 0; i < length; i++) Require(form.Byte() < 128);
            if (node != null) node.Font = form.Copy(start + 16, form.Position - start - 16);
        }

        /// <summary>Parses site for form stream padding.</summary>
        /// <param name="sites">reader that supplies the sites for this operation.</param>
        /// <param name="objects">reader that supplies the objects for this operation.</param>
        /// <param name="node">storage node that supplies the node for this operation.</param>
        private static void ParseSite(Reader sites, Reader objects, StorageNode node = null)
        {
            Reader block = sites.Block(0);
            uint mask = block.UInt32();
            Require((mask & ~0x00007bffu) == 0 && (node != null || Has(mask, 5)) && Has(mask, 7));
            uint name = block.Field(mask, 0, 4), tag = block.Field(mask, 1, 4);
            uint identity = block.Field(mask, 2, 4); block.Field(mask, 3, 4);
            uint flags = block.Field(mask, 4, 4, 0x33);
            bool streamed = (flags & 0x10) != 0;
            Require(node != null || (streamed && (flags & 0x40000) == 0));
            uint objectSize = block.Field(mask, 5, 4);
            block.Field(mask, 6, 2);
            uint type = block.Field(mask, 7, 2);
            Require(type == 17 || type == 21 || IsMorphType(type) || type == 16 || type == 47 || type == 18 || type == 12 ||
                (node != null && !streamed && (type == 14 || type == 57 || type == 7)));
            block.Field(mask, 9, 2);
            uint tooltip = block.Field(mask, 11, 4), license = block.Field(mask, 12, 4);
            uint source = block.Field(mask, 13, 4), rows = block.Field(mask, 14, 4);
            block.Align(4);
            string controlName = null;
            if (Has(mask, 0))
            {
                if (node == null) block.String(name);
                else controlName = block.StringValue(name);
            }
            if (Has(mask, 1)) block.String(tag);
            if (Has(mask, 8)) block.Skip(8);
            if (Has(mask, 11)) block.String(tooltip);
            if (Has(mask, 12)) block.String(license);
            if (Has(mask, 13)) block.String(source);
            if (Has(mask, 14)) block.String(rows);
            block.Finish();
            if (node != null) node.AddSite(identity, type, flags, Has(mask, 5), controlName);
            if (!streamed) return;
            Reader control = objects.Section(objectSize);
            if (IsMorphType(type)) ParseMorph(control, type);
            else if (type == 17 || type == 21) ParseLeaf(control, type == 21);
            else if (type == 18 && node?.Type == 57)
            {
                Require(node.Tabs == null);
                ParseTabStrip(control, node.Tabs = new TabLinks());
            }
            else Require(ParseAdditionalControl(control, type));
            control.Finish();
        }

        /// <summary>Parses leaf for form stream padding.</summary>
        /// <param name="control">reader that supplies the control for this operation.</param>
        /// <param name="label">Indicates whether label is enabled.</param>
        private static void ParseLeaf(Reader control, bool label)
        {
            Reader block = control.Block(0x0200);
            uint mask = block.UInt32();
            Require((mask & ~(label ? 0x1fffu : 0x7ffu)) == 0 && Has(mask, 5));
            Require(!Has(mask, label ? 10 : 7) && !Has(mask, label ? 12 : 10));
            block.Field(mask, 0, 4); block.Field(mask, 1, 4); block.Field(mask, 2, 4);
            uint caption = block.Field(mask, 3, 4);
            block.Field(mask, 4, 4); block.Field(mask, 6, 1);
            if (label)
            {
                block.Field(mask, 7, 4); block.Field(mask, 8, 2); block.Field(mask, 9, 2);
                block.Field(mask, 11, 2);
            }
            else block.Field(mask, 8, 2);
            block.Align(4);
            if (Has(mask, 3)) block.String(caption);
            block.Skip(8); // Required Size.
            block.Finish();
            ParseText(control);
        }

        /// <summary>Parses text for form stream padding.</summary>
        /// <param name="control">reader that supplies the control for this operation.</param>
        private static void ParseText(Reader control)
        {
            Reader block = control.Block(0x0200);
            uint mask = block.UInt32();
            Require((mask & ~0xf7u) == 0);
            uint name = block.Field(mask, 0, 4);
            block.Field(mask, 1, 4); block.Field(mask, 2, 4);
            block.Field(mask, 4, 1); block.Field(mask, 5, 1); block.Field(mask, 6, 1);
            block.Field(mask, 7, 2);
            block.Align(4);
            if (Has(mask, 0)) block.String(name);
            block.Finish();
        }

        /// <summary>Bounded little-endian reader; clears only padding reached by the property grammar.</summary>
        private sealed class Reader
        {

            /// <summary>Maintains the bytes state for reader.</summary>
            private readonly byte[] bytes;

            /// <summary>Maintains the end and origin state for reader.</summary>
            private readonly int end, origin;

            /// <summary>Gets or sets the position.</summary>
            /// <value>Current position exposed by reader.</value>
            internal int Position { get; private set; }

            /// <summary>Gets the remaining.</summary>
            /// <value>Current remaining exposed by reader.</value>
            internal int Remaining { get { return end - Position; } }

            /// <summary>Initializes a Reader instance with the supplied state.</summary>
            /// <param name="bytes">byte[] that supplies the bytes for this operation.</param>
            /// <param name="start">int that supplies the start for this operation.</param>
            /// <param name="length">int that supplies the length for this operation.</param>
            /// <param name="origin">int that supplies the origin for this operation.</param>
            internal Reader(byte[] bytes, int start, int length, int origin)
            {
                Require(start >= 0 && length >= 0 && start <= bytes.Length - length);
                this.bytes = bytes; Position = start; end = start + length; this.origin = origin;
            }

            /// <summary>Handles byte for reader.</summary>
            /// <returns>byte produced by the operation for byte on reader.</returns>
            internal byte Byte() { Require(Remaining >= 1); return bytes[Position++]; }

            /// <summary>Handles u int16 for reader.</summary>
            /// <returns>ushort produced by the operation for u int16 on reader.</returns>
            internal ushort UInt16() { uint a = Byte(); return (ushort)(a | ((uint)Byte() << 8)); }

            /// <summary>Handles u int32 for reader.</summary>
            /// <returns>uint produced by the operation for u int32 on reader.</returns>
            internal uint UInt32() { uint a = UInt16(); return a | ((uint)UInt16() << 16); }

            /// <summary>Handles skip for reader.</summary>
            /// <param name="length">int that supplies the length for this operation.</param>
            internal void Skip(int length) { Require(length >= 0 && length <= Remaining); Position += length; }

            /// <summary>Handles copy for reader.</summary>
            /// <param name="start">int that supplies the start for this operation.</param>
            /// <param name="length">int that supplies the length for this operation.</param>
            /// <returns>byte[] produced by the operation for copy on reader.</returns>
            internal byte[] Copy(int start, int length)
            {
                Require(start >= 0 && length >= 0 && start <= bytes.Length - length);
                var copy = new byte[length]; Buffer.BlockCopy(bytes, start, copy, 0, length); return copy;
            }

            /// <summary>Handles padding for reader.</summary>
            /// <param name="length">int that supplies the length for this operation.</param>
            internal void Padding(int length)
            {
                Require(length >= 0 && length <= Remaining);
                Array.Clear(bytes, Position, length); Position += length;
            }

            /// <summary>Handles align for reader.</summary>
            /// <param name="alignment">int that supplies the alignment for this operation.</param>
            internal void Align(int alignment) { Padding((alignment - ((Position - origin) % alignment)) % alignment); }

            /// <summary>Handles field for reader.</summary>
            /// <param name="mask">uint that supplies the mask for this operation.</param>
            /// <param name="bit">int that supplies the bit for this operation.</param>
            /// <param name="size">int that supplies the size for this operation.</param>
            /// <param name="defaultValue">uint that supplies the default value for this operation.</param>
            /// <returns>uint produced by the operation for field on reader.</returns>
            internal uint Field(uint mask, int bit, int size, uint defaultValue = 0)
            {
                if (!Has(mask, bit)) return defaultValue;
                Align(size);
                return size == 4 ? UInt32() : size == 2 ? UInt16() : Byte();
            }

            /// <summary>Handles string for reader.</summary>
            /// <param name="descriptor">uint that supplies the descriptor for this operation.</param>
            internal void String(uint descriptor)
            {
                uint length = descriptor & 0x7fffffffu;
                Require(length <= (uint)Remaining && ((descriptor & 0x80000000u) != 0 || (length & 1) == 0));
                Skip((int)length);
                Padding((int)((4 - (length & 3)) & 3));
            }

            /// <summary>Handles string value for reader.</summary>
            /// <param name="descriptor">uint that supplies the descriptor for this operation.</param>
            /// <returns>Text produced by the operation for string value on reader.</returns>
            internal string StringValue(uint descriptor)
            {
                uint length = descriptor & 0x7fffffffu;
                Require(length <= (uint)Remaining && ((descriptor & 0x80000000u) != 0 || (length & 1) == 0));
                string value;
                if ((descriptor & 0x80000000u) != 0)
                {
                    // Compressed fmString stores the low byte of each Unicode character, not ANSI bytes.
                    var characters = new char[(int)length];
                    for (int i = 0; i < characters.Length; i++) characters[i] = (char)bytes[Position + i];
                    value = new string(characters);
                }
                else
                {
                    try { value = new System.Text.UnicodeEncoding(false, false, true).GetString(bytes, Position, (int)length); }
                    catch (System.Text.DecoderFallbackException) { throw new UnsupportedLayoutException(); }
                }
                String(descriptor);
                return value;
            }

            /// <summary>Handles block for reader.</summary>
            /// <param name="version">ushort that supplies the version for this operation.</param>
            /// <returns>reader produced by the operation for block on reader.</returns>
            internal Reader Block(ushort version)
            {
                int start = Position;
                Require(UInt16() == version);
                uint length = UInt16();
                Require(length >= 4);
                Reader result = Section(length, start);
                return result;
            }

            /// <summary>Handles section for reader.</summary>
            /// <param name="length">uint that supplies the length for this operation.</param>
            /// <param name="alignmentOrigin">int that supplies the alignment origin for this operation.</param>
            /// <returns>reader produced by the operation for section on reader.</returns>
            internal Reader Section(uint length, int? alignmentOrigin = null)
            {
                Require(length <= (uint)Remaining);
                var result = new Reader(bytes, Position, (int)length, alignmentOrigin ?? Position);
                Position += (int)length;
                return result;
            }

            /// <summary>Handles finish for reader.</summary>
            internal void Finish() { Require(Position == end); }
        }

        /// <summary>Owns the unsupported layout exception state and operations.</summary>
        private sealed class UnsupportedLayoutException : Exception { }
    }
}

using System;

namespace VBAi
{

    /// <summary>Canonicalizes only documented MS-OFORMS padding in supported logical f/o streams.</summary>
    internal static partial class FormStreamPadding
    {
        // Primary grammar: MS-OFORMS 2.1.1.2.4, 2.2.1, 2.2.4, 2.2.10 and 2.3.
        // https://learn.microsoft.com/en-us/openspecs/office_file_formats/ms-oforms/622ed335-0723-4491-b271-e4767d7453e3
        // Property descriptors, lengths, masks, strings and extension bytes remain significant.
        /// <summary>Rejects comparison normalization when either logical stream exceeds 32 MiB.</summary>
        private const int MaxStreamBytes = 32 * 1024 * 1024;

        /// <summary>MS-OFORMS TextFont record GUID in the byte order used by the persisted stream.</summary>
        private static readonly byte[] TextFontGuid = new Guid("AFC20920-DA4E-11CE-B943-00AA006887B4").ToByteArray();

        /// <summary>MS-OFORMS StdFont record GUID in the byte order used by the persisted stream.</summary>
        private static readonly byte[] StdFontGuid = new Guid("0BE35203-8F91-11CE-9DE3-00AA004BB851").ToByteArray();

        /// <summary>
        /// Returns cloned streams with padding zeroed, or the original pair when any layout is unsupported.
        /// This is a comparison representation, never a resource to import. No input bytes are modified.
        /// </summary>
        /// <param name="form">Logical MS-OFORMS <c>f</c> stream; never modified.</param>
        /// <param name="objects">Logical MS-OFORMS <c>o</c> stream; never modified.</param>
        /// <returns>Cloned <c>f</c>/<c>o</c> streams with only documented padding zeroed, or the original pair if parsing rejects any layout.</returns>
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

        /// <summary>Tests one property-presence bit in an MS-OFORMS mask.</summary>
        /// <param name="mask">32-bit property mask.</param>
        /// <param name="bit">Zero-based bit index.</param>
        /// <returns>True when the selected bit is set.</returns>
        private static bool Has(uint mask, int bit) { return (mask & (1u << bit)) != 0; }

        /// <summary>Aborts normalization when a stream does not match the supported grammar.</summary>
        /// <param name="condition">Required layout or value predicate.</param>
        private static void Require(bool condition) { if (!condition) throw new UnsupportedLayoutException(); }

        /// <summary>Consumes one form header and its child-site records, clearing only padding validated by the supported grammar.</summary>
        /// <param name="form">Bounded reader over the logical form stream.</param>
        /// <param name="objects">Bounded reader over the matching object stream.</param>
        /// <param name="node">Optional storage node populated while parsing an imported form.</param>
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
        /// <summary>Consumes supported font records while preserving font bytes; unsupported text fonts reject import parsing.</summary>
        /// <param name="form">Reader positioned at the font record.</param>
        /// <param name="node">Optional import node receiving the exact standard-font payload.</param>
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

        /// <summary>Parses one child site and its optional object payload, rejecting unknown object types or property masks.</summary>
        /// <param name="sites">Reader over the form's bounded site section.</param>
        /// <param name="objects">Reader over the paired object stream.</param>
        /// <param name="node">Optional import storage node receiving site identity and control metadata.</param>
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

        /// <summary>Parses a supported Label or Image leaf object's fixed property blocks.</summary>
        /// <param name="control">Reader bounded to the object's declared byte length.</param>
        /// <param name="label">True selects the Label property layout; false selects Image.</param>
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

        /// <summary>Consumes the fixed text-property block and its optional name string.</summary>
        /// <param name="control">Reader bounded to the containing object's stream extent.</param>
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

            /// <summary>Backing stream bytes; padding positions may be zeroed during parsing.</summary>
            private readonly byte[] bytes;

            /// <summary>Exclusive end offset and alignment origin for this bounded reader.</summary>
            private readonly int end, origin;

            /// <summary>Gets the current absolute offset into the backing stream.</summary>
            /// <value>Offset advanced by each read and skip.</value>
            internal int Position { get; private set; }

            /// <summary>Gets the unread byte count within this reader's exclusive end.</summary>
            /// <value>Never negative for a valid reader.</value>
            internal int Remaining { get { return end - Position; } }

            /// <summary>Creates a bounded cursor over one byte-array extent; reads cannot pass its end offset.</summary>
            /// <param name="bytes">Backing logical stream.</param>
            /// <param name="start">Absolute starting offset.</param>
            /// <param name="length">Byte extent available to this reader.</param>
            /// <param name="origin">Alignment origin used by the stream grammar.</param>
            internal Reader(byte[] bytes, int start, int length, int origin)
            {
                Require(start >= 0 && length >= 0 && start <= bytes.Length - length);
                this.bytes = bytes; Position = start; end = start + length; this.origin = origin;
            }

            /// <summary>Reads the next byte and advances the cursor by one.</summary>
            /// <returns>Next byte.</returns>
            internal byte Byte() { Require(Remaining >= 1); return bytes[Position++]; }

            /// <summary>Reads an unsigned 16-bit little-endian value.</summary>
            /// <returns>Decoded value.</returns>
            internal ushort UInt16() { uint a = Byte(); return (ushort)(a | ((uint)Byte() << 8)); }

            /// <summary>Reads an unsigned 32-bit little-endian value.</summary>
            /// <returns>Decoded value.</returns>
            internal uint UInt32() { uint a = UInt16(); return a | ((uint)UInt16() << 16); }

            /// <summary>Advances over a bounded byte range without changing its contents.</summary>
            /// <param name="length">Number of bytes to consume.</param>
            internal void Skip(int length) { Require(length >= 0 && length <= Remaining); Position += length; }

            /// <summary>Copies a range from the backing stream without advancing this reader.</summary>
            /// <param name="start">Absolute offset in the backing stream.</param>
            /// <param name="length">Number of bytes to copy.</param>
            /// <returns>New byte array containing the requested range.</returns>
            internal byte[] Copy(int start, int length)
            {
                Require(start >= 0 && length >= 0 && start <= bytes.Length - length);
                var copy = new byte[length]; Buffer.BlockCopy(bytes, start, copy, 0, length); return copy;
            }

            /// <summary>Zeroes exactly the supported padding bytes and advances past them.</summary>
            /// <param name="length">Padding byte count within the current section.</param>
            internal void Padding(int length)
            {
                Require(length >= 0 && length <= Remaining);
                Array.Clear(bytes, Position, length); Position += length;
            }

            /// <summary>Aligns the cursor relative to this stream's grammar origin, clearing skipped padding bytes.</summary>
            /// <param name="alignment">Required byte alignment, typically 2 or 4.</param>
            internal void Align(int alignment) { Padding((alignment - ((Position - origin) % alignment)) % alignment); }

            /// <summary>Reads an optional property value when its presence bit is set, aligning before the value.</summary>
            /// <param name="mask">Property-presence mask.</param>
            /// <param name="bit">Presence bit for this property.</param>
            /// <param name="size">Encoded width in bytes: 1, 2, or 4.</param>
            /// <param name="defaultValue">Value returned when the property is absent.</param>
            /// <returns>Decoded unsigned value or the supplied default.</returns>
            internal uint Field(uint mask, int bit, int size, uint defaultValue = 0)
            {
                if (!Has(mask, bit)) return defaultValue;
                Align(size);
                return size == 4 ? UInt32() : size == 2 ? UInt16() : Byte();
            }

            /// <summary>Consumes a length-prefixed fmString and its documented four-byte trailing alignment.</summary>
            /// <param name="descriptor">String descriptor whose high bit selects compressed form and low 31 bits give byte length.</param>
            internal void String(uint descriptor)
            {
                uint length = descriptor & 0x7fffffffu;
                Require(length <= (uint)Remaining && ((descriptor & 0x80000000u) != 0 || (length & 1) == 0));
                Skip((int)length);
                Padding((int)((4 - (length & 3)) & 3));
            }

            /// <summary>Decodes a supported fmString, consumes its aligned extent, and rejects malformed UTF-16.</summary>
            /// <param name="descriptor">String descriptor from the parent property record.</param>
            /// <returns>Decoded string, with compressed bytes interpreted as low-byte Unicode characters.</returns>
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

            /// <summary>Opens a versioned property block with its declared length as a hard read boundary.</summary>
            /// <param name="version">Required block version identifier.</param>
            /// <returns>Reader limited to the block payload.</returns>
            internal Reader Block(ushort version)
            {
                int start = Position;
                Require(UInt16() == version);
                uint length = UInt16();
                Require(length >= 4);
                Reader result = Section(length, start);
                return result;
            }

            /// <summary>Carves a bounded child reader from the current extent and advances the parent past it.</summary>
            /// <param name="length">Section byte length, which must fit within the remaining parent bytes.</param>
            /// <param name="alignmentOrigin">Optional absolute alignment origin; defaults to the section start.</param>
            /// <returns>Child reader constrained to the declared section.</returns>
            internal Reader Section(uint length, int? alignmentOrigin = null)
            {
                Require(length <= (uint)Remaining);
                var result = new Reader(bytes, Position, (int)length, alignmentOrigin ?? Position);
                Position += (int)length;
                return result;
            }

            /// <summary>Requires exact consumption of the bounded section, rejecting unrecognized trailing bytes.</summary>
            internal void Finish() { Require(Position == end); }
        }

        /// <summary>Internal signal that parsing encountered a layout outside the explicitly supported grammar.</summary>
        private sealed class UnsupportedLayoutException : Exception { }
    }
}

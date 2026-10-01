using System;

namespace VBAi
{
    /// <summary>Canonicalizes only documented MS-OFORMS padding in supported logical f/o streams.</summary>
    internal static partial class FormStreamPadding
    {
        // Primary grammar: MS-OFORMS 2.1.1.2.4, 2.2.1, 2.2.4, 2.2.10 and 2.3.
        // https://learn.microsoft.com/en-us/openspecs/office_file_formats/ms-oforms/622ed335-0723-4491-b271-e4767d7453e3
        // Property descriptors, lengths, masks, strings and extension bytes remain significant.
        private const int MaxStreamBytes = 32 * 1024 * 1024;
        private static readonly byte[] TextFontGuid = new Guid("AFC20920-DA4E-11CE-B943-00AA006887B4").ToByteArray();
        private static readonly byte[] StdFontGuid = new Guid("0BE35203-8F91-11CE-9DE3-00AA004BB851").ToByteArray();

        /// <summary>
        /// Returns cloned streams with padding zeroed, or the original pair when any layout is unsupported.
        /// This is a comparison representation, never a resource to import. No input bytes are modified.
        /// </summary>
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

        private static bool Has(uint mask, int bit) { return (mask & (1u << bit)) != 0; }
        private static void Require(bool condition) { if (!condition) throw new UnsupportedLayoutException(); }

        private static void ParseForm(Reader form, Reader objects)
        {
            Reader block = form.Block(0x0400);
            uint mask = block.UInt32();
            Require((mask & ~0x0fffbfceu) == 0 && Has(mask, 27));
            // Pictures and mouse icons require an additional codec. Do not normalize partial forms.
            Require(!Has(mask, 15) && !Has(mask, 21));
            block.Field(mask, 1, 4); block.Field(mask, 2, 4); block.Field(mask, 3, 4);
            uint flags = block.Field(mask, 6, 4, 4);
            Require((flags & ~0x0000c004u) == 0 && (flags & 0x4000) == 0);
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
                ParseFormFont(form);
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
                Require(sites.Byte() == 0); // Flat leaf controls only.
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
            for (uint i = 0; i < count; i++) ParseSite(sites, objects);
            sites.Finish();
            form.Finish();
            objects.Finish();
        }

        // MS-OFORMS 2.4.6 / 2.4.12. Font bytes remain significant; only their exact
        // documented extent is consumed, so subsequent site padding can be parsed.
        private static void ParseFormFont(Reader form)
        {
            bool text = true, standard = true;
            for (int i = 0; i < TextFontGuid.Length; i++)
            {
                byte value = form.Byte();
                text &= value == TextFontGuid[i];
                standard &= value == StdFontGuid[i];
            }
            if (text) { ParseText(form); return; }
            Require(standard && form.Byte() == 1);
            form.UInt16(); // Signed charset, retained without interpreting its value.
            Require((form.Byte() & ~0x0e) == 0); // Bold and unused FONTFLAGS must be zero.
            Require(form.UInt16() <= 1000);
            uint height = form.UInt32();
            Require(height > 0 && height <= 655350000);
            int length = form.Byte();
            Require(length < 32);
            for (int i = 0; i < length; i++) Require(form.Byte() < 128);
        }

        private static void ParseSite(Reader sites, Reader objects)
        {
            Reader block = sites.Block(0);
            uint mask = block.UInt32();
            Require((mask & ~0x00007bffu) == 0 && Has(mask, 5) && Has(mask, 7));
            uint name = block.Field(mask, 0, 4), tag = block.Field(mask, 1, 4);
            block.Field(mask, 2, 4); block.Field(mask, 3, 4);
            uint flags = block.Field(mask, 4, 4, 0x33);
            Require((flags & 0x10) != 0 && (flags & 0x40000) == 0);
            uint objectSize = block.Field(mask, 5, 4);
            block.Field(mask, 6, 2);
            uint type = block.Field(mask, 7, 2);
            Require(type == 17 || type == 21 || IsMorphType(type) || type == 16 || type == 47 || type == 18 || type == 12);
            block.Field(mask, 9, 2);
            uint tooltip = block.Field(mask, 11, 4), license = block.Field(mask, 12, 4);
            uint source = block.Field(mask, 13, 4), rows = block.Field(mask, 14, 4);
            block.Align(4);
            if (Has(mask, 0)) block.String(name);
            if (Has(mask, 1)) block.String(tag);
            if (Has(mask, 8)) block.Skip(8);
            if (Has(mask, 11)) block.String(tooltip);
            if (Has(mask, 12)) block.String(license);
            if (Has(mask, 13)) block.String(source);
            if (Has(mask, 14)) block.String(rows);
            block.Finish();
            Reader control = objects.Section(objectSize);
            if (IsMorphType(type)) ParseMorph(control, type);
            else if (type == 17 || type == 21) ParseLeaf(control, type == 21);
            else Require(ParseAdditionalControl(control, type));
            control.Finish();
        }

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
            private readonly byte[] bytes;
            private readonly int end, origin;
            internal int Position { get; private set; }
            internal int Remaining { get { return end - Position; } }
            internal Reader(byte[] bytes, int start, int length, int origin)
            {
                Require(start >= 0 && length >= 0 && start <= bytes.Length - length);
                this.bytes = bytes; Position = start; end = start + length; this.origin = origin;
            }
            internal byte Byte() { Require(Remaining >= 1); return bytes[Position++]; }
            internal ushort UInt16() { uint a = Byte(); return (ushort)(a | ((uint)Byte() << 8)); }
            internal uint UInt32() { uint a = UInt16(); return a | ((uint)UInt16() << 16); }
            internal void Skip(int length) { Require(length >= 0 && length <= Remaining); Position += length; }
            internal void Padding(int length)
            {
                Require(length >= 0 && length <= Remaining);
                Array.Clear(bytes, Position, length); Position += length;
            }
            internal void Align(int alignment) { Padding((alignment - ((Position - origin) % alignment)) % alignment); }
            internal uint Field(uint mask, int bit, int size, uint defaultValue = 0)
            {
                if (!Has(mask, bit)) return defaultValue;
                Align(size);
                return size == 4 ? UInt32() : size == 2 ? UInt16() : Byte();
            }
            internal void String(uint descriptor)
            {
                uint length = descriptor & 0x7fffffffu;
                Require(length <= (uint)Remaining && ((descriptor & 0x80000000u) != 0 || (length & 1) == 0));
                Skip((int)length);
                Padding((int)((4 - (length & 3)) & 3));
            }
            internal Reader Block(ushort version)
            {
                int start = Position;
                Require(UInt16() == version);
                uint length = UInt16();
                Require(length >= 4);
                Reader result = Section(length, start);
                return result;
            }
            internal Reader Section(uint length, int? alignmentOrigin = null)
            {
                Require(length <= (uint)Remaining);
                var result = new Reader(bytes, Position, (int)length, alignmentOrigin ?? Position);
                Position += (int)length;
                return result;
            }
            internal void Finish() { Require(Position == end); }
        }

        private sealed class UnsupportedLayoutException : Exception { }
    }
}

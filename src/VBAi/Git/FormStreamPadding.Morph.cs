namespace VBAi
{

    /// <summary>Owns the form stream padding state and operations.</summary>
    internal static partial class FormStreamPadding
    {
        // MS-OFORMS 2.2.5.1-.8, 2.4.5 and property applicability in 2.5.
        // https://learn.microsoft.com/en-us/openspecs/office_file_formats/ms-oforms/19014e19-67fa-4060-8c81-d1463809f117
        // A comparison clone retains every property byte; only proven alignment bytes are cleared.
        /// <summary>Determines whether morph type for form stream padding.</summary>
        /// <param name="type">uint that supplies the type for this operation.</param>
        /// <returns>Boolean indicating the result of the check for is morph type on form stream padding.</returns>
        private static bool IsMorphType(uint type) { return type >= 23 && type <= 28; }

        /// <summary>Reads the shared grammar of six streamed MorphData controls without interpreting their content.</summary>
        /// <param name="control">reader that supplies the control for this operation.</param>
        /// <param name="type">uint that supplies the type for this operation.</param>
        private static void ParseMorph(Reader control, uint type)
        {
            Reader block = control.Block(0x0200);
            uint mask = block.UInt32(), highMask = block.UInt32();
            ulong properties = mask | ((ulong)highMask << 32);
            Require((mask & ~0xbff7ffffu) == 0 && (highMask & ~1u) == 0);
            Require(Has(mask, 8) && Has(mask, 31)); // Required Size and reserved bit, neither is scalar data.
            Require((properties & ~MorphProperties(type)) == 0);
            // These fields imply GuidAndPicture StreamData, whose complete codec is not supported here.
            Require(!Has(mask, 27) && !Has(mask, 28));

            block.Field(mask, 0, 4); block.Field(mask, 1, 4); block.Field(mask, 2, 4);
            block.Field(mask, 3, 4); block.Field(mask, 4, 1);
            uint scrollBars = block.Field(mask, 5, 1);
            uint style = block.Field(mask, 6, 1, 1);
            Require(style == MorphStyle(type) || (type == 25 && style == 7));
            Require(type != 24 || scrollBars == 3);
            block.Field(mask, 7, 1); block.Field(mask, 9, 2); block.Field(mask, 10, 4);
            block.Field(mask, 11, 2);
            uint textColumn = block.Field(mask, 12, 2, 0xffff);
            uint columnCount = block.Field(mask, 13, 2, 1);
            Require(textColumn == 0xffff || textColumn <= 0x7fff);
            Require(columnCount == 0xffff || columnCount <= 0x7fff);
            block.Field(mask, 14, 2);
            uint columnInfoCount = block.Field(mask, 15, 2);
            block.Field(mask, 16, 1); block.Field(mask, 17, 1); block.Field(mask, 18, 1);
            block.Field(mask, 20, 1); block.Field(mask, 21, 1);
            uint value = block.Field(mask, 22, 4), caption = block.Field(mask, 23, 4);
            block.Field(mask, 24, 4); block.Field(mask, 25, 4);
            uint effect = block.Field(mask, 26, 4, 2);
            Require(type != 28 || effect == 2);
            block.Field(mask, 29, 2);
            uint group = block.Field(highMask, 0, 4);
            block.Align(4);

            // ExtraDataBlock starts with Size, unlike Label/CommandButton, then the three optional strings.
            block.Skip(8);
            if (Has(mask, 22)) block.String(value);
            if (Has(mask, 23)) block.String(caption);
            if (Has(highMask, 0)) block.String(group);
            block.Finish();
            ParseText(control);

            // cbMorphData excludes TextProps and this array. Each column has its own version and exact extent.
            Require(columnInfoCount <= (uint)(control.Remaining / 8));
            for (uint i = 0; i < columnInfoCount; i++) ParseMorphColumn(control);
        }

        /// <summary>Restricts mask fields to those applicable to the cached control's DisplayStyle.</summary>
        /// <param name="type">uint that supplies the type for this operation.</param>
        /// <returns>ulong produced by the operation for morph properties on form stream padding.</returns>
        private static ulong MorphProperties(uint type)
        {
            const ulong common = (1UL << 0) | (1UL << 1) | (1UL << 2) | (1UL << 6) | (1UL << 7) |
                (1UL << 8) | (1UL << 22) | (1UL << 26) | (1UL << 27) | (1UL << 31);
            const ulong list = (1UL << 4) | (1UL << 10) | (1UL << 11) | (1UL << 12) |
                (1UL << 13) | (1UL << 15) | (1UL << 16) | (1UL << 17) | (1UL << 25);
            const ulong drop = (1UL << 3) | (1UL << 18) | (1UL << 20);
            const ulong button = (1UL << 21) | (1UL << 23) | (1UL << 24) | (1UL << 28) | (1UL << 29);
            switch (type)
            {
                case 23: return common | drop | (1UL << 4) | (1UL << 5) | (1UL << 9) | (1UL << 25);
                case 24: return common | list | (1UL << 5) | (1UL << 21);
                case 25: return common | list | drop | (1UL << 14);
                case 26: case 27: return common | button | (1UL << 32);
                case 28: return common | button;
                default: throw new UnsupportedLayoutException();
            }
        }

        /// <summary>Handles morph style for form stream padding.</summary>
        /// <param name="type">uint that supplies the type for this operation.</param>
        /// <returns>uint produced by the operation for morph style on form stream padding.</returns>
        private static uint MorphStyle(uint type)
        {
            return type == 23 ? 1u : type == 24 ? 2u : type == 25 ? 3u : type - 22;
        }

        /// <summary>Reads a column's optional signed Width while preserving its exact representation.</summary>
        /// <param name="control">reader that supplies the control for this operation.</param>
        private static void ParseMorphColumn(Reader control)
        {
            Reader block = control.Block(0x0200);
            uint mask = block.UInt32();
            Require((mask & ~1u) == 0);
            block.Field(mask, 0, 4); // Signed HIMETRIC width; -1 means determined by the client.
            block.Finish();
        }
    }
}

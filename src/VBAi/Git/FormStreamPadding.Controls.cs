using System;

namespace VBAi
{
    internal static partial class FormStreamPadding
    {
        private static readonly byte[] PictureGuid = new Guid("0BE35204-8F91-11CE-9DE3-00AA004BB851").ToByteArray();

        /// <summary>Parses cached flat controls; callers retain the complete original pair on unsupported grammar.</summary>
        private static bool ParseAdditionalControl(Reader control, uint type)
        {
            switch (type)
            {
                case 16: ParseSpinOrScroll(control, false); return true;
                case 47: ParseSpinOrScroll(control, true); return true;
                case 18: ParseTabStrip(control); return true;
                case 12: ParseImage(control); return true;
                default: return false;
            }
        }

        /// <summary>Preserves all numeric properties and bounds the mandatory size and optional mouse-picture envelope.</summary>
        private static void ParseSpinOrScroll(Reader control, bool scroll)
        {
            // MS-OFORMS 2.2.7 / 2.2.8: these controls have no trailing TextProps.
            // https://learn.microsoft.com/en-us/openspecs/office_file_formats/ms-oforms/e30addb5-0251-4421-b45d-c2f90ab0fcc1
            // https://learn.microsoft.com/en-us/openspecs/office_file_formats/ms-oforms/ec5e8f6b-72f7-4682-9d8c-a7fb34c71751
            Reader block = control.Block(0x0200);
            uint mask = block.UInt32();
            Require((mask & ~(scroll ? 0x1feffu : 0x7fefu)) == 0 && Has(mask, 3));
            block.Field(mask, 0, 4); block.Field(mask, 1, 4);
            uint various = block.Field(mask, 2, 4, 0x1b);
            int previous = scroll ? 9 : 8, next = previous + 1;
            bool disabled = Has(mask, 2) && (various & 2) == 0;
            Require(Has(mask, previous) == disabled && Has(mask, next) == disabled);
            if (scroll) block.Field(mask, 4, 1);
            block.Field(mask, 5, 4); block.Field(mask, 6, 4); block.Field(mask, 7, 4);
            block.Field(mask, previous, 4); block.Field(mask, next, 4);
            if (scroll)
            {
                block.Field(mask, 11, 4); block.Field(mask, 12, 4); block.Field(mask, 13, 4);
                block.Field(mask, 14, 2); block.Field(mask, 15, 4);
            }
            else
            {
                block.Field(mask, 10, 4); block.Field(mask, 11, 4); block.Field(mask, 12, 4);
            }
            int icon = scroll ? 16 : 13;
            if (Has(mask, icon)) Require(block.Field(mask, icon, 2) == 0xffff);
            if (!scroll) block.Field(mask, 14, 1);
            block.Align(4);
            block.Skip(8); // Size is mandatory and entirely significant.
            block.Finish();
            if (Has(mask, icon)) ParsePictureEnvelope(control);
        }

        /// <summary>Reads image properties and preserves exact picture payloads without OLE decoding.</summary>
        private static void ParseImage(Reader control)
        {
            // MS-OFORMS 2.2.3. Bits2/12 are value flags without stored scalar fields.
            // https://learn.microsoft.com/en-us/openspecs/office_file_formats/ms-oforms/bf65b595-9c49-40f4-b72e-b59385708972
            Reader block = control.Block(0x0200);
            uint mask = block.UInt32();
            Require((mask & ~0x7ffcu) == 0 && Has(mask, 9));
            block.Field(mask, 3, 4); block.Field(mask, 4, 4);
            block.Field(mask, 5, 1); block.Field(mask, 6, 1); block.Field(mask, 7, 1); block.Field(mask, 8, 1);
            if (Has(mask, 10)) Require(block.Field(mask, 10, 2) == 0xffff);
            block.Field(mask, 11, 1); block.Field(mask, 13, 4);
            if (Has(mask, 14)) Require(block.Field(mask, 14, 2) == 0xffff);
            block.Align(4); block.Skip(8); block.Finish();
            if (Has(mask, 10)) ParsePictureEnvelope(control);
            if (Has(mask, 14)) ParsePictureEnvelope(control);
        }

        /// <summary>Bounds a known StdPicture envelope; every payload byte remains significant.</summary>
        private static void ParsePictureEnvelope(Reader control)
        {
            // MS-OFORMS 2.4.8 / 2.4.13. No decoder is invoked and no payload byte is cleared.
            // https://learn.microsoft.com/en-us/openspecs/office_file_formats/ms-oforms/22988498-ae45-4c31-9d85-c5ddba29c750
            // https://learn.microsoft.com/en-us/openspecs/office_file_formats/ms-oforms/8c1c088a-9775-49de-9b3b-298eeb1ddc19
            for (int i = 0; i < PictureGuid.Length; i++) Require(control.Byte() == PictureGuid[i]);
            Require(control.UInt32() == 0x746c);
            uint size = control.UInt32();
            Require(size <= (uint)control.Remaining);
            control.Skip((int)size);
        }

        /// <summary>Validates all persisted tab arrays, text properties and per-tab flags within their declared extents.</summary>
        private static void ParseTabStrip(Reader control, TabLinks links = null)
        {
            // MS-OFORMS 2.2.9: main cb excludes picture, TextProps and TabStripTabFlagData.
            // https://learn.microsoft.com/en-us/openspecs/office_file_formats/ms-oforms/26809262-3501-4dfd-8792-855981a4bc74
            // https://learn.microsoft.com/en-us/openspecs/office_file_formats/ms-oforms/ecfdeaba-4137-4aa1-96c2-e6d252f47858
            Reader block = control.Block(0x0200);
            uint mask = block.UInt32();
            Require((mask & ~0x1febf77u) == 0 && Has(mask, 4) && Has(mask, 19));
            block.Field(mask, 0, 4); block.Field(mask, 1, 4); block.Field(mask, 2, 4);
            uint items = block.Field(mask, 5, 4);
            block.Field(mask, 6, 1); block.Field(mask, 8, 4); block.Field(mask, 9, 4);
            block.Field(mask, 11, 4); block.Field(mask, 12, 4);
            uint tips = block.Field(mask, 15, 4), names = block.Field(mask, 17, 4);
            block.Field(mask, 18, 4);
            uint allocated = block.Field(mask, 20, 4), tags = block.Field(mask, 21, 4);
            uint data = block.Field(mask, 22, 4), accelerators = block.Field(mask, 23, 4);
            if (Has(mask, 24)) Require(block.Field(mask, 24, 2) == 0xffff);
            block.Align(4); block.Skip(8);
            int count = -1;
            if (Has(mask, 5)) count = ParseTabArray(block, items, links?.Items);
            CheckTabArray(block, mask, 15, tips, ref count);
            CheckTabArray(block, mask, 17, names, ref count, links?.Names);
            CheckTabArray(block, mask, 21, tags, ref count);
            CheckTabArray(block, mask, 23, accelerators, ref count);
            if (Has(mask, 20) && count >= 0) Require(allocated >= (uint)count);
            if (links != null) Require(Has(mask, 5) && Has(mask, 17) && Has(mask, 22));
            block.Finish();
            if (Has(mask, 24)) ParsePictureEnvelope(control);
            ParseText(control);
            if (Has(mask, 22))
            {
                Require(data <= (uint)(control.Remaining / 4));
                if (count >= 0) Require(data == (uint)count);
                for (uint i = 0; i < data; i++) Require((control.UInt32() & ~3u) == 0);
            }
        }

        /// <summary>Checks that each optional persisted array has the same number of tabs.</summary>
        private static void CheckTabArray(Reader block, uint mask, int bit, uint length, ref int count, System.Collections.Generic.List<string> values = null)
        {
            if (!Has(mask, bit)) return;
            int actual = ParseTabArray(block, length, values);
            if (count < 0) count = actual;
            else Require(count == actual);
        }

        /// <summary>Parses character-count ArrayString entries and clears only their documented string padding.</summary>
        private static int ParseTabArray(Reader block, uint size, System.Collections.Generic.List<string> values = null)
        {
            // Ordinary fmString descriptors count BYTES; ArrayString descriptors count CHARACTERS.
            // https://learn.microsoft.com/en-us/openspecs/office_file_formats/ms-oforms/b9cf793f-2fd5-491e-9827-871ae47e9920
            Require(size > 0);
            Reader array = block.Section(size);
            int count = 0;
            while (array.Remaining != 0)
            {
                uint descriptor = array.UInt32();
                uint characters = descriptor & 0x7fffffffu;
                ulong length = (descriptor & 0x80000000u) != 0 ? characters : (ulong)characters * 2;
                Require(length <= (ulong)array.Remaining);
                if (values == null)
                {
                    array.Skip((int)length);
                    array.Padding((int)((4 - (length & 3)) & 3));
                }
                else values.Add(array.StringValue((uint)length | (descriptor & 0x80000000u)));
                count++;
            }
            array.Finish();
            return count;
        }
    }
}

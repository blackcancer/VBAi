using System;
using System.Collections;
using System.Collections.Generic;
using System.ComponentModel;
using System.Drawing;
using System.Web.Script.Serialization;

namespace VBAi
{

    /// <summary>Owns the vbe forms state and operations.</summary>
    internal sealed partial class VbeForms
    {
        // MSForms' reserved Font alias is a PROPERTYPUT-only member, not a
        // separately readable persisted resource. The ordinary Font stays in
        // every snapshot and retains its existing getter/member error checks.
        /// <summary>Attempts to describe reserved font write only for vbe forms.</summary>
        /// <param name="target">object that supplies the target for this operation.</param>
        /// <param name="descriptor">property descriptor that supplies the descriptor for this operation.</param>
        /// <param name="info">vbe property info that supplies the info for this operation.</param>
        /// <returns>Boolean indicating the result of the check for try describe reserved font write only on vbe forms.</returns>
        internal static bool TryDescribeReservedFontWriteOnly(object target,
            PropertyDescriptor descriptor, VbePropertyInfo info)
        {
            if (!IsReservedFontDescriptor(descriptor, info)) return false;
            object metadata;
            try { metadata = VbeComPropertyAccessors.Inspect(target, descriptor.Name); }
            catch { return false; } // Unknown metadata never admits the exemption.
            return TryDescribeReservedFontWriteOnly(descriptor, metadata, info);
        }

        /// <summary>Attempts to describe reserved font write only for vbe forms.</summary>
        /// <param name="descriptor">property descriptor that supplies the descriptor for this operation.</param>
        /// <param name="metadata">object that supplies the metadata for this operation.</param>
        /// <param name="info">vbe property info that supplies the info for this operation.</param>
        /// <returns>Boolean indicating the result of the check for try describe reserved font write only on vbe forms.</returns>
        internal static bool TryDescribeReservedFontWriteOnly(PropertyDescriptor descriptor,
            object metadata, VbePropertyInfo info)
        {
            if (!IsReservedFontDescriptor(descriptor, info) ||
                !ReservedFontMetadataIsWriteOnly(metadata)) return false;
            info.Kind = "writeOnly";
            info.SetterStatus = "GetterUnavailable"; // Existing truthful catalogue status.
            info.Display = "(write-only COM alias)";
            info.Value = null;
            info.Digest = null;
            info.Members = null;
            info.Error = null;
            return true;
        }

        /// <summary>Determines whether reserved font descriptor for vbe forms.</summary>
        /// <param name="descriptor">property descriptor that supplies the descriptor for this operation.</param>
        /// <param name="info">vbe property info that supplies the info for this operation.</param>
        /// <returns>Boolean indicating the result of the check for is reserved font descriptor on vbe forms.</returns>
        private static bool IsReservedFontDescriptor(PropertyDescriptor descriptor, VbePropertyInfo info)
        {
            return descriptor != null && info != null &&
                descriptor.Name == "_Font_Reserved" && info.Name == "_Font_Reserved" &&
                descriptor.PropertyType == typeof(Font);
        }

        /// <summary>Handles reserved font metadata is write only for vbe forms.</summary>
        /// <param name="metadata">object that supplies the metadata for this operation.</param>
        /// <returns>Boolean indicating the result of the check for reserved font metadata is write only on vbe forms.</returns>
        internal static bool ReservedFontMetadataIsWriteOnly(object metadata)
        {
            if (metadata == null) return false;
            Dictionary<string, object> fields;
            try
            {
                var serializer = new JavaScriptSerializer();
                fields = serializer.Deserialize<Dictionary<string, object>>(serializer.Serialize(metadata));
            }
            catch { return false; }
            object value;
            if (fields == null || !fields.TryGetValue("Property", out value) ||
                !(value is string) || (string)value != "_Font_Reserved" ||
                !fields.TryGetValue("Discovery", out value) ||
                !(value is string) || (string)value != "IProvideClassInfo/ITypeInfo" ||
                !fields.TryGetValue("MetadataComplete", out value) || !(value is bool) || !(bool)value ||
                !fields.TryGetValue("SetterDeclared", out value) || !(value is bool) || !(bool)value ||
                !fields.TryGetValue("Errors", out value) || !(value is IList) || ((IList)value).Count != 0 ||
                !fields.TryGetValue("Interfaces", out value) || !(value is IList) || ((IList)value).Count == 0)
                return false;
            foreach (object item in (IList)value)
                if (!(item is string) || string.IsNullOrWhiteSpace((string)item)) return false;
            if (!fields.TryGetValue("Accessors", out value) || !(value is IList) || ((IList)value).Count == 0)
                return false;
            foreach (object item in (IList)value)
            {
                var accessor = item as IDictionary<string, object>;
                if (accessor == null || !accessor.TryGetValue("Name", out value) ||
                    !(value is string) || (string)value != "_Font_Reserved" ||
                    !accessor.TryGetValue("DispId", out value) || !(value is int) || (int)value != 0x7ffffdff ||
                    !accessor.TryGetValue("InvocationKind", out value) || !(value is string) ||
                    ((string)value != "INVOKE_PROPERTYPUT" && (string)value != "INVOKE_PROPERTYPUTREF"))
                    return false;
            }
            return true;
        }
    }
}

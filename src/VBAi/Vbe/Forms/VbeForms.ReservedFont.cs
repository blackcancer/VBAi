using System;
using System.Collections;
using System.Collections.Generic;
using System.ComponentModel;
using System.Drawing;
using System.Web.Script.Serialization;

namespace VBAi
{

    /// <summary>Handles the MSForms reserved Font property alias when COM metadata proves it is write-only.</summary>
    internal sealed partial class VbeForms
    {
        // MSForms' reserved Font alias is a PROPERTYPUT-only member, not a
        // separately readable persisted resource. The ordinary Font stays in
        // every snapshot and retains its existing getter/member error checks.
        /// <summary>Marks the reserved alias as write-only only after reading its live COM type information.</summary>
        /// <param name="target">Live control whose COM metadata is inspected; inspection failure leaves it unclassified.</param>
        /// <param name="descriptor">Reflected property descriptor that must identify the reserved Font alias.</param>
        /// <param name="info">Catalogue entry updated with getter-unavailable status when metadata is conclusive.</param>
        /// <returns><see langword="true"/> only when the target, descriptor, and complete write-only metadata agree.</returns>
        internal static bool TryDescribeReservedFontWriteOnly(object target,
            PropertyDescriptor descriptor, VbePropertyInfo info)
        {
            if (!IsReservedFontDescriptor(descriptor, info)) return false;
            object metadata;
            try { metadata = VbeComPropertyAccessors.Inspect(target, descriptor.Name); }
            catch { return false; } // Unknown metadata never admits the exemption.
            return TryDescribeReservedFontWriteOnly(descriptor, metadata, info);
        }

        /// <summary>Updates the catalogue entry when supplied metadata proves the reserved alias has setters only.</summary>
        /// <param name="descriptor">Reflected property descriptor for the candidate member.</param>
        /// <param name="metadata">Serialized COM property metadata; incomplete, malformed, or mismatched metadata is rejected.</param>
        /// <param name="info">Catalogue entry that receives write-only status and has no readable value or digest.</param>
        /// <returns><see langword="true"/> when the exact reserved property and setter-only accessors are verified.</returns>
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

        /// <summary>Checks that both reflection and catalogue identify the reserved Font alias.</summary>
        /// <param name="descriptor">Reflected member, which must be named <c>_Font_Reserved</c> and have type <see cref="Font"/>.</param>
        /// <param name="info">Catalogue entry whose name must independently match the reserved alias.</param>
        /// <returns><see langword="true"/> only when both references are non-null and identify that exact member.</returns>
        private static bool IsReservedFontDescriptor(PropertyDescriptor descriptor, VbePropertyInfo info)
        {
            return descriptor != null && info != null &&
                descriptor.Name == "_Font_Reserved" && info.Name == "_Font_Reserved" &&
                descriptor.PropertyType == typeof(Font);
        }

        /// <summary>Validates complete COM type information for a setter-only <c>_Font_Reserved</c> property.</summary>
        /// <param name="metadata">Metadata containing the exact property name, IProvideClassInfo discovery, no errors, interfaces, and property-put accessors.</param>
        /// <returns><see langword="true"/> only when every required field and accessor is present and consistent; malformed data returns <see langword="false"/>.</returns>
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

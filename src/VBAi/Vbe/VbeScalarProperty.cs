using System;
using System.ComponentModel;
using System.Globalization;
using System.Reflection;
using System.Runtime.ExceptionServices;
using System.Runtime.InteropServices;

namespace VBAi
{
    /// <summary>Writes already validated scalar values without the Framework COM descriptor's undersized VARIANT buffer.</summary>
    internal static class VbeScalarProperty
    {
        /// <summary>Identifies native COM targets; replaceable by isolated dispatch contract tests.</summary>
        internal static Func<object, bool> NativeObject = Marshal.IsComObject;

        /// <summary>Preserves managed descriptors while routing native properties through the CLR COM binder.</summary>
        /// <param name="target">Validated property owner.</param>
        /// <param name="descriptor">Descriptor used by the caller for validation and readback.</param>
        /// <param name="value">Scalar value already converted by the caller.</param>
        internal static void Set(object target, PropertyDescriptor descriptor, object value)
        {
            if (NativeObject(target)) SetNative(target, descriptor.Name, value);
            else descriptor.SetValue(target, value);
        }

        /// <summary>Invokes a native scalar setter using CLR-owned, architecture-correct COM argument marshaling.</summary>
        /// <param name="target">Native property owner.</param>
        /// <param name="name">Exact validated descriptor name.</param>
        /// <param name="value">Scalar value already converted by the caller.</param>
        internal static void SetNative(object target, string name, object value)
        {
            try
            {
                target.GetType().InvokeMember(name, BindingFlags.SetProperty | BindingFlags.Public | BindingFlags.Instance,
                    null, target, new[] { value }, CultureInfo.InvariantCulture);
            }
            catch (TargetInvocationException error) when (error.InnerException != null)
            {
                ExceptionDispatchInfo.Capture(error.InnerException).Throw();
                throw;
            }
        }
    }
}

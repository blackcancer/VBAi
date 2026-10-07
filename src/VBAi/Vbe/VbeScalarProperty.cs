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

        /// <summary>Identifies the attempted write or its subsequent verification without implying rollback.</summary>
        internal enum FailurePhase
        {

            /// <summary>Identifies the setter invocation case of failure phase.</summary>
            SetterInvocation,

            /// <summary>Identifies the retention readback case of failure phase.</summary>
            RetentionReadback
        }

        /// <summary>Private exception-data identity prevents unrelated annotations from being reported as scalar phases.</summary>
        private static readonly object FailurePhaseKey = new object();

        /// <summary>Annotates the original exception; diagnostic failures must never replace it.</summary>
        /// <param name="error">Original exception to annotate; diagnostic annotation failures are swallowed.</param>
        /// <param name="phase">Setter-invocation or retention-readback phase; other phases are ignored.</param>
        internal static void AnnotateFailure(Exception error, FailurePhase phase)
        {
            try
            {
                if (phase == FailurePhase.SetterInvocation || phase == FailurePhase.RetentionReadback)
                    error.Data[FailurePhaseKey] = phase;
            }
            catch (Exception) { }
        }

        /// <summary>Adds bounded phase information at an error-response boundary, leaving direct callers' exception unchanged.</summary>
        /// <param name="error">Exception describing the error failure.</param>
        /// <returns>Text produced by the operation for format failure on vbe scalar property.</returns>
        internal static string FormatFailure(Exception error)
        {
            try
            {
                Exception original = error;
                FailurePhase phase;
                for (int depth = 0; ; depth++)
                {
                    object annotation = original.Data[FailurePhaseKey];
                    if (annotation is FailurePhase observed &&
                        (observed == FailurePhase.SetterInvocation || observed == FailurePhase.RetentionReadback))
                    {
                        phase = observed;
                        break;
                    }
                    if (depth == 8 || !(original is TargetInvocationException) || original.InnerException == null)
                        return error.Message;
                    original = original.InnerException;
                }
                string type = original.GetType().FullName ?? original.GetType().Name;
                if (type.Length > 128) type = type.Substring(0, 128);
                return original.Message + " [Scalar property phase: " + phase + "; exception: " + type +
                    "; HRESULT: 0x" + unchecked((uint)original.HResult).ToString("X8", CultureInfo.InvariantCulture) +
                    ". The value may already have changed; do not retry automatically.]";
            }
            catch (Exception) { return error.Message; }
        }

        /// <summary>Identifies native COM targets; replaceable by isolated dispatch contract tests.</summary>
        internal static Func<object, bool> NativeObject = Marshal.IsComObject;

        /// <summary>Isolates reflection dispatch so tests can observe the one attempted scalar PROPERTYPUT.</summary>
        internal static Action<object, string, BindingFlags, object> NativeSetterInvocation = (target, name, flags, value) =>
            target.GetType().InvokeMember(name, flags, null, target, new[] { value }, CultureInfo.InvariantCulture);

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
                // SetProperty requests both PUT and PUTREF for COM. These validated scalars are value setters;
                // request PROPERTYPUT alone without trying another mutation if its result is uncertain.
                NativeSetterInvocation(target, name, BindingFlags.PutDispProperty | BindingFlags.Public | BindingFlags.Instance, value);
            }
            catch (TargetInvocationException error) when (error.InnerException != null)
            {
                ExceptionDispatchInfo.Capture(error.InnerException).Throw();
                throw;
            }
        }
    }
}

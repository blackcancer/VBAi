using System;
using System.Globalization;
using System.IO;
using System.Runtime.InteropServices;
using System.Threading.Tasks;

namespace VBAi
{

    /// <summary>Runs the asynchronous native General dialog route for VBE project metadata.</summary>
    internal sealed partial class VbeProjectComponents
    {

        /// <summary>Creates the owner-thread native General dialog adapter for a VBE root HWND.</summary>
        internal Func<IntPtr, Action, VbeProjectGeneralOperation.INative> GeneralNativeFactory = (root, context) => new VbeProjectGeneralNative(root, context);

        /// <summary>Creates the asynchronous General operation state machine around its native adapter.</summary>
        internal Func<VbeProjectGeneralOperation.INative, VbeProjectGeneralOperation> GeneralOperationFactory = native => new VbeProjectGeneralOperation(native);

        /// <summary>Canonical COM identity comparison used for project and active-selection revalidation.</summary>
        internal Func<object, object, bool> GeneralProjectIdentity = SameGeneralProject;
        // Additive route only. Existing set_project_property is never called by this workflow.
        /// <summary>Inspects or writes HelpContextID/HelpFile through one native General dialog with repeated identity and authorization checks.</summary>
        /// <param name="source">Request containing exact project revision, mode, caption, property/value, and async authorization callback.</param>
        /// <param name="write">True for the supported HelpContextID or existing canonical CHM HelpFile write; false for read-only inspection.</param>
        /// <param name="captureExactCommand">Captures the exact native command and returns its one-time open callback.</param>
        /// <param name="durableClaim">Persists command/field/commit attempt milestones before their native side effects.</param>
        /// <param name="requireNativeContext">Checks that the current host project context remains valid on the VBE owner thread.</param>
        /// <returns>Terminal report with readback metadata only when safe to publish; uncertain outcomes suppress metadata and are never replayed.</returns>
        internal async Task<object> ProjectGeneralAsync(Request source, bool write,
            Func<Request, Action, Action<Action>> captureExactCommand,
            Action<VbeProjectGeneralOperation.Result> durableClaim, Action requireNativeContext)
        {
            if (source == null || string.IsNullOrWhiteSpace(source.Project) || string.IsNullOrWhiteSpace(source.ExpectedProjectVersion) ||
                source.ExpectedMode != 2 || string.IsNullOrWhiteSpace(source.ControlCaption) || captureExactCommand == null)
                throw new ArgumentException("Project, ExpectedMode=2, ExpectedProjectVersion and exact ControlCaption are required.");
            if (write && ((source.Property != "HelpContextID" && source.Property != "HelpFile") || source.Value == null ||
                string.IsNullOrWhiteSpace(source.ExpectedOptionsVersion)))
                throw new ArgumentException("Only HelpContextID or rooted CHM HelpFile with non-null Value and ExpectedOptionsVersion is supported.");
            // Capture request values/delegate now; do not consult a mutable caller request during the modal.
            var request = new Request { Project = source.Project, ExpectedProjectVersion = source.ExpectedProjectVersion, ExpectedMode = 2,
                ControlCaption = source.ControlCaption, ExpectedOptionsVersion = source.ExpectedOptionsVersion };
            int? value = write && source.Property == "HelpContextID" ? (int?)RequireGeneralInt32(source.Value) : null;
            string helpFile = write && source.Property == "HelpFile" ? RequireGeneralHelpFile(source.Value) : null;
            Action<bool> authorization = source.RevalidateProjectPropertyAuthorization;
            if (authorization == null) throw new InvalidOperationException("The async General route requires an original request authorization callback for reads and writes.");
            object original = (object)GetDesignProject(request.Project);
            string name = (string)((dynamic)original).Name;
            IntPtr root = new IntPtr(Convert.ToInt64(vbe.MainWindow.HWnd));
            var native = GeneralNativeFactory(root, requireNativeContext);
            Action live = () => {
                native.RequireOwner(); authorization(true);
                if (!GeneralProjectIdentity(original, (object)GetDesignProject(request.Project)) || (int)((dynamic)original).Protection != 0 ||
                    !GeneralProjectIdentity(original, (object)vbe.ActiveVBProject) || (string)((dynamic)original).Name != name ||
                    new IntPtr(Convert.ToInt64(vbe.MainWindow.HWnd)) != root)
                    throw new InvalidOperationException("General original canonical project/active selection/name/protection/VBE root changed.");
                AssertProjectVersion(request, original);
                // Version resolution performs host reads; compare the selector/active identity again afterwards.
                if (!GeneralProjectIdentity(original, (object)GetDesignProject(request.Project)) || !GeneralProjectIdentity(original, (object)vbe.ActiveVBProject) ||
                    (int)((dynamic)original).Protection != 0 || (string)((dynamic)original).Name != name || new IntPtr(Convert.ToInt64(vbe.MainWindow.HWnd)) != root)
                    throw new InvalidOperationException("General final version resolved a different project.");
                AssertProjectVersion(request, original);
                if (!GeneralProjectIdentity(original, (object)GetDesignProject(request.Project)) || !GeneralProjectIdentity(original, (object)vbe.ActiveVBProject))
                    throw new InvalidOperationException("General last revision read resolved a different project.");
                authorization(false); native.RequireOwner();
            };
            Action pure = () => { authorization(false); native.RequireOwner(); };
            live(); Action<Action> open = captureExactCommand(request, live);
            var operation = GeneralOperationFactory(native);
            // Caller retains the async authorization delegate until await settles. No COM RCW is passed to a worker.
            var result = await operation.RunAsync(name, value, request.ExpectedOptionsVersion, live, pure, open, durableClaim, helpFile);
            // An uncertain original modal must not trigger another host read/recovery action.
            if (result.Uncertain) { ClearGeneralMetadata(result); RecordGeneralPublication(result, durableClaim); return result; }
            if (result.Error != null && !result.RefusedBeforeWrite) return result;
            try
            {
                if (!result.Terminal || !result.OriginalExecuteReturned || !result.DialogClosed)
                    throw new InvalidOperationException("Original General Execute and modal closure are not settled.");
                if (result.RefusedBeforeWrite && (result.FieldAttempts != 0 || result.OkAttempts != 0 || result.MutationInvoked ||
                    result.CommittedRequested || result.ControlValueVerified || result.CancelAttempts != 1))
                    throw new InvalidOperationException("The original no-write representation refusal and single Cancel are not proved.");
                if (!write || result.RefusedBeforeWrite) live();
                else
                {
                    native.RequireOwner(); authorization(true);
                    if (!GeneralProjectIdentity(original, (object)GetDesignProject(request.Project)) || !GeneralProjectIdentity(original, (object)vbe.ActiveVBProject) ||
                        (int)((dynamic)original).Protection != 0 || (string)((dynamic)original).Name != name || new IntPtr(Convert.ToInt64(vbe.MainWindow.HWnd)) != root)
                        throw new InvalidOperationException("General final publication no longer owns the approved original project.");
                    // The one successful native OK changed metadata. Do not compare its obsolete pre-write revision.
                    if (!GeneralProjectIdentity(original, (object)GetDesignProject(request.Project)) || !GeneralProjectIdentity(original, (object)vbe.ActiveVBProject))
                        throw new InvalidOperationException("General publication target changed during final host reads.");
                    authorization(false); native.RequireOwner();
                }
            }
            catch (Exception error)
            {
                result.Error = error.ToString(); result.Available = false;
                result.Uncertain |= result.RefusedBeforeWrite || result.MutationInvoked || !result.OriginalExecuteReturned || !result.DialogClosed;
                ClearGeneralMetadata(result);
            }
            if (result.RefusedBeforeWrite) ClearGeneralMetadata(result);
            RecordGeneralPublication(result, durableClaim);
            return result;
        }

        /// <summary>Attempts to publish the terminal operation receipt and marks the result unavailable if persistence fails.</summary>
        /// <param name="result">Completed native General operation state being published.</param>
        /// <param name="durableClaim">Receipt callback that stores the final attempt counts and outcome.</param>
        private static void RecordGeneralPublication(VbeProjectGeneralOperation.Result result, Action<VbeProjectGeneralOperation.Result> durableClaim)
        {
            try { durableClaim(result); }
            catch (Exception error)
            {
                result.Error = (result.Error == null ? "" : result.Error + " | ") + "Final publication receipt: " + error;
                result.Uncertain |= result.CommandEntered; result.Available = false; ClearGeneralMetadata(result);
            }
        }

        /// <summary>Compares two project objects by canonical IUnknown identity.</summary>
        /// <param name="first">Previously captured project COM object.</param>
        /// <param name="second">Current project COM object to compare.</param>
        /// <returns>False for null inputs; otherwise true only when both IUnknown pointers match.</returns>
        private static bool SameGeneralProject(object first, object second)
        {
            if (first == null || second == null) return false;
            IntPtr a = Marshal.GetIUnknownForObject(first), b = IntPtr.Zero;
            try { b = Marshal.GetIUnknownForObject(second); return a == b; }
            finally { if (b != IntPtr.Zero) Marshal.Release(b); Marshal.Release(a); }
        }

        /// <summary>Removes all native General values and versions from a result that cannot be safely published.</summary>
        /// <param name="result">Operation result to mark unavailable and clear of sensitive/stale metadata.</param>
        private static void ClearGeneralMetadata(VbeProjectGeneralOperation.Result result)
        { result.Available = false; result.Name = result.Description = result.HelpFile = result.HelpContextText = result.ConditionalCompilation = result.OptionsVersion = null; }

        /// <summary>Converts supported numeric values without truncation or culture-dependent parsing.</summary>
        /// <param name="value">Nonnegative integral Int32 value or invariant integer string; booleans, enums, chars, fractions, and overflow are rejected.</param>
        /// <returns>Exact HelpContextID value in the range 0 through Int32.MaxValue.</returns>
        internal static int RequireGeneralInt32(object value)
        {
            if (value == null || value is bool || value is char || value.GetType().IsEnum)
                throw new ArgumentException("General HelpContextID requires an exact Int32 value.");
            if (value is string text)
            {
                if (int.TryParse(text, NumberStyles.Integer, CultureInfo.InvariantCulture, out int parsed) && parsed >= 0) return parsed;
                throw new ArgumentException("General HelpContextID requires an invariant Int32 integer string.");
            }
            decimal number;
            try
            {
                switch (Type.GetTypeCode(value.GetType()))
                {
                    case TypeCode.SByte: case TypeCode.Byte: case TypeCode.Int16: case TypeCode.UInt16:
                    case TypeCode.Int32: case TypeCode.UInt32: case TypeCode.Int64: case TypeCode.UInt64:
                    case TypeCode.Single: case TypeCode.Double: case TypeCode.Decimal:
                        number = Convert.ToDecimal(value, CultureInfo.InvariantCulture); break;
                    default: throw new ArgumentException("General HelpContextID requires an exact numeric Int32 value.");
                }
            }
            catch (OverflowException error) { throw new ArgumentException("General HelpContextID is outside the Int32 range.", error); }
            // Check the source floating number before Decimal conversion can round an almost-integral value.
            if (value is double floating && (double.IsNaN(floating) || double.IsInfinity(floating) || Math.Truncate(floating) != floating))
                throw new ArgumentException("General HelpContextID must not be fractional or non-finite.");
            if (value is float single && (float.IsNaN(single) || float.IsInfinity(single) || Math.Truncate(single) != single))
                throw new ArgumentException("General HelpContextID must not be fractional or non-finite.");
            if (number < 0 || number > int.MaxValue || decimal.Truncate(number) != number)
                throw new ArgumentException("General HelpContextID must be an exact in-range Int32.");
            return (int)number;
        }

        /// <summary>Requires an existing canonical absolute CHM path and returns it without text conversion.</summary>
        /// <param name="value">String path with rooted spelling, no NUL, and a .chm extension.</param>
        /// <returns>The unchanged path when it is already canonical and names an existing file.</returns>
        internal static string RequireGeneralHelpFile(object value)
        {
            if (!(value is string path) || string.IsNullOrWhiteSpace(path) || path.IndexOf('\0') >= 0 || !Path.IsPathRooted(path) ||
                !string.Equals(Path.GetExtension(path), ".chm", StringComparison.OrdinalIgnoreCase))
                throw new ArgumentException("General HelpFile requires an explicit rooted CHM file path string.");
            string canonical = Path.GetFullPath(path);
            if (!string.Equals(canonical, path, StringComparison.Ordinal) || !File.Exists(path))
                throw new ArgumentException("General HelpFile requires an existing canonical rooted CHM path, without text conversion.");
            return path;
        }
    }
}

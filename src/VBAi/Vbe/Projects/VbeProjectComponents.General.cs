using System;
using System.Globalization;
using System.IO;
using System.Runtime.InteropServices;
using System.Threading.Tasks;

namespace VBAi
{

    /// <summary>Owns the vbe project components state and operations.</summary>
    internal sealed partial class VbeProjectComponents
    {

        /// <summary>Maintains the general native factory state for vbe project components.</summary>
        internal Func<IntPtr, Action, VbeProjectGeneralOperation.INative> GeneralNativeFactory = (root, context) => new VbeProjectGeneralNative(root, context);

        /// <summary>Maintains the general operation factory state for vbe project components.</summary>
        internal Func<VbeProjectGeneralOperation.INative, VbeProjectGeneralOperation> GeneralOperationFactory = native => new VbeProjectGeneralOperation(native);

        /// <summary>Maintains the general project identity state for vbe project components.</summary>
        internal Func<object, object, bool> GeneralProjectIdentity = SameGeneralProject;
        // Additive route only. Existing set_project_property is never called by this workflow.
        /// <summary>Handles project general async for vbe project components.</summary>
        /// <param name="source">request that supplies the source for this operation.</param>
        /// <param name="write">Indicates whether write is enabled.</param>
        /// <param name="captureExactCommand">func&lt;request, action, action&lt;action&gt;&gt; that supplies the capture exact command for this operation.</param>
        /// <param name="durableClaim">result&gt; that supplies the durable claim for this operation.</param>
        /// <param name="requireNativeContext">action that supplies the require native context for this operation.</param>
        /// <returns>task&lt;object&gt; produced by the operation for project general async on vbe project components.</returns>
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

        /// <summary>Handles record general publication for vbe project components.</summary>
        /// <param name="result">result that supplies the result for this operation.</param>
        /// <param name="durableClaim">result&gt; that supplies the durable claim for this operation.</param>
        private static void RecordGeneralPublication(VbeProjectGeneralOperation.Result result, Action<VbeProjectGeneralOperation.Result> durableClaim)
        {
            try { durableClaim(result); }
            catch (Exception error)
            {
                result.Error = (result.Error == null ? "" : result.Error + " | ") + "Final publication receipt: " + error;
                result.Uncertain |= result.CommandEntered; result.Available = false; ClearGeneralMetadata(result);
            }
        }

        /// <summary>Compares general project for vbe project components.</summary>
        /// <param name="first">object that supplies the first for this operation.</param>
        /// <param name="second">object that supplies the second for this operation.</param>
        /// <returns>Boolean indicating the result of the check for same general project on vbe project components.</returns>
        private static bool SameGeneralProject(object first, object second)
        {
            if (first == null || second == null) return false;
            IntPtr a = Marshal.GetIUnknownForObject(first), b = IntPtr.Zero;
            try { b = Marshal.GetIUnknownForObject(second); return a == b; }
            finally { if (b != IntPtr.Zero) Marshal.Release(b); Marshal.Release(a); }
        }

        /// <summary>Clears general metadata for vbe project components.</summary>
        /// <param name="result">result that supplies the result for this operation.</param>
        private static void ClearGeneralMetadata(VbeProjectGeneralOperation.Result result)
        { result.Available = false; result.Name = result.Description = result.HelpFile = result.HelpContextText = result.ConditionalCompilation = result.OptionsVersion = null; }

        /// <summary>Requires general int32 for vbe project components.</summary>
        /// <param name="value">object that supplies the value for this operation.</param>
        /// <returns>int produced by the operation for require general int32 on vbe project components.</returns>
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

        /// <summary>Requires general help file for vbe project components.</summary>
        /// <param name="value">object that supplies the value for this operation.</param>
        /// <returns>Text produced by the operation for require general help file on vbe project components.</returns>
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

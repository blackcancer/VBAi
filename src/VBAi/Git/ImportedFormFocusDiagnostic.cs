using System;
using System.Globalization;
using System.Text;
using System.Threading;

namespace VBAi
{

    /// <summary>Bounded failure-only observations; never authorizes rendering or changes native navigation.</summary>
    internal static class ImportedFormFocusDiagnostic
    {

        /// <summary>Maintains the field limit state for imported form focus diagnostic.</summary>
        internal const int FieldLimit = 128;

        /// <summary>Owns the snapshot state and operations.</summary>
        internal sealed class Snapshot
        {

            /// <summary>Maintains the state and active identity and active type and active handle and active caption and expected caption and root enabled and active matches previous state for snapshot.</summary>
            internal string State, ActiveIdentity, ActiveType, ActiveHandle, ActiveCaption, ExpectedCaption,
                RootEnabled, ActiveMatchesPrevious;

            /// <summary>Maintains the previous present state for snapshot.</summary>
            internal bool PreviousPresent;
        }

        /// <summary>Reads each additional operand once on the owning STA; an unavailable operand does not hide its peers.</summary>
        /// <param name="activePresent">Indicates whether active present is enabled.</param>
        /// <param name="activeMatchesExpected">Indicates whether active matches expected is enabled.</param>
        /// <param name="previousPresent">Indicates whether previous present is enabled.</param>
        /// <param name="rootPresent">Indicates whether root present is enabled.</param>
        /// <param name="expectedType">int that supplies the expected type for this operation.</param>
        /// <param name="expectedHandle">long that supplies the expected handle for this operation.</param>
        /// <param name="activeType">func&lt;int&gt; that supplies the active type for this operation.</param>
        /// <param name="activeHandle">func&lt;long&gt; that supplies the active handle for this operation.</param>
        /// <param name="activeCaption">func&lt;string&gt; that supplies the active caption for this operation.</param>
        /// <param name="expectedCaption">func&lt;string&gt; that supplies the expected caption for this operation.</param>
        /// <param name="rootEnabled">func&lt;bool&gt; that supplies the root enabled for this operation.</param>
        /// <param name="activeMatchesPrevious">func&lt;bool&gt; that supplies the active matches previous for this operation.</param>
        /// <returns>snapshot produced by the operation for observe on imported form focus diagnostic.</returns>
        internal static Snapshot Observe(bool activePresent, bool activeMatchesExpected, bool previousPresent,
            bool rootPresent, int expectedType, long expectedHandle, Func<int> activeType, Func<long> activeHandle,
            Func<string> activeCaption, Func<string> expectedCaption, Func<bool> rootEnabled, Func<bool> activeMatchesPrevious)
        {
            var value = new Snapshot {
                State = "observed-after-refusal", ActiveIdentity = !activePresent ? "null" : activeMatchesExpected ? "expected" : "different",
                PreviousPresent = previousPresent
            };
            if (Thread.CurrentThread.GetApartmentState() != ApartmentState.STA)
            {
                value.State = "not-observed-non-STA";
                return value;
            }
            value.ExpectedCaption = Read(expectedCaption);
            if (!activePresent)
                value.ActiveType = value.ActiveHandle = value.ActiveCaption = "not-observed-null-active";
            else if (activeMatchesExpected)
            {
                // These exact operands were already read by the guard from the same COM identity.
                value.ActiveType = expectedType.ToString(CultureInfo.InvariantCulture);
                value.ActiveHandle = expectedHandle.ToString(CultureInfo.InvariantCulture);
                value.ActiveCaption = value.ExpectedCaption;
            }
            else
            {
                value.ActiveType = Read(activeType);
                value.ActiveHandle = Read(activeHandle);
                value.ActiveCaption = Read(activeCaption);
            }
            value.RootEnabled = rootPresent ? Read(rootEnabled) : "not-observed-zero-root";
            value.ActiveMatchesPrevious = !activePresent ? "not-observed-null-active" :
                previousPresent ? Read(activeMatchesPrevious) : "not-observed-no-previous";
            return value;
        }

        /// <summary>Reads  for imported form focus diagnostic.</summary>
        /// <typeparam name="T">The type used for t.</typeparam>
        /// <param name="read">func&lt;t&gt; that supplies the read for this operation.</param>
        /// <returns>Text produced by the operation for read on imported form focus diagnostic.</returns>
        private static string Read<T>(Func<T> read)
        {
            if (read == null) return "unavailable-missing-getter";
            try { T result = read(); return Bounded((object)result == null ? null : Convert.ToString(result, CultureInfo.InvariantCulture)); }
            catch (Exception error)
            {
                // No exception message: it can contain paths, control characters or unbounded native data.
                return "unavailable(" + Bounded(error.GetType().Name) + ":0x" +
                    error.HResult.ToString("X8", CultureInfo.InvariantCulture) + ")";
            }
        }

        /// <summary>Formats already captured scalars only, independently of native objects and getters.</summary>
        /// <param name="value">snapshot that supplies the value for this operation.</param>
        /// <returns>Text produced by the operation for format on imported form focus diagnostic.</returns>
        internal static string Format(Snapshot value)
        {
            if (value == null) throw new ArgumentNullException(nameof(value));
            return " FocusObservation=" + Bounded(value.State) + ";ActiveIdentity=" + Bounded(value.ActiveIdentity) +
                ";ActiveType=" + Bounded(value.ActiveType) + ";ActiveHandle=" + Bounded(value.ActiveHandle) +
                ";ActiveCaption=" + Bounded(value.ActiveCaption) + ";ExpectedCaption=" + Bounded(value.ExpectedCaption) +
                ";RootEnabled=" + Bounded(value.RootEnabled) + ";PreviousPresent=" + value.PreviousPresent +
                ";ActiveMatchesPrevious=" + Bounded(value.ActiveMatchesPrevious) + ".";
        }

        /// <summary>One line with bounded fields; captions cannot inject apparent diagnostic operands.</summary>
        /// <param name="text">Text that supplies the text value. Use the format required by the calling operation.</param>
        /// <returns>Text produced by the operation for bounded on imported form focus diagnostic.</returns>
        internal static string Bounded(string text)
        {
            if (text == null) return "<null>";
            int length = Math.Min(text.Length, FieldLimit);
            if (length > 0 && char.IsHighSurrogate(text[length - 1])) length--;
            var result = new StringBuilder(length);
            for (int i = 0; i < length; i++)
            {
                char c = text[i];
                result.Append(char.IsControl(c) || c == ';' || c == '=' || c == '|' || c == '\u2028' || c == '\u2029' ? ' ' : c);
            }
            return result.ToString();
        }
    }
}

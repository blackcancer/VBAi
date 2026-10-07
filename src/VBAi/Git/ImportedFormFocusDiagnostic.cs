using System;
using System.Globalization;
using System.Text;
using System.Threading;

namespace VBAi
{

    /// <summary>Bounded failure-only observations; never authorizes rendering or changes native navigation.</summary>
    internal static class ImportedFormFocusDiagnostic
    {

        /// <summary>Maximum UTF-16 characters emitted for any single diagnostic field.</summary>
        internal const int FieldLimit = 128;

        /// <summary>Stores bounded scalar observations captured after an imported-form focus refusal.</summary>
        internal sealed class Snapshot
        {

            /// <summary>State label, active identity/type/HWND/caption, expected caption, root enabled flag, and previous-window match.</summary>
            internal string State, ActiveIdentity, ActiveType, ActiveHandle, ActiveCaption, ExpectedCaption,
                RootEnabled, ActiveMatchesPrevious;

            /// <summary>Whether a previous active window existed when the observation was taken.</summary>
            internal bool PreviousPresent;
        }

        /// <summary>Reads each additional operand once on the owning STA; an unavailable operand does not hide its peers.</summary>
        /// <param name="activePresent">Whether the active VBE window object was available.</param>
        /// <param name="activeMatchesExpected">Whether it is the exact imported designer window already inspected by the guard.</param>
        /// <param name="previousPresent">Whether a prior active window was captured.</param>
        /// <param name="rootPresent">Whether the VBE root window handle is available.</param>
        /// <param name="expectedType">Known type of the expected designer window.</param>
        /// <param name="expectedHandle">Known HWND of the expected designer window.</param>
        /// <param name="activeType">Optional getter for a different active window's type.</param>
        /// <param name="activeHandle">Optional getter for a different active window's HWND.</param>
        /// <param name="activeCaption">Optional getter for a different active window's caption.</param>
        /// <param name="expectedCaption">Getter for the expected caption.</param>
        /// <param name="rootEnabled">Getter for whether the VBE root window is enabled.</param>
        /// <param name="activeMatchesPrevious">Getter comparing active and previously captured windows.</param>
        /// <returns>Scalar-only observation; failed getters become bounded unavailable markers.</returns>
        internal static Snapshot Observe(bool activePresent, bool activeMatchesExpected, bool previousPresent,
            bool rootPresent, int expectedType, long expectedHandle, Func<int> activeType, Func<long> activeHandle,
            Func<string> activeCaption, Func<string> expectedCaption, Func<bool> rootEnabled, Func<bool> activeMatchesPrevious)
        {
            var value = new Snapshot
            {
                State = "observed-after-refusal",
                ActiveIdentity = !activePresent ? "null" : activeMatchesExpected ? "expected" : "different",
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

        /// <summary>Reads one optional diagnostic getter without propagating its native failure or exception message.</summary>
        /// <typeparam name="T">Getter result type.</typeparam>
        /// <param name="read">One getter to invoke.</param>
        /// <returns>Bounded invariant-culture value or a bounded unavailable marker with exception type/HRESULT.</returns>
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
        /// <param name="value">Previously captured scalar snapshot.</param>
        /// <returns>One-line diagnostic; it performs no COM access or native getter calls.</returns>
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
        /// <param name="text">Untrusted caption, exception type, or other diagnostic field.</param>
        /// <returns>At most 128 characters with control and delimiter characters replaced by spaces.</returns>
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

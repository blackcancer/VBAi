using System;
using System.Collections.Generic;
using System.Runtime.InteropServices;
using System.Security.Principal;

namespace VBAi
{

    /// <summary>Opt-in qualification evidence: public metadata of the calling thread's effective token only.</summary>
    internal static class EffectiveTokenObservation
    {

        /// <summary>Win32 ERROR_NO_TOKEN code that permits a fallback to the process primary token.</summary>
        internal const int ErrorNoToken = 1008;

        /// <summary>Defines read-only token acquisition and metadata operations; implementations never impersonate or alter a token.</summary>
        internal interface ITokenReader
        {

            /// <summary>Attempts to open the calling thread's effective token for query-only access.</summary>
            /// <param name="token">Receives the token handle on success.</param>
            /// <param name="error">Receives the last Win32 error from the open attempt.</param>
            /// <returns><see langword="true"/> when the thread token was opened.</returns>
            bool OpenThread(out IntPtr token, out int error);

            /// <summary>Attempts to open the current process primary token for query-only access.</summary>
            /// <param name="token">Receives the token handle on success.</param>
            /// <param name="error">Receives the last Win32 error from the open attempt.</param>
            /// <returns><see langword="true"/> when the primary token was opened.</returns>
            bool OpenPrimary(out IntPtr token, out int error);

            /// <summary>Reads public identity, restriction, container, type, and authentication metadata from an already-open token.</summary>
            /// <param name="token">Query-only token handle that the caller must close.</param>
            /// <returns>Named metadata values; individual query failures may be reported in <c>MetadataErrors</c>.</returns>
            IDictionary<string, object> Metadata(IntPtr token);

            /// <summary>Closes a token handle obtained by one of the open methods.</summary>
            /// <param name="token">Token handle to release.</param>
            void Close(IntPtr token);
        }

        /// <summary>Falls back to the primary token only for ERROR_NO_TOKEN; never impersonates or changes a token.</summary>
        /// <param name="reader">Optional token-reader implementation; null uses the query-only Win32 reader.</param>
        /// <returns>Observation dictionary with READ, PARTIAL, or UNVERIFIED state and only successfully read metadata.</returns>
        internal static IDictionary<string, object> Read(ITokenReader reader = null)
        {
            reader = reader ?? new NativeTokenReader();
            var result = new Dictionary<string, object> { ["State"] = "UNVERIFIED", ["OpenAsSelf"] = true };
            IntPtr token = IntPtr.Zero;
            try
            {
                if (reader.OpenThread(out token, out int error))
                {
                    result["Source"] = "Thread";
                }
                else
                {
                    result["OpenThreadTokenError"] = error;
                    if (error != ErrorNoToken) return result;
                    if (!reader.OpenPrimary(out token, out error))
                    { result["OpenProcessTokenError"] = error; return result; }
                    result["Source"] = "PrimaryAfterErrorNoToken";
                }
                foreach (var item in reader.Metadata(token)) result[item.Key] = item.Value;
                result["State"] = result.ContainsKey("MetadataErrors") ? "PARTIAL" : "READ";
            }
            catch (Exception error) { result["ErrorType"] = error.GetType().FullName; result["HResult"] = error.HResult; }
            finally
            {
                if (token != IntPtr.Zero)
                {
                    try { reader.Close(token); }
                    catch (Exception error) { result["CloseErrorType"] = error.GetType().FullName; result["State"] = "PARTIAL"; }
                }
            }
            return result;
        }

        /// <summary>Implements query-only access to the current thread token and current process primary token.</summary>
        private sealed class NativeTokenReader : ITokenReader
        {

            /// <summary>Opens the current thread token with TOKEN_QUERY and OPENAS_SELF enabled.</summary>
            /// <param name="token">Receives the opened token handle.</param>
            /// <param name="error">Receives the last Win32 error, including ERROR_NO_TOKEN when not impersonating.</param>
            /// <returns><see langword="true"/> when OpenThreadToken succeeds.</returns>
            public bool OpenThread(out IntPtr token, out int error)
            {
                bool ok = OpenThreadToken(GetCurrentThread(), 0x8 /* TOKEN_QUERY */, true, out token);
                error = Marshal.GetLastWin32Error(); return ok;
            }

            /// <summary>Opens the current process primary token with TOKEN_QUERY access.</summary>
            /// <param name="token">Receives the opened token handle.</param>
            /// <param name="error">Receives the last Win32 error.</param>
            /// <returns><see langword="true"/> when OpenProcessToken succeeds.</returns>
            public bool OpenPrimary(out IntPtr token, out int error)
            {
                bool ok = OpenProcessToken(GetCurrentProcess(), 0x8 /* TOKEN_QUERY */, out token);
                error = Marshal.GetLastWin32Error(); return ok;
            }

            /// <summary>Releases a token handle and throws a Win32Exception if CloseHandle fails.</summary>
            /// <param name="token">Token handle to close.</param>
            public void Close(IntPtr token)
            {
                if (!CloseHandle(token)) throw new System.ComponentModel.Win32Exception(Marshal.GetLastWin32Error());
            }

            /// <summary>Reads selected token information classes and records per-class failures without elevating access.</summary>
            /// <param name="token">Token handle opened with TOKEN_QUERY.</param>
            /// <returns>Metadata for restriction, user/integrity SID, app-container state, token type, impersonation level, and IDs.</returns>
            public IDictionary<string, object> Metadata(IntPtr token)
            {
                var result = new Dictionary<string, object>();
                var errors = new Dictionary<string, int>();
                result["Restricted"] = IsTokenRestricted(token);
                result["UserSid"] = Information(token, 1, "UserSid", errors, IntPtr.Size, p => new SecurityIdentifier(Marshal.ReadIntPtr(p)).Value);
                result["IntegritySid"] = Information(token, 25, "IntegritySid", errors, IntPtr.Size, p => new SecurityIdentifier(Marshal.ReadIntPtr(p)).Value);
                result["AppContainer"] = Information(token, 29, "AppContainer", errors, 4, p => (object)(Marshal.ReadInt32(p) != 0));
                result["TokenType"] = Information(token, 8, "TokenType", errors, 4, p => (object)Marshal.ReadInt32(p));
                if (Convert.ToInt32(result["TokenType"] ?? 0) == 2)
                    result["ImpersonationLevel"] = Information(token, 9, "ImpersonationLevel", errors, 4, p => (object)Marshal.ReadInt32(p));
                else result["ImpersonationLevel"] = null;
                var statistics = Information(token, 10, "Statistics", errors, Marshal.SizeOf<TokenStatistics>(), p => (object)Marshal.PtrToStructure<TokenStatistics>(p));
                if (statistics is TokenStatistics stats)
                {
                    result["AuthenticationId"] = stats.AuthenticationId.ToString();
                    result["TokenId"] = stats.TokenId.ToString();
                }
                if (errors.Count > 0) result["MetadataErrors"] = errors;
                return result;
            }

            /// <summary>Queries one token information class with a bounded native buffer and stores any Win32 failure under its metadata key.</summary>
            /// <param name="token">Query-only token handle.</param>
            /// <param name="kind">TOKEN_INFORMATION_CLASS numeric value requested from Win32.</param>
            /// <param name="name">Stable key used in the result and error dictionaries.</param>
            /// <param name="errors">Receives the Win32 error when sizing, retrieval, or returned length is invalid.</param>
            /// <param name="minimum">Minimum structure size accepted for this information class.</param>
            /// <param name="read">Decoder that converts the validated native buffer into a managed value.</param>
            /// <returns>Decoded value, or null when the query fails validation.</returns>
            private static object Information(IntPtr token, int kind, string name, IDictionary<string, int> errors, int minimum, Func<IntPtr, object> read)
            {
                bool sized = GetTokenInformation(token, kind, IntPtr.Zero, 0, out int size);
                int error = Marshal.GetLastWin32Error();
                if (sized || error != 122 || size < minimum || size > 65536) { errors[name] = error; return null; }
                IntPtr buffer = Marshal.AllocHGlobal(size);
                try
                {
                    bool ok = GetTokenInformation(token, kind, buffer, size, out int returned);
                    error = Marshal.GetLastWin32Error();
                    if (!ok || returned < minimum || returned > size) { errors[name] = error; return null; }
                    return read(buffer);
                }
                finally { Marshal.FreeHGlobal(buffer); }
            }
        }

        /// <summary>Native LUID layout used for token and authentication identifiers returned by TOKEN_STATISTICS.</summary>
        [StructLayout(LayoutKind.Sequential)]
        private struct Luid
        {

            /// <summary>Unsigned low-order 32 bits of the LUID.</summary>
            internal uint Low;

            /// <summary>Signed high-order 32 bits of the LUID, reinterpreted as unsigned when formatted.</summary>
            internal int High;

            /// <summary>Formats the LUID as fixed-width hexadecimal high and low words.</summary>
            /// <returns>Eight uppercase hex digits, a colon, then eight uppercase hex digits.</returns>
            public override string ToString() => unchecked((uint)High).ToString("X8") + ":" + Low.ToString("X8");
        }

        /// <summary>Native TOKEN_STATISTICS layout; only token ID and authentication ID are surfaced in the observation.</summary>
        [StructLayout(LayoutKind.Sequential)]
        private struct TokenStatistics
        {

            /// <summary>Unique identifier assigned to this token.</summary>
            internal Luid TokenId, AuthenticationId;

            /// <summary>Native expiration timestamp field required to preserve TOKEN_STATISTICS layout.</summary>
            internal long ExpirationTime;

            /// <summary>Native token type and impersonation-level fields required to preserve TOKEN_STATISTICS layout.</summary>
            internal int TokenType, ImpersonationLevel;

            /// <summary>Native dynamic allocation and collection counts required to preserve TOKEN_STATISTICS layout.</summary>
            internal uint DynamicCharged, DynamicAvailable, GroupCount, PrivilegeCount;

            /// <summary>Identifier changed when the token's assigned privileges are modified.</summary>
            internal Luid ModifiedId;
        }

        /// <summary>Gets a pseudo-handle for the calling thread, used only to query its effective token.</summary>
        /// <returns>Current-thread pseudo-handle; it is not closed.</returns>
        [DllImport("kernel32.dll")] private static extern IntPtr GetCurrentThread();

        /// <summary>Gets a pseudo-handle for the current process, used only to query its primary token.</summary>
        /// <returns>Current-process pseudo-handle; it is not closed.</returns>
        [DllImport("kernel32.dll")] private static extern IntPtr GetCurrentProcess();

        /// <summary>Releases an opened token handle.</summary>
        /// <param name="handle">Token handle returned by OpenThreadToken or OpenProcessToken.</param>
        /// <returns><see langword="true"/> when Windows closes the handle.</returns>
        [DllImport("kernel32.dll", SetLastError = true)][return: MarshalAs(UnmanagedType.Bool)] private static extern bool CloseHandle(IntPtr handle);

        /// <summary>Opens a thread's effective token with the requested access and caller security context.</summary>
        /// <param name="thread">Thread whose effective token is queried; callers pass the current-thread pseudo-handle.</param>
        /// <param name="access">Requested token rights; this observer passes TOKEN_QUERY only.</param>
        /// <param name="openAsSelf">When true, performs the access check using the process security context.</param>
        /// <param name="token">Receives the opened token handle.</param>
        /// <returns><see langword="true"/> on success; otherwise false and GetLastError identifies the failure.</returns>
        [DllImport("advapi32.dll", SetLastError = true)][return: MarshalAs(UnmanagedType.Bool)] private static extern bool OpenThreadToken(IntPtr thread, uint access, [MarshalAs(UnmanagedType.Bool)] bool openAsSelf, out IntPtr token);

        /// <summary>Opens the process primary token with the requested access.</summary>
        /// <param name="process">Process whose primary token is queried; callers pass the current-process pseudo-handle.</param>
        /// <param name="access">Requested token rights; this observer passes TOKEN_QUERY only.</param>
        /// <param name="token">Receives the opened token handle.</param>
        /// <returns><see langword="true"/> on success; otherwise false and GetLastError identifies the failure.</returns>
        [DllImport("advapi32.dll", SetLastError = true)][return: MarshalAs(UnmanagedType.Bool)] private static extern bool OpenProcessToken(IntPtr process, uint access, out IntPtr token);

        /// <summary>Queries a selected TOKEN_INFORMATION_CLASS into a caller-owned buffer.</summary>
        /// <param name="token">Token handle opened with TOKEN_QUERY.</param>
        /// <param name="kind">TOKEN_INFORMATION_CLASS numeric selector.</param>
        /// <param name="information">Output buffer, or null for the required-size query.</param>
        /// <param name="size">Buffer size in bytes.</param>
        /// <param name="returned">Receives required or actual bytes, depending on the query phase.</param>
        /// <returns><see langword="true"/> when the information is returned; otherwise false and GetLastError supplies the reason.</returns>
        [DllImport("advapi32.dll", SetLastError = true)][return: MarshalAs(UnmanagedType.Bool)] private static extern bool GetTokenInformation(IntPtr token, int kind, IntPtr information, int size, out int returned);

        /// <summary>Checks whether Windows marks the token as restricted.</summary>
        /// <param name="token">Token handle to inspect.</param>
        /// <returns><see langword="true"/> when the token is restricted.</returns>
        [DllImport("advapi32.dll")][return: MarshalAs(UnmanagedType.Bool)] private static extern bool IsTokenRestricted(IntPtr token);
    }
}

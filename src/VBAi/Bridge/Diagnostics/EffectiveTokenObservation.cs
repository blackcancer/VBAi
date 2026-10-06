using System;
using System.Collections.Generic;
using System.Runtime.InteropServices;
using System.Security.Principal;

namespace VBAi
{

    /// <summary>Opt-in qualification evidence: public metadata of the calling thread's effective token only.</summary>
    internal static class EffectiveTokenObservation
    {

        /// <summary>Maintains the error no token state for effective token observation.</summary>
        internal const int ErrorNoToken = 1008;

        /// <summary>Defines the i token reader contract.</summary>
        internal interface ITokenReader
        {

            /// <summary>Opens thread for i token reader.</summary>
            /// <param name="token">Native handle that supplies the token for this operation.</param>
            /// <param name="error">int that supplies the error for this operation.</param>
            /// <returns>Boolean indicating the result of the check for open thread on i token reader.</returns>
            bool OpenThread(out IntPtr token, out int error);

            /// <summary>Opens primary for i token reader.</summary>
            /// <param name="token">Native handle that supplies the token for this operation.</param>
            /// <param name="error">int that supplies the error for this operation.</param>
            /// <returns>Boolean indicating the result of the check for open primary on i token reader.</returns>
            bool OpenPrimary(out IntPtr token, out int error);

            /// <summary>Handles metadata for i token reader.</summary>
            /// <param name="token">Native handle that supplies the token for this operation.</param>
            /// <returns>i dictionary&lt;string, object&gt; produced by the operation for metadata on i token reader.</returns>
            IDictionary<string, object> Metadata(IntPtr token);

            /// <summary>Closes  for i token reader.</summary>
            /// <param name="token">Native handle that supplies the token for this operation.</param>
            void Close(IntPtr token);
        }

        /// <summary>Falls back to the primary token only for ERROR_NO_TOKEN; never impersonates or changes a token.</summary>
        /// <param name="reader">i token reader that supplies the reader for this operation.</param>
        /// <returns>i dictionary&lt;string, object&gt; produced by the operation for read on effective token observation.</returns>
        internal static IDictionary<string, object> Read(ITokenReader reader = null)
        {
            reader = reader ?? new NativeTokenReader();
            var result = new Dictionary<string, object> { ["State"] = "UNVERIFIED", ["OpenAsSelf"] = true };
            IntPtr token = IntPtr.Zero;
            try
            {
                int error;
                if (reader.OpenThread(out token, out error)) result["Source"] = "Thread";
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

        /// <summary>Owns the native token reader state and operations.</summary>
        private sealed class NativeTokenReader : ITokenReader
        {

            /// <summary>Opens thread for native token reader.</summary>
            /// <param name="token">Native handle that supplies the token for this operation.</param>
            /// <param name="error">int that supplies the error for this operation.</param>
            /// <returns>Boolean indicating the result of the check for open thread on native token reader.</returns>
            public bool OpenThread(out IntPtr token, out int error)
            {
                bool ok = OpenThreadToken(GetCurrentThread(), 0x8 /* TOKEN_QUERY */, true, out token);
                error = Marshal.GetLastWin32Error(); return ok;
            }

            /// <summary>Opens primary for native token reader.</summary>
            /// <param name="token">Native handle that supplies the token for this operation.</param>
            /// <param name="error">int that supplies the error for this operation.</param>
            /// <returns>Boolean indicating the result of the check for open primary on native token reader.</returns>
            public bool OpenPrimary(out IntPtr token, out int error)
            {
                bool ok = OpenProcessToken(GetCurrentProcess(), 0x8 /* TOKEN_QUERY */, out token);
                error = Marshal.GetLastWin32Error(); return ok;
            }

            /// <summary>Closes  for native token reader.</summary>
            /// <param name="token">Native handle that supplies the token for this operation.</param>
            public void Close(IntPtr token)
            {
                if (!CloseHandle(token)) throw new System.ComponentModel.Win32Exception(Marshal.GetLastWin32Error());
            }

            /// <summary>Handles metadata for native token reader.</summary>
            /// <param name="token">Native handle that supplies the token for this operation.</param>
            /// <returns>i dictionary&lt;string, object&gt; produced by the operation for metadata on native token reader.</returns>
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

            /// <summary>Handles information for native token reader.</summary>
            /// <param name="token">Native handle that supplies the token for this operation.</param>
            /// <param name="kind">int that supplies the kind for this operation.</param>
            /// <param name="name">Text that supplies the name value. Use the format required by the calling operation.</param>
            /// <param name="errors">i dictionary&lt;string, int&gt; that supplies the errors for this operation.</param>
            /// <param name="minimum">int that supplies the minimum for this operation.</param>
            /// <param name="read">func&lt;int ptr, object&gt; that supplies the read for this operation.</param>
            /// <returns>object produced by the operation for information on native token reader.</returns>
            private static object Information(IntPtr token, int kind, string name, IDictionary<string, int> errors, int minimum, Func<IntPtr, object> read)
            {
                int size;
                bool sized = GetTokenInformation(token, kind, IntPtr.Zero, 0, out size);
                int error = Marshal.GetLastWin32Error();
                if (sized || error != 122 || size < minimum || size > 65536) { errors[name] = error; return null; }
                IntPtr buffer = Marshal.AllocHGlobal(size);
                try
                {
                    int returned;
                    bool ok = GetTokenInformation(token, kind, buffer, size, out returned);
                    error = Marshal.GetLastWin32Error();
                    if (!ok || returned < minimum || returned > size) { errors[name] = error; return null; }
                    return read(buffer);
                }
                finally { Marshal.FreeHGlobal(buffer); }
            }
        }

        /// <summary>Carries the luid values passed between operations.</summary>
        [StructLayout(LayoutKind.Sequential)] private struct Luid
        {

            /// <summary>Maintains the low state for luid.</summary>
            internal uint Low;

/// <summary>Maintains the high state for luid.</summary>
internal int High;

            /// <summary>Handles to string for luid.</summary>
            /// <returns>Text produced by the operation for to string on luid.</returns>
            public override string ToString() => unchecked((uint)High).ToString("X8") + ":" + Low.ToString("X8");
        }

        /// <summary>Carries the token statistics values passed between operations.</summary>
        [StructLayout(LayoutKind.Sequential)] private struct TokenStatistics
        {

            /// <summary>Identifies the token id and authentication id associated with token statistics.</summary>
            internal Luid TokenId, AuthenticationId;

            /// <summary>Maintains the expiration time state for token statistics.</summary>
            internal long ExpirationTime;

            /// <summary>Maintains the token type and impersonation level state for token statistics.</summary>
            internal int TokenType, ImpersonationLevel;

            /// <summary>Counts the dynamic charged and dynamic available and group count and privilege count maintained by token statistics.</summary>
            internal uint DynamicCharged, DynamicAvailable, GroupCount, PrivilegeCount;

            /// <summary>Identifies the modified id associated with token statistics.</summary>
            internal Luid ModifiedId;
        }

        /// <summary>Returns current thread for effective token observation.</summary>
        /// <returns>int ptr produced by the operation for get current thread on effective token observation.</returns>
        [DllImport("kernel32.dll")] private static extern IntPtr GetCurrentThread();

        /// <summary>Returns current process for effective token observation.</summary>
        /// <returns>int ptr produced by the operation for get current process on effective token observation.</returns>
        [DllImport("kernel32.dll")] private static extern IntPtr GetCurrentProcess();

        /// <summary>Closes handle for effective token observation.</summary>
        /// <param name="handle">Native handle that supplies the handle for this operation.</param>
        /// <returns>Boolean indicating the result of the check for close handle on effective token observation.</returns>
        [DllImport("kernel32.dll", SetLastError = true)] [return: MarshalAs(UnmanagedType.Bool)] private static extern bool CloseHandle(IntPtr handle);

        /// <summary>Opens thread token for effective token observation.</summary>
        /// <param name="thread">Native handle that supplies the thread for this operation.</param>
        /// <param name="access">uint that supplies the access for this operation.</param>
        /// <param name="openAsSelf">Indicates whether open as self is enabled.</param>
        /// <param name="token">Native handle that supplies the token for this operation.</param>
        /// <returns>Boolean indicating the result of the check for open thread token on effective token observation.</returns>
        [DllImport("advapi32.dll", SetLastError = true)] [return: MarshalAs(UnmanagedType.Bool)] private static extern bool OpenThreadToken(IntPtr thread, uint access, [MarshalAs(UnmanagedType.Bool)] bool openAsSelf, out IntPtr token);

        /// <summary>Opens process token for effective token observation.</summary>
        /// <param name="process">Native handle that supplies the process for this operation.</param>
        /// <param name="access">uint that supplies the access for this operation.</param>
        /// <param name="token">Native handle that supplies the token for this operation.</param>
        /// <returns>Boolean indicating the result of the check for open process token on effective token observation.</returns>
        [DllImport("advapi32.dll", SetLastError = true)] [return: MarshalAs(UnmanagedType.Bool)] private static extern bool OpenProcessToken(IntPtr process, uint access, out IntPtr token);

        /// <summary>Returns token information for effective token observation.</summary>
        /// <param name="token">Native handle that supplies the token for this operation.</param>
        /// <param name="kind">int that supplies the kind for this operation.</param>
        /// <param name="information">Native handle that supplies the information for this operation.</param>
        /// <param name="size">int that supplies the size for this operation.</param>
        /// <param name="returned">int that supplies the returned for this operation.</param>
        /// <returns>Boolean indicating the result of the check for get token information on effective token observation.</returns>
        [DllImport("advapi32.dll", SetLastError = true)] [return: MarshalAs(UnmanagedType.Bool)] private static extern bool GetTokenInformation(IntPtr token, int kind, IntPtr information, int size, out int returned);

        /// <summary>Determines whether token restricted for effective token observation.</summary>
        /// <param name="token">Native handle that supplies the token for this operation.</param>
        /// <returns>Boolean indicating the result of the check for is token restricted on effective token observation.</returns>
        [DllImport("advapi32.dll")] [return: MarshalAs(UnmanagedType.Bool)] private static extern bool IsTokenRestricted(IntPtr token);
    }
}

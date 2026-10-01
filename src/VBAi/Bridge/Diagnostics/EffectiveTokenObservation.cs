using System;
using System.Collections.Generic;
using System.Runtime.InteropServices;
using System.Security.Principal;

namespace VBAi
{
    /// <summary>Opt-in qualification evidence: public metadata of the calling thread's effective token only.</summary>
    internal static class EffectiveTokenObservation
    {
        internal const int ErrorNoToken = 1008;
        internal interface ITokenReader
        {
            bool OpenThread(out IntPtr token, out int error);
            bool OpenPrimary(out IntPtr token, out int error);
            IDictionary<string, object> Metadata(IntPtr token);
            void Close(IntPtr token);
        }

        /// <summary>Falls back to the primary token only for ERROR_NO_TOKEN; never impersonates or changes a token.</summary>
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

        private sealed class NativeTokenReader : ITokenReader
        {
            public bool OpenThread(out IntPtr token, out int error)
            {
                bool ok = OpenThreadToken(GetCurrentThread(), 0x8 /* TOKEN_QUERY */, true, out token);
                error = Marshal.GetLastWin32Error(); return ok;
            }
            public bool OpenPrimary(out IntPtr token, out int error)
            {
                bool ok = OpenProcessToken(GetCurrentProcess(), 0x8 /* TOKEN_QUERY */, out token);
                error = Marshal.GetLastWin32Error(); return ok;
            }
            public void Close(IntPtr token)
            {
                if (!CloseHandle(token)) throw new System.ComponentModel.Win32Exception(Marshal.GetLastWin32Error());
            }
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

        [StructLayout(LayoutKind.Sequential)] private struct Luid
        {
            internal uint Low; internal int High;
            public override string ToString() => unchecked((uint)High).ToString("X8") + ":" + Low.ToString("X8");
        }
        [StructLayout(LayoutKind.Sequential)] private struct TokenStatistics
        {
            internal Luid TokenId, AuthenticationId;
            internal long ExpirationTime;
            internal int TokenType, ImpersonationLevel;
            internal uint DynamicCharged, DynamicAvailable, GroupCount, PrivilegeCount;
            internal Luid ModifiedId;
        }
        [DllImport("kernel32.dll")] private static extern IntPtr GetCurrentThread();
        [DllImport("kernel32.dll")] private static extern IntPtr GetCurrentProcess();
        [DllImport("kernel32.dll", SetLastError = true)] [return: MarshalAs(UnmanagedType.Bool)] private static extern bool CloseHandle(IntPtr handle);
        [DllImport("advapi32.dll", SetLastError = true)] [return: MarshalAs(UnmanagedType.Bool)] private static extern bool OpenThreadToken(IntPtr thread, uint access, [MarshalAs(UnmanagedType.Bool)] bool openAsSelf, out IntPtr token);
        [DllImport("advapi32.dll", SetLastError = true)] [return: MarshalAs(UnmanagedType.Bool)] private static extern bool OpenProcessToken(IntPtr process, uint access, out IntPtr token);
        [DllImport("advapi32.dll", SetLastError = true)] [return: MarshalAs(UnmanagedType.Bool)] private static extern bool GetTokenInformation(IntPtr token, int kind, IntPtr information, int size, out int returned);
        [DllImport("advapi32.dll")] [return: MarshalAs(UnmanagedType.Bool)] private static extern bool IsTokenRestricted(IntPtr token);
    }
}

using System;
using System.ComponentModel;
using System.Linq;
using System.Runtime.InteropServices;
using System.Security.AccessControl;
using System.Security.Principal;

namespace VBAi.Tests.Integration
{
    /// <summary>Copies raw source DACL ACEs without translating generic masks into FileSystemRights.</summary>
    internal static class NativeProbeDaclCopy
    {
        [DllImport("advapi32.dll", CharSet = CharSet.Unicode, ExactSpelling = true, SetLastError = true)]
        [return: MarshalAs(UnmanagedType.Bool)]
        private static extern bool GetFileSecurityW(string path, uint information, [Out] byte[] descriptor, uint length, out uint needed);
        [DllImport("advapi32.dll", CharSet = CharSet.Unicode, ExactSpelling = true, SetLastError = true)]
        [return: MarshalAs(UnmanagedType.Bool)]
        private static extern bool SetFileSecurityW(string path, uint information, [In] byte[] descriptor);

        /// <summary>Reads only the public owner and DACL, preserving raw ACE masks and flags without an ACL adapter.</summary>
        internal static RawSecurityDescriptor ReadAccessAndOwner(string path)
        {
            const uint ownerAndDacl = 0x00000005;
            uint needed;
            bool unexpectedSuccess = GetFileSecurityW(path, ownerAndDacl, null, 0, out needed);
            int error = Marshal.GetLastWin32Error();
            if (unexpectedSuccess || error != 122 || needed == 0 || needed > 65536)
                throw new Win32Exception(error, "Native owner/DACL sizing failed or exceeded the disposable probe bound.");
            var bytes = new byte[needed];
            if (!GetFileSecurityW(path, ownerAndDacl, bytes, (uint)bytes.Length, out needed))
                throw new Win32Exception(Marshal.GetLastWin32Error(), "Native owner/DACL read failed.");
            return new RawSecurityDescriptor(bytes, 0);
        }

        /// <summary>Writes only the DACL and its protection bit; never sets owner/group/SACL or requests elevation.</summary>
        internal static void WriteProtectedAccess(string freshOwnedChild, RawSecurityDescriptor descriptor)
        {
            const uint daclAndProtectionOnly = 0x80000004;
            if (!SetFileSecurityW(freshOwnedChild, daclAndProtectionOnly, Bytes(descriptor)))
                throw new Win32Exception(Marshal.GetLastWin32Error(), "Fresh-child DACL write failed; no elevation or permission retry is allowed.");
        }

        /// <summary>Preserves target identity/SACL and source ACE order, types, masks and flags except Inherited.</summary>
        internal static RawSecurityDescriptor Create(byte[] sourceBytes, byte[] targetBytes)
        {
            var source = new RawSecurityDescriptor(sourceBytes, 0);
            var target = new RawSecurityDescriptor(targetBytes, 0);
            if (source.DiscretionaryAcl == null || (source.ControlFlags & ControlFlags.DiscretionaryAclPresent) == 0)
                throw new ArgumentException("Source must have a non-null DACL; do not replace it with an unrestricted descriptor.", nameof(sourceBytes));
            var copied = new RawAcl(source.DiscretionaryAcl.Revision, source.DiscretionaryAcl.Count);
            for (int index = 0; index < source.DiscretionaryAcl.Count; index++)
            {
                var ace = GenericAce.CreateFromBinaryForm(Bytes(source.DiscretionaryAcl[index]), 0);
                ace.AceFlags &= ~AceFlags.Inherited;
                copied.InsertAce(index, ace);
            }
            ControlFlags flags = target.ControlFlags & ~(ControlFlags.DiscretionaryAclAutoInherited |
                ControlFlags.DiscretionaryAclAutoInheritRequired | ControlFlags.DiscretionaryAclDefaulted);
            flags |= ControlFlags.DiscretionaryAclPresent | ControlFlags.DiscretionaryAclProtected;
            return new RawSecurityDescriptor(flags, target.Owner, target.Group, target.SystemAcl, copied);
        }

        /// <summary>Checks a direct source allow ACE; never inserts a new grant for the current user.</summary>
        internal static bool SourceAllowsDirectFullControl(byte[] sourceBytes, SecurityIdentifier user)
        {
            var source = new RawSecurityDescriptor(sourceBytes, 0);
            if (source.DiscretionaryAcl == null) return false;
            return source.DiscretionaryAcl.Cast<GenericAce>().OfType<QualifiedAce>().Any(ace =>
                ace.AceQualifier == AceQualifier.AccessAllowed && ace.SecurityIdentifier.Equals(user) &&
                (ace.AceFlags & AceFlags.InheritOnly) == 0 &&
                (((uint)ace.AccessMask & 0x10000000U) != 0 ||
                 (ace.AccessMask & (int)FileSystemRights.FullControl) == (int)FileSystemRights.FullControl));
        }

        /// <summary>Accepts only Windows adding the reported AI control flag; DACL bytes and identity must remain exact.</summary>
        internal static void VerifyReadback(RawSecurityDescriptor expected, RawSecurityDescriptor actual)
        {
            if (actual.DiscretionaryAcl == null || !Bytes(expected.DiscretionaryAcl).SequenceEqual(Bytes(actual.DiscretionaryAcl)))
                throw new InvalidOperationException("DACL readback changed raw ACE bytes, masks, order, types or flags.");
            if (!Equals(expected.Owner, actual.Owner) || !Equals(expected.Group, actual.Group))
                throw new InvalidOperationException("DACL readback changed the target owner or group.");
            ControlFlags difference = expected.ControlFlags ^ actual.ControlFlags;
            bool addedAiOnly = difference == ControlFlags.DiscretionaryAclAutoInherited &&
                (expected.ControlFlags & ControlFlags.DiscretionaryAclAutoInherited) == 0;
            if (difference != 0 && !addedAiOnly)
                throw new InvalidOperationException("DACL readback changed control flags beyond adding AutoInherited.");
            if ((actual.ControlFlags & (ControlFlags.DiscretionaryAclPresent | ControlFlags.DiscretionaryAclProtected)) !=
                (ControlFlags.DiscretionaryAclPresent | ControlFlags.DiscretionaryAclProtected))
                throw new InvalidOperationException("Copied DACL must be present and protected against inheritance.");
        }

        internal static byte[] Bytes(GenericAce value)
        { var bytes = new byte[value.BinaryLength]; value.GetBinaryForm(bytes, 0); return bytes; }
        internal static byte[] Bytes(RawAcl value)
        { var bytes = new byte[value.BinaryLength]; value.GetBinaryForm(bytes, 0); return bytes; }
        internal static byte[] Bytes(RawSecurityDescriptor value)
        { var bytes = new byte[value.BinaryLength]; value.GetBinaryForm(bytes, 0); return bytes; }
    }
}

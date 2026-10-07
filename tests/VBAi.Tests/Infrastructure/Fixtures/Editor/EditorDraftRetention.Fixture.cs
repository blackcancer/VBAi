using Microsoft.Win32.SafeHandles;
using System;
using System.IO;
using System.Runtime.InteropServices;
using System.Text;

namespace VBAi.Tests.Infrastructure
{
    internal static class OwnedDraftJunction
    {
        [DllImport("kernel32.dll", CharSet = CharSet.Unicode, SetLastError = true)]
        private static extern SafeFileHandle CreateFile(string path, uint access, uint share, IntPtr security, uint creation, uint flags, IntPtr template);
        [DllImport("kernel32.dll", SetLastError = true)]
        private static extern bool DeviceIoControl(SafeFileHandle handle, uint code, byte[] input, int inputLength, IntPtr output, int outputLength, out int returned, IntPtr overlapped);
        internal static void Create(string path, string target)
        {
            Directory.CreateDirectory(path);
            byte[] substitute = Encoding.Unicode.GetBytes(@"\??\" + Path.GetFullPath(target));
            byte[] printable = Encoding.Unicode.GetBytes(Path.GetFullPath(target));
            using (var stream = new MemoryStream())
            using (var writer = new BinaryWriter(stream))
            {
                writer.Write(0xA0000003u); writer.Write((ushort)(8 + substitute.Length + printable.Length + 4)); writer.Write((ushort)0);
                writer.Write((ushort)0); writer.Write((ushort)substitute.Length); writer.Write((ushort)(substitute.Length + 2)); writer.Write((ushort)printable.Length);
                writer.Write(substitute); writer.Write((ushort)0); writer.Write(printable); writer.Write((ushort)0);
                byte[] buffer = stream.ToArray();
                using (var handle = CreateFile(path, 0x40000000, 7, IntPtr.Zero, 3, 0x02200000, IntPtr.Zero))
                {
                    if (handle.IsInvalid || !DeviceIoControl(handle, 0x000900A4, buffer, buffer.Length, IntPtr.Zero, 0, out int returned, IntPtr.Zero))
                        throw new System.ComponentModel.Win32Exception(Marshal.GetLastWin32Error());
                }
            }
        }
    }
}
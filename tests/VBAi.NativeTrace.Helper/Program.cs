using System;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;

namespace VBAi.NativeTrace.Helper
{
    /// <summary>Disposable non-Office syscall target with fixed stdin commands and one bounded synthetic file operation.</summary>
    internal static class Program
    {
        [DllImport("kernel32.dll", CharSet = CharSet.Unicode, SetLastError = true)]
        private static extern IntPtr CreateFileW(string path, uint access, uint share, IntPtr security, uint disposition, uint flags, IntPtr template);
        [DllImport("kernel32.dll", SetLastError = true)] private static extern bool WriteFile(IntPtr file, byte[] bytes, uint count, out uint written, IntPtr overlapped);
        [DllImport("kernel32.dll", SetLastError = true)] private static extern bool ReadFile(IntPtr file, byte[] bytes, uint count, out uint read, IntPtr overlapped);
        [DllImport("kernel32.dll", SetLastError = true)] private static extern uint SetFilePointer(IntPtr file, int distance, IntPtr high, uint method);
        [DllImport("kernel32.dll")] private static extern bool CloseHandle(IntPtr handle);

        private static int Main(string[] args)
        {
            if (args.Length != 1 || !Path.IsPathRooted(args[0]) || !Directory.Exists(args[0]) ||
                !Guid.TryParseExact(Path.GetFileName(args[0]), "N", out _)) return 2;
            string destination = Path.Combine(args[0], "synthetic-owned.txt");
            if (File.Exists(destination)) return 3;
            using (var process = Process.GetCurrentProcess())
                Console.WriteLine("READY " + process.Id + " " + process.StartTime.ToUniversalTime().ToString("o"));
            bool opened = false;
            for (string command; (command = Console.ReadLine()) != null;)
            {
                if (command == "STOP") { Console.WriteLine("STOPPED"); return 0; }
                if (command != "OPEN" || opened) { Console.WriteLine("REFUSED"); continue; }
                opened = true;
                IntPtr file = CreateFileW(destination, 0xC0000000, 0, IntPtr.Zero, 1, 0x80, IntPtr.Zero);
                int error = Marshal.GetLastWin32Error();
                if (file == new IntPtr(-1)) { Console.WriteLine("OPEN_FAILED " + error); continue; }
                try
                {
                    byte[] bytes = Encoding.ASCII.GetBytes("VBAi owned diagnostic fixture");
                    bool written = WriteFile(file, bytes, (uint)bytes.Length, out uint count, IntPtr.Zero) && count == bytes.Length;
                    bool seek = SetFilePointer(file, 0, IntPtr.Zero, 0) == 0;
                    byte[] actual = new byte[bytes.Length];
                    bool read = ReadFile(file, actual, (uint)actual.Length, out uint received, IntPtr.Zero) && received == bytes.Length;
                    Console.WriteLine(written && seek && read && bytes.SequenceEqual(actual) ? "OPEN_VERIFIED" : "OPEN_READBACK_FAILED");
                }
                finally { CloseHandle(file); }
            }
            return 4; // An EOF is not the owned normal-shutdown command.
        }
    }
}

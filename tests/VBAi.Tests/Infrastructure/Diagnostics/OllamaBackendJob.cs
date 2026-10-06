using System;
using System.ComponentModel;
using System.Diagnostics;
using System.Runtime.InteropServices;
using System.Threading;

namespace VBAi.Tests.Integration
{
    /// <summary>Contains only a freshly launched qualification backend and its future workers.</summary>
    /// <remarks>No kill-on-close flag: uncertain native work must remain available for diagnosis.</remarks>
    public sealed class OllamaBackendJob : IDisposable
    {
        private IntPtr job;
        private bool attached;
        private bool stopEntered;

        [DllImport("kernel32.dll", CharSet = CharSet.Unicode, SetLastError = true)]
        private static extern IntPtr CreateJobObject(IntPtr attributes, string name);
        [DllImport("kernel32.dll", SetLastError = true)]
        private static extern bool AssignProcessToJobObject(IntPtr job, IntPtr process);
        [DllImport("kernel32.dll", SetLastError = true)]
        private static extern bool IsProcessInJob(IntPtr process, IntPtr job, out bool member);
        [DllImport("kernel32.dll", SetLastError = true)]
        private static extern bool QueryInformationJobObject(IntPtr job, int informationClass, IntPtr buffer, uint size, out uint returned);
        [DllImport("kernel32.dll", SetLastError = true)]
        private static extern bool TerminateJobObject(IntPtr job, uint exitCode);
        [DllImport("kernel32.dll")]
        private static extern bool CloseHandle(IntPtr handle);

        /// <summary>Creates a private unnamed job without global limits or automatic termination.</summary>
        public OllamaBackendJob()
        {
            job = CreateJobObject(IntPtr.Zero, null);
            if (job == IntPtr.Zero) throw new Win32Exception(Marshal.GetLastWin32Error());
        }

        /// <summary>Attaches the caller's original backend handle before any model request.</summary>
        public void Attach(Process originalBackend)
        {
            RequireOpen();
            if (attached || stopEntered) throw new InvalidOperationException("Only one newly created backend can enter this job.");
            if (originalBackend == null) throw new ArgumentNullException(nameof(originalBackend));
            if (originalBackend.HasExited) throw new InvalidOperationException("Backend already exited.");
            if (!AssignProcessToJobObject(job, originalBackend.Handle)) throw new Win32Exception(Marshal.GetLastWin32Error());
            attached = true;
            bool member;
            if (!IsProcessInJob(originalBackend.Handle, job, out member) || !member)
                throw new InvalidOperationException("Original backend job membership was not proven.");
        }

        /// <summary>Reads the kernel's current membership, without a process-name or PID search.</summary>
        public int[] ReadProcessIds()
        {
            RequireOpen();
            // JOBOBJECT_BASIC_PROCESS_ID_LIST: two DWORD counts followed by ULONG_PTR IDs.
            // Refuse an unexpectedly large/incomplete set rather than dropping unknown members.
            const int capacity = 256;
            int size = 8 + capacity * IntPtr.Size;
            IntPtr buffer = Marshal.AllocHGlobal(size);
            try
            {
                uint returned;
                if (!QueryInformationJobObject(job, 3, buffer, (uint)size, out returned))
                    throw new Win32Exception(Marshal.GetLastWin32Error());
                int assigned = Marshal.ReadInt32(buffer), count = Marshal.ReadInt32(buffer, 4);
                if (count < 0 || count > capacity || assigned != count)
                    throw new InvalidOperationException("Backend job membership snapshot is incomplete.");
                var ids = new int[count];
                for (int i = 0; i < count; i++) ids[i] = checked((int)Marshal.ReadIntPtr(buffer, 8 + i * IntPtr.Size).ToInt64());
                return ids;
            }
            finally { Marshal.FreeHGlobal(buffer); }
        }

        /// <summary>Stops this synthetic job once, after requests settle and all Office hosts exit.</summary>
        /// <returns>The membership immediately before termination. It is not an Office exit receipt.</returns>
        public int[] Stop(bool requestsSettled, bool officeAbsent)
        {
            RequireOpen();
            if (!requestsSettled || !officeAbsent) throw new InvalidOperationException("Unsettled provider or native work forbids backend teardown.");
            if (!attached || stopEntered) throw new InvalidOperationException("Backend shutdown requires one attachment and permits no retry.");
            var members = ReadProcessIds();
            stopEntered = true;
            if (!TerminateJobObject(job, 1)) throw new Win32Exception(Marshal.GetLastWin32Error());
            var watch = Stopwatch.StartNew();
            while (ReadProcessIds().Length != 0)
            {
                if (watch.ElapsedMilliseconds >= 15000) throw new TimeoutException("Backend job still contains live workers.");
                Thread.Sleep(50);
            }
            return members;
        }

        private void RequireOpen()
        {
            if (job == IntPtr.Zero) throw new ObjectDisposedException(nameof(OllamaBackendJob));
        }

        /// <summary>Closes only the job handle; it deliberately leaves uncertain processes alive.</summary>
        public void Dispose()
        {
            if (job == IntPtr.Zero) return;
            CloseHandle(job);
            job = IntPtr.Zero;
        }
    }
}

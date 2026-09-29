using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace VBAi
{
    /// <summary>Copies native debugger text only when the existing clipboard can be preserved.</summary>
    internal sealed class VbeDebugClipboard
    {
        [DllImport("user32.dll")] private static extern uint GetClipboardSequenceNumber();
        [DllImport("user32.dll")] private static extern IntPtr GetClipboardOwner();
        [DllImport("user32.dll")] private static extern uint GetWindowThreadProcessId(IntPtr window, out uint processId);

        internal Func<uint> Sequence = GetClipboardSequenceNumber;
        internal Func<Task> YieldNative = () => Task.Delay(20);
        private int reading;
        internal Func<IDataObject> ReadData = Clipboard.GetDataObject;
        internal Action<DataObject> WriteData = data => {
            if (data == null) Clipboard.Clear();
            else Clipboard.SetDataObject(data, true, 3, 20);
        };
        internal Func<bool> IsHostOwner = () => {
            GetWindowThreadProcessId(GetClipboardOwner(), out uint processId);
            return processId == (uint)Process.GetCurrentProcess().Id;
        };

        /// <summary>Runs a validated Copy command once, reads its new text and restores the captured formats.</summary>
        internal async Task<string> ReadAsync(Action prepareSelection, Action copy, Action validateTarget = null)
        {
            if (Thread.CurrentThread.GetApartmentState() != ApartmentState.STA || SynchronizationContext.Current == null)
                throw new InvalidOperationException("Native debugger copy requires the VBE STA synchronization context.");
            if (Interlocked.CompareExchange(ref reading, 1, 0) != 0)
                throw new InvalidOperationException("A native Immediate capture is already in progress.");
            try { return await ReadCoreAsync(prepareSelection, copy, validateTarget); }
            finally { Interlocked.Exchange(ref reading, 0); }
        }

        private async Task<string> ReadCoreAsync(Action prepareSelection, Action copy, Action validateTarget)
        {
            int ownerThread = Thread.CurrentThread.ManagedThreadId;
            Action verifyThread = () => {
                if (Thread.CurrentThread.ManagedThreadId != ownerThread || Thread.CurrentThread.GetApartmentState() != ApartmentState.STA)
                    throw new InvalidOperationException("Native clipboard capture left the owning STA thread.");
            };
            uint before = Sequence();
            var backup = Snapshot.Capture(ReadData());
            if (Sequence() != before)
                throw new InvalidOperationException("Clipboard changed during backup; native Copy was not attempted.");
            prepareSelection();
            // CommandBars.Execute can enqueue the native action and return before it runs.
            // Yield to the owning message loop instead of blocking it or pumping recursively.
            await YieldNative();
            verifyThread();
            validateTarget?.Invoke();
            if (Sequence() != before)
                throw new InvalidOperationException("Clipboard changed before native Copy; it was not attempted.");

            uint copied = before;
            bool ownsCopy = false;
            try
            {
                Exception commandError = null;
                try { copy(); }
                catch (Exception ex) { commandError = ex; }
                for (int attempt = 0; attempt < 25; attempt++)
                {
                    await YieldNative();
                    verifyThread();
                    copied = Sequence();
                    ownsCopy = copied != before && IsHostOwner();
                    if (copied != before) break;
                    validateTarget?.Invoke();
                }
                if (commandError != null) throw new InvalidOperationException("Native Copy failed; it was not retried.", commandError);
                if (!ownsCopy)
                    throw new InvalidOperationException("Native Copy produced no verifiable clipboard update from this host.");
                validateTarget?.Invoke();
                var data = ReadData();
                string text = data?.GetData(DataFormats.UnicodeText, false) as string
                    ?? data?.GetData(DataFormats.Text, false) as string;
                if (Sequence() != copied || !IsHostOwner())
                    throw new InvalidOperationException("Clipboard changed while reading native output; that output was discarded.");
                if (text == null || text.Length > 1024 * 1024)
                    throw new InvalidOperationException("Native Copy did not produce text within the one-million-character limit.");
                return text;
            }
            finally
            {
                if (ownsCopy)
                {
                    verifyThread();
                    if (Sequence() != copied || !IsHostOwner())
                        throw new InvalidOperationException("Clipboard changed after native Copy. The newer content was preserved; the earlier clipboard was not restored.");
                    WriteData(backup.CreateDataObject());
                    uint restored = Sequence();
                    if (!backup.Matches(ReadData()) || restored != Sequence())
                        throw new InvalidOperationException("The original clipboard could not be verified after restoration.");
                }
            }
        }

        /// <summary>A bounded, independent copy; no unsupported or unavailable format is silently dropped.</summary>
        internal sealed class Snapshot
        {
            internal const int MaximumBytes = 8 * 1024 * 1024;
            private readonly Dictionary<string, object> values = new Dictionary<string, object>(StringComparer.Ordinal);

            internal static Snapshot Capture(IDataObject data)
            {
                var snapshot = new Snapshot();
                int total = 0;
                foreach (string format in data?.GetFormats(false) ?? new string[0])
                {
                    object value = data.GetData(format, false);
                    long length;
                    if (value is string text) length = (long)text.Length * 2;
                    else if (value is MemoryStream stream) length = stream.Length;
                    else throw new InvalidOperationException("Clipboard format cannot be preserved: " + format + ". Native Copy was not attempted.");
                    if (length > MaximumBytes - total)
                        throw new InvalidOperationException("Clipboard backup exceeds 8 MiB; native Copy was not attempted.");
                    total += (int)length;
                    snapshot.values.Add(format, value is MemoryStream bytes ? (object)bytes.ToArray() : value);
                }
                return snapshot;
            }

            internal DataObject CreateDataObject()
            {
                if (values.Count == 0) return null;
                var data = new DataObject();
                foreach (var value in values)
                    data.SetData(value.Key, false, value.Value is byte[] bytes ? (object)new MemoryStream(bytes, false) : value.Value);
                return data;
            }

            internal bool Matches(IDataObject data)
            {
                if (values.Count == 0) return data == null || data.GetFormats(false).Length == 0;
                if (data == null) return false;
                if (!new HashSet<string>(data.GetFormats(false), StringComparer.Ordinal).SetEquals(values.Keys)) return false;
                foreach (var value in values)
                {
                    object actual = data.GetData(value.Key, false);
                    if (value.Value is byte[] bytes)
                    {
                        if (!(actual is MemoryStream stream) || !bytes.SequenceEqual(stream.ToArray())) return false;
                    }
                    else if (!Equals(value.Value, actual)) return false;
                }
                return true;
            }
        }
    }
}

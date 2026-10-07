using System.Threading;

namespace VBAi
{
    /// <summary>Cancels only queued compilation callbacks; admitted native work is never replayed or declared cancelled.</summary>
    internal sealed class VbeCompilationAdmission
    {
        private int state;
        internal bool TryBegin() => Interlocked.CompareExchange(ref state, 1, 0) == 0;
        internal bool CancelPending() => Interlocked.CompareExchange(ref state, 2, 0) == 0;
    }
}

using System;
using System.Collections.Generic;

namespace VBAi.Tests.Integration
{
    /// <summary>Bounds native enumeration and retains only visible windows on the exact owned UI thread.</summary>
    internal static class WordChatTopWindowInventory
    {
        internal delegate bool Visitor(IntPtr window, IntPtr state);
        internal delegate bool Enumerator(Visitor visitor, IntPtr state);

        internal sealed class Identity
        {
            internal int ProcessId;
            internal uint ThreadId;
            internal bool Visible;
        }

        internal sealed class Receipt
        {
            internal bool ApiReturned;
            internal int VisitedTotal, OwnedProcessCount, ExactThreadVisibleCount;
            internal bool GlobalBoundHit;
            internal string FailureStatus;
        }

        internal static IntPtr[] Read(Enumerator enumerate, Func<IntPtr, Identity> identify,
            int processId, uint threadId, Action<Receipt> recordFailure)
        {
            if (enumerate == null) throw new ArgumentNullException(nameof(enumerate));
            if (identify == null) throw new ArgumentNullException(nameof(identify));
            if (recordFailure == null) throw new ArgumentNullException(nameof(recordFailure));
            var selected = new List<IntPtr>();
            var receipt = new Receipt();
            Visitor visitor = (window, unused) => {
                if (++receipt.VisitedTotal > 4096)
                {
                    receipt.GlobalBoundHit = true;
                    return false;
                }
                Identity identity = identify(window);
                if (identity != null && identity.ProcessId == processId)
                {
                    receipt.OwnedProcessCount++;
                    if (identity.ThreadId == threadId && identity.Visible)
                    {
                        receipt.ExactThreadVisibleCount++;
                        if (selected.Count < 64) selected.Add(window);
                    }
                }
                return true;
            };
            receipt.ApiReturned = enumerate(visitor, IntPtr.Zero);
            receipt.FailureStatus = receipt.GlobalBoundHit ? "GlobalBound" :
                !receipt.ApiReturned ? "ApiFailure" :
                receipt.ExactThreadVisibleCount > 64 ? "TargetBound" : null;
            if (receipt.FailureStatus != null)
            {
                recordFailure(receipt);
                throw new InvalidOperationException("Bounded Word top-window inventory failed: " + receipt.FailureStatus + ".");
            }
            return selected.ToArray();
        }
    }
}

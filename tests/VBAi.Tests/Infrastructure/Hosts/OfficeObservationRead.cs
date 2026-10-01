using System;
using System.Runtime.InteropServices;
using System.Threading;

namespace VBAi.Tests.Integration
{
    /// <summary>Bounds retries of explicitly read-only native observations on their calling STA.</summary>
    internal static class OfficeObservationRead
    {
        /// <summary>Retries only a rejected COM getter; never accepts a save, close, or mutation delegate.</summary>
        internal static T Getter<T>(Func<T> getter, Action<int, int> recordRetry = null, Action<int> pause = null)
        {
            if (getter == null) throw new ArgumentNullException(nameof(getter));
            int[] delays = { 50, 100, 200, 400, 800 };
            for (int attempt = 0; ; attempt++)
            {
                try { return getter(); }
                catch (COMException error) when (error.HResult == unchecked((int)0x80010001) && attempt < delays.Length)
                {
                    recordRetry?.Invoke(attempt + 1, delays[attempt]);
                    (pause ?? Thread.Sleep)(delays[attempt]);
                }
            }
        }
    }
}

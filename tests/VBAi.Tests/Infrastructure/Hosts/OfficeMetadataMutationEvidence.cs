using System;
using System.Runtime.ExceptionServices;

namespace VBAi.Tests.Integration
{
    /// <summary>Observes a returned mutation once, including failure, without replacing or retrying its original outcome.</summary>
    internal static class OfficeMetadataMutationEvidence
    {
        internal static void Run(Action mutate, Action observe)
        {
            Exception failure = null;
            try { mutate(); }
            catch (Exception original) { failure = original; }
            try { observe(); }
            catch (Exception readback) { failure = failure == null ? readback :
                new AggregateException("The original metadata mutation and its read-only diagnostic both failed; no mutation was replayed.", failure, readback); }
            if (failure != null) ExceptionDispatchInfo.Capture(failure).Throw();
        }
    }
}

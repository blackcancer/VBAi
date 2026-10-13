namespace VBAi
{
    internal sealed partial class VbeTestExplorerService
    {
        /// <summary>Returns a bridge reply without throwing for an expected source-revision refusal.</summary>
        /// <param name="request">The request admitted by the owning session and its privacy checks.</param>
        /// <param name="executionGuard">Authorization callback revalidated before native execution.</param>
        /// <returns>The unchanged success payload or an explicit failure with no native dispatch.</returns>
        /// <remarks>Unexpected failures still propagate to the session's normal error handler. The
        /// strict internal command API retains its exception contract. A stale request does not
        /// become a caught first-chance exception in an attached host debugger.</remarks>
        internal Response CommandResponse(Request request, System.Action executionGuard = null)
        {
            object result = CommandCore(request, executionGuard, true);
            return result as Response ?? Response.Success(result);
        }
    }
}

using System;
using System.Threading.Tasks;

namespace VBAi
{
    internal static partial class VbeUiTask
    {
        /// <summary>Completes native work and publishes its terminal state on the same owner STA.</summary>
        /// <typeparam name="T">Operation result type.</typeparam>
        /// <param name="operation">Work whose awaits use the private owner dispatcher.</param>
        /// <param name="completed">Synchronous callback receiving any operation or cleanup failure.</param>
        /// <returns>The whole operation task, including terminal publication.</returns>
        /// <remarks>The callback runs after dispatcher disposal. It must finish synchronously and cannot post through that dispatcher.</remarks>
        internal static Task<T> RunWithCompletion<T>(Func<Task<T>> operation, Action<Exception> completed)
        {
            if (completed == null) throw new ArgumentNullException(nameof(completed));
            return RunOwned(operation, completed);
        }

        /// <summary>Attempts cleanup and terminal publication once while retaining every distinct failure.</summary>
        /// <param name="operationFailure">The original work failure, if any.</param>
        /// <param name="cleanup">Owner-thread cleanup, invoked once.</param>
        /// <param name="completed">Optional synchronous publisher, invoked after cleanup even when cleanup fails.</param>
        /// <returns>The original failure or an aggregate retaining additional cleanup/publication failures.</returns>
        internal static Exception FinishOwnerOperation(Exception operationFailure, Action cleanup, Action<Exception> completed)
        {
            if (cleanup == null) throw new ArgumentNullException(nameof(cleanup));
            Exception failure = operationFailure;
            try { cleanup(); }
            catch (Exception error) { failure = CombineOwnerFailures(failure, error); }
            try { completed?.Invoke(failure); }
            catch (Exception error) { failure = CombineOwnerFailures(failure, error); }
            return failure;
        }

        /// <summary>Preserves an original exception and adds a distinct later failure without duplicating its identity.</summary>
        /// <param name="first">Failure already captured, if any.</param>
        /// <param name="next">Newly observed failure.</param>
        /// <returns>The existing exception, the new exception, or an ordered aggregate of both.</returns>
        private static Exception CombineOwnerFailures(Exception first, Exception next)
            => first == null ? next : ReferenceEquals(first, next) ? first :
                new AggregateException("Owner operation and terminal cleanup failed.", first, next);
    }
}

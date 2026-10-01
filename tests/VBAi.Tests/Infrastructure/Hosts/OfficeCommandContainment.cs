using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Runtime.ExceptionServices;

namespace VBAi.Tests.Integration
{
    /// <summary>Records synthetic Office requests before dispatch and refuses native work after uncertain delivery.</summary>
    internal sealed class OfficeCommandContainment
    {
        private int sequence;
        internal bool Pending { get; private set; }
        internal bool Uncertain { get; private set; }
        internal string Command { get; private set; }
        internal Exception Failure { get; private set; }

        internal void RequireTerminal()
        {
            if (Pending || Uncertain)
                throw new InvalidOperationException("The owned Office bridge command " + Command + " has no confirmed terminal response; no native cleanup, reset, reopen or request replay is permitted.");
        }

        internal IDictionary<string, object> Send(string command, object request,
            Action<object> append, Action persist, Func<IDictionary<string, object>> dispatch, Action retain)
        {
            RequireTerminal();
            if (sequence >= 512) throw new InvalidOperationException("The synthetic Office command evidence limit was reached before dispatch.");
            var record = new Dictionary<string, object> {
                ["Command"] = command, ["Request"] = request, ["Sequence"] = ++sequence,
                ["StartedUtc"] = DateTime.UtcNow.ToString("o"), ["State"] = "Prepared; terminal response has not been observed"
            };
            append(record);
            persist(); // A failed preparation write never emits a request.
            Command = command;
            Pending = true;
            var timing = Stopwatch.StartNew();
            IDictionary<string, object> response = null;
            Exception failure = null;
            try
            {
                response = dispatch();
                if (response == null || !response.TryGetValue("Ok", out var ok) || !(ok is bool))
                    throw new InvalidDataException("The owned Office bridge did not return a terminal protocol response; retain the host without retry.");
                Pending = false;
                record["State"] = "TerminalResponse";
            }
            catch (Exception error)
            {
                failure = error;
                Uncertain = true; // A timeout does not cancel native work already emitted.
                record["State"] = "Uncertain; owned host retained without native cleanup";
                try { retain(); }
                catch (Exception retention) { failure = new AggregateException("Office delivery and ownership retention both failed.", error, retention); }
            }
            record["Response"] = response;
            record["Error"] = failure?.ToString();
            record["Pending"] = Pending;
            record["Uncertain"] = Uncertain;
            record["ElapsedMilliseconds"] = timing.ElapsedMilliseconds;
            record["FinishedUtc"] = DateTime.UtcNow.ToString("o");
            Failure = failure;
            try { persist(); }
            catch (Exception evidence)
            {
                if (failure != null)
                {
                    Failure = new AggregateException("Office dispatch and durable result persistence both failed; no request was replayed.", failure, evidence);
                    throw Failure;
                }
                throw;
            }
            if (failure != null) ExceptionDispatchInfo.Capture(failure).Throw();
            return response;
        }
    }
}

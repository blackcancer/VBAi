using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;

namespace VBAi
{

    /// <summary>Pauses and resumes HTTP provider tool rounds without replaying completed actions.</summary>
    internal sealed partial class ChatWindow
    {

        /// <summary>Current HTTP provider stream ID used to route incremental UI updates.</summary>
        private string providerStreamId;

        /// <summary>Retains only the latest HTTP stream protocol metadata for local diagnostics.</summary>
        private StreamDiagnostics lastHttpStreamDiagnostics;

        /// <summary>Runs a pending tool call and records its result for the paused provider turn.</summary>
        /// <param name="name">Provider tool name or the invoke_tool gateway name.</param>
        /// <param name="arguments">JSON arguments passed once to the VBE tool handler.</param>
        /// <returns>Serialized tool result; interrupted/uncertain attempts are recorded for inspection, not replay.</returns>
        private Task<string> ExecuteBudgetTool(string name, string arguments) => ExecuteProjectedBudgetTool(name, arguments, null);

        /// <summary>Projects one admitted HTTP tool call while preserving the existing result/replay bookkeeping.</summary>
        /// <param name="name">Declared tool name.</param>
        /// <param name="arguments">Complete argument JSON.</param>
        /// <param name="activityId">Scoped call identity; null when Copilot owns the call's activity events.</param>
        /// <returns>Original result supplied to the provider.</returns>
        private async Task<string> ExecuteProjectedBudgetTool(string name, string arguments, string activityId)
        {
            var ownerSession = currentSession;
            string ownerTurn = activeTurnId;
            Action<CodexAgentActivity> publish = activity =>
            {
                if (currentSession == ownerSession && activeTurnId == ownerTurn && !IsDisposed) ReceiveAgentActivity(activity);
            };
            string label = name;
            if (name == "invoke_tool")
            {
                if (json.DeserializeObject(arguments) is IDictionary<string, object> gateway && gateway.TryGetValue("ToolName", out var target)) label = Convert.ToString(target);
            }
            if (activityId != null) publish(ProviderActivityProjection.Tool(activityId, name, arguments, "inProgress"));
            try
            {
                string result = await InvokeTool(tools, name, arguments);
                if (activityId != null) publish(ProviderActivityProjection.Tool(activityId, name, arguments,
                    ProviderActivityProjection.ToolOutcome(result), ProviderActivityProjection.ToolDiagnostic(result)));
                bool success = false;
                try { success = json.Deserialize<Response>(result)?.Ok == true; } catch { }
                if (ownerSession.CompletedToolActions == null) ownerSession.CompletedToolActions = new List<string>();
                ownerSession.CompletedToolActions.Add(label + UiText.Get(success ? " — response received; inspect returned state" : " — refused or failed"));
                return result;
            }
            catch (Exception error)
            {
                if (activityId != null) publish(ProviderActivityProjection.Tool(activityId, name, arguments,
                    error is OperationCanceledException ? "interrupted" : "failed", error.Message));
                if (ownerSession.CompletedToolActions == null) ownerSession.CompletedToolActions = new List<string>();
                ownerSession.CompletedToolActions.Add(label + UiText.Get(" — interrupted; verify live state before retrying"));
                throw;
            }
        }

        /// <summary>Persists the current provider turn so it can resume after its execution budget is extended.</summary>
        /// <param name="provider">HTTP provider whose turn must resume.</param>
        /// <param name="model">Exact model ID selected when the turn was paused.</param>
        private void PauseBudget(LlmProvider provider, string model)
        {
            currentSession.BudgetPaused = true;
            currentSession.PausedTurnId = activeTurnId;
            currentSession.PausedProvider = provider.Name;
            currentSession.PausedModel = model;
            currentSession.PausedEffort = (effortPicker.SelectedItem as LlmEffortOption)?.Id;
            currentSession.PausedMode = currentSession.Mode;
            string actions = string.Join("\n", currentSession.CompletedToolActions ?? new List<string>());
            string remaining = verifyAfterEdit.Checked && codeChanges.Any(c => c.TurnId == activeTurnId)
                ? UiText.Get("Pending: automatic verification and final response.") : UiText.Get("Pending: final response.");
            AddEntry(new ChatEntry
            {
                Speaker = "Assistant",
                TurnId = activeTurnId,
                Text = UiText.Get("Safety pause: repeated rounds without progress or the intervention ceiling was reached. Resume from saved results; completed actions will not be replayed.") + "\n" + remaining + "\n\n" + actions
            });
            SetStatus(UiText.Get("Paused — resume when ready"));
        }

        /// <summary>Updates the budget controls to reflect whether the current turn is paused or resumable.</summary>
        private void UpdateBudgetControls()
        {
            bool hasScope = !loadingScope && !sessionViewUnavailable && (scopeSession == null || scopePicker?.SelectedItem is MacroScope);
            if (resumeTurn != null) resumeTurn.Enabled = hasScope && !busy && currentSession?.BudgetPaused == true;
            if (send == null || prompt == null) return;
            bool hasText = !string.IsNullOrWhiteSpace(prompt.Text);
            send.Symbol = busy && !hasText ? UiSymbol.Stop : currentSession?.BudgetPaused == true && !hasText ? UiSymbol.Play : UiSymbol.Upload;
            send.Text = busy ? UiText.Get(hasText ? "Queue ↑" : "Stop ■") :
                UiText.Get(currentSession?.BudgetPaused == true && !hasText ? "Resume ▶" : "Send ↑");
            toolTips.SetToolTip(send, UiText.Get(busy ? (hasText ? "Queue this message after the current response." : "Stop the current response. Changes already applied can still be undone in the chat.") : "Send the message and its context to the agent."));
            send.Enabled = busy ? (hasText ? hasScope : !stopRequested) : hasScope;
        }

        /// <summary>Resumes the saved provider turn after completing any pending tool responses.</summary>
        /// <returns>task produced by the operation for resume budget async on chat window.</returns>
        private async Task ResumeBudgetAsync()
        {
            if (busy || currentSession?.BudgetPaused != true) return;
            if (!(providerPicker.SelectedItem is LlmProvider provider) || !(modelPicker.SelectedItem is LlmModelOption model) || provider.IsCodex || provider.Name != currentSession.PausedProvider || model.Id != currentSession.PausedModel ||
                currentSession.Mode != currentSession.PausedMode || (effortPicker.SelectedItem as LlmEffortOption)?.Id != currentSession.PausedEffort)
            { SetStatus(UiText.Get("Restore the paused provider, model, effort and mode before resuming.")); return; }
            try { EnsureCurrentScope(); }
            catch (Exception error) { SetStatus(error.Message); return; }
            activeTurnId = currentSession.PausedTurnId; stopRequested = false; tools.Mode = currentSession.Mode;
            tools.CurrentProviderName = provider.Name; providerStreamId = null;
            var first = messages.Count == 0 ? null : json.DeserializeObject(json.Serialize(messages[0])) as IDictionary<string, object>;
            if (first != null && first.TryGetValue("role", out var role) && Convert.ToString(role) == "system") messages.RemoveAt(0);
            messages.Insert(0, new { role = "system", content = LlmVbeContext.DeveloperInstructions });
            int previousChanges = codeChanges.Count;
            currentSession.BudgetPaused = false; SetBusy(true);
            bool completed = false;
            try { completed = await RunHttpBudgetAsync(provider, model.Id); }
            catch (Exception error)
            {
                CompletePendingToolResponses(); currentSession.BudgetPaused = true;
                Append("Assistant", UiText.Get("Continuation interrupted. Completed results were kept; verify live state before retrying.") + "\n" + error.Message);
                SetStatus(UiText.Get("Paused — resume when ready"));
            }
            finally
            {
                activeHttpClient = null;
                if (!IsDisposed)
                {
                    int applied = codeChanges.Skip(previousChanges).Count(c => c.TurnId == activeTurnId);
                    if (applied > 0) AddEntry(new ChatEntry
                    {
                        Speaker = "Intervention",
                        TurnId = activeTurnId,
                        Text = applied + UiText.Get(" change(s) applied. Review the files and undo this turn below.")
                    });
                    if (!currentSession.BudgetPaused && verifyAfterEdit.Checked && codeChanges.Any(c => c.TurnId == activeTurnId)) await VerifyProjectAsync();
                    activeTurnId = null; SetBusy(false); SaveCurrentSession();
                    await DispatchPendingAsync(completed);
                }
            }
        }

        /// <summary>Completes persisted tool actions before a paused provider turn resumes.</summary>
        private void CompletePendingToolResponses()
        {
            var records = messages.Select(m => json.DeserializeObject(json.Serialize(m)) as IDictionary<string, object>).Where(m => m != null).ToArray();
            var answered = new HashSet<string>(records.Where(m => m.ContainsKey("tool_call_id")).Select(m => Convert.ToString(m["tool_call_id"])));
            foreach (var record in records)
            {
                if (!record.TryGetValue("tool_calls", out var value) || !(value is object[] calls)) continue;
                foreach (var item in calls)
                    if (item is IDictionary<string, object> call && call.TryGetValue("id", out var id) && answered.Add(Convert.ToString(id)))
                        messages.Add(new { role = "tool", tool_call_id = Convert.ToString(id), content = json.Serialize(Response.Failure("Interrupted before a result was recorded. Execution is unconfirmed; inspect live state and do not replay automatically.")) });
            }
        }

        /// <summary>Runs the HTTP provider turn while applying its configured token and time budgets.</summary>
        /// <param name="provider">HTTP provider used for every continuation round.</param>
        /// <param name="model">Model ID frozen for this provider turn.</param>
        /// <returns>True after a final assistant response; false when the configured loop budget pauses the turn.</returns>
        private async Task<bool> RunHttpBudgetAsync(LlmProvider provider, string model)
        {
            using (var client = new LlmChatClient(provider, settings,
                model, HttpHandlerOverride?.Invoke(),
                (modelPicker.SelectedItem as LlmModelOption)?.Id == model
                    ? ((LlmModelOption)modelPicker.SelectedItem).Capabilities : LlmModelCapabilities.Unknown))
            {
                activeHttpClient = client;
                var ownerSession = currentSession;
                string ownerTurn = activeTurnId;
                string activityScope = "provider:" + ownerTurn + ":" + Guid.NewGuid().ToString("N");
                bool acceptingActivities = true;
                client.ActivityUpdate = activity =>
                {
                    if (acceptingActivities && currentSession == ownerSession && activeTurnId == ownerTurn && !IsDisposed)
                        ReceiveAgentActivity(ProviderActivityProjection.WithScope(activity, activityScope));
                };
                try
                {
                    client.ToolHandler = async (name, arguments) =>
                    {
                        if (stopRequested) throw new OperationCanceledException();
                        return await ExecuteBudgetTool(name, arguments);
                    };
                    SetStatus(client.DisplayName + UiText.Get(" — working"));
                    var observedResults = new HashSet<string>(StringComparer.Ordinal);
                    int stalledRounds = 0;
                    // Eight rounds without new successful tool results are a fallback; progressing work continues.
                    // A separate ceiling keeps even continuously changing tool loops bounded.
                    for (int turn = 0; turn < 64; turn++)
                    {
                        if (stopRequested) throw new OperationCanceledException();
                        providerStreamId = "http-" + Guid.NewGuid().ToString("N");
                        string streamId = providerStreamId;
                        bool receivedText = false;
                        client.TextDelta = fragment =>
                        {
                            if (!acceptingActivities || currentSession != ownerSession || activeTurnId != ownerTurn ||
                                providerStreamId != streamId || stopRequested || IsDisposed) return;
                            receivedText = true;
                            ReceiveChatUpdate("final", streamId, fragment, false);
                        };
                        IDictionary<string, object> message;
                        lastHttpStreamDiagnostics = null;
                        try { message = await client.CompleteAsync(messages, tools.CatalogForProvider()); }
                        finally { lastHttpStreamDiagnostics = client.LastStreamDiagnostics; }
                        if (stopRequested) throw new OperationCanceledException();
                        if (receivedText) ReceiveChatUpdate("final", streamId, Convert.ToString(message["content"]), true);
                        providerStreamId = null;
                        messages.Add(message);
                        var calls = message.TryGetValue("tool_calls", out object rawCalls) ? rawCalls as object[] : null;
                        if (calls == null || calls.Length == 0)
                        {
                            string answer = message.ContainsKey("content") ? Convert.ToString(message["content"]) : "";
                            CompleteAssistantResponse(string.IsNullOrWhiteSpace(answer) ? UiText.Get("No text response.") : answer);
                            currentSession.ResumeContext = null;
                            currentSession.BudgetPaused = false;
                            SetStatus(client.DisplayName + UiText.Get(" — ready"));
                            return true;
                        }
                        bool progressed = false;
                        foreach (object rawCall in calls)
                        {
                            if (stopRequested) throw new OperationCanceledException();
                            var call = rawCall as IDictionary<string, object>;
                            var function = call != null && call.ContainsKey("function") ? call["function"] as IDictionary<string, object> : null;
                            if (function == null || !call.ContainsKey("id")) throw new InvalidOperationException("Invalid tool call.");
                            string name = Convert.ToString(function["name"]);
                            string arguments = Convert.ToString(function["arguments"]);
                            string callId = Convert.ToString(call["id"]);
                            bool recorded = messages.Select(m => json.DeserializeObject(json.Serialize(m)) as IDictionary<string, object>)
                                .Any(m => m != null && m.TryGetValue("tool_call_id", out var existingId) && Convert.ToString(existingId) == callId);
                            string result = recorded
                                ? json.Serialize(Response.Failure("This tool call ID already has a recorded result. No action was replayed. Read the prior result and live state before proposing a new action."))
                                : await ExecuteProjectedBudgetTool(name, arguments, activityScope + ":tool:" + callId);
                            if (recorded) ReceiveAgentActivity(ProviderActivityProjection.Tool(activityScope + ":tool:" + callId + ":duplicate:" + turn, name, arguments, "declined", ProviderActivityProjection.ToolDiagnostic(result)));
                            if (!recorded)
                            {
                                try
                                {
                                    if (json.Deserialize<Response>(result)?.Ok == true && observedResults.Add(EditorDocument.Hash(name + "\n" + arguments + "\n" + result))) progressed = true;
                                }
                                catch { }
                            }
                            messages.Add(new { role = "tool", tool_call_id = Convert.ToString(call["id"]), content = result });
                            SaveCurrentSession();
                        }
                        stalledRounds = progressed ? 0 : stalledRounds + 1;
                        if (stalledRounds >= 8) { PauseBudget(provider, model); return false; }
                    }
                    PauseBudget(provider, model); return false;
                }
                finally { acceptingActivities = false; }
            }
        }
    }
}

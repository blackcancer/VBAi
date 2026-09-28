using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;

namespace CodexVBE
{
    /// <summary>Provides the chat window implementation.</summary>
internal sealed partial class ChatWindow
    {
        /// <summary>Stores the provider stream id used by ChatWindow.</summary>
private string providerStreamId;
        /// <summary>Performs the execute budget tool operation for ChatWindow.</summary>
/// <param name="name">Text containing the name.</param>
/// <param name="arguments">Text containing the arguments.</param>
/// <returns>The result produced by this operation.</returns>
private async Task<string> ExecuteBudgetTool(string name, string arguments)
        {
            string label = name;
            if (name == "invoke_tool")
            {
                var gateway = json.DeserializeObject(arguments) as IDictionary<string, object>;
                if (gateway != null && gateway.TryGetValue("ToolName", out var target)) label = Convert.ToString(target);
            }
            try
            {
                string result = await InvokeTool(tools, name, arguments);
                bool success = false;
                try { success = json.Deserialize<Response>(result)?.Ok == true; } catch { }
                if (currentSession.CompletedToolActions == null) currentSession.CompletedToolActions = new List<string>();
                currentSession.CompletedToolActions.Add(label + UiText.Get(success ? " — response received; inspect returned state" : " — refused or failed"));
                return result;
            }
            catch
            {
                if (currentSession.CompletedToolActions == null) currentSession.CompletedToolActions = new List<string>();
                currentSession.CompletedToolActions.Add(label + UiText.Get(" — interrupted; verify live state before retrying"));
                throw;
            }
        }
        /// <summary>Performs the pause budget operation for ChatWindow.</summary>
/// <param name="provider">The provider used by this operation.</param>
/// <param name="model">Text containing the model.</param>
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
            AddEntry(new ChatEntry { Speaker = "Assistant", TurnId = activeTurnId,
                Text = UiText.Get("Safety pause: repeated rounds without progress or the intervention ceiling was reached. Resume from saved results; completed actions will not be replayed.") + "\n" + remaining + "\n\n" + actions });
            SetStatus(UiText.Get("Paused — resume when ready"));
        }
        /// <summary>Performs the update budget controls operation for ChatWindow.</summary>
private void UpdateBudgetControls()
        {
            if (resumeTurn != null) resumeTurn.Enabled = !busy && currentSession?.BudgetPaused == true;
            if (send == null || prompt == null) return;
            bool hasText = !string.IsNullOrWhiteSpace(prompt.Text);
            send.Text = busy ? UiText.Get(hasText ? "Queue ↑" : "Stop ■") :
                UiText.Get(currentSession?.BudgetPaused == true && !hasText ? "Resume ▶" : "Send ↑");
            toolTips.SetToolTip(send, UiText.Get(busy ? (hasText ? "Queue this message after the current response." : "Stop the current response. Changes already applied can still be undone in the chat.") : "Send the message and its context to the agent."));
            send.Enabled = !busy || !stopRequested || hasText;
        }
        /// <summary>Performs the resume budget async operation for ChatWindow.</summary>
/// <returns>The result produced by this operation.</returns>
private async Task ResumeBudgetAsync()
        {
            if (busy || currentSession?.BudgetPaused != true) return;
            var provider = providerPicker.SelectedItem as LlmProvider;
            var model = modelPicker.SelectedItem as LlmModelOption;
            if (provider == null || model == null || provider.IsCodex || provider.Name != currentSession.PausedProvider || model.Id != currentSession.PausedModel ||
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
                    if (applied > 0) AddEntry(new ChatEntry { Speaker = "Intervention", TurnId = activeTurnId,
                        Text = applied + UiText.Get(" change(s) applied. Review the files and undo this turn below.") });
                    if (!currentSession.BudgetPaused && verifyAfterEdit.Checked && codeChanges.Any(c => c.TurnId == activeTurnId)) await VerifyProjectAsync();
                    activeTurnId = null; SetBusy(false); SaveCurrentSession();
                    await DispatchPendingAsync(completed);
                }
            }
        }
        /// <summary>Performs the complete pending tool responses operation for ChatWindow.</summary>
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
        /// <summary>Performs the run http budget async operation for ChatWindow.</summary>
/// <param name="provider">The provider used by this operation.</param>
/// <param name="model">Text containing the model.</param>
/// <returns>The result produced by this operation.</returns>
private async Task<bool> RunHttpBudgetAsync(LlmProvider provider, string model)
        {
                using (var client = new LlmChatClient(provider, settings,
                    model, HttpHandlerOverride?.Invoke()))
                {
                    activeHttpClient = client;
                    client.ToolHandler = async (name, arguments) =>
                    {
                        if (stopRequested) throw new OperationCanceledException();
                        Append("Outil", name);
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
                            if (stopRequested || IsDisposed) return;
                            receivedText = true;
                            ReceiveChatUpdate("final", streamId, fragment, false);
                        };
                        var message = await client.CompleteAsync(messages, tools.CatalogForProvider());
                        if (stopRequested) throw new OperationCanceledException();
                        if (receivedText) ReceiveChatUpdate("final", streamId, Convert.ToString(message["content"]), true);
                        providerStreamId = null;
                        messages.Add(message);
                        object rawCalls;
                        var calls = message.TryGetValue("tool_calls", out rawCalls) ? rawCalls as object[] : null;
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
                            Append("Outil", name);
                            string callId = Convert.ToString(call["id"]);
                            bool recorded = messages.Select(m => json.DeserializeObject(json.Serialize(m)) as IDictionary<string, object>)
                                .Any(m => m != null && m.TryGetValue("tool_call_id", out var existingId) && Convert.ToString(existingId) == callId);
                            string result = recorded
                                ? json.Serialize(Response.Failure("This tool call ID already has a recorded result. No action was replayed. Read the prior result and live state before proposing a new action."))
                                : await ExecuteBudgetTool(name, arguments);
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
        }
    }
}

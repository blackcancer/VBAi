using System;
using System.Collections.Generic;
using System.IO;
using System.Text;
using System.Web.Script.Serialization;

namespace VBAi.Tests.Integration
{
    /// <summary>Publishes one immutable owner plan and never repeats an emitted qualification step.</summary>
    internal sealed class OwnerGitQualificationScope : IDisposable
    {
        private readonly string previousEnvironment;
        private readonly HashSet<string> claimed = new HashSet<string>(StringComparer.Ordinal);
        private bool published, disposed;
        private string[] orderedIds;
        private int nextClaim;
        internal string EvidenceRoot { get; }
        internal string ManifestPath { get; }
        internal IList<string> TerminalReceipts { get; } = new List<string>();

        internal OwnerGitQualificationScope(string evidenceRoot)
        {
            if (string.IsNullOrWhiteSpace(evidenceRoot) || !Path.IsPathRooted(evidenceRoot))
                throw new ArgumentException("An absolute disposable evidence root is required.");
            EvidenceRoot = Path.GetFullPath(evidenceRoot);
            if (!Directory.Exists(EvidenceRoot)) throw new DirectoryNotFoundException(EvidenceRoot);
            previousEnvironment = Environment.GetEnvironmentVariable(OwnerGitQualificationManifest.EnvironmentName);
            if (!string.IsNullOrEmpty(previousEnvironment))
                throw new InvalidOperationException("An owner Git qualification is already configured; no overlapping scope.");
            ManifestPath = Path.Combine(EvidenceRoot, Guid.NewGuid().ToString("N") + ".owner-git.json");
            Environment.SetEnvironmentVariable(OwnerGitQualificationManifest.EnvironmentName, ManifestPath);
        }

        internal static OwnerGitQualificationStep Step(string verb, string expectedDirectory, string targetDirectory,
            string checkpoint = null, string expectedError = null)
        {
            return new OwnerGitQualificationStep {
                Id = Guid.NewGuid().ToString("N"), Verb = verb,
                ExpectedSnapshotDirectory = Path.GetFullPath(expectedDirectory),
                ExpectedSnapshotSha256 = OwnerGitQualificationManifest.SnapshotDirectoryHash(expectedDirectory),
                TargetSnapshotDirectory = Path.GetFullPath(targetDirectory),
                TargetSnapshotSha256 = OwnerGitQualificationManifest.SnapshotDirectoryHash(targetDirectory),
                CheckpointId = checkpoint, ExpectedErrorSubstring = expectedError
            };
        }

        internal void Publish(ExcelVbeFixture host, string workbookPath, string repositoryChild, string branch,
            OwnerGitQualificationStep[] steps, string remote = null, string remoteCommit = null)
        {
            Publish(host.OwnerGitIdentity(workbookPath, EvidenceRoot, repositoryChild, branch, steps, remote, remoteCommit));
        }

        internal void Publish(OwnerGitQualificationManifest manifest)
        {
            if (published || disposed) throw new InvalidOperationException("One immutable owner Git plan only.");
            if (!string.Equals(Environment.GetEnvironmentVariable(OwnerGitQualificationManifest.EnvironmentName), ManifestPath, StringComparison.Ordinal))
                throw new InvalidOperationException("The host opt-in was replaced concurrently.");
            if (manifest == null || !string.Equals(manifest.EvidenceRoot, EvidenceRoot, StringComparison.Ordinal))
                throw new InvalidOperationException("The plan does not own this evidence root.");
            string json = new JavaScriptSerializer().Serialize(manifest);
            OwnerGitQualificationManifest.Parse(json);
            using (var stream = new FileStream(ManifestPath, FileMode.CreateNew, FileAccess.Write, FileShare.Read))
            {
                byte[] bytes = new UTF8Encoding(false, true).GetBytes(json);
                stream.Write(bytes, 0, bytes.Length);
            }
            published = true;
            orderedIds = Array.ConvertAll(manifest.Steps, step => step.Id);
        }

        internal IDictionary<string, object> Execute(ExcelVbeFixture host, OwnerGitQualificationStep step, string revision)
        {
            Claim(step.Id, revision);
            string terminalPath = Path.Combine(EvidenceRoot, "owner-git-" + step.Id + ".terminal.json");
            IDictionary<string, object> reply;
            try { reply = host.Command(new { Command = OwnerGitQualificationManifest.CommandName, Action = step.Id, ExpectedSha256 = revision }); }
            catch { host.PreserveForDiagnosticRecovery = true; throw; }
            IDictionary<string, object> terminal = null;
            bool classified = false;
            try
            {
                if (System.IO.File.Exists(terminalPath))
                {
                    terminal = new JavaScriptSerializer().DeserializeObject(System.IO.File.ReadAllText(terminalPath)) as IDictionary<string, object>;
                    TerminalReceipts.Add(terminalPath);
                }
                bool retain = MustRetain(reply, terminal);
                if (retain) host.PreserveForDiagnosticRecovery = true;
                ValidateTerminal(step.Id, terminal);
                classified = true;
                if (retain && reply != null && reply.ContainsKey("Ok") && Equals(reply["Ok"], true))
                    throw new IOException("A successful bridge response has an uncertain durable outcome.");
                if (reply == null) throw new IOException("The owner step has no terminal bridge response; it was not repeated.");
                if (!reply.ContainsKey("Ok") || !Equals(reply["Ok"], true))
                    throw new InvalidOperationException(Convert.ToString(reply.ContainsKey("Error") ? reply["Error"] : "Owner Git step failed without a classified response."));
                if (terminal == null) throw new IOException("An emitted owner step has no durable terminal receipt.");
                return ValidateSuccess(step.Id, terminalPath, reply);
            }
            catch (Exception error)
            {
                if (!classified || error is IOException) host.PreserveForDiagnosticRecovery = true;
                throw;
            }
        }

        internal void Claim(string id, string revision)
        {
            Guid value;
            if (!published || disposed || !Guid.TryParseExact(id, "N", out value) || !OwnerGitQualificationManifest.Sha(revision) ||
                nextClaim >= orderedIds.Length || orderedIds[nextClaim] != id)
                throw new InvalidOperationException("A published plan, exact step id and revision are required before emission.");
            if (!claimed.Add(id)) throw new InvalidOperationException("An owner Git step was already emitted; no retry.");
            nextClaim++;
        }

        /// <summary>Only a durable classified prewrite refusal permits normal fixture cleanup after failure.</summary>
        internal static bool MustRetain(IDictionary<string, object> reply, IDictionary<string, object> terminal)
        {
            if (reply == null || !reply.ContainsKey("Ok") || !(reply["Ok"] is bool) ||
                terminal == null || !terminal.ContainsKey("Outcome") || !terminal.ContainsKey("MutationStarted")) return true;
            string outcome = Convert.ToString(terminal["Outcome"]);
            bool ok = reply.ContainsKey("Ok") && Equals(reply["Ok"], true);
            if (ok) return outcome != "Succeeded" || !Equals(terminal["MutationStarted"], true);
            return !Equals(terminal["MutationStarted"], false) ||
                !terminal.ContainsKey("RecoveryPending") || !Equals(terminal["RecoveryPending"], false) ||
                (outcome != "FailedBeforeMutation" && outcome != "ExpectedPrewriteRefusal");
        }

        internal static void ValidateTerminal(string stepId, IDictionary<string, object> terminal)
        {
            if (terminal == null || !terminal.ContainsKey("StepId") || !Equals(terminal["StepId"], stepId) ||
                !terminal.ContainsKey("Outcome") || !(terminal["Outcome"] is string) ||
                !terminal.ContainsKey("MutationStarted") || !(terminal["MutationStarted"] is bool))
                throw new IOException("The emitted step has no matching durable classification.");
        }

        internal static IDictionary<string, object> ValidateSuccess(string stepId, string receiptPath, IDictionary<string, object> reply)
        {
            var data = reply != null && reply.ContainsKey("Data") ? reply["Data"] as IDictionary<string, object> : null;
            if (data == null || !data.ContainsKey("StepId") || !Equals(data["StepId"], stepId) ||
                !data.ContainsKey("Outcome") || !Equals(data["Outcome"], "Succeeded") ||
                !data.ContainsKey("ReceiptPath") || !Equals(data["ReceiptPath"], receiptPath) ||
                !data.ContainsKey("Result") || !(data["Result"] is IDictionary<string, object>))
                throw new IOException("The owner response does not match the emitted step and receipt.");
            return data;
        }

        public void Dispose()
        {
            if (disposed) return;
            disposed = true;
            if (!string.Equals(Environment.GetEnvironmentVariable(OwnerGitQualificationManifest.EnvironmentName), ManifestPath, StringComparison.Ordinal))
                throw new InvalidOperationException("Concurrent owner Git configuration changed; it was preserved.");
            Environment.SetEnvironmentVariable(OwnerGitQualificationManifest.EnvironmentName, previousEnvironment);
        }
    }
}

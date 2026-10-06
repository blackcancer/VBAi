using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Threading;
using System.Web.Script.Serialization;

namespace VBAi
{

    /// <summary>Persists LLM settings with optimistic field merging and a cross-process write lock.</summary>
    internal sealed partial class LlmSettings
    {

        /// <summary>Deep JSON snapshot loaded before local edits, used as the merge base.</summary>
        private IDictionary<string, object> baseline;

        /// <summary>Canonical settings path associated with the remembered baseline.</summary>
        private string baselinePath;

        /// <summary>Whether the settings file existed when the current baseline was captured.</summary>
        private bool baselineExisted;

        /// <summary>Reads the legacy approval default without altering encrypted values.</summary>
        /// <param name="content">Serialized settings JSON to deserialize.</param>
        /// <returns>Settings object with the legacy VBE approval default supplied when that field is absent.</returns>
        private static LlmSettings DecodeSettings(string content)
        {
            var serializer = new JavaScriptSerializer();
            var fields = serializer.DeserializeObject(content) as IDictionary<string, object>;
            var result = serializer.Deserialize<LlmSettings>(content) ?? new LlmSettings();
            if (fields == null || !fields.ContainsKey(nameof(VbeEditApproval))) result.VbeEditApproval = "AskEachTime";
            return result;
        }

        /// <summary>Creates a detached JSON-shaped copy of the current public settings fields.</summary>
        /// <returns>Dictionary snapshot used for change detection and merging.</returns>
        private IDictionary<string, object> Snapshot()
        {
            var serializer = new JavaScriptSerializer();
            return (IDictionary<string, object>)serializer.DeserializeObject(serializer.Serialize(this));
        }

        /// <summary>Replaces the merge base after a successful load or atomic save.</summary>
        /// <param name="path">Canonical file path associated with this baseline.</param>
        /// <param name="exists">Whether that file existed at the time the baseline was captured.</param>
        private void RememberBaseline(string path, bool exists)
        {
            baseline = Snapshot();
            baselinePath = path;
            baselineExisted = exists;
        }

        /// <summary>Merges local changes against the loaded snapshot under a cross-process file lock.</summary>
        private void SaveMerged()
        {
            string path = Path.GetFullPath(FilePath);
            if (baselinePath != null && !string.Equals(path, baselinePath, StringComparison.OrdinalIgnoreCase))
                throw new IOException("The settings location changed. Reload settings before saving.");
            Directory.CreateDirectory(Path.GetDirectoryName(path));
            using (AcquireWriteLock(path))
            {
                bool exists = File.Exists(path);
                if (baselineExisted && !exists)
                    throw new IOException("The settings file was removed. Reload settings before saving.");
                if (baseline == null && exists)
                    throw new IOException("Load the existing settings before saving changes.");

                var serializer = new JavaScriptSerializer();
                var desired = Snapshot();
                IDictionary<string, object> merged = desired;
                if (exists)
                {
                    string content = File.ReadAllText(path, Encoding.UTF8);
                    // Normalized known fields preserve legacy defaults; unknown fields survive the rewrite.
                    merged = serializer.DeserializeObject(content) as IDictionary<string, object>
                        ?? new Dictionary<string, object>();
                    var current = DecodeSettings(content).Snapshot();
                    foreach (var key in desired.Keys)
                    {
                        baseline.TryGetValue(key, out var before);
                        current.TryGetValue(key, out var stored);
                        merged[key] = MergeValue(before, desired[key], stored, key);
                    }
                }

                string payload = serializer.Serialize(merged);
                var saved = DecodeSettings(payload);
                UpdatePaths.WriteAtomic(path, payload);
                // Keep this long-lived instance current, so later edits do not revert merged fields.
                foreach (var property in typeof(LlmSettings).GetProperties(BindingFlags.Instance | BindingFlags.Public))
                    if (property.CanRead && property.CanWrite) property.SetValue(this, property.GetValue(saved, null), null);
                RememberBaseline(path, true);
                ApplyProviderLabels();
            }
        }

        /// <summary>Merges a field when only one side changed from baseline, recursively merging JSON objects.</summary>
        /// <param name="before">Value in the baseline loaded before local edits.</param>
        /// <param name="desired">Current in-memory value to save.</param>
        /// <param name="stored">Value most recently read from disk under the write lock.</param>
        /// <param name="field">Setting name included in a conflict error; values are never included.</param>
        /// <returns>Nonconflicting merged value; throws when both local and external changes disagree.</returns>
        private static object MergeValue(object before, object desired, object stored, string field)
        {
            if (ValuesEqual(before, desired)) return stored;
            if (ValuesEqual(before, stored) || ValuesEqual(desired, stored)) return desired;
            var oldMap = before as IDictionary<string, object>;
            var newMap = desired as IDictionary<string, object>;
            var diskMap = stored as IDictionary<string, object>;
            if (oldMap != null && newMap != null && diskMap != null)
            {
                var result = new Dictionary<string, object>(diskMap);
                foreach (string key in oldMap.Keys.Union(newMap.Keys))
                {
                    bool oldPresent = oldMap.TryGetValue(key, out var oldValue);
                    bool newPresent = newMap.TryGetValue(key, out var newValue);
                    bool diskPresent = diskMap.TryGetValue(key, out var diskValue);
                    if (oldPresent == newPresent && ValuesEqual(oldValue, newValue)) continue;
                    if (diskPresent != oldPresent || !ValuesEqual(diskValue, oldValue))
                    {
                        if (diskPresent == newPresent && ValuesEqual(diskValue, newValue)) continue;
                        throw SettingsConflict(field);
                    }
                    if (newPresent) result[key] = newValue; else result.Remove(key);
                }
                return result;
            }
            throw SettingsConflict(field);
        }

        /// <summary>Compares scalar values or recursively compares JSON object fields.</summary>
        /// <param name="left">First scalar or dictionary value.</param>
        /// <param name="right">Second scalar or dictionary value.</param>
        /// <returns>True when scalar equality or all nested dictionary keys and values match.</returns>
        private static bool ValuesEqual(object left, object right)
        {
            var a = left as IDictionary<string, object>;
            var b = right as IDictionary<string, object>;
            if (a != null && b != null)
                return a.Count == b.Count && a.All(pair => b.TryGetValue(pair.Key, out var value) && ValuesEqual(pair.Value, value));
            return Equals(left, right);
        }

        /// <summary>Sets tings conflict for llm settings.</summary>
        /// <param name="field">Conflicting setting key; its value is deliberately omitted to protect credentials and endpoints.</param>
        /// <returns>io exception produced by the operation for settings conflict on llm settings.</returns>
        private static IOException SettingsConflict(string field)
        {
            // Never include values: the field may contain a credential or a private endpoint.
            return new IOException("Settings changed in another host (" + field + "). Reopen settings and retry.");
        }

        /// <summary>Acquires an exclusive lock file, retrying briefly while another process holds it.</summary>
        /// <param name="path">Canonical settings path; the lock uses the adjacent <c>.lock</c> filename.</param>
        /// <returns>Open exclusive FileStream whose disposal releases the cross-process lock.</returns>
        private static FileStream AcquireWriteLock(string path)
        {
            var elapsed = Stopwatch.StartNew();
            while (true)
            {
                try { return new FileStream(path + ".lock", FileMode.OpenOrCreate, FileAccess.ReadWrite, FileShare.None); }
                catch (IOException) when (elapsed.ElapsedMilliseconds < 500) { Thread.Sleep(15); }
            }
        }
    }
}

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
    internal sealed partial class LlmSettings
    {
        private IDictionary<string, object> baseline;
        private string baselinePath;
        private bool baselineExisted;

        /// <summary>Reads the legacy approval default without altering encrypted values.</summary>
        private static LlmSettings DecodeSettings(string content)
        {
            var serializer = new JavaScriptSerializer();
            var fields = serializer.DeserializeObject(content) as IDictionary<string, object>;
            var result = serializer.Deserialize<LlmSettings>(content) ?? new LlmSettings();
            if (fields == null || !fields.ContainsKey(nameof(VbeEditApproval))) result.VbeEditApproval = "AskEachTime";
            return result;
        }

        private IDictionary<string, object> Snapshot()
        {
            var serializer = new JavaScriptSerializer();
            return (IDictionary<string, object>)serializer.DeserializeObject(serializer.Serialize(this));
        }

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

        private static bool ValuesEqual(object left, object right)
        {
            var a = left as IDictionary<string, object>;
            var b = right as IDictionary<string, object>;
            if (a != null && b != null)
                return a.Count == b.Count && a.All(pair => b.TryGetValue(pair.Key, out var value) && ValuesEqual(pair.Value, value));
            return Equals(left, right);
        }

        private static IOException SettingsConflict(string field)
        {
            // Never include values: the field may contain a credential or a private endpoint.
            return new IOException("Settings changed in another host (" + field + "). Reopen settings and retry.");
        }

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

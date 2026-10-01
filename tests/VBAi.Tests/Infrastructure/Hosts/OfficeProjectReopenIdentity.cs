using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace VBAi.Tests.Integration
{
    /// <summary>Checks document/project identity across a fresh native reopen, independently of selector canonicalization.</summary>
    internal static class OfficeProjectReopenIdentity
    {
        /// <summary>Contains terminal read-only observations from the already bound owned fixture.</summary>
        internal sealed class Evidence
        {
            public string Host { get; set; }
            public string DocumentPath { get; set; }
            public string Selector { get; set; }
            public int ProcessId { get; set; }
            public string[] Metadata { get; set; }
            public IDictionary<string, object> Observation { get; set; }
            public IDictionary<string, object> Status { get; set; }
        }

        /// <summary>Accepts a name/path selector change only with stable metadata and exact native host/document proof.</summary>
        internal static void Require(string host, string path, string previousSelector, int previousPid,
            string[] expectedMetadata, Guid expectedMvid, Evidence actual)
        {
            Assert.IsNotNull(actual);
            Assert.AreEqual(host, actual.Host, "The reopened fixture changed host kind.");
            Assert.IsTrue(previousPid > 0 && actual.ProcessId > 0, "Both owned process IDs are required.");
            if (host == "Access" || host == "Publisher")
                Assert.AreNotEqual(previousPid, actual.ProcessId, "These adapters require an independent fresh process.");
            RequirePath(path, actual.DocumentPath);
            string name = MetadataName(expectedMetadata), currentName = MetadataName(actual.Metadata);
            Assert.AreEqual(name, currentName, "The persisted native project Name changed.");
            RequireSelector(previousSelector, name, path);
            RequireSelector(actual.Selector, name, path);
            Assert.IsNotNull(actual.Status);
            Assert.AreEqual(expectedMvid.ToString("D"), Field(actual.Status, "AssemblyModuleVersionId"));
            RequirePid(actual.ProcessId, Field(actual.Status, "HostProcessId"));
            Assert.IsNotNull(actual.Observation);
            Assert.IsNull(Field(actual.Observation, "ObservationError"), "Fresh native observation failed.");
            RequirePid(actual.ProcessId, Field(actual.Observation, "ProcessId"));
            RequirePid(actual.ProcessId, Field(actual.Observation, "ApplicationOwnerPid"));
            RequirePath(path, Field(actual.Observation, "DocumentPath") as string);
            var persistence = Field(actual.Observation, "Persistence") as IDictionary<string, object>;
            Assert.IsNotNull(persistence, "Native persistence identity evidence is required.");
            Assert.AreEqual(host, Field(persistence, "Host"));
            Assert.AreEqual(true, Field(persistence, "HostAvailable"));
            Assert.AreEqual(true, Field(persistence, "IdentityVerified"));
            RequirePid(actual.ProcessId, Field(persistence, "OwnerProcessId"));
            RequirePath(path, Field(persistence, "HostPath") as string);
            Assert.AreEqual(actual.Selector, Field(persistence, "Project"), "Native proof targeted another selector.");
        }

        private static string MetadataName(string[] metadata)
        {
            Assert.IsNotNull(metadata);
            var names = metadata.Where(value => value != null && value.StartsWith("Name=", StringComparison.Ordinal)).ToArray();
            Assert.AreEqual(1, names.Length, "Exactly one native metadata Name is required.");
            string name = names[0].Substring("Name=".Length);
            Assert.IsFalse(string.IsNullOrWhiteSpace(name), "A blank native project Name is not identity proof.");
            return name;
        }

        private static void RequireSelector(string selector, string name, string path)
        {
            if (string.Equals(name, selector, StringComparison.Ordinal)) return;
            RequirePath(path, selector);
        }

        private static void RequirePath(string expected, string actual)
        {
            Assert.IsFalse(string.IsNullOrWhiteSpace(expected));
            Assert.IsFalse(string.IsNullOrWhiteSpace(actual));
            Assert.IsTrue(Path.IsPathRooted(expected) && Path.IsPathRooted(actual), "Identity paths must be absolute.");
            Assert.AreEqual(Path.GetFullPath(expected), Path.GetFullPath(actual), true, "The reopened native document is not the disposable file.");
        }

        private static void RequirePid(int expected, object actual)
        {
            Assert.IsNotNull(actual, "A missing PID is not ownership proof.");
            Assert.AreEqual(expected, Convert.ToInt32(actual), "Native evidence belongs to another host process.");
        }

        private static object Field(IDictionary<string, object> value, string key)
        { object result; return value.TryGetValue(key, out result) ? result : null; }
    }
}

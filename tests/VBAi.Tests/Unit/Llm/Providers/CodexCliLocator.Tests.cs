using Microsoft.VisualStudio.TestTools.UnitTesting;
using System;
using System.IO;

namespace VBAi.Tests.Unit
{
    [TestClass, TestCategory("Unit")]
    public sealed class CodexCliLocatorTests
    {
        [TestMethod]
        public void LocatorRespectsOverrideLegacyVersionsAndDeterministicTimestampTies()
        {
            string previous = Environment.GetEnvironmentVariable("VBAi_CODEX_CLI");
            try
            {
                Environment.SetEnvironmentVariable("VBAi_CODEX_CLI", "configured.exe");
                Assert.AreEqual("configured.exe", CodexCliLocator.Resolve(_ => throw new Exception("not needed"), _ => throw new Exception("not needed"), _ => throw new Exception("not needed")));
                Environment.SetEnvironmentVariable("VBAi_CODEX_CLI", " ");
                string legacy = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData), "Programs", "OpenAI", "Codex", "bin", "codex.exe");
                Assert.AreEqual(legacy, CodexCliLocator.Resolve(_ => true, _ => throw new Exception("not needed"), _ => throw new Exception("not needed")));
                string first = Path.Combine("a", "codex.exe"), last = Path.Combine("z", "codex.exe");
                Assert.AreEqual(first, CodexCliLocator.Resolve(p => p != legacy && p != Path.Combine("missing", "codex.exe"), _ => new[] { "z", "missing", "a" }, _ => new DateTime(2026, 9, 28)));
                Assert.AreEqual(last, CodexCliLocator.Resolve(p => p != legacy, _ => new[] { "a", "z" }, p => p == last ? DateTime.MaxValue : DateTime.MinValue));
                Assert.AreEqual("codex.exe", CodexCliLocator.Resolve(_ => false, _ => new string[0], _ => DateTime.MinValue));
                Assert.AreEqual("codex.exe", CodexCliLocator.Resolve(_ => false, _ => throw new IOException("missing directory"), _ => DateTime.MinValue));
                Assert.AreEqual("codex.exe", CodexCliLocator.Resolve(_ => false, _ => throw new UnauthorizedAccessException("protected directory"), _ => DateTime.MinValue));
                Assert.ThrowsException<InvalidOperationException>(() => CodexCliLocator.Resolve(_ => false, _ => throw new InvalidOperationException("unexpected failure"), _ => DateTime.MinValue));
            }
            finally { Environment.SetEnvironmentVariable("VBAi_CODEX_CLI", previous); }
        }
    }
}

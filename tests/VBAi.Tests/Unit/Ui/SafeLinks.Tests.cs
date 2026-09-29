using System;
using System.Diagnostics;
using VBAi;
using Microsoft.VisualStudio.TestTools.UnitTesting;
namespace VBAi.Tests.Unit
{
    /// <summary>Vérifie la validation des liens externes autorisés par l’interface.</summary>
    [TestClass]
    public sealed class SafeLinksTests
    {
        /// <summary>Accepte les liens Web sans identifiants et rejette credentials et schémas non Web.</summary>
        [TestMethod]
        public void HttpAndHttpsLinksRejectCredentialsAndNonWebSchemes()
        {
            foreach (var text in new[] { "https://example.invalid/path", "http://example.invalid", "HTTPS://example.invalid" }) Assert.IsTrue(SafeLinks.Allowed(text),text);
            foreach (var text in new[] { null, "", "invalid", "mailto:a@example.invalid", "file:///C:/tmp", "https://user:pass@example.invalid", "http://user@example.invalid" }) Assert.IsFalse(SafeLinks.Allowed(text),text);
            Assert.ThrowsException<ArgumentException>(() => SafeLinks.Open("file:///C:/tmp"));
            ProcessStartInfo observed = null;
            SafeLinks.Open("https://example.invalid/path", p => { observed=p; return null; });
            Assert.AreEqual("https://example.invalid/path",observed.FileName);
            Assert.IsTrue(observed.UseShellExecute);
        }
    }
}

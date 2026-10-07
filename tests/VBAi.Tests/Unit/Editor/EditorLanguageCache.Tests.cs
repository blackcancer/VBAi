using Microsoft.VisualStudio.TestTools.UnitTesting;
using System;
using System.Linq;
using System.Threading;

namespace VBAi.Tests.Unit.Editor
{
    [TestClass, TestCategory("Unit")]
    public sealed class EditorLanguageCacheTests
    {
        [TestMethod]
        public void OnlyChangedModulesAreRebuiltAndRemovedDeclarationsDisappear()
        {
            var cache = new EditorLanguageCache();
            var a = new EditorSource { Module = "A", ComponentType = 1, Text = "Public first As Long" };
            var b = new EditorSource { Module = "B", ComponentType = 2, Text = "Public second As String" };
            var original = cache.Build(new[] { a, b }, Array.Empty<string>(), CancellationToken.None);
            var warm = cache.Build(new[] { a, b }, Array.Empty<string>(), CancellationToken.None);
            Assert.AreEqual(original.Key, warm.Key); Assert.AreEqual(2, cache.ParsedModules);
            Assert.AreSame(original.Parts[0].Symbols, warm.Parts[0].Symbols);
            b.Text = "Public changed As String";
            var changed = cache.Build(new[] { a, b }, Array.Empty<string>(), CancellationToken.None);
            Assert.AreEqual(3, cache.ParsedModules); Assert.AreNotEqual(original.Key, changed.Key);
            Assert.AreEqual(original.Parts[0].Key, changed.Parts[0].Key);
            Assert.AreNotEqual(original.Parts[1].Key, changed.Parts[1].Key);
            Assert.IsFalse(changed.Symbols.Any(s => s.Name == "second"));
            var removed = cache.Build(new[] { a }, Array.Empty<string>(), CancellationToken.None);
            Assert.AreEqual(1, removed.Parts.Length); Assert.IsFalse(removed.Symbols.Any(s => s.Module == "B"));
            a.Module = "Renamed";
            var renamed = cache.Build(new[] { a }, Array.Empty<string>(), CancellationToken.None);
            Assert.IsTrue(renamed.Symbols.All(s => s.Module == "Renamed"));
        }
        [TestMethod]
        public void CancelledIndexingDoesNotParseAndTypeChangesInvalidateTheModule()
        {
            var cache = new EditorLanguageCache();
            var source = new EditorSource { Module = "SameName", ComponentType = 1, Text = "Public member As Long" };
            Assert.ThrowsException<OperationCanceledException>(() => cache.Build(new[] { source }, Array.Empty<string>(), new CancellationToken(true)));
            Assert.AreEqual(0, cache.ParsedModules);
            var module = cache.Build(new[] { source }, Array.Empty<string>(), CancellationToken.None);
            source.ComponentType = 2;
            var cls = cache.Build(new[] { source }, Array.Empty<string>(), CancellationToken.None);
            Assert.AreNotEqual(module.Key, cls.Key); Assert.AreEqual("Class", cls.Symbols[0].Kind);
        }
    }
}

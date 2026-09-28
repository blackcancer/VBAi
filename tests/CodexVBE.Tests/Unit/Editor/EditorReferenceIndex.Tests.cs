using System;
using System.IO;
using System.Linq;
using CodexVBE;
using Microsoft.VisualStudio.TestTools.UnitTesting;
namespace CodexVBE.Tests.Unit.Editor
{
    [TestClass, TestCategory("Unit")]
    public sealed class EditorReferenceIndexTests
    {
        [TestMethod]
        public void ReadsInstalledScriptingMetadataWithoutCreatingAutomationObjects()
        {
            string library = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.System), "scrrun.dll");
            Assert.IsTrue(File.Exists(library));
            var members = EditorReferenceIndex.Read(new[] { library }, new[] { "Scripting.Dictionary" });
            Assert.IsTrue(members.Any(s => s.Name == "Add" && s.Parameters.Length == 2));
            Assert.IsTrue(members.Any(s => s.Name == "Count" && s.Kind == "Property"));
            Assert.IsTrue(members.All(s => s.External && s.Module == "Dictionary"));
        }
    }
}

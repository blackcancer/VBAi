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
        public void FollowsReturnedObjectTypesAndKeepsLibraryIdentity()
        {
            string library = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.System), "scrrun.dll");
            var members = EditorReferenceIndex.Read(new[] { library }, new[] { "Scripting.FileSystemObject" });
            var file = members.First(s => s.Module == "FileSystemObject" && s.Name == "GetFile");
            Assert.AreEqual("Scripting.IFile", file.TypeName);
            Assert.IsTrue(members.Any(s => s.Module == "IFile" && s.Name == "OpenAsTextStream" && s.Library == "Scripting"));
            Assert.IsTrue(members.Any(s => s.Module == "ITextStream" && s.Name == "WriteLine"));
        }
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

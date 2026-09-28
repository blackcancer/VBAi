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
        [TestMethod]
        public void CacheMissingLibrariesAndNestedPointerArrayMetadataRemainReadOnly()
        {
            var paths = new[] { Path.Combine(Path.GetTempPath(), "owned-missing-" + Guid.NewGuid().ToString("N") + ".tlb") };
            var first = EditorReferenceIndex.Read(paths, new string[0]); Assert.AreEqual(0, first.Length); Assert.AreSame(first, EditorReferenceIndex.Read(paths, new string[0]));
            var scalar = new System.Runtime.InteropServices.ComTypes.TYPEDESC { vt = (short)System.Runtime.InteropServices.VarEnum.VT_I4 };
            IntPtr pointer = System.Runtime.InteropServices.Marshal.AllocHGlobal(System.Runtime.InteropServices.Marshal.SizeOf(scalar));
            try
            {
                System.Runtime.InteropServices.Marshal.StructureToPtr(scalar, pointer, false);
                var method = typeof(EditorReferenceIndex).GetMethod("ReturnType", System.Reflection.BindingFlags.NonPublic | System.Reflection.BindingFlags.Static);
                foreach (var kind in new[] { System.Runtime.InteropServices.VarEnum.VT_PTR, System.Runtime.InteropServices.VarEnum.VT_SAFEARRAY })
                    Assert.IsNull(method.Invoke(null, new object[] { null, new System.Runtime.InteropServices.ComTypes.TYPEDESC { vt = (short)kind, lpValue = pointer } }));
            }
            finally { System.Runtime.InteropServices.Marshal.FreeHGlobal(pointer); }
        }
        [TestMethod]
        public void NamelessNativeMetadataIsSkippedAndDescriptorsAreAlwaysReleased()
        {
            var info = new CodexVBE.Tests.Infrastructure.OwnedNamelessTypeInfo(); var symbols = new System.Collections.Generic.List<EditorSymbol>();
            typeof(EditorReferenceIndex).GetMethod("ReadMembers", System.Reflection.BindingFlags.NonPublic | System.Reflection.BindingFlags.Static)
                .Invoke(null, new object[] { info.Info, "Owned", symbols, new System.Collections.Generic.HashSet<Guid>(), 0 });
            Assert.AreEqual(0, symbols.Count); Assert.AreEqual(1, info.NamesRead); Assert.AreEqual(1, info.FunctionsReleased); Assert.AreEqual(1, info.TypesReleased);
        }
    }
}

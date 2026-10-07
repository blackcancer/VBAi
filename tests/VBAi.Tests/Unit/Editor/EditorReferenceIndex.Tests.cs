using Microsoft.VisualStudio.TestTools.UnitTesting;
using System;
using System.IO;
using System.Linq;
namespace VBAi.Tests.Unit.Editor
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
            Assert.IsTrue(members.All(s => s.External));
            Assert.IsTrue(members.Where(s => (s.Kind == "Property" || s.Kind == "Procedure") && !s.Global).All(s => s.Module == "Dictionary"));
            Assert.IsTrue(members.Any(s => s.Name == "Dictionary" && s.Kind == "Class"));
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
                    Assert.AreEqual("Long", method.Invoke(null, new object[] { null, new System.Runtime.InteropServices.ComTypes.TYPEDESC { vt = (short)kind, lpValue = pointer } }));
            }
            finally { System.Runtime.InteropServices.Marshal.FreeHGlobal(pointer); }
        }
        [TestMethod]
        public void NamelessNativeMetadataIsSkippedAndDescriptorsAreAlwaysReleased()
        {
            var info = new VBAi.Tests.Infrastructure.OwnedNamelessTypeInfo(); var symbols = new System.Collections.Generic.List<EditorSymbol>();
            typeof(EditorReferenceIndex).GetMethod("ReadMembers", System.Reflection.BindingFlags.NonPublic | System.Reflection.BindingFlags.Static)
                .Invoke(null, new object[] { info.Info, "Owned", symbols, new System.Collections.Generic.HashSet<Guid>(), 0 });
            Assert.AreEqual(0, symbols.Count); Assert.AreEqual(1, info.NamesRead); Assert.AreEqual(1, info.FunctionsReleased); Assert.AreEqual(1, info.TypesReleased);
        }
        private static EditorSymbol[] Metadata(VBAi.Tests.Infrastructure.OwnedReferenceMetadata fixture)
        {
            var symbols = new System.Collections.Generic.List<EditorSymbol>();
            typeof(EditorReferenceIndex).GetMethod("ReadMembers", System.Reflection.BindingFlags.NonPublic | System.Reflection.BindingFlags.Static)
                .Invoke(null, new object[] { fixture.Info, "Owned", symbols, new System.Collections.Generic.HashSet<Guid>(), 0 });
            return symbols.ToArray();
        }
        [TestMethod]
        public void ParametersReturnValuesOptionalFlagsAndNativeHelpProduceCompleteSignatures()
        {
            foreach (bool fallbackName in new[] { false, true })
                foreach (bool property in new[] { false, true })
                    using (var fixture = new VBAi.Tests.Infrastructure.OwnedReferenceMetadata())
                    {
                        fixture.NameCount = fallbackName ? 1 : 5;
                        fixture.Invocation = property ? System.Runtime.InteropServices.ComTypes.INVOKEKIND.INVOKE_PROPERTYGET : System.Runtime.InteropServices.ComTypes.INVOKEKIND.INVOKE_FUNC;
                        fixture.Kind = property ? System.Runtime.InteropServices.ComTypes.TYPEKIND.TKIND_MODULE : System.Runtime.InteropServices.ComTypes.TYPEKIND.TKIND_DISPATCH;
                        fixture.MemberId = property ? 0 : 1;
                        fixture.FunctionFlags = property ? (short)System.Runtime.InteropServices.ComTypes.FUNCFLAGS.FUNCFLAG_FDEFAULTBIND : (short)0;
                        fixture.Parameters = new[] {
                    VBAi.Tests.Infrastructure.OwnedReferenceMetadata.Parameter(System.Runtime.InteropServices.VarEnum.VT_I4, System.Runtime.InteropServices.ComTypes.PARAMFLAG.PARAMFLAG_FIN),
                    VBAi.Tests.Infrastructure.OwnedReferenceMetadata.Parameter(System.Runtime.InteropServices.VarEnum.VT_BSTR, System.Runtime.InteropServices.ComTypes.PARAMFLAG.PARAMFLAG_FOPT | System.Runtime.InteropServices.ComTypes.PARAMFLAG.PARAMFLAG_FOUT),
                    VBAi.Tests.Infrastructure.OwnedReferenceMetadata.Parameter(System.Runtime.InteropServices.VarEnum.VT_I4, System.Runtime.InteropServices.ComTypes.PARAMFLAG.PARAMFLAG_FLCID),
                    VBAi.Tests.Infrastructure.OwnedReferenceMetadata.Parameter(System.Runtime.InteropServices.VarEnum.VT_BOOL, System.Runtime.InteropServices.ComTypes.PARAMFLAG.PARAMFLAG_FRETVAL) };
                        var member = Metadata(fixture).Single();
                        Assert.AreEqual("Boolean", member.TypeName); Assert.AreEqual(2, member.Parameters.Length);
                        StringAssert.Contains(member.Parameters[0], "ByVal " + (fallbackName ? "argument1" : "value1") + " As Long");
                        StringAssert.Contains(member.Parameters[1], "Optional ByRef ");
                        Assert.AreEqual(property, member.DefaultMember); Assert.AreEqual(property, member.Global);
                        Assert.AreEqual("Native member documentation", member.Documentation); Assert.AreEqual("Owned.chm", member.HelpFile); Assert.AreEqual(42, member.HelpContext);
                        Assert.AreEqual(1, fixture.FunctionReleases); Assert.AreEqual(1, fixture.TypeReleases);
                    }
        }
        [TestMethod]
        public void ScalarTypesAndUnknownDescriptorsRetainTheirNativeMeaning()
        {
            var method = typeof(EditorReferenceIndex).GetMethod("ReturnType", System.Reflection.BindingFlags.NonPublic | System.Reflection.BindingFlags.Static);
            foreach (var pair in new[] { Tuple.Create(System.Runtime.InteropServices.VarEnum.VT_BSTR, "String"), Tuple.Create(System.Runtime.InteropServices.VarEnum.VT_BOOL, "Boolean"), Tuple.Create(System.Runtime.InteropServices.VarEnum.VT_VOID, "Void"), Tuple.Create(System.Runtime.InteropServices.VarEnum.VT_ERROR, (string)null) })
                Assert.AreEqual(pair.Item2, method.Invoke(null, new object[] { null, new System.Runtime.InteropServices.ComTypes.TYPEDESC { vt = (short)pair.Item1 } }));
        }
        [TestMethod]
        public void EmptyAndUnknownReturnsAndUnavailableParameterNamesRetainAUsableSignature()
        {
            foreach (var type in new[] { System.Runtime.InteropServices.VarEnum.VT_ERROR, System.Runtime.InteropServices.VarEnum.VT_VOID })
                foreach (bool defaultBinding in new[] { false, true })
                    using (var fixture = new VBAi.Tests.Infrastructure.OwnedReferenceMetadata())
                    {
                        fixture.Return = type; fixture.NameCount = 1; fixture.FunctionFlags = defaultBinding ? (short)System.Runtime.InteropServices.ComTypes.FUNCFLAGS.FUNCFLAG_FDEFAULTBIND : (short)0;
                        fixture.Parameters = new[] { VBAi.Tests.Infrastructure.OwnedReferenceMetadata.Parameter(System.Runtime.InteropServices.VarEnum.VT_ERROR, 0) };
                        var member = Metadata(fixture).Single(); Assert.AreEqual(type == System.Runtime.InteropServices.VarEnum.VT_ERROR ? "Variant" : "Void", member.TypeName);
                        Assert.AreEqual(defaultBinding, member.DefaultMember); StringAssert.Contains(member.Parameters[0], "argument1 As Variant");
                    }
        }
        [TestMethod]
        public void VariablesEnumsConstantsHiddenMembersAndFailuresReleaseEveryNativeDescriptor()
        {
            foreach (var kind in new[] { System.Runtime.InteropServices.ComTypes.TYPEKIND.TKIND_RECORD, System.Runtime.InteropServices.ComTypes.TYPEKIND.TKIND_ENUM, System.Runtime.InteropServices.ComTypes.TYPEKIND.TKIND_MODULE })
                foreach (string state in new[] { "field", "constant", "hidden", "restricted", "empty", "error", "unknown" })
                    using (var fixture = new VBAi.Tests.Infrastructure.OwnedReferenceMetadata())
                    {
                        fixture.Function = false; fixture.Kind = kind; fixture.EmptyVariableName = state == "empty"; fixture.FailDocumentation = state == "error";
                        fixture.Variables = new[] { new System.Runtime.InteropServices.ComTypes.VARDESC { memid = 100,
                    varkind = state == "constant" ? System.Runtime.InteropServices.ComTypes.VARKIND.VAR_CONST : System.Runtime.InteropServices.ComTypes.VARKIND.VAR_PERINSTANCE,
                    wVarFlags = state == "hidden" ? (short)System.Runtime.InteropServices.ComTypes.VARFLAGS.VARFLAG_FHIDDEN : state == "restricted" ? (short)System.Runtime.InteropServices.ComTypes.VARFLAGS.VARFLAG_FRESTRICTED : (short)0,
                    elemdescVar = VBAi.Tests.Infrastructure.OwnedReferenceMetadata.Parameter(state == "unknown" ? System.Runtime.InteropServices.VarEnum.VT_ERROR : System.Runtime.InteropServices.VarEnum.VT_I4, 0) } };
                        if (state == "error") Assert.ThrowsException<System.Reflection.TargetInvocationException>(() => Metadata(fixture));
                        else
                        {
                            var members = Metadata(fixture); Assert.AreEqual(state == "hidden" || state == "restricted" || state == "empty" ? 0 : 1, members.Length);
                            if (members.Length != 0) { Assert.AreEqual(kind != System.Runtime.InteropServices.ComTypes.TYPEKIND.TKIND_RECORD, members[0].Global); Assert.AreEqual(state == "unknown" ? "Variant" : "Long", members[0].TypeName); }
                        }
                        Assert.AreEqual(1, fixture.VariableReleases); Assert.AreEqual(1, fixture.TypeReleases);
                    }
        }
        [TestMethod]
        public void NativeLibrariesExposeTypesHelpAndInvalidateWhenReferencesDisappear()
        {
            string library = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.System), "scrrun.dll");
            var first = EditorReferenceIndex.Read(new[] { library, library }, new[] { "Scripting.Dictionary" });
            var identity = first.Single(s => s.Kind == "Library"); Assert.AreEqual("Scripting", identity.Library); Assert.AreEqual(library, identity.LibraryPath); Assert.IsNotNull(identity.LibraryDescription);
            Assert.IsTrue(first.Any(s => s.Name == "Add" && s.Parameters.Any(p => p.Contains(" As "))));
            var removed = EditorReferenceIndex.Read(new string[0], new string[0]); Assert.AreEqual(0, removed.Length);
            var restored = EditorReferenceIndex.Read(new[] { library }, new[] { "Scripting.Dictionary" }); Assert.IsTrue(restored.Any(s => s.Name == "Add"));
        }
    }
}

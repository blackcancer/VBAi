using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using System.Runtime.InteropServices.ComTypes;
using TYPEATTR = System.Runtime.InteropServices.ComTypes.TYPEATTR;
using FUNCDESC = System.Runtime.InteropServices.ComTypes.FUNCDESC;
using FUNCFLAGS = System.Runtime.InteropServices.ComTypes.FUNCFLAGS;
using INVOKEKIND = System.Runtime.InteropServices.ComTypes.INVOKEKIND;
using IMPLTYPEFLAGS = System.Runtime.InteropServices.ComTypes.IMPLTYPEFLAGS;
using TYPEKIND = System.Runtime.InteropServices.ComTypes.TYPEKIND;

namespace CodexVBE
{
    /// <summary>Reads only type-library metadata, never instantiates referenced automation classes.</summary>
    internal static class EditorReferenceIndex
    {
        [DllImport("oleaut32.dll", CharSet = CharSet.Unicode, PreserveSig = false)]
        private static extern void LoadTypeLibEx(string path, int registration, out ITypeLib library);
        [ThreadStatic] private static string cacheKey;
        [ThreadStatic] private static EditorSymbol[] cacheValue;
        internal static EditorSymbol[] Read(string[] paths, string[] requestedTypes)
        {
            string key = string.Join("|", paths.Select(p => p + ":" + System.IO.File.GetLastWriteTimeUtc(p).Ticks)) + ":" + string.Join("|", requestedTypes);
            if (key == cacheKey) return cacheValue;
            var symbols = new List<EditorSymbol>();
            var requested = new HashSet<string>(requestedTypes.Select(n => n.Split('.').Last()), StringComparer.OrdinalIgnoreCase);
            foreach (string path in paths.Distinct(StringComparer.OrdinalIgnoreCase))
            {
                ITypeLib library = null;
                try
                {
                    LoadTypeLibEx(path, 2, out library);
                    int count = Math.Min(library.GetTypeInfoCount(), 10000);
                    for (int i = 0; i < count; i++)
                    {
                        library.GetDocumentation(i, out string name, out string description, out int help, out string file);
                        if (!requested.Contains(name)) continue;
                        library.GetTypeInfo(i, out ITypeInfo info);
                        try { ReadMembers(info, name, symbols, new HashSet<Guid>(), 0); }
                        finally { Marshal.ReleaseComObject(info); }
                    }
                }
                catch (COMException error) { LoadLog.Write("Monaco type metadata: " + error.ErrorCode); }
                finally { if (library != null) Marshal.ReleaseComObject(library); }
            }
            cacheValue = symbols.ToArray(); cacheKey = key; return cacheValue;
        }
        private static void ReadMembers(ITypeInfo info, string owner, List<EditorSymbol> symbols, HashSet<Guid> visited, int depth)
        {
            if (depth > 8) return;
            info.GetTypeAttr(out IntPtr pointer);
            TYPEATTR attr;
            try { attr = (TYPEATTR)Marshal.PtrToStructure(pointer, typeof(TYPEATTR)); }
            finally { info.ReleaseTypeAttr(pointer); }
            if (!visited.Add(attr.guid)) return;
            for (int i = 0; i < Math.Min((int)attr.cFuncs, 4096); i++)
            {
                info.GetFuncDesc(i, out pointer);
                try
                {
                    var func = (FUNCDESC)Marshal.PtrToStructure(pointer, typeof(FUNCDESC));
                    if ((func.wFuncFlags & (short)(FUNCFLAGS.FUNCFLAG_FHIDDEN | FUNCFLAGS.FUNCFLAG_FRESTRICTED)) != 0) continue;
                    var names = new string[Math.Max(1, Math.Min(64, func.cParams + 1))];
                    info.GetNames(func.memid, names, names.Length, out int found);
                    if (found == 0) continue;
                    string[] parameters = names.Skip(1).Take(found - 1).ToArray();
                    symbols.Add(new EditorSymbol { Name = names[0], Module = owner, Scope = "Module", Kind = func.invkind == INVOKEKIND.INVOKE_FUNC ? "Procedure" : "Property", Declaration = owner + "." + names[0] + "(" + string.Join(", ", parameters) + ")", Parameters = parameters, External = true });
                }
                finally { info.ReleaseFuncDesc(pointer); }
            }
            for (int i = 0; i < attr.cImplTypes; i++)
            {
                info.GetImplTypeFlags(i, out IMPLTYPEFLAGS flags);
                if ((flags & IMPLTYPEFLAGS.IMPLTYPEFLAG_FSOURCE) != 0) continue;
                if (attr.typekind == TYPEKIND.TKIND_COCLASS && (flags & IMPLTYPEFLAGS.IMPLTYPEFLAG_FDEFAULT) == 0) continue;
                info.GetRefTypeOfImplType(i, out int reference); info.GetRefTypeInfo(reference, out ITypeInfo inherited);
                try { ReadMembers(inherited, owner, symbols, visited, depth + 1); }
                finally { Marshal.ReleaseComObject(inherited); }
            }
        }
    }
}

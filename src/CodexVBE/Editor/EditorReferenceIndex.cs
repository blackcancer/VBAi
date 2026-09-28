using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using System.Runtime.InteropServices.ComTypes;
using TYPEDESC = System.Runtime.InteropServices.ComTypes.TYPEDESC;
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
        /// <summary>Charge une bibliothèque de types depuis un fichier sans consulter le registre COM.</summary>
        /// <param name="path">Chemin du fichier de bibliothèque de types.</param>
        /// <param name="registration">Mode de chargement OLE Automation.</param>
        /// <param name="library">Reçoit l’interface ITypeLib chargée.</param>
        /// <exception cref="COMException">Le fichier ne peut pas être chargé comme bibliothèque de types.</exception>
[DllImport("oleaut32.dll", CharSet = CharSet.Unicode, PreserveSig = false)]
        private static extern void LoadTypeLibEx(string path, int registration, out ITypeLib library);
        /// <summary>Clé de cache du dernier index sur le thread courant.</summary>
[ThreadStatic] private static string cacheKey;
        /// <summary>Symboles issus du dernier index correspondant à la clé de cache.</summary>
[ThreadStatic] private static EditorSymbol[] cacheValue;
        /// <summary>Lit les membres accessibles des types demandés dans les fichiers de référence.</summary>
        /// <param name="paths">Chemins des bibliothèques de types référencées.</param>
        /// <param name="requestedTypes">Noms complets des types utilisés par le projet.</param>
        /// <returns>Symboles externes avec leur propriétaire et leurs paramètres.</returns>
internal static EditorSymbol[] Read(string[] paths, string[] requestedTypes)
        {
            string key = string.Join("|", paths.Select(p => p + ":" + System.IO.File.GetLastWriteTimeUtc(p).Ticks)) + ":" + string.Join("|", requestedTypes);
            if (key == cacheKey) return cacheValue;
            var symbols = new List<EditorSymbol>();
            var requested = new HashSet<string>(requestedTypes.Select(n => n.Split('.').Last()), StringComparer.OrdinalIgnoreCase);
            requested.Add("Application"); requested.Add("_Application"); requested.Add("_Global");
            var indexed = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            for (int pass = 0; pass < 6; pass++)
            {
                int previousCount = requested.Count;
                foreach (string path in paths.Distinct(StringComparer.OrdinalIgnoreCase))
                {
                    ITypeLib library = null;
                    try
                    {
                        LoadTypeLibEx(path, 2, out library);
                        library.GetDocumentation(-1, out string libraryName, out string libraryHelp, out int libraryContext, out string libraryFile);
                        int count = Math.Min(library.GetTypeInfoCount(), 10000);
                        for (int i = 0; i < count; i++)
                        {
                            library.GetDocumentation(i, out string name, out string description, out int help, out string file);
                            if (!requested.Contains(name) || !indexed.Add(path + "|" + name)) continue;
                            library.GetTypeInfo(i, out ITypeInfo info);
                            try
                            {
                                int start = symbols.Count;
                                ReadMembers(info, name, symbols, new HashSet<Guid>(), 0);
                                for (int member = start; member < symbols.Count; member++) symbols[member].Library = libraryName;
                            }
                            finally { Marshal.ReleaseComObject(info); }
                        }
                    }
                    catch (COMException error) { LoadLog.Write("Monaco type metadata: " + error.ErrorCode); }
                    finally { if (library != null) Marshal.ReleaseComObject(library); }
                }
                foreach (var symbol in symbols.Where(item => !string.IsNullOrEmpty(item.TypeName))) requested.Add(symbol.TypeName.Split('.').Last());
                if (requested.Count == previousCount) break;
            }
            cacheValue = symbols.ToArray(); cacheKey = key; return cacheValue;
        }
        /// <summary>Performs the return type operation for EditorReferenceIndex.</summary>
/// <param name="info">The info used by this operation.</param>
/// <param name="description">The description used by this operation.</param>
/// <returns>The result produced by this operation.</returns>
private static string ReturnType(ITypeInfo info, TYPEDESC description)
        {
            var kind = (VarEnum)description.vt;
            if (kind == VarEnum.VT_PTR || kind == VarEnum.VT_SAFEARRAY)
                return ReturnType(info, (TYPEDESC)Marshal.PtrToStructure(description.lpValue, typeof(TYPEDESC)));
            if (kind != VarEnum.VT_USERDEFINED) return null;
            info.GetRefTypeInfo(unchecked((int)description.lpValue.ToInt64()), out ITypeInfo target);
            try
            {
                target.GetDocumentation(-1, out string name, out string help, out int context, out string file);
                target.GetContainingTypeLib(out ITypeLib library, out int index);
                try { library.GetDocumentation(-1, out string owner, out string text, out int topic, out string path); return owner + "." + name; }
                finally { Marshal.ReleaseComObject(library); }
            }
            finally { Marshal.ReleaseComObject(target); }
        }
        /// <summary>Ajoute les fonctions et propriétés visibles d’un type, puis parcourt ses interfaces héritées.</summary>
        /// <param name="info">Informations COM du type à lire.</param>
        /// <param name="owner">Nom du type qui possédera les symboles indexés.</param>
        /// <param name="symbols">Collection de résultats enrichie pendant le parcours.</param>
        /// <param name="visited">Identifiants de types déjà visités, pour éviter les cycles.</param>
        /// <param name="depth">Profondeur d’héritage courante, limitée pour borner la récursion.</param>
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
                    symbols.Add(new EditorSymbol { Name = names[0], Module = owner, Scope = "Module", Kind = func.invkind == INVOKEKIND.INVOKE_FUNC ? "Procedure" : "Property", Declaration = owner + "." + names[0] + "(" + string.Join(", ", parameters) + ")", Parameters = parameters, TypeName = ReturnType(info, func.elemdescFunc.tdesc), External = true });
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

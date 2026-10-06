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
using ELEMDESC = System.Runtime.InteropServices.ComTypes.ELEMDESC;
using PARAMFLAG = System.Runtime.InteropServices.ComTypes.PARAMFLAG;
using VARDESC = System.Runtime.InteropServices.ComTypes.VARDESC;
using VARFLAGS = System.Runtime.InteropServices.ComTypes.VARFLAGS;

namespace VBAi
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

        /// <summary>Correspondance des types Automation scalaires avec leur nom VBA.</summary>
        private static readonly Dictionary<VarEnum, string> PrimitiveTypes = new Dictionary<VarEnum, string> {
            { VarEnum.VT_EMPTY, "Variant" }, { VarEnum.VT_VOID, "Void" }, { VarEnum.VT_VARIANT, "Variant" },
            { VarEnum.VT_I2, "Integer" }, { VarEnum.VT_I4, "Long" }, { VarEnum.VT_I8, "LongLong" },
            { VarEnum.VT_UI1, "Byte" }, { VarEnum.VT_R4, "Single" }, { VarEnum.VT_R8, "Double" },
            { VarEnum.VT_CY, "Currency" }, { VarEnum.VT_DATE, "Date" }, { VarEnum.VT_BSTR, "String" },
            { VarEnum.VT_BOOL, "Boolean" }, { VarEnum.VT_DISPATCH, "Object" }, { VarEnum.VT_UNKNOWN, "Object" }
        };

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
                        if (pass == 0) symbols.Add(new EditorSymbol { Name = libraryName, Module = libraryName, Scope = "Module", Kind = "Library", TypeName = libraryName, Declaration = libraryName, Documentation = libraryHelp, Library = libraryName, LibraryDescription = libraryHelp, LibraryPath = path, HelpFile = libraryFile, HelpContext = libraryContext, External = true });
                        int count = Math.Min(library.GetTypeInfoCount(), 10000);
                        for (int i = 0; i < count; i++)
                        {
                            library.GetDocumentation(i, out string name, out string description, out int help, out string file);
                            library.GetTypeInfoType(i, out TYPEKIND typeKind);
                            if (pass == 0 && !name.StartsWith("_", StringComparison.Ordinal))
                            {
                                symbols.Add(new EditorSymbol { Name = name, Module = libraryName, Scope = "Module", Kind = typeKind == TYPEKIND.TKIND_COCLASS ? "Class" : typeKind == TYPEKIND.TKIND_ENUM ? "Enum" : "Type", TypeName = libraryName + "." + name, Declaration = libraryName + "." + name, Documentation = description, Library = libraryName, LibraryDescription = libraryHelp, LibraryPath = path, HelpFile = file, HelpContext = help, External = true });
                            }
                            if ((!requested.Contains(name) && typeKind != TYPEKIND.TKIND_MODULE && typeKind != TYPEKIND.TKIND_ENUM) || !indexed.Add(path + "|" + name)) continue;
                            library.GetTypeInfo(i, out ITypeInfo info);
                            try
                            {
                                int start = symbols.Count;
                                ReadMembers(info, name, symbols, new HashSet<Guid>(), 0);
                                for (int member = start; member < symbols.Count; member++)
                                { symbols[member].Library = libraryName; symbols[member].LibraryDescription = libraryHelp; symbols[member].LibraryPath = path; }
                            }
                            finally { Marshal.ReleaseComObject(info); }
                        }
                    }
                    catch (COMException error) { LoadLog.Write("Monaco type metadata: " + error.ErrorCode); }
                    finally { if (library != null) Marshal.ReleaseComObject(library); }
                }
                // Catalogued type names are offered in completion without eagerly expanding every type.
                foreach (var symbol in symbols.Where(item => new[] { "Procedure", "Property", "Field", "EnumMember", "Constant" }.Contains(item.Kind))) requested.Add(symbol.TypeName.Split('.').Last());
                if (requested.Count == previousCount) break;
            }
            cacheValue = symbols.ToArray(); cacheKey = key; return cacheValue;
        }

        /// <summary>Resolves a COM type description to a VBA-facing primitive or qualified user-defined type name.</summary>
        /// <param name="info">Type information that owns the referenced type descriptor.</param>
        /// <param name="description">COM descriptor for the return or member type; pointer and SAFEARRAY wrappers are unwrapped recursively.</param>
        /// <returns>Mapped type name, or <see langword="null"/> when the descriptor has no supported primitive mapping.</returns>
        private static string ReturnType(ITypeInfo info, TYPEDESC description)
        {
            var kind = (VarEnum)description.vt;
            if (kind == VarEnum.VT_PTR || kind == VarEnum.VT_SAFEARRAY)
                return ReturnType(info, (TYPEDESC)Marshal.PtrToStructure(description.lpValue, typeof(TYPEDESC)));
            if (kind != VarEnum.VT_USERDEFINED)
                return PrimitiveTypes.TryGetValue(kind, out string typeName) ? typeName : null;
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
                    var parameters = new List<string>();
                    string resultType = ReturnType(info, func.elemdescFunc.tdesc) ?? "Variant";
                    for (int parameter = 0; parameter < func.cParams; parameter++)
                    {
                        var element = (ELEMDESC)Marshal.PtrToStructure(IntPtr.Add(func.lprgelemdescParam, parameter * Marshal.SizeOf(typeof(ELEMDESC))), typeof(ELEMDESC));
                        var flags = element.desc.paramdesc.wParamFlags;
                        string parameterType = ReturnType(info, element.tdesc) ?? "Variant";
                        if ((flags & PARAMFLAG.PARAMFLAG_FRETVAL) != 0) { resultType = parameterType; continue; }
                        if ((flags & PARAMFLAG.PARAMFLAG_FLCID) != 0) continue;
                        string parameterName = parameter + 1 < found ? names[parameter + 1] : "argument" + (parameter + 1);
                        parameters.Add(((flags & PARAMFLAG.PARAMFLAG_FOPT) != 0 ? "Optional " : "") + ((flags & PARAMFLAG.PARAMFLAG_FOUT) != 0 || (VarEnum)element.tdesc.vt == VarEnum.VT_PTR ? "ByRef " : "ByVal ") + parameterName + " As " + parameterType);
                    }
                    info.GetDocumentation(func.memid, out string memberName, out string documentation, out int helpContext, out string helpFile);
                    string[] parameterNames = parameters.ToArray();
                    symbols.Add(new EditorSymbol { Name = names[0], Module = owner, Scope = "Module", Kind = func.invkind == INVOKEKIND.INVOKE_FUNC ? "Procedure" : "Property", Declaration = owner + "." + names[0] + "(" + string.Join(", ", parameterNames) + ")" + (resultType == "Void" ? "" : " As " + resultType), Parameters = parameterNames, TypeName = resultType, Documentation = documentation, HelpFile = helpFile, HelpContext = helpContext, DefaultMember = func.memid == 0 || (func.wFuncFlags & (short)FUNCFLAGS.FUNCFLAG_FDEFAULTBIND) != 0, Global = attr.typekind == TYPEKIND.TKIND_MODULE, External = true });
                }
                finally { info.ReleaseFuncDesc(pointer); }
            }
            for (int i = 0; i < Math.Min((int)attr.cVars, 4096); i++)
            {
                info.GetVarDesc(i, out pointer);
                try
                {
                    var variable = (VARDESC)Marshal.PtrToStructure(pointer, typeof(VARDESC));
                    if ((variable.wVarFlags & (short)(VARFLAGS.VARFLAG_FHIDDEN | VARFLAGS.VARFLAG_FRESTRICTED)) != 0) continue;
                    info.GetDocumentation(variable.memid, out string name, out string documentation, out int context, out string file);
                    if (string.IsNullOrEmpty(name)) continue;
                    string typeName = ReturnType(info, variable.elemdescVar.tdesc) ?? "Variant";
                    string kind = attr.typekind == TYPEKIND.TKIND_ENUM ? "EnumMember" : variable.varkind == VARKIND.VAR_CONST ? "Constant" : "Field";
                    symbols.Add(new EditorSymbol { Name = name, Module = owner, Scope = "Module", Kind = kind, TypeName = typeName, Declaration = owner + "." + name + " As " + typeName, Documentation = documentation, HelpFile = file, HelpContext = context, Global = attr.typekind == TYPEKIND.TKIND_MODULE || attr.typekind == TYPEKIND.TKIND_ENUM, External = true });
                }
                finally { info.ReleaseVarDesc(pointer); }
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

using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Runtime.InteropServices.ComTypes;
using System.Security.Cryptography;
using System.Text;
using FUNCDESC = System.Runtime.InteropServices.ComTypes.FUNCDESC;
using IMPLTYPEFLAGS = System.Runtime.InteropServices.ComTypes.IMPLTYPEFLAGS;
using TYPEATTR = System.Runtime.InteropServices.ComTypes.TYPEATTR;
using TYPEKIND = System.Runtime.InteropServices.ComTypes.TYPEKIND;
using TYPELIBATTR = System.Runtime.InteropServices.ComTypes.TYPELIBATTR;
using VARDESC = System.Runtime.InteropServices.ComTypes.VARDESC;

namespace VBAi
{

    /// <summary>Énumère les types et membres d’une référence de projet à partir de ses seules métadonnées de bibliothèque COM.</summary>
    internal sealed class VbeReferenceTypes
    {

        /// <summary>Nombre maximal d’éléments retournés par une page.</summary>
        private const int MaxPageSize = 50;

        /// <summary>Limite de sécurité pour le nombre de types inspectés.</summary>
        private const int MaxTypes = 10000;

        /// <summary>Limite de sécurité pour le nombre de membres inspectés.</summary>
        private const int MaxMembers = 10000;

        /// <summary>Instance VBE utilisée pour résoudre les projets et références sélectionnés.</summary>
        private readonly dynamic vbe;

        /// <summary>Chargeur injectable d’une bibliothèque depuis un fichier.</summary>
        private readonly Func<string, ITypeLib> loadFile;

        /// <summary>Chargeur injectable d’une bibliothèque enregistrée.</summary>
        private readonly Func<Guid, ushort, ushort, ITypeLib> loadRegistered;

        /// <summary>Charge une bibliothèque de types depuis un fichier sans l’enregistrer dans le système.</summary>
        /// <param name="file">Chemin du fichier de bibliothèque.</param>
        /// <param name="regKind">Mode de chargement COM.</param>
        /// <param name="typeLib">Reçoit la bibliothèque chargée.</param>
        [DllImport("oleaut32.dll", CharSet = CharSet.Unicode, PreserveSig = false)]
        private static extern void LoadTypeLibEx(string file, int regKind, out ITypeLib typeLib);

        /// <summary>Charge la bibliothèque enregistrée identifiée par son GUID et sa version.</summary>
        /// <param name="guid">Identifiant de la bibliothèque.</param>
        /// <param name="major">Version majeure.</param>
        /// <param name="minor">Version mineure.</param>
        /// <param name="lcid">Identifiant de langue.</param>
        /// <param name="typeLib">Reçoit la bibliothèque chargée.</param>
        [DllImport("oleaut32.dll", PreserveSig = false)]
        private static extern void LoadRegTypeLib(ref Guid guid, ushort major, ushort minor,
            int lcid, out ITypeLib typeLib);

        /// <summary>Crée le lecteur avec les chargeurs COM du système.</summary>
        /// <param name="vbe">Instance VBE servant à résoudre le projet.</param>
        public VbeReferenceTypes(object vbe)
            : this(vbe, LoadSelectedFile, LoadRegisteredLibrary) { }

        /// <summary>Crée le lecteur avec les chargeurs de fichiers et de bibliothèques enregistrées fournis.</summary>
        /// <param name="vbe">Instance VBE servant à résoudre le projet.</param>
        /// <param name="loadFile">Fonction de chargement depuis un fichier.</param>
        /// <param name="loadRegistered">Fonction de chargement depuis le registre COM.</param>
        /// <exception cref="ArgumentNullException">Un des chargeurs requis est nul.</exception>
        internal VbeReferenceTypes(object vbe, Func<string, ITypeLib> loadFile,
            Func<Guid, ushort, ushort, ITypeLib> loadRegistered)
        {
            this.vbe = vbe;
            this.loadFile = loadFile ?? throw new ArgumentNullException(nameof(loadFile));
            this.loadRegistered = loadRegistered ?? throw new ArgumentNullException(nameof(loadRegistered));
        }

        /// <summary>Charge le fichier de référence avec le mode REGKIND_NONE pour éviter toute inscription.</summary>
        /// <param name="path">Chemin du fichier de bibliothèque.</param>
        /// <returns>Bibliothèque de types chargée.</returns>
        private static ITypeLib LoadSelectedFile(string path)
        {
            LoadTypeLibEx(path, 2, out ITypeLib library); // REGKIND_NONE
            return library;
        }

        /// <summary>Charge une bibliothèque de types par son identité enregistrée.</summary>
        /// <param name="guid">Identifiant de la bibliothèque.</param>
        /// <param name="major">Version majeure.</param>
        /// <param name="minor">Version mineure.</param>
        /// <returns>Bibliothèque enregistrée correspondante.</returns>
        private static ITypeLib LoadRegisteredLibrary(Guid guid, ushort major, ushort minor)
        {
            LoadRegTypeLib(ref guid, major, minor, 0, out ITypeLib library);
            return library;
        }

        /// <summary>Retourne une page de types de la référence explicitement sélectionnée par le projet.</summary>
        /// <param name="request">Requête contenant projet, identité de référence, décalage et limite.</param>
        /// <returns>Objet résultat avec métadonnées de référence, pagination et types lus ou erreurs par type.</returns>
        /// <exception cref="ArgumentException">Les paramètres de page ou l’identité de référence sont invalides.</exception>
        /// <exception cref="InvalidOperationException">La bibliothèque dépasse les limites ou la référence ne peut pas être ouverte.</exception>
        public object ListTypes(Request request)
        {
            int offset = ValidatePage(request);
            ReferenceSource reference = FindReference(request);
            LoadedLibrary loaded = OpenLibrary(reference);
            ITypeLib library = loaded.Library;
            int count = library.GetTypeInfoCount();
            if (count < 0 || count > MaxTypes)
                throw new InvalidOperationException("Type library type count exceeds the inspection limit.");
            var types = new List<object>();
            int end = Math.Min(count, offset + PageSize(request));
            for (int index = offset; index < end; index++)
            {
                try
                {
                    library.GetTypeInfo(index, out ITypeInfo typeInfo);
                    types.Add(ReadType(typeInfo, index, reference));
                }
                catch (Exception ex)
                {
                    types.Add(new { TypeIndex = index, Error = Error(ex) });
                }
            }
            return new
            {
                request.Project,
                Reference = reference,
                loaded.Source,
                loaded.FallbackError,
                TotalTypes = count,
                Offset = offset,
                Returned = types.Count,
                HasMore = end < count,
                Types = types
            };
        }

        /// <summary>Retourne une page de membres pour un type précédemment listé, après vérification de son identité.</summary>
        /// <param name="request">Requête identifiant le type et la page de membres.</param>
        /// <returns>Objet résultat avec le type, l’interface choisie, la pagination et les membres ou erreurs par membre.</returns>
        /// <exception cref="ArgumentException">L’identité du type manque ou les paramètres de page sont invalides.</exception>
        /// <exception cref="ArgumentOutOfRangeException">L’index de type est hors limites.</exception>
        /// <exception cref="InvalidOperationException">L’identité a changé, la référence est invalide ou la limite est dépassée.</exception>
        public object ListMembers(Request request)
        {
            ValidatePage(request);
            if (string.IsNullOrWhiteSpace(request.TypeIdentity))
                throw new ArgumentException("TypeIdentity from list_reference_types is required.");
            ReferenceSource reference = FindReference(request);
            LoadedLibrary loaded = OpenLibrary(reference);
            ITypeLib library = loaded.Library;
            int typeCount = library.GetTypeInfoCount();
            if (typeCount < 0 || typeCount > MaxTypes || request.TypeIndex < 0 ||
                request.TypeIndex >= typeCount)
                throw new ArgumentOutOfRangeException("TypeIndex");
            library.GetTypeInfo(request.TypeIndex, out ITypeInfo typeInfo);
            TypeMetadata metadata = ReadTypeMetadata(typeInfo, request.TypeIndex, reference);
            if (!string.Equals(metadata.Identity, request.TypeIdentity, StringComparison.OrdinalIgnoreCase))
                throw new InvalidOperationException("The type changed since list_reference_types; list it again.");
            ITypeInfo memberType = typeInfo;
            TypeMetadata memberMetadata = metadata;
            string resolution = "DirectType";
            if (metadata.Kind == TYPEKIND.TKIND_COCLASS.ToString())
            {
                memberType = DefaultNonSourceInterface(typeInfo);
                if (memberType != null)
                {
                    memberMetadata = ReadTypeMetadata(memberType, request.TypeIndex, reference);
                    resolution = "DefaultNonSourceInterface";
                }
                else
                    resolution = "NoDefaultNonSourceInterface";
            }
            int count = memberType == null ? 0 : memberMetadata.FunctionCount + memberMetadata.VariableCount;
            if (count > MaxMembers)
                throw new InvalidOperationException("Type member count exceeds the inspection limit.");
            int end = Math.Min(count, request.Offset + PageSize(request));
            var members = new List<object>();
            for (int index = request.Offset; index < end; index++)
            {
                try
                {
                    members.Add(index < memberMetadata.FunctionCount
                        ? ReadFunction(memberType, index)
                        : ReadVariable(memberType, index - memberMetadata.FunctionCount, index));
                }
                catch (Exception ex)
                {
                    members.Add(new { MemberIndex = index, Error = Error(ex) });
                }
            }
            return new
            {
                request.Project,
                Reference = reference,
                Type = metadata,
                MemberInterface = memberMetadata,
                Resolution = resolution,
                loaded.Source,
                loaded.FallbackError,
                Scope = "Raw COM type members and accessors; inherited interfaces and VBE Object Browser filtering are not expanded.",
                TotalMembers = count,
                request.Offset,
                Returned = members.Count,
                HasMore = end < count,
                Members = members
            };
        }

        /// <summary>Valide le décalage et la limite de pagination d’une requête.</summary>
        /// <param name="request">Requête à contrôler.</param>
        /// <returns>Décalage validé.</returns>
        /// <exception cref="ArgumentOutOfRangeException">Le décalage ou la limite dépasse les bornes admises.</exception>
        private static int ValidatePage(Request request)
        {
            if (request.Offset < 0 || request.Offset > MaxTypes)
                throw new ArgumentOutOfRangeException("Offset", "Offset must be between 0 and 10000.");
            if (request.Limit < 0 || request.Limit > MaxPageSize)
                throw new ArgumentOutOfRangeException("Limit", "Limit must be between 1 and 50; omit for 50.");
            return request.Offset;
        }

        /// <summary>Détermine la taille de page effective, en utilisant la limite maximale si la requête omet la limite.</summary>
        /// <param name="request">Requête paginée.</param>
        /// <returns>Limite explicite ou limite maximale par défaut.</returns>
        private static int PageSize(Request request) { return request.Limit == 0 ? MaxPageSize : request.Limit; }

        /// <summary>Résout l’unique référence de projet correspondant exactement au GUID et à la version demandés.</summary>
        /// <param name="request">Requête identifiant le projet et la référence.</param>
        /// <returns>Métadonnées et empreinte de fichier de la référence retenue.</returns>
        /// <exception cref="ArgumentException">Le projet, le GUID ou la version est invalide.</exception>
        /// <exception cref="InvalidOperationException">La référence est ambiguë, cassée ou absente du projet.</exception>
        /// <exception cref="FileNotFoundException">Le fichier de bibliothèque n’est pas disponible.</exception>
        private ReferenceSource FindReference(Request request)
        {
            if (string.IsNullOrWhiteSpace(request.Project) ||
                !Guid.TryParse(request.Guid, out Guid expectedGuid) || request.Major < 0 || request.Minor < 0 ||
                request.Major > ushort.MaxValue || request.Minor > ushort.MaxValue)
                throw new ArgumentException("Project, reference Guid, Major and Minor are required.");
            dynamic project = VbeProjectResolver.Resolve(vbe, request.Project);
            ReferenceSource found = null;
            foreach (dynamic reference in project.References)
            {
                if (!Guid.TryParse((string)reference.GUID, out Guid actualGuid) ||
                    actualGuid != expectedGuid || (int)reference.Major != request.Major ||
                    (int)reference.Minor != request.Minor) continue;
                if (found != null) throw new InvalidOperationException("Reference identity is ambiguous.");
                if ((bool)reference.IsBroken)
                    throw new InvalidOperationException("The selected reference is broken.");
                string path = (string)reference.FullPath;
                if (string.IsNullOrWhiteSpace(path) || !Path.IsPathRooted(path) || !File.Exists(path))
                    throw new FileNotFoundException("Reference type library file is unavailable.", path);
                var file = new FileInfo(path);
                found = new ReferenceSource
                {
                    Name = (string)reference.Name,
                    Guid = actualGuid.ToString("B"),
                    Major = request.Major,
                    Minor = request.Minor,
                    FullPath = path,
                    FileLength = file.Length,
                    FileLastWriteUtc = file.LastWriteTimeUtc.ToString("O")
                };
            }
            if (found == null) throw new InvalidOperationException("The exact reference is not selected by this project.");
            return found;
        }

        /// <summary>Ouvre le fichier de référence, utilise le registre en repli et vérifie l’identité de la bibliothèque chargée.</summary>
        /// <param name="reference">Référence résolue dans le projet.</param>
        /// <returns>Bibliothèque COM, source de chargement et erreur de repli éventuelle.</returns>
        /// <exception cref="InvalidOperationException">Le chargement échoue ou l’identité de la bibliothèque ne correspond pas à la référence.</exception>
        private LoadedLibrary OpenLibrary(ReferenceSource reference)
        {
            ITypeLib library;
            string source = "SelectedReferenceFile: LoadTypeLibEx(REGKIND_NONE)";
            object fallbackError = null;
            try { library = loadFile(reference.FullPath); }
            catch (COMException ex)
            {
                fallbackError = Error(ex);
                Guid guid = Guid.Parse(reference.Guid);
                try
                {
                    library = loadRegistered(guid, (ushort)reference.Major, (ushort)reference.Minor);
                }
                catch (COMException registryError)
                {
                    throw new InvalidOperationException("Type library failed from selected file (HRESULT " +
                        HResultHex(ex) + ") and registry (HRESULT " + HResultHex(registryError) + ").",
                        registryError);
                }
                source = "RegisteredTypeLibrary: LoadRegTypeLib";
            }
            if (library == null) throw new InvalidOperationException("LoadTypeLibEx returned no library.");
            IntPtr pointer = IntPtr.Zero;
            try
            {
                library.GetLibAttr(out pointer);
                var attr = (TYPELIBATTR)Marshal.PtrToStructure(pointer, typeof(TYPELIBATTR));
                if (!string.Equals(attr.guid.ToString("B"), reference.Guid, StringComparison.OrdinalIgnoreCase) ||
                    attr.wMajorVerNum != reference.Major || attr.wMinorVerNum != reference.Minor)
                    throw new InvalidOperationException("The file type library identity differs from the selected VBE reference.");
            }
            finally { if (pointer != IntPtr.Zero) library.ReleaseTLibAttr(pointer); }
            return new LoadedLibrary
            {
                Library = library,
                Source = source,
                FallbackError = fallbackError
            };
        }

        /// <summary>Choisit l’interface par défaut non source d’une coclasse, si elle en expose une.</summary>
        /// <param name="coclass">Informations de type de la coclasse.</param>
        /// <returns>Informations de l’interface choisie, ou nul si aucune interface admissible n’existe.</returns>
        private static ITypeInfo DefaultNonSourceInterface(ITypeInfo coclass)
        {
            IntPtr pointer = IntPtr.Zero;
            try
            {
                coclass.GetTypeAttr(out pointer);
                var attr = (TYPEATTR)Marshal.PtrToStructure(pointer, typeof(TYPEATTR));
                int selected = -1;
                for (int index = 0; index < attr.cImplTypes; index++)
                {
                    coclass.GetImplTypeFlags(index, out IMPLTYPEFLAGS flags);
                    if ((flags & IMPLTYPEFLAGS.IMPLTYPEFLAG_FSOURCE) != 0) continue;
                    if ((flags & IMPLTYPEFLAGS.IMPLTYPEFLAG_FDEFAULT) != 0)
                    { selected = index; break; }
                    if (selected == -1) selected = index;
                }
                if (selected == -1) return null;
                coclass.GetRefTypeOfImplType(selected, out int href);
                coclass.GetRefTypeInfo(href, out ITypeInfo result);
                return result;
            }
            finally { if (pointer != IntPtr.Zero) coclass.ReleaseTypeAttr(pointer); }
        }

        /// <summary>Projette les métadonnées d’un type dans l’objet de résultat retourné par l’API.</summary>
        /// <param name="typeInfo">Informations COM du type.</param>
        /// <param name="index">Index du type dans la bibliothèque.</param>
        /// <param name="reference">Identité et version du fichier de référence.</param>
        /// <returns>Objet exposant les informations et l’identité du type.</returns>
        private static object ReadType(ITypeInfo typeInfo, int index, ReferenceSource reference)
        {
            TypeMetadata type = ReadTypeMetadata(typeInfo, index, reference);
            return new
            {
                type.TypeIndex,
                type.Name,
                type.Guid,
                type.Kind,
                type.FunctionCount,
                type.VariableCount,
                type.ImplementedInterfaceCount,
                TypeIdentity = type.Identity
            };
        }

        /// <summary>Lit les attributs, le nom et le GUID d’un type et calcule son identité de pagination.</summary>
        /// <param name="typeInfo">Informations COM du type.</param>
        /// <param name="index">Index du type dans la bibliothèque.</param>
        /// <param name="reference">Métadonnées de la référence d’origine.</param>
        /// <returns>Métadonnées de type et identité stable pour la référence lue.</returns>
        private static TypeMetadata ReadTypeMetadata(ITypeInfo typeInfo, int index, ReferenceSource reference)
        {
            IntPtr pointer = IntPtr.Zero;
            try
            {
                typeInfo.GetTypeAttr(out pointer);
                var attr = (TYPEATTR)Marshal.PtrToStructure(pointer, typeof(TYPEATTR));
                typeInfo.GetDocumentation(-1, out string name, out string description, out int helpContext, out string helpFile);
                string identity = Hash(reference.Guid + "|" + reference.Major + "|" + reference.Minor +
                    "|" + reference.FileLength + "|" + reference.FileLastWriteUtc + "|" +
                    index + "|" + attr.guid + "|" + attr.typekind + "|" + name);
                return new TypeMetadata
                {
                    TypeIndex = index,
                    Name = name,
                    Guid = attr.guid.ToString("B"),
                    Kind = attr.typekind.ToString(),
                    FunctionCount = attr.cFuncs,
                    VariableCount = attr.cVars,
                    ImplementedInterfaceCount = attr.cImplTypes,
                    Identity = identity
                };
            }
            finally { if (pointer != IntPtr.Zero) typeInfo.ReleaseTypeAttr(pointer); }
        }

        /// <summary>Lit le nom, le DISPID, le genre d’invocation et les noms de paramètres d’une fonction COM.</summary>
        /// <param name="typeInfo">Informations COM du type membre.</param>
        /// <param name="index">Index de fonction.</param>
        /// <returns>Objet membre destiné au résultat paginé.</returns>
        private static object ReadFunction(ITypeInfo typeInfo, int index)
        {
            IntPtr pointer = IntPtr.Zero;
            try
            {
                typeInfo.GetFuncDesc(index, out pointer);
                var descriptor = (FUNCDESC)Marshal.PtrToStructure(pointer, typeof(FUNCDESC));
                string[] names = new string[Math.Min(32, Math.Max(1, descriptor.cParams + 1))];
                typeInfo.GetNames(descriptor.memid, names, names.Length, out int count);
                return new
                {
                    MemberIndex = index,
                    Kind = "Function",
                    Name = count > 0 ? names[0] : null,
                    DispId = descriptor.memid,
                    InvocationKind = descriptor.invkind.ToString(),
                    FunctionFlags = descriptor.wFuncFlags,
                    ParameterCount = (int)descriptor.cParams,
                    ParameterNames = names.Skip(1).Take(Math.Max(0, count - 1)).ToArray(),
                    ParameterNamesTruncated = descriptor.cParams + 1 > names.Length
                };
            }
            finally { if (pointer != IntPtr.Zero) typeInfo.ReleaseFuncDesc(pointer); }
        }

        /// <summary>Lit le nom, le DISPID et le genre d’une variable de type COM.</summary>
        /// <param name="typeInfo">Informations COM du type membre.</param>
        /// <param name="index">Index de variable dans le descripteur COM.</param>
        /// <param name="memberIndex">Index global de membre utilisé dans le résultat.</param>
        /// <returns>Objet membre destiné au résultat paginé.</returns>
        private static object ReadVariable(ITypeInfo typeInfo, int index, int memberIndex)
        {
            IntPtr pointer = IntPtr.Zero;
            try
            {
                typeInfo.GetVarDesc(index, out pointer);
                var descriptor = (VARDESC)Marshal.PtrToStructure(pointer, typeof(VARDESC));
                typeInfo.GetDocumentation(descriptor.memid, out string name, out string description,
                    out int helpContext, out string helpFile);
                return new
                {
                    MemberIndex = memberIndex,
                    Kind = "Variable",
                    Name = name,
                    DispId = descriptor.memid,
                    VariableKind = descriptor.varkind.ToString()
                };
            }
            finally { if (pointer != IntPtr.Zero) typeInfo.ReleaseVarDesc(pointer); }
        }

        /// <summary>Calcule l’empreinte SHA-256 hexadécimale minuscule d’une chaîne UTF-8.</summary>
        /// <param name="value">Texte à hacher.</param>
        /// <returns>Empreinte composée de 64 chiffres hexadécimaux.</returns>
        private static string Hash(string value)
        {
            using (var sha = SHA256.Create())
                return BitConverter.ToString(sha.ComputeHash(Encoding.UTF8.GetBytes(value)))
                    .Replace("-", "").ToLowerInvariant();
        }

        /// <summary>Projette une exception en type, HRESULT et message pour un résultat sérialisable.</summary>
        /// <param name="error">Exception à décrire.</param>
        /// <returns>Objet anonyme contenant les informations d’erreur.</returns>
        private static object Error(Exception error)
        {
            return new
            {
                Type = error.GetType().Name,
                HResult = HResultHex(error),
                error.Message
            };
        }

        /// <summary>Formate le HRESULT d’une exception sur huit chiffres hexadécimaux préfixés par <c>0x</c>.</summary>
        /// <param name="error">Exception source.</param>
        /// <returns>Représentation hexadécimale non signée du HRESULT.</returns>
        private static string HResultHex(Exception error)
        {
            return "0x" + unchecked((uint)error.HResult).ToString("X8");
        }

        /// <summary>Identité exacte et métadonnées de fichier d’une référence du projet.</summary>
        private sealed class ReferenceSource
        {

            /// <summary>Obtient ou définit le nom de la référence.</summary>
            /// <value>le nom de la référence.</value>
            public string Name { get; set; }

            /// <summary>Obtient ou définit le GUID au format avec accolades.</summary>
            /// <value>le GUID au format avec accolades.</value>
            public string Guid { get; set; }

            /// <summary>Obtient ou définit la version majeure.</summary>
            /// <value>la version majeure.</value>
            public int Major { get; set; }

            /// <summary>Obtient ou définit la version mineure.</summary>
            /// <value>la version mineure.</value>
            public int Minor { get; set; }

            /// <summary>Obtient ou définit le chemin complet du fichier.</summary>
            /// <value>le chemin complet du fichier.</value>
            public string FullPath { get; set; }

            /// <summary>Obtient ou définit la taille du fichier en octets.</summary>
            /// <value>la taille du fichier en octets.</value>
            public long FileLength { get; set; }

            /// <summary>Obtient ou définit la date de modification UTC au format rond.</summary>
            /// <value>la date de modification UTC au format rond.</value>
            public string FileLastWriteUtc { get; set; }
        }

        /// <summary>Résultat de l’ouverture d’une bibliothèque et indication de sa source.</summary>
        private sealed class LoadedLibrary
        {

            /// <summary>Obtient ou définit la bibliothèque COM chargée.</summary>
            /// <value>la bibliothèque COM chargée.</value>
            public ITypeLib Library { get; set; }

            /// <summary>Obtient ou définit le mécanisme de chargement utilisé.</summary>
            /// <value>le mécanisme de chargement utilisé.</value>
            public string Source { get; set; }

            /// <summary>Obtient ou définit l’erreur du chargement fichier si le repli registre a réussi.</summary>
            /// <value>l’erreur du chargement fichier si le repli registre a réussi.</value>
            public object FallbackError { get; set; }
        }

        /// <summary>Métadonnées de type COM utilisées pour la liste et l’inspection paginée.</summary>
        private sealed class TypeMetadata
        {

            /// <summary>Obtient ou définit l’index du type dans sa bibliothèque.</summary>
            /// <value>l’index du type dans sa bibliothèque.</value>
            public int TypeIndex { get; set; }

            /// <summary>Obtient ou définit le nom du type.</summary>
            /// <value>le nom du type.</value>
            public string Name { get; set; }

            /// <summary>Obtient ou définit le GUID du type.</summary>
            /// <value>le GUID du type.</value>
            public string Guid { get; set; }

            /// <summary>Obtient ou définit la catégorie COM du type.</summary>
            /// <value>la catégorie COM du type.</value>
            public string Kind { get; set; }

            /// <summary>Obtient ou définit le nombre de fonctions exposées.</summary>
            /// <value>le nombre de fonctions exposées.</value>
            public int FunctionCount { get; set; }

            /// <summary>Obtient ou définit le nombre de variables exposées.</summary>
            /// <value>le nombre de variables exposées.</value>
            public int VariableCount { get; set; }

            /// <summary>Obtient ou définit le nombre d’interfaces implémentées.</summary>
            /// <value>le nombre d’interfaces implémentées.</value>
            public int ImplementedInterfaceCount { get; set; }

            /// <summary>Obtient ou définit l’empreinte d’identité utilisée pour vérifier que le type n’a pas changé.</summary>
            /// <value>l’empreinte d’identité utilisée pour vérifier que le type n’a pas changé.</value>
            public string Identity { get; set; }
        }
    }
}

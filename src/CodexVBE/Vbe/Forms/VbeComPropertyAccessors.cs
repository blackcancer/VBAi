using System;
using System.Collections.Generic;
using System.Runtime.InteropServices;
using System.Runtime.InteropServices.ComTypes;
using INVOKEKIND = System.Runtime.InteropServices.ComTypes.INVOKEKIND;
using TYPEATTR = System.Runtime.InteropServices.ComTypes.TYPEATTR;
using FUNCDESC = System.Runtime.InteropServices.ComTypes.FUNCDESC;
using IMPLTYPEFLAGS = System.Runtime.InteropServices.ComTypes.IMPLTYPEFLAGS;

namespace CodexVBE
{
    /// <summary>Inspecte les accesseurs COM déclarés pour une propriété sans prétendre vérifier un setter en exécution.</summary>
    internal static class VbeComPropertyAccessors
    {
        /// <summary>Parcourt ITypeInfo et rapporte les interfaces et accesseurs de la propriété demandée.</summary>
        /// <param name="target">Objet COM à inspecter.</param>
        /// <param name="propertyName">Nom de la propriété recherché sans distinction de casse.</param>
        /// <returns>Rapport de métadonnées qui distingue le setter déclaré de sa vérification réelle.</returns>
        /// <exception cref="ArgumentException">La cible COM est nulle ou le nom est vide.</exception>
        public static object Inspect(object target, string propertyName)
        {
            if (target == null || string.IsNullOrWhiteSpace(propertyName))
                throw new ArgumentException("A COM target and property name are required.");
            var interfaces = new List<string>();
            var accessors = new List<object>();
            var errors = new List<string>();
            var provider = target as IProvideClassInfo;
            if (provider == null)
                return new { Property = propertyName, Discovery = "IProvideClassInfo/ITypeInfo",
                    MetadataComplete = false, SetterDeclared = (bool?)null,
                    ActualSetterVerified = false, Interfaces = interfaces,
                    Accessors = accessors, Errors = new[] { "IProvideClassInfo is unavailable." } };

            ITypeInfo classInfo = null;
            try
            {
                provider.GetClassInfo(out classInfo);
                if (classInfo == null) throw new InvalidOperationException("GetClassInfo returned null.");
                var visited = new HashSet<Guid>();
                ReadType(classInfo, propertyName, 0, visited, interfaces, accessors, errors);
            }
            catch (Exception ex)
            {
                errors.Add(ex.GetType().Name + ": " + ex.Message);
            }
            bool complete = errors.Count == 0 && interfaces.Count > 0;
            bool setter = false;
            foreach (dynamic accessor in accessors)
                if ((string)accessor.InvocationKind == INVOKEKIND.INVOKE_PROPERTYPUT.ToString() ||
                    (string)accessor.InvocationKind == INVOKEKIND.INVOKE_PROPERTYPUTREF.ToString())
                    setter = true;
            return new { Property = propertyName, Discovery = "IProvideClassInfo/ITypeInfo",
                MetadataComplete = complete, SetterDeclared = complete ? (bool?)setter : null,
                ActualSetterVerified = false, Interfaces = interfaces,
                Accessors = accessors, Errors = errors };
        }

        /// <summary>Inspecte un type COM puis ses interfaces héritées non-source en évitant les GUID déjà visités.</summary>
        /// <param name="info">Informations de type COM à parcourir.</param>
        /// <param name="propertyName">Nom de propriété recherché.</param>
        /// <param name="depth">Profondeur courante de l’héritage.</param>
        /// <param name="visited">GUID déjà inspectés afin d’éviter les cycles.</param>
        /// <param name="interfaces">Liste recevant les noms et GUID d’interfaces.</param>
        /// <param name="accessors">Liste recevant les accesseurs correspondants.</param>
        /// <param name="errors">Liste recevant les erreurs locales de parcours.</param>
        private static void ReadType(ITypeInfo info, string propertyName, int depth,
            HashSet<Guid> visited, List<string> interfaces, List<object> accessors,
            List<string> errors)
        {
            if (depth > 8) { errors.Add("Type inheritance exceeds depth 8."); return; }
            IntPtr attrPointer = IntPtr.Zero;
            try
            {
                info.GetTypeAttr(out attrPointer);
                var attr = (TYPEATTR)Marshal.PtrToStructure(attrPointer, typeof(TYPEATTR));
                if (!visited.Add(attr.guid)) return;
                if (attr.cFuncs > 2048 || attr.cImplTypes > 64)
                { errors.Add("Type exceeds accessor inspection limit: " + attr.guid); return; }
                string name, description, helpFile;
                int helpContext;
                info.GetDocumentation(-1, out name, out description, out helpContext, out helpFile);
                interfaces.Add(name + " (" + attr.guid.ToString("D") + ")");
                for (int index = 0; index < attr.cFuncs; index++)
                {
                    IntPtr functionPointer = IntPtr.Zero;
                    try
                    {
                        info.GetFuncDesc(index, out functionPointer);
                        var function = (FUNCDESC)Marshal.PtrToStructure(functionPointer, typeof(FUNCDESC));
                        var names = new string[1];
                        int count;
                        info.GetNames(function.memid, names, 1, out count);
                        if (count > 0 && string.Equals(names[0], propertyName, StringComparison.OrdinalIgnoreCase))
                            accessors.Add(new { Interface = name, Name = names[0],
                                DispId = function.memid, InvocationKind = function.invkind.ToString() });
                    }
                    catch (Exception ex)
                    {
                        errors.Add(name + " function " + index + ": " + ex.Message);
                    }
                    finally
                    {
                        if (functionPointer != IntPtr.Zero) info.ReleaseFuncDesc(functionPointer);
                    }
                }
                for (int index = 0; index < attr.cImplTypes; index++)
                {
                    try
                    {
                        IMPLTYPEFLAGS flags;
                        info.GetImplTypeFlags(index, out flags);
                        if ((flags & IMPLTYPEFLAGS.IMPLTYPEFLAG_FSOURCE) != 0) continue;
                        int reference;
                        info.GetRefTypeOfImplType(index, out reference);
                        ITypeInfo inherited;
                        info.GetRefTypeInfo(reference, out inherited);
                        ReadType(inherited, propertyName, depth + 1, visited, interfaces, accessors, errors);
                    }
                    catch (Exception ex)
                    {
                        errors.Add(name + " inherited interface " + index + ": " + ex.Message);
                    }
                }
            }
            finally
            {
                if (attrPointer != IntPtr.Zero) info.ReleaseTypeAttr(attrPointer);
            }
        }
    }
}

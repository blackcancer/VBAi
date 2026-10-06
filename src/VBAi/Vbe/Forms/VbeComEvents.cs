using System;
using System.Collections.Generic;
using System.Runtime.InteropServices;
using System.Runtime.InteropServices.ComTypes;
using IMPLTYPEFLAGS = System.Runtime.InteropServices.ComTypes.IMPLTYPEFLAGS;
using TYPEATTR = System.Runtime.InteropServices.ComTypes.TYPEATTR;
using TYPEKIND = System.Runtime.InteropServices.ComTypes.TYPEKIND;
using FUNCDESC = System.Runtime.InteropServices.ComTypes.FUNCDESC;

namespace VBAi
{

    /// <summary>Interface COM qui expose les informations de type de la classe d’un objet.</summary>
    [ComImport, Guid("B196B283-BAB4-101A-B69C-00AA00341D07"),
        InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
    internal interface IProvideClassInfo
    {

        /// <summary>Retourne la description Automation de la classe COM.</summary>
        /// <param name="typeInfo">Reçoit les informations de type.</param>
        void GetClassInfo(out ITypeInfo typeInfo);
    }

    /// <summary>Découvre les événements des interfaces COM source exposées par un contrôle.</summary>
    internal static class VbeComEvents
    {

        /// <summary>Lit les interfaces source via IProvideClassInfo et rapporte les événements ainsi que les limites de la découverte.</summary>
        /// <param name="target">Objet COM à inspecter.</param>
        /// <returns>Rapport contenant les interfaces source, événements, erreurs et indicateurs de complétude.</returns>
        public static object Read(object target)
        {
            var events = new List<object>();
            var sources = new List<object>();
            var errors = new List<string>();
            var provider = target as IProvideClassInfo;
            if (provider == null)
                return new { Discovery = "IProvideClassInfo", SourceInterfacesComplete = false,
                    VbeEventCatalogComplete = false,
                    Scope = "COM source interfaces only; VBE and VBA runtime events may be absent.",
                    Reason = "The COM object does not expose IProvideClassInfo.",
                    Sources = sources, Events = events, Errors = errors };

            // The runtime owns returned COM interfaces. Only GetTypeAttr/GetFuncDesc
            // descriptors require explicit ReleaseTypeAttr/ReleaseFuncDesc calls.
            ITypeInfo classInfo = null;
            try
            {
                provider.GetClassInfo(out classInfo);
                if (classInfo == null)
                    throw new InvalidOperationException("GetClassInfo returned no type information.");
                ReadClass(classInfo, events, sources, errors);
            }
            catch (Exception ex)
            {
                errors.Add(ex.GetType().Name + ": " + ex.Message);
            }
            bool complete = errors.Count == 0 && sources.Count > 0;
            return new { Discovery = "IProvideClassInfo/ITypeInfo source interfaces",
                SourceInterfacesComplete = complete, VbeEventCatalogComplete = false,
                Scope = "COM source interfaces only; VBE and VBA runtime events may be absent (for example UserForm.Initialize).",
                Reason = complete ? null : "No complete COM source event interface was read.",
                Sources = sources, Events = events, Errors = errors };
        }

        /// <summary>Vérifie la coclasse et parcourt ses interfaces source COM dans la limite de 64 interfaces.</summary>
        /// <param name="classInfo">Informations de type de la coclasse.</param>
        /// <param name="events">Liste recevant les événements découverts.</param>
        /// <param name="sources">Liste recevant les interfaces source.</param>
        /// <param name="errors">Liste recevant les erreurs propres aux interfaces.</param>
        private static void ReadClass(ITypeInfo classInfo, List<object> events,
            List<object> sources, List<string> errors)
        {
            IntPtr classAttrPointer = IntPtr.Zero;
            try
            {
                classInfo.GetTypeAttr(out classAttrPointer);
                var classAttr = (TYPEATTR)Marshal.PtrToStructure(classAttrPointer, typeof(TYPEATTR));
                if (classAttr.typekind != TYPEKIND.TKIND_COCLASS)
                    throw new InvalidOperationException("GetClassInfo did not return a coclass type.");
                if (classAttr.cImplTypes > 64)
                    throw new InvalidOperationException("Coclass exposes more than 64 interfaces.");
                for (int index = 0; index < classAttr.cImplTypes; index++)
                {
                    IMPLTYPEFLAGS flags;
                    classInfo.GetImplTypeFlags(index, out flags);
                    if ((flags & IMPLTYPEFLAGS.IMPLTYPEFLAG_FSOURCE) == 0) continue;
                    int href;
                    classInfo.GetRefTypeOfImplType(index, out href);
                    ITypeInfo source = null;
                    try
                    {
                        classInfo.GetRefTypeInfo(href, out source);
                        ReadSource(source, flags, events, sources);
                    }
                    catch (Exception ex)
                    {
                        errors.Add("Source interface " + index + ": " + ex.Message);
                    }
                }
            }
            finally
            {
                if (classAttrPointer != IntPtr.Zero) classInfo.ReleaseTypeAttr(classAttrPointer);
            }
        }

        /// <summary>Ajoute les métadonnées de l’interface source et ses membres événementiels, avec des limites de catalogue.</summary>
        /// <param name="source">Informations de type de l’interface source.</param>
        /// <param name="flags">Indicateurs COM associés à l’interface implémentée.</param>
        /// <param name="events">Liste recevant les événements décrits.</param>
        /// <param name="sources">Liste recevant les interfaces source.</param>
        private static void ReadSource(ITypeInfo source, IMPLTYPEFLAGS flags,
            List<object> events, List<object> sources)
        {
            IntPtr sourceAttrPointer = IntPtr.Zero;
            try
            {
                source.GetTypeAttr(out sourceAttrPointer);
                var attr = (TYPEATTR)Marshal.PtrToStructure(sourceAttrPointer, typeof(TYPEATTR));
                if (attr.cFuncs > 256 || events.Count + attr.cFuncs > 512)
                    throw new InvalidOperationException("Event interface exceeds catalog limit.");
                string interfaceName;
                string description;
                int helpContext;
                string helpFile;
                source.GetDocumentation(-1, out interfaceName, out description,
                    out helpContext, out helpFile);
                sources.Add(new { Name = interfaceName, Guid = attr.guid.ToString("D"),
                    Kind = attr.typekind.ToString(), Flags = flags.ToString(), Count = attr.cFuncs });
                for (int index = 0; index < attr.cFuncs; index++)
                {
                    IntPtr functionPointer = IntPtr.Zero;
                    try
                    {
                        source.GetFuncDesc(index, out functionPointer);
                        var function = (FUNCDESC)Marshal.PtrToStructure(functionPointer, typeof(FUNCDESC));
                        var names = new string[Math.Min(function.cParams + 1, 64)];
                        int namesCount;
                        source.GetNames(function.memid, names, names.Length, out namesCount);
                        if (namesCount == 0 || string.IsNullOrWhiteSpace(names[0]))
                            throw new InvalidOperationException("Event name is absent for member " + function.memid + ".");
                        events.Add(new { Name = names[0], DispId = function.memid,
                            SourceInterface = interfaceName, SourceGuid = attr.guid.ToString("D"),
                            ParameterCount = (int)function.cParams,
                            InvocationKind = function.invkind.ToString() });
                    }
                    finally
                    {
                        if (functionPointer != IntPtr.Zero) source.ReleaseFuncDesc(functionPointer);
                    }
                }
            }
            finally
            {
                if (sourceAttrPointer != IntPtr.Zero) source.ReleaseTypeAttr(sourceAttrPointer);
            }
        }
    }
}

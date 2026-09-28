using System;
using System.Runtime.InteropServices;

namespace CodexVBE
{
    /// <summary>Contrat COM des notifications du cycle de vie d’un complément Office.</summary>
    [ComImport]
    [Guid("B65AD801-ABAF-11D0-BB8B-00A0C90F2744")]
    [InterfaceType(ComInterfaceType.InterfaceIsDual)]
    public interface IDTExtensibility2
    {
        /// <summary>Notifie le complément de sa connexion à l’application hôte.</summary>
        /// <param name="application">Objet Automation de l’application hôte.</param>
        /// <param name="connectMode">Mode de connexion défini par l’hôte.</param>
        /// <param name="addInInstance">Instance Automation du complément enregistrée par l’hôte.</param>
        /// <param name="custom">Arguments personnalisés associés à la connexion.</param>
        [DispId(1)]
        void OnConnection([In, MarshalAs(UnmanagedType.IDispatch)] object application,
            [In] int connectMode,
            [In, MarshalAs(UnmanagedType.IDispatch)] object addInInstance,
            [In, Out, MarshalAs(UnmanagedType.SafeArray, SafeArraySubType = VarEnum.VT_VARIANT)] ref object[] custom);

        /// <summary>Notifie le complément de son retrait de l’application.</summary>
        /// <param name="removeMode">Mode de retrait défini par l’hôte.</param>
        /// <param name="custom">Arguments personnalisés associés au retrait.</param>
        [DispId(2)]
        void OnDisconnection([In] int removeMode,
            [In, Out, MarshalAs(UnmanagedType.SafeArray, SafeArraySubType = VarEnum.VT_VARIANT)] ref object[] custom);

        /// <summary>Notifie le complément d’une mise à jour de la collection de compléments.</summary>
        /// <param name="custom">Arguments personnalisés transmis par l’hôte.</param>
        [DispId(3)]
        void OnAddInsUpdate([In, Out, MarshalAs(UnmanagedType.SafeArray, SafeArraySubType = VarEnum.VT_VARIANT)] ref object[] custom);

        /// <summary>Notifie le complément que le démarrage de l’application est terminé.</summary>
        /// <param name="custom">Arguments personnalisés transmis par l’hôte.</param>
        [DispId(4)]
        void OnStartupComplete([In, Out, MarshalAs(UnmanagedType.SafeArray, SafeArraySubType = VarEnum.VT_VARIANT)] ref object[] custom);

        /// <summary>Notifie le complément que l’application commence son arrêt.</summary>
        /// <param name="custom">Arguments personnalisés transmis par l’hôte.</param>
        [DispId(5)]
        void OnBeginShutdown([In, Out, MarshalAs(UnmanagedType.SafeArray, SafeArraySubType = VarEnum.VT_VARIANT)] ref object[] custom);
    }
}

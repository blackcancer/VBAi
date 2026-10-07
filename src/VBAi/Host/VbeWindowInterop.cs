using System;
using System.Runtime.InteropServices;

namespace VBAi
{
    // The interface IDs and member order come from the installed VBE Extensibility type library.
    // CreateToolWindow requires the typed AddIn interface; late-bound object/IDispatch
    // arguments are rejected by VBE with DISP_E_TYPEMISMATCH.
    /// <summary>Interface COM typée du complément utilisée lors de la création d’une fenêtre VBE.</summary>
    [ComImport]
    [Guid("DA936B64-AC8B-11D1-B6E5-00A0C90F2744")]
    [InterfaceType(ComInterfaceType.InterfaceIsDual)]
    internal interface IVbeAddIn { }

    /// <summary>Projection COM de la collection de fenêtres VBE et de sa fabrique de fenêtres outil.</summary>
    [ComImport]
    [Guid("F57B7ED0-D8AB-11D1-85DF-00C04F98F42C")]
    [InterfaceType(ComInterfaceType.InterfaceIsDual)]
    internal interface IVbeWindows
    {

        /// <summary>Retourne l’instance VBE propriétaire de la collection.</summary>
        /// <value>Instance VBE via l’interface COM.</value>
        [DispId(1)] object VBE { [return: MarshalAs(UnmanagedType.Interface)] get; }

        /// <summary>Retourne le parent COM de la collection de fenêtres.</summary>
        /// <value>Objet parent via l’interface COM.</value>
        [DispId(2)] object Parent { [return: MarshalAs(UnmanagedType.Interface)] get; }

        /// <summary>Accède à une fenêtre par l’index accepté par la collection COM.</summary>
        /// <param name="index">Index ou clé de la fenêtre.</param>
        /// <returns>Fenêtre COM correspondante.</returns>
        [DispId(0)]
        [return: MarshalAs(UnmanagedType.Interface)]
        object Item(
            [In, MarshalAs(UnmanagedType.Struct)] object index);

        /// <summary>Retourne le nombre de fenêtres de la collection.</summary>
        /// <value>Nombre d’éléments.</value>
        [DispId(201)] int Count { get; }

        /// <summary>Fournit l’énumérateur COM de la collection.</summary>
        /// <returns>Énumérateur des fenêtres de la collection.</returns>
        [DispId(-4)] System.Collections.IEnumerator GetEnumerator();

        /// <summary>Crée une fenêtre outil VBE pour le complément et le document fournis.</summary>
        /// <param name="addIn">Interface COM typée du complément enregistré.</param>
        /// <param name="progId">ProgID du contrôle document hébergé.</param>
        /// <param name="caption">Légende de la fenêtre outil.</param>
        /// <param name="position">Position initiale demandée au VBE.</param>
        /// <param name="document">Référence du document COM reçue ou mise à jour par le VBE.</param>
        /// <returns>Fenêtre outil nouvellement créée.</returns>
        [DispId(300)]
        [return: MarshalAs(UnmanagedType.Interface)]
        object CreateToolWindow(
            [In, MarshalAs(UnmanagedType.Interface)] IVbeAddIn addIn,
            [In, MarshalAs(UnmanagedType.BStr)] string progId,
            [In, MarshalAs(UnmanagedType.BStr)] string caption,
            [In, MarshalAs(UnmanagedType.BStr)] string position,
            [In, Out, MarshalAs(UnmanagedType.IDispatch)] ref object document);
    }
}

using System;
using System.IO;
using System.Runtime.InteropServices;
using System.Text.RegularExpressions;

namespace CodexVBE
{
    /// <summary>Interface Automation de lecture des propriétés d’une image OLE.</summary>
    [ComImport, Guid("7BF80981-BF32-101A-8BBB-00AA00300CAB"),
     InterfaceType(ComInterfaceType.InterfaceIsIDispatch)]
    internal interface IOlePictureDisp
    {
        /// <summary>Handle natif de l’image OLE.</summary>
        /// <value>Handle natif de l’image OLE.</value>
        [DispId(0)] int Handle { get; }
        /// <summary>Type déclaré de l’image OLE.</summary>
        /// <value>Type déclaré de l’image OLE.</value>
        [DispId(3)] short Type { get; }
        /// <summary>Largeur exposée par l’image OLE.</summary>
        /// <value>Largeur exposée par l’image OLE.</value>
        [DispId(4)] int Width { get; }
        /// <summary>Hauteur exposée par l’image OLE.</summary>
        /// <value>Hauteur exposée par l’image OLE.</value>
        [DispId(5)] int Height { get; }
    }

    /// <summary>Charge des fichiers d’image comme objets OLE pris en charge par le concepteur.</summary>
    internal static class OlePictureLoader
    {
        /// <summary>Charge une image avec l’API OleAut et retourne son objet IDispatch.</summary>
        /// <param name="fileName">Chemin absolu du fichier image à charger.</param>
        /// <param name="picture">Objet image COM fourni par oleaut32.</param>
        [DllImport("oleaut32.dll", EntryPoint = "OleLoadPictureFile", PreserveSig = false)]
        private static extern void OleLoadPictureFile(
            [MarshalAs(UnmanagedType.Struct)] object fileName,
            [MarshalAs(UnmanagedType.IDispatch)] out object picture);

        /// <summary>Valide le chemin puis charge l’image en objet OLE.</summary>
        /// <param name="path">Chemin absolu du fichier image local ou UNC.</param>
        /// <returns>Objet COM image chargé depuis le chemin validé.</returns>
        public static object Load(string path)
        {
            if (string.IsNullOrWhiteSpace(path) ||
                !Regex.IsMatch(path, @"^(?:[A-Za-z]:[\\/]|\\\\[^\\/]+[\\/][^\\/]+[\\/])"))
                throw new ArgumentException("Picture path must be a fully qualified local or UNC path.");
            string fullPath = System.IO.Path.GetFullPath(path);
            if (!File.Exists(fullPath)) throw new FileNotFoundException("Picture file not found.", fullPath);
            OleLoadPictureFile(fullPath, out object picture);
            if (picture == null || !Marshal.IsComObject(picture))
                throw new InvalidOperationException("Windows did not load an OLE picture from the supplied file.");
            return picture;
        }

        /// <summary>Construit une empreinte à partir du type, des dimensions et du handle OLE.</summary>
        /// <param name="picture">Objet COM de type image OLE à interroger.</param>
        /// <returns>Chaîne combinant type, largeur, hauteur et handle OLE.</returns>
        public static string Fingerprint(object picture)
        {
            if (picture == null) return null;
            var typed = picture as IOlePictureDisp;
            if (typed == null)
                throw new InvalidOperationException("The UserForm Picture is not an OLE picture.");
            return typed.Type + ":" + typed.Width + ":" + typed.Height + ":" + typed.Handle;
        }
    }
}

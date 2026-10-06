using System;
using System.IO;
using System.Globalization;
using System.Runtime.InteropServices;
using System.Runtime.InteropServices.ComTypes;
using System.Security.Cryptography;
using System.Text.RegularExpressions;

namespace VBAi
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

    // IPersistStream inherits IPersist: keep the complete native vtable order.
        /// <summary>Managed declaration of the native IPersistStream vtable used by OLE pictures.</summary>
    [ComImport, Guid("00000109-0000-0000-C000-000000000046"), InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
    internal interface IOlePicturePersistence
    {

        /// <summary>Reads the COM class identifier of the persisted object.</summary>
        /// <param name="classId">Receives the OLE picture class identifier when the HRESULT succeeds.</param>
        /// <returns>The native HRESULT; the declaration preserves it for explicit error handling.</returns>
        [PreserveSig] int GetClassID(out Guid classId);

        /// <summary>Reports whether the persistence object has unsaved changes.</summary>
        /// <returns>The native HRESULT: S_OK when dirty, S_FALSE when clean, or a failure code.</returns>
        [PreserveSig] int IsDirty();

        /// <summary>Loads the object's persisted state from the supplied COM stream.</summary>
        /// <param name="stream">Readable stream containing the object's native persistence format.</param>
        /// <returns>The native HRESULT from the persistence implementation.</returns>
        [PreserveSig] int Load([MarshalAs(UnmanagedType.Interface)] IStream stream);

        /// <summary>Writes the object's persisted state to the supplied COM stream.</summary>
        /// <param name="stream">Writable stream that receives the native persistence bytes.</param>
        /// <param name="clearDirty">Whether the implementation may clear its dirty state after saving.</param>
        /// <returns>The native HRESULT from the persistence implementation.</returns>
        [PreserveSig] int Save([MarshalAs(UnmanagedType.Interface)] IStream stream, [MarshalAs(UnmanagedType.Bool)] bool clearDirty);

        /// <summary>Returns the maximum number of bytes required to persist the object.</summary>
        /// <param name="size">Receives the maximum byte count when the HRESULT succeeds.</param>
        /// <returns>The native HRESULT from the persistence implementation.</returns>
        [PreserveSig] int GetSizeMax(out ulong size);
    }

    /// <summary>Charge des fichiers d’image comme objets OLE pris en charge par le concepteur.</summary>
    internal static class OlePictureLoader
    {

        /// <summary>Maximum accepted serialized image size, in bytes, before allocating persistence buffers.</summary>
        internal const long MaximumPersistenceBytes = 64L * 1024 * 1024;

        /// <summary>Creates an IStream backed by an HGLOBAL according to the OLE ownership flag.</summary>
        /// <param name="memory">HGLOBAL to wrap, or zero to let OLE allocate the backing memory.</param>
        /// <param name="deleteOnRelease">Whether releasing the returned stream also frees its HGLOBAL.</param>
        /// <param name="stream">Receives the COM stream that owns or wraps the specified memory.</param>
        [DllImport("ole32.dll", PreserveSig = false)]
        private static extern void CreateStreamOnHGlobal(IntPtr memory, [MarshalAs(UnmanagedType.Bool)] bool deleteOnRelease,
            [MarshalAs(UnmanagedType.Interface)] out IStream stream);

        /// <summary>Contrat du chargeur natif d’une image OLE Automation.</summary>
        /// <param name="fileName">Chemin du fichier image à lire.</param>
        /// <param name="picture">Objet image OLE retourné.</param>
        internal delegate void PictureReader(object fileName, out object picture);

        /// <summary>Charge l’image avec OleAut par défaut.</summary>
        internal static PictureReader ReadPicture = OleLoadPictureFile;

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
            object picture;
            ReadPicture(fullPath, out picture);
            if (picture == null || !Marshal.IsComObject(picture))
                throw new InvalidOperationException("Windows did not load an OLE picture from the supplied file.");
            return picture;
        }

        /// <summary>Empreinte du contenu persisté OLE, indépendante du handle GDI emprunté.</summary>
        /// <param name="picture">Objet COM de type image OLE à interroger.</param>
        /// <returns>Type, dimensions et SHA-256 de la persistance ; null pour une image vide.</returns>
        public static string Fingerprint(object picture)
        {
            if (picture == null) return null;
            var typed = picture as IOlePictureDisp;
            if (typed == null)
                throw new InvalidOperationException("The UserForm Picture is not an OLE picture.");
            short type = typed.Type;
            if (type == 0 || type == -1) return null;
            int width = typed.Width, height = typed.Height;
            if (type < 1 || type > 4 || width <= 0 || height <= 0)
                throw new InvalidOperationException("Unsupported OLE picture type or dimensions.");
            var persistence = picture as IOlePicturePersistence;
            if (persistence == null) throw new InvalidOperationException("The OLE picture cannot persist its content.");
            ulong maximum;
            Marshal.ThrowExceptionForHR(persistence.GetSizeMax(out maximum));
            if (maximum == 0 || maximum > (ulong)MaximumPersistenceBytes)
                throw new InvalidOperationException("OLE picture persistence exceeds the bounded content budget.");
            int dirtyBefore = persistence.IsDirty();
            Marshal.ThrowExceptionForHR(dirtyBefore);
            IStream stream = null;
            try
            {
                // This stream alone is owned here; the picture and its GDI handle remain borrowed.
                CreateStreamOnHGlobal(IntPtr.Zero, true, out stream);
                Marshal.ThrowExceptionForHR(persistence.Save(stream, false));
                System.Runtime.InteropServices.ComTypes.STATSTG stat;
                stream.Stat(out stat, 1);
                if (stat.cbSize <= 0 || stat.cbSize > MaximumPersistenceBytes || (ulong)stat.cbSize > maximum)
                    throw new InvalidOperationException("OLE picture persisted size differs from its bounded size contract.");
                stream.Seek(0, 0, IntPtr.Zero);
                var bytes = new byte[(int)stat.cbSize];
                IntPtr count = Marshal.AllocCoTaskMem(sizeof(int));
                try
                {
                    Marshal.WriteInt32(count, 0);
                    stream.Read(bytes, bytes.Length, count);
                    if (Marshal.ReadInt32(count) != bytes.Length)
                        throw new InvalidOperationException("OLE picture persistence returned incomplete content.");
                }
                finally { Marshal.FreeCoTaskMem(count); }
                int dirtyAfter = persistence.IsDirty();
                Marshal.ThrowExceptionForHR(dirtyAfter);
                if (dirtyBefore != dirtyAfter || typed.Type != type || typed.Width != width || typed.Height != height)
                    throw new InvalidOperationException("OLE picture metadata changed during content inspection.");
                using (var hash = SHA256.Create())
                    return "ole-persist-v1:" + type.ToString(CultureInfo.InvariantCulture) + ":" +
                        width.ToString(CultureInfo.InvariantCulture) + ":" + height.ToString(CultureInfo.InvariantCulture) + ":" +
                        BitConverter.ToString(hash.ComputeHash(bytes)).Replace("-", string.Empty).ToLowerInvariant();
            }
            finally
            {
                // Never release the borrowed picture RCW or delete its bitmap/icon/metafile.
                if (stream != null) Marshal.ReleaseComObject(stream);
            }
        }
    }
}

using System;
using System.IO;
using System.Runtime.InteropServices;
using System.Security.Cryptography;
using Microsoft.Win32;

namespace CodexVBE
{
    /// <summary>Vérifie la signature du projet VBA enregistré via les SIP Office Microsoft et WinVerifyTrust.</summary>
    internal sealed class VbeSignatureVerifier
    {
        /// <summary>Identifiant du SIP VBA des fichiers Office binaires.</summary>
        internal static readonly Guid LegacySubject = new Guid("01F45160-3E3E-11D3-B49A-00104B2CF645");
        /// <summary>Identifiant du SIP VBA des fichiers Office Open XML.</summary>
        internal static readonly Guid XmlSubject = new Guid("6E64D5BD-CEB0-4B66-B4A0-15AC71775C48");
        /// <summary>Politique Authenticode générique Windows.</summary>
        private static readonly Guid Policy = new Guid("00AAC56B-CD44-11D0-8CC2-00C04FC295EE");

        /// <summary>Informations natives du fichier ; le GUID force le SIP VBA plutôt qu'une signature de document.</summary>
        [StructLayout(LayoutKind.Sequential, CharSet = CharSet.Unicode)]
        internal struct TrustFile
        {
            /// <summary>Taille native de la structure.</summary>
            internal uint Size;
            /// <summary>Chemin absolu Unicode du document.</summary>
            [MarshalAs(UnmanagedType.LPWStr)] internal string Path;
            /// <summary>Handle conservé ouvert en lecture pendant la vérification.</summary>
            internal IntPtr Handle;
            /// <summary>Pointeur vers l'identifiant SIP VBA sélectionné.</summary>
            internal IntPtr Subject;
        }

        /// <summary>Paramètres et état de la politique WinVerifyTrust.</summary>
        [StructLayout(LayoutKind.Sequential)]
        internal struct TrustData
        {
            /// <summary>Taille de la structure.</summary>
            internal uint Size;
            /// <summary>Données facultatives de politique, nulles.</summary>
            internal IntPtr PolicyData;
            /// <summary>Données facultatives du SIP, nulles.</summary>
            internal IntPtr SipData;
            /// <summary>Choix d'interface : WTD_UI_NONE.</summary>
            internal uint Ui;
            /// <summary>Vérification de révocation supplémentaire.</summary>
            internal uint Revocation;
            /// <summary>Type d'entrée : WTD_CHOICE_FILE.</summary>
            internal uint Choice;
            /// <summary>Pointeur vers les informations du fichier.</summary>
            internal IntPtr File;
            /// <summary>Action de vérification ou de libération de l'état.</summary>
            internal uint Action;
            /// <summary>État alloué par la politique Windows.</summary>
            internal IntPtr State;
            /// <summary>Référence URL réservée, nulle.</summary>
            internal IntPtr Url;
            /// <summary>Révocation en cache et interdiction des téléchargements.</summary>
            internal uint Flags;
            /// <summary>Contexte d'interface, zéro.</summary>
            internal uint Context;
            /// <summary>Paramètres facultatifs de signature Windows, nuls.</summary>
            internal IntPtr SignatureSettings;
        }

                /// <summary>Frontière injectable de WinVerifyTrust, conservant les structures natives.</summary>
        /// <param name="owner">Handle propriétaire transmis à WinVerifyTrust.</param>
        /// <param name="action">Identifiant de la politique d’action, passé par référence.</param>
        /// <param name="data">Structure native des options et de l’état de vérification.</param>
        /// <returns>Code HRESULT natif de WinVerifyTrust.</returns>
        internal delegate int VerifyCall(IntPtr owner, ref Guid action, ref TrustData data);
                /// <summary>Appel Windows de vérification et de fermeture.</summary>
        /// <param name="owner">Handle propriétaire de l’appel.</param>
        /// <param name="action">Identifiant de la politique d’action, passé par référence.</param>
        /// <param name="data">Données natives du fichier et options de vérification.</param>
        /// <returns>Code HRESULT retourné par WinVerifyTrust.</returns>
        [DllImport("wintrust.dll", ExactSpelling = true, PreserveSig = true)]
        private static extern int WinVerifyTrust(IntPtr owner, ref Guid action, ref TrustData data);
        /// <summary>Appel natif par défaut, injectable pour les erreurs de politique.</summary>
        internal VerifyCall Native = WinVerifyTrust;
        /// <summary>Lecture du SIP enregistré, sans modification du registre.</summary>
        internal Func<Guid, string> Provider = ReadProvider;

        /// <summary>Vérifie le fichier sur disque sans exécuter VBA, accéder aux clés privées ou enregistrer le document.</summary>
        /// <param name="path">Chemin absolu du document Office enregistré.</param>
        /// <returns>SHA du fichier, disponibilité du SIP et résultat distinct de la présence d'une signature.</returns>
        internal object Verify(string path)
        {
            if (string.IsNullOrWhiteSpace(path) || !System.IO.Path.IsPathRooted(path))
                throw new ArgumentException("An absolute saved Office document Path is required.");
            path = System.IO.Path.GetFullPath(path);
            string extension = System.IO.Path.GetExtension(path).ToLowerInvariant();
            Guid subject;
            string library;
            switch (extension)
            {
                case ".xlam": case ".xlsb": case ".xlsm": case ".xltm":
                case ".potm": case ".ppam": case ".ppsm": case ".pptm":
                case ".vsdm": case ".vssm": case ".vstm": case ".docm": case ".dotm":
                    subject = XmlSubject; library = "msosipx.dll"; break;
                case ".xla": case ".xls": case ".xlt": case ".pot": case ".ppa": case ".pps": case ".ppt":
                case ".mpp": case ".mpt": case ".pub": case ".vdw": case ".vsd": case ".vss": case ".vst":
                case ".doc": case ".dot": case ".wiz":
                    subject = LegacySubject; library = "msosip.dll"; break;
                default: throw new NotSupportedException("The Microsoft Office VBA SIP does not support this format; standalone SWP and Access are not treated as Office signatures.");
            }
            using (var stream = new FileStream(path, FileMode.Open, FileAccess.Read, FileShare.Read))
            {
                if (stream.Length == 0 || stream.Length > 512L * 1024 * 1024)
                    throw new InvalidOperationException("Signature inspection accepts nonempty saved files up to 512 MiB.");
                string hash;
                using (var sha = SHA256.Create()) hash = BitConverter.ToString(sha.ComputeHash(stream)).Replace("-", "").ToLowerInvariant();
                stream.Position = 0;
                string provider = Provider(subject);
                if (string.IsNullOrWhiteSpace(provider) || !System.IO.Path.IsPathRooted(provider) ||
                    !System.IO.Path.GetFileName(provider).Equals(library, StringComparison.OrdinalIgnoreCase) || !File.Exists(provider))
                    return new { Path = path, FileSha256 = hash, Available = false, Status = "VerifierUnavailable",
                        Trusted = (bool?)null, SignatureValid = (bool?)null, SipSubject = subject.ToString("B"),
                        RequiredComponent = library, Reason = "Register the Microsoft Office x64 Subject Interface Package before VBA signature verification. No certificate-only check is substituted." };
                int code = VerifyNative(path, stream.SafeFileHandle.DangerousGetHandle(), subject);
                bool? valid = code == 0 ? true : code == unchecked((int)0x80096010) || code == unchecked((int)0x80096004) ? (bool?)false : null;
                return new { Path = path, FileSha256 = hash, Available = true, Status = Status(code),
                    Trusted = code == 0, SignatureValid = valid, NativeStatus = "0x" + unchecked((uint)code).ToString("X8"),
                    SipSubject = subject.ToString("B"), Provider = provider, NetworkRetrieval = false,
                    Verification = "Windows Authenticode policy with Microsoft Office VBA SIP; strongest present VBA signature",
                    Limit = "Saved disk file only; unsaved VBE edits are not verified. Microsoft SIP selects V3, then agile, then legacy. A nonzero policy result is never reported as valid or trusted; cached revocation may be indeterminate." };
            }
        }

                /// <summary>Prépare les structures SIP, vérifie le fichier et libère toujours l’état WinTrust.</summary>
        /// <param name="path">Chemin du document enregistré.</param>
        /// <param name="handle">Handle ouvert en lecture du fichier.</param>
        /// <param name="subject">Identifiant du sujet Office SIP approprié au format.</param>
        /// <returns>Code HRESULT de la vérification de confiance native.</returns>
        private int VerifyNative(string path, IntPtr handle, Guid subject)
        {
            IntPtr subjectBuffer = Marshal.AllocHGlobal(Marshal.SizeOf(typeof(Guid)));
            IntPtr fileBuffer = Marshal.AllocHGlobal(Marshal.SizeOf(typeof(TrustFile)));
            var file = new TrustFile { Size = (uint)Marshal.SizeOf(typeof(TrustFile)), Path = path, Handle = handle, Subject = subjectBuffer };
            Marshal.StructureToPtr(subject, subjectBuffer, false);
            Marshal.StructureToPtr(file, fileBuffer, false);
            var data = new TrustData { Size = (uint)Marshal.SizeOf(typeof(TrustData)), Ui = 2, Choice = 1,
                File = fileBuffer, Action = 1, Flags = 0x1000 | 0x80 };
            var policy = Policy;
            try { return Native(new IntPtr(-1), ref policy, ref data); }
            finally
            {
                try { data.Action = 2; Native(new IntPtr(-1), ref policy, ref data); }
                finally { Marshal.DestroyStructure(fileBuffer, typeof(TrustFile)); Marshal.FreeHGlobal(fileBuffer); Marshal.FreeHGlobal(subjectBuffer); }
            }
        }

                /// <summary>Lit la bibliothèque du SIP Office x64 enregistré pour la vérification du digest.</summary>
        /// <param name="subject">Identifiant du sujet SIP Office.</param>
        /// <returns>Chemin configuré de la bibliothèque, ou <see langword="null"/> si la valeur est absente ou illisible.</returns>
        private static string ReadProvider(Guid subject)
        {
            using (var registry = RegistryKey.OpenBaseKey(RegistryHive.LocalMachine, RegistryView.Registry64))
            using (var key = registry.OpenSubKey(@"SOFTWARE\Microsoft\Cryptography\OID\EncodingType 0\CryptSIPDllVerifyIndirectData\" + subject.ToString("B")))
                return key?.GetValue("Dll") as string;
        }

                /// <summary>Classe les résultats sans confondre absence, altération et confiance incomplète.</summary>
        /// <param name="result">Code HRESULT produit par WinVerifyTrust.</param>
        /// <returns>Étiquette stable correspondant aux codes reconnus, ou <c>VerificationFailed</c>.</returns>
        private static string Status(int result)
        {
            switch (unchecked((uint)result))
            {
                case 0: return "Trusted";
                case 0x800B0100: return "NoSignature";
                case 0x80096010: return "BadDigest";
                case 0x80096004: return "InvalidSignature";
                case 0x800B0109: return "UntrustedRoot";
                case 0x800B0101: return "ExpiredCertificate";
                case 0x800B010C: return "RevokedCertificate";
                case 0x80092013: case 0x800B010E: return "RevocationIndeterminate";
                default: return "VerificationFailed";
            }
        }
    }
}

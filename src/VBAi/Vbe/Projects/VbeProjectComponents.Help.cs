using System;
using System.Globalization;
using System.IO;
using System.Runtime.InteropServices;

namespace VBAi
{
    /// <summary>Ouvre l’aide locale configurée dans un projet VBA par l’API Windows HTML Help.</summary>
    internal sealed partial class VbeProjectComponents
    {
        /// <summary>Appelle l’API native HTML Help pour un fichier et un contexte de rubrique.</summary>
        /// <param name="owner">Handle de la fenêtre propriétaire, nul pour cet appel.</param>
        /// <param name="file">Chemin du fichier d’aide CHM.</param>
        /// <param name="command">Commande HTML Help à exécuter.</param>
        /// <param name="data">Contexte numérique de la rubrique.</param>
        /// <returns>Handle de la fenêtre d’aide retourné par HTML Help.</returns>
        [DllImport("hhctrl.ocx", EntryPoint = "HtmlHelpW", CharSet = CharSet.Unicode)]
        private static extern IntPtr HtmlHelp(IntPtr owner, string file, uint command, UIntPtr data);

        /// <summary>Frontière de l'aide CHM native ; un handle n'atteste pas la lecture de la rubrique.</summary>
        internal Func<string, uint, IntPtr> HelpLauncher = (path, context) => NativeHelp(IntPtr.Zero, path, context == 0 ? 0U : 15U, new UIntPtr(context));
        /// <summary>Invokes HtmlHelp by default; permits owned boundary checks without opening a help window.</summary>
        internal static Func<IntPtr, string, uint, UIntPtr, IntPtr> NativeHelp = HtmlHelp;

                /// <summary>Ouvre le fichier CHM et le contexte configurés dans le projet après contrôle de version.</summary>
                /// <param name="request">Requête contenant le projet et la version attendue de ses propriétés.</param>
                /// <returns>Le chemin et le contexte invoqués, ainsi que l’indication de création d’une fenêtre.</returns>
        public object OpenProjectHelp(Request request)
        {
            if (request == null || string.IsNullOrWhiteSpace(request.Project) || string.IsNullOrWhiteSpace(request.ExpectedProjectVersion))
                throw new ArgumentException("Project and ExpectedProjectVersion are required.");
            dynamic project = GetProject(request.Project);
            AssertProjectVersion(request, project);
            string file = (string)project.HelpFile;
            if (string.IsNullOrWhiteSpace(file) || !Path.IsPathRooted(file) || !Path.GetExtension(file).Equals(".chm", StringComparison.OrdinalIgnoreCase))
                throw new InvalidOperationException("The project must reference an absolute CHM help file. Legacy HLP and web help are not substituted.");
            file = Path.GetFullPath(file);
            if (!File.Exists(file)) throw new FileNotFoundException("The project help file is absent.", file);
            if (!uint.TryParse(Convert.ToString(project.HelpContextID, CultureInfo.InvariantCulture), NumberStyles.None, CultureInfo.InvariantCulture, out uint context))
                throw new InvalidOperationException("The project help context is not an unsigned integer.");
            IntPtr window = HelpLauncher(file, context);
            return new { request.Project, HelpFile = file, HelpContextID = context,
                Invoked = true, HelpWindowCreated = window != IntPtr.Zero, TopicVerified = false,
                Limit = "A help window handle does not verify that the requested context exists or that CHM security permits its content." };
        }
    }
}

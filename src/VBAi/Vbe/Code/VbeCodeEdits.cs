using System;
using System.Collections.Generic;
using System.Linq;
using System.Security.Cryptography;
using System.Text;

namespace VBAi
{

    /// <summary>Applique des modifications VBA bornées et conserve un historique de session distinct de l’historique natif VBE.</summary>
    internal sealed partial class VbeCodeEdits
    {

        /// <summary>Exécuteur des commandes de lecture et remplacement des modules.</summary>
        private readonly Func<Request, Response> execute;

        /// <summary>Entrées d’édition disponibles pour annulation.</summary>
        private readonly List<Entry> undo = new List<Entry>();

        /// <summary>Entrées d’édition disponibles pour rétablissement.</summary>
        private readonly List<Entry> redo = new List<Entry>();

        /// <summary>Indique qu’une écriture provient d’une annulation ou d’un rétablissement.</summary>
        private bool replaying;

        /// <summary>Transition de source conservée dans l’historique local à la session.</summary>
        private sealed class Entry
        {

            /// <summary>Stores the project and module identity with complete source snapshots before and after the edit.</summary>
            internal string Project, Module, Before, After;
        }

        /// <summary>Crée l’éditeur transactionnel avec un exécuteur de commandes VBE.</summary>
        /// <param name="execute">Transport des commandes VBE.</param>
        internal VbeCodeEdits(Func<Request, Response> execute) { this.execute = execute; }

        /// <summary>Ajoute une transition de code à l’historique borné de session et vide le rétablissement.</summary>
        /// <param name="project">Projet associé à l’édition.</param>
        /// <param name="module">Module associé à l’édition.</param>
        /// <param name="before">Texte source avant modification.</param>
        /// <param name="after">Texte source après modification.</param>
        internal void Record(string project, string module, string before, string after)
        {
            if (replaying || before == after) return;
            undo.Add(new Entry { Project = project, Module = module, Before = before, After = after });
            redo.Clear();
            // Bounded, session-local history. Never pretend this is the native VBE undo stack.
            while (undo.Count > 50 || undo.Sum(x => (long)x.Before.Length + x.After.Length) > 4 * 1024 * 1024) undo.RemoveAt(0);
        }

        /// <summary>Transforme une plage explicite du module ou renvoie un aperçu avant écriture.</summary>
        /// <param name="request">Action textuelle, plage sélectionnée et empreinte source attendue.</param>
        /// <param name="preview">Si true, retourne les sources avant/après sans mutation.</param>
        /// <returns>Aperçu détaillé ou résultat de remplacement du module.</returns>
        internal object Edit(Request request, bool preview)
        {
            string before = Read(request.Project, request.Module);
            Check(before, request.ExpectedSha256);
            string after = VbaTextEdits.Transform(before, request);
            if (preview) return new
            {
                request.Project,
                request.Module,
                Before = before,
                After = after,
                ExpectedSha256 = Hash(before),
                Changed = before != after,
                Scope = "Explicit module line range; identifier replacements are lexical, not semantic refactoring."
            };
            return Write(request.Project, request.Module, before, after);
        }

        /// <summary>Prévisualise ou applique un renommage local lié à une déclaration et à la plage VBIDE.</summary>
        /// <param name="request">Module, procédure, déclaration exacte, nouveau nom et versions attendues.</param>
        /// <param name="preview">Si true, retourne le plan sans écrire le module.</param>
        /// <returns>Plan avant/après ou résultat vérifié de l’écriture.</returns>
        /// <exception cref="InvalidOperationException">La procédure ou déclaration est ambiguë, absente ou a changé.</exception>
        internal object RenameLocal(Request request, bool preview)
        {
            string before = Read(request.Project, request.Module); Check(before, request.ExpectedSha256);
            Response catalog = execute(new Request { Command = "list_procedures", Project = request.Project, Module = request.Module });
            if (!catalog.Ok) throw new InvalidOperationException(catalog.Error);
            dynamic data = catalog.Data;
            if (!string.Equals((string)data.Sha256, request.ExpectedSha256, StringComparison.OrdinalIgnoreCase))
                throw new InvalidOperationException("The module changed before resolving the procedure.");
            dynamic selected = null;
            foreach (dynamic procedure in data.Procedures)
                if ((int)procedure.Kind == request.ProcKind && string.Equals((string)procedure.Name, request.Procedure, StringComparison.OrdinalIgnoreCase))
                { if (selected != null) throw new InvalidOperationException("The procedure is ambiguous."); selected = procedure; }
            if (selected == null) throw new InvalidOperationException("The exact procedure is absent.");
            string after = VbaLocalRename.Transform(before, request, (int)selected.BodyLine, (int)selected.EndLine);
            if (preview) return new
            {
                request.Project,
                request.Module,
                request.Procedure,
                Before = before,
                After = after,
                ExpectedSha256 = Hash(before),
                Changed = before != after,
                Scope = "One explicit local variable/constant in its VBIDE procedure range; members/types/labels/named arguments are excluded. Parameters, conditional code and project-wide refactoring are refused."
            };
            if (request.ExpectedMode != 2) throw new ArgumentException("ExpectedMode=2 is required to rename a local declaration.");
            Response state = execute(new Request { Command = "debug_state", Project = request.Project });
            if (!state.Ok || (int)((dynamic)state.Data).Mode != 2) throw new InvalidOperationException("Renaming requires design mode.");
            return Write(request.Project, request.Module, before, after);
        }

        /// <summary>Annule ou rétablit la dernière édition de session compatible avec le module et son SHA actuel.</summary>
        /// <param name="request">Projet, module et empreinte du code courant.</param>
        /// <param name="forward">Si true, rejoue une entrée d’annulation; sinon, annule une entrée de rétablissement.</param>
        /// <returns>Résultat de l’écriture et déplacement de l’entrée entre les piles d’historique.</returns>
        /// <exception cref="InvalidOperationException">Aucune entrée ne correspond ou la source diffère de la transition attendue.</exception>
        internal object Replay(Request request, bool forward)
        {
            var source = forward ? redo : undo; var destination = forward ? undo : redo;
            var entry = source.LastOrDefault(x => string.Equals(x.Project, request.Project, StringComparison.OrdinalIgnoreCase) && string.Equals(x.Module, request.Module, StringComparison.OrdinalIgnoreCase)) ?? throw new InvalidOperationException("No VBAi code edit is available for this module in this session.");
            string current = Read(request.Project, request.Module);
            Check(current, request.ExpectedSha256);
            Check(current, Hash(forward ? entry.Before : entry.After));
            replaying = true;
            try
            {
                object result = Write(request.Project, request.Module, current, forward ? entry.After : entry.Before);
                source.Remove(entry); destination.Add(entry); return result;
            }
            finally { replaying = false; }
        }

        /// <summary>Remplace tout le code avec l’empreinte attendue, puis exige une relecture identique.</summary>
        /// <param name="project">Projet cible.</param>
        /// <param name="module">Module cible.</param>
        /// <param name="before">Texte source protégé par l’empreinte.</param>
        /// <param name="after">Nouveau texte source complet.</param>
        /// <returns>Données de résultat du transport VBE après vérification de la relecture.</returns>
        /// <exception cref="InvalidOperationException">L’écriture échoue ou la source relue diffère du texte demandé.</exception>
        private object Write(string project, string module, string before, string after)
        {
            Response result = execute(new Request
            {
                Command = "replace_lines",
                Project = project,
                Module = module,
                ExpectedSha256 = Hash(before),
                StartLine = 1,
                Count = CodeRollback.Lines(before).Length,
                Text = after
            });
            if (!result.Ok) throw new InvalidOperationException(result.Error);
            string readback = Read(project, module);
            if (!string.Equals(readback, after, StringComparison.Ordinal))
                throw new InvalidOperationException("VBE text differs from the requested edit. Read the module before continuing.");
            return result.Data;
        }

        /// <summary>Lit le code d’un module par le transport VBE et propage ses erreurs.</summary>
        /// <param name="project">Projet propriétaire.</param>
        /// <param name="module">Nom du module.</param>
        /// <returns>Code courant du module.</returns>
        private string Read(string project, string module)
        {
            Response result = execute(new Request { Command = "read_module", Project = project, Module = module });
            if (!result.Ok) throw new InvalidOperationException(result.Error);
            return (string)((dynamic)result.Data).Code;
        }

        /// <summary>Compare le code à l’empreinte attendue et refuse les modifications concurrentes.</summary>
        /// <param name="code">Texte courant du module.</param>
        /// <param name="expected">SHA-256 attendu.</param>
        /// <exception cref="InvalidOperationException">L’empreinte manque ou ne correspond pas.</exception>
        private static void Check(string code, string expected)
        {
            if (string.IsNullOrWhiteSpace(expected) || !string.Equals(Hash(code), expected, StringComparison.OrdinalIgnoreCase))
                throw new InvalidOperationException("The module changed. Read its current code and SHA before editing or replaying history.");
        }

        /// <summary>Calcule l’empreinte SHA-256 hexadécimale minuscule du texte UTF-8.</summary>
        /// <param name="text">Code source à empreinter.</param>
        /// <returns>Empreinte de 64 caractères hexadécimaux.</returns>
        private static string Hash(string text)
        { using (var sha = SHA256.Create()) return BitConverter.ToString(sha.ComputeHash(Encoding.UTF8.GetBytes(text))).Replace("-", "").ToLowerInvariant(); }
    }
}

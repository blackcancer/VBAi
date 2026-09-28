using System;
using System.Security.Cryptography;
using System.Text;

namespace CodexVBE
{
    /// <summary>Contrat d’accès au code d’un module natif de l’éditeur Visual Basic.</summary>
internal interface IEditorModule
    {
        /// <summary>Obtient le nom affiché du module.</summary>
        /// <value>Nom fourni par l’adaptateur natif.</value>
string Name { get; }
        /// <summary>Obtient la clé stable utilisée pour retrouver le document.</summary>
        /// <value>Clé du module dans la session de l’éditeur.</value>
string Key { get; }
        /// <summary>Lit le code actuel dans l’hôte natif.</summary>
        /// <returns>Texte source du module.</returns>
string Read();
        /// <summary>Indique si le module peut être modifié dans l’état actuel de l’hôte.</summary>
        /// <value><see langword="true"/> si l’adaptateur autorise l’écriture.</value>
bool CanWrite { get; }
        /// <summary>Remplace le code si sa version native correspond encore au texte attendu.</summary>
        /// <param name="expected">Texte natif servant de garde contre les modifications concurrentes.</param>
        /// <param name="text">Nouveau texte source.</param>
        /// <returns>Texte relu après la tentative d’écriture.</returns>
string Write(string expected, string text);
        /// <summary>Affiche le module natif à la position demandée.</summary>
        /// <param name="line">Numéro de ligne, indexé à partir de un.</param>
        /// <param name="column">Numéro de colonne, indexé à partir de un.</param>
void ShowNative(int line, int column);
    }

    /// <summary>A revisioned buffer. External edits never silently replace a dirty buffer.</summary>
    internal sealed class EditorDocument
    {
        /// <summary>Longueur maximale acceptée pour un document source.</summary>
internal const int MaxLength = 2 * 1024 * 1024;
        /// <summary>Identifiant propre à cette instance de document.</summary>
internal readonly string Id = Guid.NewGuid().ToString("N");
        /// <summary>Adaptateur du module natif associé au document.</summary>
internal readonly IEditorModule Module;
        /// <summary>Clé permettant de retrouver le brouillon enregistré après rechargement.</summary>
        /// <value>Clé de récupération dérivée de l’adaptateur associé.</value>
internal string RecoveryKey { get; private set; }
        /// <summary>Version du texte sur laquelle le brouillon est basé.</summary>
        /// <value>Texte source servant de base aux comparaisons et à la synchronisation.</value>
internal string Baseline { get; private set; }
        /// <summary>Texte actuellement présenté dans l’éditeur.</summary>
        /// <value>Contenu normalisé du brouillon.</value>
internal string Text { get; private set; }
        /// <summary>Dernier texte observé dans le module natif.</summary>
        /// <value>Contenu normalisé lu lors de la dernière observation.</value>
internal string Native { get; private set; }
        /// <summary>Indique que le texte natif et le brouillon ont divergé depuis la base commune.</summary>
        /// <value><see langword="true"/> lorsqu’une résolution explicite est nécessaire.</value>
internal bool Conflict { get; private set; }
        /// <summary>Indique que le brouillon diffère de sa version de base.</summary>
        /// <value><see langword="true"/> lorsque le document contient des modifications locales.</value>
internal bool Dirty => Text != Baseline;
        /// <summary>Indique si l’adaptateur autorise actuellement une écriture native.</summary>
        /// <value>Valeur courante de la capacité d’écriture du module.</value>
internal bool Writable => Module.CanWrite;
        /// <summary>Crée un document à partir du code actuellement lu dans le module.</summary>
        /// <param name="module">Adaptateur du module source.</param>
internal EditorDocument(IEditorModule module)
        { Module = module; RecoveryKey = module.Key; Baseline = Text = Native = Normalize(module.Read()); Validate(Text); }
        /// <summary>Valide et adopte le texte modifié par l’utilisateur.</summary>
        /// <param name="value">Nouveau contenu du brouillon.</param>
internal void Edit(string value) { Validate(value); Text = Normalize(value); }
        /// <summary>Restaure la base et le brouillon sauvegardés, puis observe le module natif.</summary>
        /// <param name="baseline">Version de base enregistrée.</param>
        /// <param name="draft">Contenu de brouillon enregistré.</param>
internal void Restore(string baseline, string draft)
        { Validate(baseline); Validate(draft); Baseline = Normalize(baseline); Text = Normalize(draft); Observe(); }
        /// <summary>Relit le module, actualise l’état de conflit et détecte une mise à jour distante propre.</summary>
        /// <returns>Nouveau texte natif à appliquer lorsque le brouillon est propre, sinon <see langword="null"/>.</returns>
internal string Observe()
        {
            Native = Normalize(Module.Read()); RecoveryKey = Module.Key;
            Conflict = Dirty && Native != Baseline && Native != Text;
            return !Dirty && Native != Baseline ? Native : null;
        }
        // Call only after the renderer accepted the guarded replacement.
        /// <summary>Accepte une nouvelle version native déjà validée par le rendu.</summary>
        /// <param name="code">Texte source accepté.</param>
internal void AcceptRemote(string code) { Baseline = Text = Native = Normalize(code); Conflict = false; }
        /// <summary>Enregistre le résultat natif d’une écriture sans écraser un brouillon modifié depuis sa capture.</summary>
        /// <param name="code">Texte relu après l’écriture.</param>
        /// <param name="capturedText">Texte du brouillon au moment où l’écriture a commencé.</param>
internal void Acknowledge(string code, string capturedText)
        { Baseline = Native = Normalize(code); if (Text == capturedText) Text = Baseline; Conflict = false; }
        /// <summary>Synchronise le brouillon avec le module après vérification des conflits et du plan éventuel.</summary>
        /// <param name="plan">Plan préparé facultatif, qui doit correspondre exactement à la base et au brouillon courants.</param>
        /// <returns>Texte effectivement relu après synchronisation.</returns>
        /// <exception cref="InvalidOperationException">Un conflit, une indisponibilité ou un changement du plan empêche l’écriture.</exception>
internal string Synchronize(EditorSyncPlan plan = null)
        {
            Observe();
            if (!Dirty) return Text;
            if (Conflict) throw new InvalidOperationException("The module changed in VBA. Resolve the conflict first.");
            if (!Writable) throw new InvalidOperationException("VBA is running, paused or unavailable. Your draft is preserved.");
            if (Native == Text) { Baseline = Text; return Text; }
            if (plan != null && (plan.Before != Baseline || plan.After != Text))
                throw new InvalidOperationException("The draft changed while synchronization was being prepared.");
            string actual = Normalize(Module is EditorVbeModule nativeModule
                ? nativeModule.WritePrepared(Baseline, Text, plan) : Module.Write(Baseline, Text));
            // Advance the baseline only after a successful write and readback.
            Baseline = Native = actual;
            return actual;
        }
        /// <summary>Résout un conflit en remplaçant la version native explicitement comparée par le brouillon.</summary>
        /// <param name="reviewedNative">Texte natif que l’utilisateur a comparé et choisi de remplacer.</param>
        /// <returns>Texte relu après résolution.</returns>
        /// <exception cref="InvalidOperationException">Le module a changé depuis la comparaison ou ne peut pas être écrit.</exception>
internal string ResolveWithDraft(string reviewedNative)
        {
            Observe();
            if (Native != reviewedNative) throw new InvalidOperationException("The module changed again. Compare with VBA before resolving.");
            if (!Writable) throw new InvalidOperationException("VBA is running, paused or unavailable. Your draft is preserved.");
            string actual = Normalize(Module.Write(reviewedNative, Text));
            Baseline = Native = actual; Conflict = false; return actual;
        }
        /// <summary>Convertit toute fin de ligne en LF pour comparer les sources sans dépendre de la plateforme.</summary>
        /// <param name="text">Texte à normaliser.</param>
        /// <returns>Texte normalisé, ou chaîne vide si l’entrée est nulle.</returns>
internal static string Normalize(string text) => (text ?? "").Replace("\r\n", "\n").Replace("\r", "\n");
        /// <summary>Calcule l’empreinte SHA-256 hexadécimale minuscule du texte UTF-8.</summary>
        /// <param name="text">Texte à hacher.</param>
        /// <returns>Empreinte de 64 caractères.</returns>
internal static string Hash(string text)
        { using (var sha = SHA256.Create()) return BitConverter.ToString(sha.ComputeHash(Encoding.UTF8.GetBytes(text))).Replace("-", "").ToLowerInvariant(); }
        /// <summary>Refuse les textes nuls, trop longs ou contenant un caractère nul.</summary>
        /// <param name="text">Texte source à contrôler.</param>
        /// <exception cref="InvalidOperationException">Le texte ne respecte pas les limites du document.</exception>
internal static void Validate(string text)
        { if (text == null || text.Length > MaxLength || text.IndexOf('\0') >= 0) throw new InvalidOperationException("The editor document is invalid or too large."); }
        /// <summary>Construit le plus petit remplacement contigu entre deux sources en conservant leur préfixe et suffixe communs.</summary>
        /// <param name="before">Source avant modification.</param>
        /// <param name="after">Source après modification.</param>
        /// <returns>Début de ligne indexé à un, nombre de lignes à remplacer et texte de remplacement CRLF.</returns>
internal static Tuple<int, int, string> Difference(string before, string after)
        {
            var a = before.Length == 0 ? new string[0] : Normalize(before).Split('\n');
            var b = after.Length == 0 ? new string[0] : Normalize(after).Split('\n');
            int prefix = 0, suffix = 0;
            while (prefix < a.Length && prefix < b.Length && a[prefix] == b[prefix]) prefix++;
            while (suffix < a.Length - prefix && suffix < b.Length - prefix && a[a.Length - 1 - suffix] == b[b.Length - 1 - suffix]) suffix++;
            return Tuple.Create(prefix + 1, a.Length - prefix - suffix, string.Join("\r\n", b, prefix, b.Length - prefix - suffix));
        }
    }
}

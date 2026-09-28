using System.Web.Script.Serialization;
namespace CodexVBE
{
    /// <summary>Décrit le résultat d’une opération de restauration d’un formulaire découpé.</summary>
internal sealed class FormCutChange
    {
        /// <summary>Projet contenant le formulaire.</summary>
        /// <value>Nom ou sélecteur du projet concerné.</value>
public string Project { get; set; }
        /// <summary>Nom du formulaire concerné.</summary>
        /// <value>Nom de composant du formulaire.</value>
public string Form { get; set; }
        /// <summary>Chemin du contrôle parent utilisé pour localiser le formulaire.</summary>
        /// <value>Chemin hiérarchique du contrôle parent.</value>
public string ParentPath { get; set; }
        /// <summary>Nombre de contrôles contenus dans la modification.</summary>
        /// <value>Nombre observé ou restauré de contrôles.</value>
public int ControlCount { get; set; }
        /// <summary>Indique si la restauration a été confirmée.</summary>
        /// <value><see langword="true"/> lorsque la restauration est vérifiée.</value>
public bool Restored { get; set; }
        /// <summary>Indique si une tentative de restauration a été effectuée.</summary>
        /// <value><see langword="true"/> dès qu’une opération de restauration a été tentée.</value>
public bool Attempted { get; set; }
        /// <summary>Identifiant interne de récupération exclu de la sérialisation JSON.</summary>
        /// <value>Identifiant du dossier de récupération associé.</value>
[ScriptIgnore] public string RecoveryId { get; set; }
        /// <summary>Outil interne propriétaire de l’opération, exclu du JSON.</summary>
        /// <value>Instance qui a construit cette description de résultat.</value>
[ScriptIgnore] public LlmVbeTools Owner { get; set; }
    }
}

# Organisation du code

Le projet `CodexVBE.csproj` produit l'assembly COM `CodexVBE.dll`. Les dossiers sous `src` décrivent les responsabilités, sans changer le namespace public `CodexVBE` ni le ProgID, le CLSID et la bibliothèque de types utilisés par l'installeur.

| Dossier | Responsabilité |
| --- | --- |
| `src/Host` | Connexion du complément à VBE, menus natifs et journal de chargement. |
| `src/Bridge` | Protocole local et serveur de commandes dans le processus hôte. |
| `src/Vbe/Code` | Navigation et recherche dans les modules VBA. |
| `src/Vbe/Debug` | Commandes natives du débogueur et lecture de ses fenêtres. |
| `src/Vbe/Forms` | UserForms, contrôles, propriétés, conteneurs et listes. |
| `src/Vbe/Projects` | Composants et références des projets. |
| `src/Vbe/Windows` | Fenêtres et volets de code VBIDE. |
| `src/Vbe/VbeSession.cs` | Coordination des commandes VBE. |
| `src/Llm/Chat` | Conversation, contexte VBE et outils appelables par le modèle. |
| `src/Llm/Providers` | Fournisseurs, compte Codex et catalogues de modèles. |
| `src/Llm/Settings` | Configuration et son formulaire WinForms. |

Les paires `.cs`, `.Designer.cs` et `.resx` restent réunies dans leur dossier ; le projet conserve les métadonnées `SubType` et `DependentUpon` afin que Visual Studio ouvre les deux formulaires dans son concepteur. La règle IDE0130 est désactivée seulement sous `src` : le namespace `CodexVBE` reste stable malgré les dossiers fonctionnels.

## Séparation éventuelle en assemblies

Une extraction de bibliothèques distinctes serait utile si les contrats entre VBIDE, passerelle et interface étaient stables et testables sans l'hôte. Aujourd'hui, `VbeSession` et `LlmVbeTools` traversent ces frontières et l'installeur enregistre un chemin `CodeBase` vers une seule DLL. Une séparation immédiate imposerait de nouveaux contrats publics, des références de projets et le déploiement vérifié de DLL supplémentaires dans Excel et SOLIDWORKS. Le regroupement physique clarifie d'abord ces frontières tout en gardant le chargement COM actuel. Une extraction future doit être accompagnée d'un test d'installation et de chargement réel dans les deux hôtes.

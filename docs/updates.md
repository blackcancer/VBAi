# Mises à jour par releases GitHub

## Utilisation

Ouvrir **Outils → Mises à jour VBAi**, ou **À propos → Vérifier les mises à jour**.
La fenêtre contient les préférences, les notes de version, la vérification manuelle,
le téléchargement, la planification de l’installation, l’ignorance d’une version et
l’annulation d’une mise à jour encore en attente. Ses contrôles fixes sont déclarés
dans le Designer ; les données et les états sont renseignés à l’exécution.

Par défaut : vérification automatique activée, téléchargement automatique activé,
installation automatique désactivée, canal stable. Après activation et enregistrement
de l’installation automatique, les releases suivantes se préparent sans nouvelle
validation dans VBAi. Les éventuelles exigences UAC de Windows restent applicables.

Sur un déploiement géré par l’installeur, le complément enregistre l’hôte dès sa
connexion et vérifie au démarrage, puis réévalue l’échéance chaque heure. Une
vérification réussie vaut pour 24 heures ; un échec réseau/authentification permet
une nouvelle tentative après une heure. Un verrou de fichier évite les opérations
automatiques concurrentes entre les hôtes du même utilisateur. La déconnexion
arrête le minuteur et annule les opérations réseau en cours. Les vérifications
manuelles restent possibles à tout moment.

Les notes de la dernière release détectée sont conservées dans le cache : la fenêtre
peut montrer une mise à jour déjà téléchargée sans relancer une requête. Les drafts,
les versions plus anciennes ou égales à la version produit et, sur le canal stable,
les préversions sont ignorées. L’option Ignorer s’applique à la version sélectionnée.

## Dépôt privé, puis public

La source produit est `blackcancer/CodexVBE`, indépendamment du dépôt d’une macro.
La lecture commence sans authentification. Sur réponse 401/404, VBAi essaie une fois
les identifiants Git Credential Manager du compte sélectionné dans les paramètres
GitHub, sans ouvrir de connexion interactive. Quand le dépôt sera public, une
réponse publique réussie n’utilisera aucun identifiant.

Les assets sont demandés via leur identifiant à l’API GitHub. Les redirections
HTTPS sont limitées aux CDN GitHub utilisés par les releases ; le jeton n’est jamais
transmis au CDN. La taille annoncée, la limite de 512 Mio et le SHA-256 de l’asset
sont vérifiés avant que le fichier temporaire devienne un installeur prêt.
Les réponses de catalogue sont plafonnées à 4 Mio ; les requêtes de catalogue et
les téléchargements ont des délais globaux de deux et dix minutes.

Le contrat `digest: sha256:…` et le téléchargement binaire des assets sont décrits
par la [documentation GitHub des assets de release](https://docs.github.com/en/rest/releases/assets).
Une release sans installeur compatible ou sans digest est visible mais ne peut pas
être installée par VBAi. Les archives de sources ne constituent pas des installeurs.

## Contrat du futur installeur

Publier un tag de version produit, de préférence `v1.2.3` ou `v1.2.3-beta.1`, avec
un asset **`VBAi-Setup-win-x64.msi`** ou **`VBAi-Setup-win-x64.exe`**. Si les deux
existent, le MSI est préféré. Les versions sont comparées numériquement, avec les
règles de précédence des préversions ; les métadonnées de build sont ignorées.

L’installeur doit :

- Être signé avec une signature Authenticode embarquée reconnue par Windows.
  Le programme d’application revérifie le SHA-256 et la confiance Windows avant
  toute exécution, via [WinVerifyTrust](https://learn.microsoft.com/en-us/windows/win32/api/wintrust/nf-wintrust-winverifytrust).
- Déployer `CodexVBE.dll`, sa TLB, toutes les dépendances de rendu, `VBAi.Updater.exe`
  et sa configuration dans un répertoire d’installation stable.
- Inscrire les identités COM actuelles et le complément VBE 64 bits, conserver les
  paramètres et données utilisateur, puis gérer les fichiers occupés et le rollback
  de sa transaction. Les GUID/ProgID et la version d’assembly COM `0.1.0.0` restent
  inchangés ; **`ProductVersion`** est la version de livraison.
- Écrire `vbai-installation.json` à côté de la DLL et préserver `InstallationId`
  lors des mises à jour :

```json
{
  "Product": "VBAi",
  "Architecture": "win-x64",
  "UpdateProtocol": 1,
  "InstallationId": "00000000-0000-0000-0000-000000000001"
}
```

Le MSI est lancé avec `/i <fichier> /quiet /norestart /L*v <journal>`.
L’EXE doit implémenter **`/update /quiet /norestart`**, écrire son journal, attendre
la fin réelle de l’installation avant de quitter et renvoyer un code fiable.
Les codes `0`, `3010` et `1641` sont traités comme réussite ; les deux derniers
signalent un redémarrage Windows. Voir les [codes Windows Installer](https://learn.microsoft.com/en-us/windows/win32/msi/error-codes).
Après code 0, la version produit de la DLL installée doit correspondre à la cible.

VBAi ne crée pas un installeur factice et ne remplace pas directement les DLL :
l’installeur doit encore être développé et qualifié. Sans son marqueur, les
checkouts de développement autorisent les vérifications/téléchargements manuels,
mais ne démarrent aucune opération automatique et refusent l’application.

## Application différée et récupération

`VBAi.Updater` est une application WinForms .NET Framework 4.8 x64 indépendante.
Elle partage les sources du protocole, les langues et le thème, sans référence à
`CodexVBE.dll`. Avant lancement, le complément copie son programme dans le cache,
pour que l’exécutable en cours n’empêche pas l’installeur de remplacer sa version
déployée. Un mutex utilisateur empêche plusieurs programmes d’application concurrents.

Chaque hôte chargé possède une inscription PID/heure de naissance/répertoire.
L’inscription persiste après déconnexion COM, car le CLR peut garder la DLL chargée.
Le programme attend la disparition des processus correspondants ; il ne ferme ni
ne tue Excel, SOLIDWORKS ou une autre application. Un PID réutilisé ne bloque pas
la mise à jour. Une inscription illisible ou inaccessible entraîne une attente,
car la fermeture de l’hôte ne peut pas être démontrée. Le futur installeur doit aussi
traiter les fichiers occupés si un hôte démarre entre la vérification et l’application.

L’annulation d’un travail en attente est accessible dans la fenêtre de mise à jour.
Un verrou empêche d’annuler pendant l’installation. Une interruption après le début
de l’installation produit un état incertain au redémarrage du programme externe :
aucune relance automatique de cette version n’est effectuée. Les travaux en attente
reprennent au chargement suivant ; les métadonnées terminées restent disponibles.

Tout se conserve sous `%LocalAppData%\CodexVBE\Updates` : préférences, catalogue,
inscriptions d’hôtes, installeurs vérifiés, travail en attente et programme externe.
Aucun fichier n’est ajouté près d’une macro et aucun code VBA n’est transmis.

## Préparer une livraison

```powershell
powershell.exe -NoProfile -File .\tools\Prepare-Release.ps1 -Version 1.2.3
```

Cette commande compile les payloads Release, exporte la TLB, rassemble les
composants et génère le marqueur sous `artifacts/releases/1.2.3/package`.
Le futur installeur consommera ce répertoire. Après création et signature de
l’installeur, l’option `-InstallerPath` prépare l’asset normalisé et `SHA256SUMS.txt`.
La commande **ne publie aucune release** et ne modifie aucune inscription COM.

## Qualification locale

Les appels GitHub, les identifiants et les lancements d’installeurs sont simulés.
La suite Unit complète passe : **1 168 tests, aucun échec ni test ignoré**.
La compilation de la solution et la préparation Release passent sans avertissement
ni erreur. Les 31 surfaces Designer du complément chargent et se redimensionnent ;
la surface du programme externe passe également. Les captures françaises claires,
anglaises sombres et arabes sombres ne présentent aucun contrôle débordant.
Les vérifications natives de PID et de signature sont en lecture seule. Les
captures des interfaces sont autonomes ; elles ne lancent aucun hôte ni macro.
Le paquet de préversion `0.1.1-beta.1` a été préparé avec sa TLB : version produit
correcte, identité COM conservée et programme externe présent.

La vérification complète avec un véritable installeur signé publié dans une release,
l’application/rollback et les redémarrages des hôtes restent **NOT_RUN**. L’édition et
l’enregistrement manuel du source sous Visual Studio restent également à qualifier.

Preuves locales : `artifacts/updates/tests/unit-final.trx`,
`artifacts/updates/designers-final/`, `artifacts/updates/designers-worker/`,
`artifacts/updates/screens/` et `artifacts/updates/release-final/`.

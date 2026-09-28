# Propriétés d'un UserForm vierge dans Excel

> Archive conservée le 28 septembre 2026. Ce document contient des observations et des décisions de sa période de rédaction ; ses états « à faire » et ses anciens chiffres ne constituent pas le bilan actuel. Voir [la documentation actuelle](../../README.md) et [les travaux restants](../../roadmap.md).

Inventaire généré le 26 septembre 2026 dans le VBE Excel 64 bits, sur le formulaire jetable `CodexInventoryForm` du projet `VBAProject`. Le fichier [CSV](../../reference/excel-userform-properties.csv) contient les **51 propriétés** renvoyées par `VBComponent.Properties`, leur type obtenu sur le concepteur, la valeur initiale et l'indication de lecture seule lorsqu'un descripteur la fournit.

`ReadOnly=unknown` signifie que le descripteur du concepteur n'a pas été trouvé ; cela ne prouve ni qu'une écriture est impossible ni qu'elle est sûre. `Kind=object` demande un parcours typé distinct. `_Font_Reserved` produit une erreur de lecture COM sur ce poste. `Font` expose un `System.Drawing.Font` dont les membres sont en lecture seule dans le descripteur .NET, tandis que la police du formulaire peut être changée via l'objet COM du concepteur : l'essai `Font.Name=Segoe UI` a été relu avec une nouvelle révision du formulaire. `Picture` et `MouseIcon` sont des objets ; leur chargement depuis un fichier n'est pas encore validé dans Excel.

Cette capture décrit **ce UserForm dans cet hôte**. La couverture complète du VBE exige aussi les propriétés de chaque type de contrôle, les conteneurs imbriqués, les modules, projets, références, fenêtres, commandes et états de débogage, avec une relecture réelle après modification.

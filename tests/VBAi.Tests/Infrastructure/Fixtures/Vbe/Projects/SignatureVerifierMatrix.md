# Matrice de vérification VBA préparée avant exécution

- Chemin absolu, extension Office documentée, fichier absent/vide/trop grand ; fichier maintenu ouvert en lecture, SHA des octets réellement inspectés.
- SIP OOXML et binaire exacts ; aucun SIP, chemin relatif, nom de DLL inattendu et DLL absente : VerifierUnavailable, aucune substitution par certificat ou signature de document.
- Politique native : GUID VBA, handle et chemin corrects, aucune interface, cache uniquement, fermeture de l'état pour succès, erreur HRESULT, exception de vérification et exception de fermeture.
- Résultats : trusted, absence de signature, digest modifié, signature invalide, racine non approuvée, certificat expiré/révoqué, révocation indéterminée, autre échec.
- LLM : inspection en Discussion/Lecture seule, champ Path requis, routage de session sans projet imposé.
- Qualification native : unsigned Office ; document VBA signé jetable ; copie avec projet VBA altéré et signature conservée ; comparaison indépendante SignTool. Le SIP choisit V3 puis agile puis ancien, pas une vérification séparée des trois signatures.

Les bibliothèques Microsoft de test restent sous artifacts, exclues de Git. SWP et Access sont explicitement non pris en charge par le SIP Office.

# Matrice trust, rename et persistance autonome après 1b4eb25

Scénarios ajoutés avant première exécution du lot :
- CertificateTrust : vrai certificat jetable jamais installé ; flags cache-only et EKU Authenticode, états Trusted/NotTrusted/IndeterminateOffline par statuts natifs distincts, erreurs GetChain/Policy, allocation OOM et libération de la chaîne. Frontières chain/policy/free/allocator typées, défauts crypt32/Marshal identiques, restauration finally.
- Rename : host nonExcel, scope absent/relatif, Value type/syntaxe/keyword, protection illisible, path modifié au readback COM avant setter, version périmée, setter ignoré/idempotence, nombre/nom/code des modules mutés, path canoniquement identique mais représentation différente, budgets lignes négatives/100001 et caractères>4MiB, source vide.
- StandalonePersistence : Type, format SWP, path null/vide/relatif/absent/présent/read-only, exception native Type/FileName, SaveAs absent/relatif/autreformat/existant/parent inexistant, révision et ExpectedHostPath, sauvegarde rejetée ou readback path absent/relatif/autre/Savedfalse/fichierabsent/vide, vraie écriture temporaire sans Excel/SW.
- Session certificate_trust : mauvais SHA syntaxique, store CurrentUser/My readonly, zéro/deux/un certificat public match, Dispose et réponse offline ; aucune clé/credential utilisateur.
Les fixtures VBIDE sont normalement construites ; les mutations se produisent aux frontières de lecture/écriture natives et les réponses sont vérifiées par effets, versions et limites.

Validation finale : 82 tests verts, zéro ignoré ; build sans warning ni erreur.
| Source | Lignes | Branches |
| --- | --- | --- |
| VbeCertificateTrust.cs | 34/34 | 24/24 |
| VbeProjectComponents.Rename.cs | 48/48 | 114/114 |
| VbeProjectComponents.StandalonePersistence.cs | 52/52 | 94/94 |
Rapport : artifacts/coverage/project-trust-toolbar-validated/a91d44f3-58bd-46bf-b1cc-3e645717c388/coverage.cobertura.xml.
Exclusion collector unique [ProviderTests]*, aucune exclusion production. Aucun changement de garde. Le véritable appel crypt32 du certificat auto-signé reste testé en cache-only ; aucune installation de certificat, aucune lecture de clé privée, aucun Excel/SOLIDWORKS réel.

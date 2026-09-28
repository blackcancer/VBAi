# Archives de documentation

Archives regroupées le **28 septembre 2026**. Ces documents sont conservés pour leurs observations, décisions et preuves datées. Les anciens pourcentages, listes de fonctions et mentions « à faire » ne constituent pas l’état actuel : voir [l’index courant](../README.md), [l’état du projet](../project.md) et [les travaux restants](../roadmap.md).

## Exploration du VBE

- [État initial et diagnostic de chargement](exploration/project.md)
- [Intégration LLM : journal des fonctionnalités](exploration/llm-integration.md)
- [Qualifications VBE](exploration/vbe-coverage.md)
- [Inventaire et expériences successives](exploration/vbe-remaining-coverage.md)
- [Propriétés MSForms](exploration/forms-property-coverage.md)
- [Duplication des contrôles](exploration/forms-duplication-coverage.md)
- [Capture d’un UserForm Excel](exploration/excel-userform-properties.md)

Les données CSV correspondantes sont conservées dans [reference/](../reference/designer.md).

## Conception et comparaisons

- [Contrat initial de l’interface LLM](design/llm-interface.md)
- [Cible inspirée de Copilot Visual Studio](design/copilot-visual-studio-parity.md)
- [Revue de l’interface GitHub](design/ui-github-review.md)

Ces intentions sont à confronter aux guides actuels [Conversation](../chat-ui.md), [GitHub](../github-integration.md) et [WinForms](../winforms-designer.md).

## Inventaires de tests

- [ChatWindow](test-inventories/chat-window-coverage-inventory.md)
- [Codex app-server](test-inventories/codex-app-server-test-inventory.md)
- [Fenêtres de debug : tests](test-inventories/vbe-debug-windows-test-inventory.md)
- [Fenêtres de debug : frontières natives](test-inventories/vbe-debug-windows-native-inventory.md)
- [Formulaires : branches du cœur](test-inventories/vbe-forms-core-branch-inventory.md)
- [Formulaires : duplication](test-inventories/vbe-forms-duplication-inventory.md)
- [Dispatch de session](test-inventories/vbe-session-dispatch-inventory.md)
- [Mesure à 100 % avant chat-ux](test-coverage-pre-chat-ux.md)
- [Couverture à 100 % après isolation CLI](test-coverage-session-isolation.md)
- [Validation initiale de la fusion chat-ux](test-coverage-chat-ux-integration.md)

La [mesure globale courante](../test-coverage.md) prend en compte les sources ajoutées par `chat-ux` ; les inventaires archivés ne la remplacent pas.

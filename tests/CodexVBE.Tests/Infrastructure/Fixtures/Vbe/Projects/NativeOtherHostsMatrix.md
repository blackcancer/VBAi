# Adaptateurs natifs Word/PowerPoint — matrice avant exécution

- Résolution : processus non reconnu, ROT owned/étranger/COM/binder, NativeOM nul/HRESULT/COM/binder/étranger/absent, filtres PID et classe exacts pour deux hôtes.
- PID : fenêtre réelle owned/nulle, Word sans fenêtre, owners divergents/inconnus, catalogue 1001 ; PowerPoint HWND exact.
- Documents : catalogue Word/présentations, 1001, VBProject, état saved/read-only sous bool et MsoTriState, document non sauvegardé.
- Sauvegarde : Save/SaveAs2/SaveAs, destination et format exacts, une invocation sans hôte Office.
- Identité : IUnknown mêmes/différents, objets gérés/nuls, wrappers libérés aux deux positions ; références acquises toujours relâchées.
- Fichiers : existence/taille sur répertoire jetable owned. Les preuves restent des contrats locaux, qualification réelle NOT_RUN.

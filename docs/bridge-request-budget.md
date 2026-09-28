# Réception du pont local

Le serveur de canal nommé accepte toujours une requête par connexion et conserve son worker unique. La lecture de la ligne est maintenant effectuée par `BridgeRequestReader` sur un canal asynchrone.

- Taille maximale de la trame : **10 Mio de données UTF-8**, contrôlée pendant la réception, avant construction d'une ligne complète et avant désérialisation JSON.
- Temps maximal de réception : **10 secondes au total** après connexion. Recevoir quelques octets ne réinitialise pas ce délai.
- Une ligne doit être terminée par LF, éventuellement précédé de CR. EOF sans LF, dépassement et UTF-8 invalide sont refusés.
- À expiration, la connexion est fermée et la lecture asynchrone libérée ; le worker peut accepter le client suivant.
- La durée d'exécution d'une commande VBE n'est pas limitée par ce budget de réception.

Les tests utilisent de vrais canaux nommés locaux : client silencieux, ligne partielle, interruption, trame trop longue sans retour à la ligne, JSON mal formé et UTF-8 invalide, suivis d'un client sain. Les bornes exactes et les contrats historiques du serveur sont vérifiés aussi. Il s'agit de robustesse du transport local ; cette modification ne constitue pas une démonstration de vulnérabilité distante.

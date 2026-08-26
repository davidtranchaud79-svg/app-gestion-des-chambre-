# Gestion des chambres

Micro-application de reservation reliee au Google Sheet de gestion des chambres.

## Role de l'application

- Afficher les chambres disponibles en temps reel.
- Lire les disponibilites depuis les onglets `Chambres` et `Registre`.
- Enregistrer les demandes publiques dans `Reservations_Public`.
- Garder une validation manuelle par la reception.
- Eviter les doubles demandes simultanees avec `LockService`.

## Structure

| Fichier | Role |
| --- | --- |
| `src/Code.gs` | Backend Apps Script |
| `src/Index.html` | Interface publique |
| `src/appsscript.json` | Configuration Apps Script |
| `.github/workflows/deploy-apps-script.yml` | Deploiement GitHub Actions avec clasp |

## Installation rapide

1. Ouvrir le Google Sheet.
2. Aller dans `Extensions > Apps Script`.
3. Coller `src/Code.gs` dans `Code.gs`.
4. Creer un fichier HTML nomme `Index`.
5. Coller `src/Index.html` dedans.
6. Deployer comme application web.

Parametres de deploiement conseilles :

- Executer en tant que : `Moi`
- Acces : `Toute personne avec le lien`

## GitHub Actions

Pour deployer automatiquement avec GitHub Actions, ajouter ces secrets :

| Secret | Contenu |
| --- | --- |
| `CLASP_SCRIPT_ID` | ID du projet Apps Script |
| `CLASP_CREDENTIALS` | Contenu du fichier `~/.clasprc.json` apres `clasp login` |

Ne jamais committer `.clasprc.json` ni `.clasp.json` reel avec les identifiants.

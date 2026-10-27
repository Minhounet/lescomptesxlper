# CLAUDE.md — Les comptes (suivi budget d'un couple)

## Confidentialité (règle absolue)
- Le dépôt GitHub est PUBLIC. Ne jamais y commiter de données réelles
  (montants, noms, comptes, IBAN, salaires).
- Seuls les modèles `suivi_comptes_vierge.xlsx` et `suivi_comptes_example.xlsx` sont versionnés ;
  ils ne contiennent que des données fictives (Alex et Sam).
- Ce fichier est versionné (il ne contient aucune donnée personnelle). Les fichiers réels restent hors Git
  (`.gitignore`) : vérifier `git status` avant tout commit ; ils sont livrés dans la conversation, jamais poussés.

## Projet
- Classeur Excel / Google Sheets pour deux personnes : charges, revenus, répartition,
  virements à programmer vers les comptes joints, reste libre, imprévus. Voir `README.md`.
- Langue : français. Pas d'accès Google Drive (refus de l'utilisateur) : livrer des `.xlsx`
  à importer dans Sheets.

## Conventions du modèle
- Cellules jaunes à police bleue = saisie ; le reste est calculé. Police Arial.
- Formules compatibles Excel, LibreOffice et Google Sheets (SUMIFS, INDEX/MATCH, IFERROR) ;
  pas de XLOOKUP, FILTER, etc. Recalculer et vérifier 0 erreur avant de livrer.
- Deux personnes exactement ; une charge = un compte débité.
- Principe de conception : centraliser sur les comptes joints et lisser (enveloppes annuelles, provision imprévus),
  ex. épargne enfants = virement du compte joint principal vers un livret par enfant.

## Évolutions et migrations
- **Toute modification du modèle, des exemples ou des scripts entraîne la mise à jour, dans le même commit,
  de `README.md`, `CLAUDE.md` et `MIGRATION.md` si concernés** (fonctionnement, conventions, zones de saisie,
  version, historique). Ne jamais l'oublier ; relire ces trois fichiers avant de commiter.
- Lire `MIGRATION.md` avant toute évolution : versionnement (`Paramètres!B25`), zones de saisie, `scripts/build_modeles.py`
  (génère vierge + example) et `scripts/migrer.py`. Le tenir à jour à chaque changement de structure.
- Migration d'un fichier réel : hors Git, vers la dernière version du modèle.

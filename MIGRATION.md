# Évolutions du modèle et migrations

Le modèle est versionné (`Paramètres!B25`, format `majeur.mineur.correctif`). Les modèles
`suivi_comptes_vierge.xlsx` et `suivi_comptes_example.xlsx` sont générés par `scripts/build_modeles.py`
(données fictives uniquement). Les fichiers réels ne sont **jamais** dans ce dépôt (public).

## Règles pour faire évoluer le modèle
1. Modifier `scripts/build_modeles.py`, régénérer vierge **et** example (`python3 scripts/build_modeles.py` et `... blank`).
2. Recalculer et vérifier 0 erreur de formule.
3. Incrémenter `VERSION` : *correctif* (libellé, mise en forme), *mineur* (colonne ou onglet ajouté sans casser les zones de saisie),
   *majeur* (zone de saisie déplacée ou supprimée).
4. Si une zone de saisie change : mettre à jour `ZONES` dans `scripts/migrer.py` et le tableau ci-dessous, puis ajouter une entrée à l'historique.

## Zones de saisie (v1.0.0)
| Feuille | Plages recopiées |
|---|---|
| Paramètres | B4:B7, B15:C19, B22 |
| Comptes | A5:D14 |
| Charges | A5:E54, G5:K54 (F, L, M, N sont calculées) |
| Revenus | A5:E24, G5:H24 |
| Charges perso | A5:E44, G5:I44 |
| Imprévus | B5:B6, A13:D62 |

## Migrer un fichier réel
`python3 scripts/migrer.py ancien.xlsx nouveau.xlsx` copie ces zones vers le modèle courant. Écrire `nouveau.xlsx`
hors du dépôt. Puis rouvrir dans Excel/Sheets (recalcul) et comparer la Synthèse avec l'ancienne.
Pour un changement **majeur**, adapter d'abord `ZONES` à l'ancienne structure (table de correspondance) avant de lancer.

## Historique
- **1.0.1** : textes d'aide corrigés ; le modèle vierge contient la ligne « Provision imprévus » (montant = % de Paramètres × salaires, à 0 par défaut). Zones de saisie inchangées.
- **1.0.0** : première version (Synthèse, Paramètres, Comptes, Charges, Revenus, Charges perso, Imprévus).

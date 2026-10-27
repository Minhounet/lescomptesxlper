# Les comptes — suivi du budget d'un foyer de deux personnes

Classeur Excel / Google Sheets pour piloter le budget d'un couple : qui paie quoi, depuis quel compte,
combien chacun doit virer sur les comptes joints, et ce qu'il reste de libre.

## Fichiers

| Fichier | Rôle |
|---|---|
| `suivi_comptes_vierge.xlsx` | Modèle vide, prêt à remplir |
| `suivi_comptes_example.xlsx` | Même modèle avec des données fictives (Alex et Sam) pour voir le résultat |

Pour Google Sheets : importer le `.xlsx` dans Drive, puis « Enregistrer sous Google Sheets ».
Formules et listes déroulantes sont conservées.

> **Confidentialité** : ce dépôt est public. Ne jamais y mettre de vraies données (montants, comptes, noms).
> Seuls les deux modèles ci-dessus y sont versionnés.

## Principe

Le fichier décrit **un mois type** : les charges récurrentes sont ramenées en mensuel (une taxe annuelle est divisée
par 12, un trimestre par 3). On ne le touche que lorsque la situation change (salaire, nouvelle charge).
Les dépenses irrégulières (vêtements, médecin) sont **lissées** en enveloppes. Les comptes joints sont
alimentés par des virements permanents calculés par le fichier.

Convention : cellules **jaunes à police bleue = à saisir**, tout le reste est calculé.

## Onglets

1. **Synthèse** (rien à saisir)
   - *Par personne* : salaires, participations aux comptes joints, taux d'effort, charges perso, épargne, **reste libre**.
   - *Par compte* : charges, aides reçues, **« À virer par … »** chaque mois pour alimenter chaque compte joint, solde prévisionnel.
   - *Fixe vs variable* : part des charges communes incompressibles.
2. **Paramètres** : prénoms, compte perso de chacun, périodicités, part de chacun dans les salaires (prorata),
   budget perso (mode des charges perso, épargne en % ou en montant), % de provision pour imprévus.
3. **Comptes** : jusqu'à 10 comptes nommés librement, de type Joint, Personnel ou Épargne.
4. **Charges** : jusqu'à 50 lignes. Pour chacune : nature (fixe/variable), périodicité, montant, mode
   (prélèvement, virement, carte, chèque, espèces), compte débité, compte destinataire si virement, jour,
   **répartition** (50/50, prorata des salaires, personne 1 seule, personne 2 seule).
5. **Revenus** : salaires (sur le compte de chacun) et aides type CAF (sur le compte de votre choix).
6. **Charges perso** : dépenses propres à chacun. Mode *Détail* : les lignes sont additionnées ;
   mode *Forfait* (Paramètres) : une somme arbitraire, le détail reste privé.
7. **Imprévus** : journal des dépenses imprévues comparé à la provision mensuelle (livret dédié) ; le solde se met à jour seul.

## Utilisation

1. Paramètres → prénoms et comptes perso. 2. Comptes. 3. Revenus. 4. Charges.
5. Optionnel : Charges perso, épargne perso, Imprévus.
6. Lire la Synthèse et programmer les virements « À virer par … » vers chaque compte joint.

Un compte joint à solde 0 est équilibré. Un compte perso à solde négatif signifie que le salaire ne suffit pas.

## Hypothèses et limites

- Foyer de **deux personnes** exactement.
- Une charge débitée sur un compte perso est supportée en totalité par son titulaire (pas de compensation).
- Les aides sont réparties entre les deux au prorata de leurs parts de charges sur le compte qui les reçoit.
- Une charge = un seul compte débité ; scinder en deux lignes si elle est répartie sur deux comptes.
- Le fichier est un budget prévisionnel : il ne se connecte pas aux banques et ne mémorise pas l'historique.

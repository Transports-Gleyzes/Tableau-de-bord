# Tableau de bord en ligne — Gleyzes / LPB

Site privé pour suivre la santé de l'entreprise en temps réel, construit sur un Google Sheet.

| Onglet du site | Ce qu'il montre | D'où viennent les données |
|---|---|---|
| **Accueil** | CA, marge, livraisons, retards, véhicules indisponibles, factures en attente, heures chauffeurs, puis le bloc **🔴 ATTENTION** (documents qui expirent, clients à relancer, marge d'une activité en baisse…) | tout ce qui suit |
| **Finances** | CA, charges par poste, résultat, marge et coût au km, par camion et par société | onglet `FINANCES` (importé depuis `TABLEAU_DE_BORD_FLOTTE.xlsx`) |
| **Flotte** | Contrôle technique, entretien, assurance, chronotachygraphe | onglet `FLOTTE` (saisie manuelle) |
| **Salariés** | Permis, FIMO/FCO, visite médicale, carte conducteur, formations, absences | onglets `SALARIES` et `ABSENCES` |
| **Carburant & livraisons** | Livraisons par jour, CA des tournées, litres, consommation L/100 km | onglets `LIVRAISONS` et `CARBURANT` |
| **Factures** | Factures en attente, échues, clients à relancer ; boutons « Relancé » et « Payée » | onglet `FACTURES` |
| **Alertes** | Toutes les alertes, filtrables par gravité | calculées automatiquement |
| **Saisie rapide** | Formulaires (aussi sur téléphone) : livraison (avec retards), plein, heures chauffeur, facture, absence | écrit dans le Google Sheet |

Chaque matin vers 7h, un **mail d'alerte** part si une échéance approche (30 j / 7 j / dépassée), si un camion est en perte ou sous la marge minimale, ou si un camion consomme trop. Le même mail n'est jamais envoyé deux fois. Le lundi, un récapitulatif complet est envoyé.

Comme dans le classeur Excel : **une donnée manquante n'est jamais remplacée par 0**. Elle s'affiche « N/D » et les totaux concernés sont marqués « Incomplet ».

---

## Installation (une seule fois, environ 15 minutes)

Il faut un compte Google (une adresse Gmail ou Google Workspace).

### 1. Créer le Google Sheet et coller le code
1. Sur [sheets.google.com](https://sheets.google.com), créez un classeur vide nommé **TABLEAU DE BORD ENTREPRISE**.
2. Menu **Extensions > Apps Script**.
3. Dans l'éditeur qui s'ouvre :
   - remplacez le contenu de `Code.gs` par celui du fichier [`Code.gs`](Code.gs) ;
   - cliquez sur **+ > HTML**, nommez le fichier `Dashboard` (sans « .html ») et collez le contenu de [`Dashboard.html`](Dashboard.html) ;
   - roue dentée **Paramètres du projet**, cochez « Afficher le fichier manifeste appsscript.json », puis remplacez son contenu par [`appsscript.json`](appsscript.json).
4. Enregistrez (icône disquette).

### 2. Créer les onglets
1. Revenez au Google Sheet et rechargez la page : un menu **Tableau de bord** apparaît.
2. **Tableau de bord > 1. Installer / compléter les onglets**. Google demande d'autoriser le script : acceptez (écran « Application non validée » > *Paramètres avancés* > *Accéder au projet*. C'est normal, c'est votre propre script).
3. Les onglets sont créés. Vous pouvez relancer cette commande à tout moment : elle ne supprime jamais rien.

### 3. Remplir les données de base
- **PARAMETRES** : vérifiez `EMAIL_ALERTES` (plusieurs adresses possibles, séparées par des virgules, par exemple `transports-gleyzes@orange.fr`) et ajustez les seuils.
- **ID_CLASSEUR_ECHEANCES** (onglet PARAMETRES, colonne B) : collez le lien de votre Google Sheet « Échéances flotte ». Ses onglets « Toutes les échéances », « CONGES - ABSENCES - FORMATION » et « Pense-bête mensuel » sont alors lus directement : on continue à les remplir là-bas, le tableau de bord suit. Relancez ensuite **1. Installer / compléter les onglets** : les véhicules et chauffeurs de ce tableau sont ajoutés à FLOTTE et SALARIES.
- **ID_PLANNING_INTER**, **ID_PLANNING_CARBURANT**, **ID_LITRAGES** : liens des Google Sheets « Planning Inter (ITM) », « Planning Carburant » et « Litrages véhicules ». Les deux plannings restent distincts partout (onglet **Activité** : CA, tournées, livraisons, chauffeurs, clients) ; les litrages alimentent l'onglet **Consommation** et l'alerte « le camion X consomme trop » (2 dernières semaines regroupées, réglable avec NB_SEMAINES_CONSO). Un fichier Excel doit d'abord être converti : dans Google Drive, ouvrez-le puis **Fichier > Enregistrer au format Google Sheets**.
- **Colonne CAMION des plannings** (menu **Tableau de bord > Remplir la colonne CAMION des plannings**, puis chaque matin automatiquement) : chaque planning reçoit un onglet **CAMIONS** (chauffeur → camion habituel, proposé d'après les 10 dernières tournées) et, dans les onglets du mois en cours et des mois suivants, une colonne CAMION qui se remplit seule dès que le chauffeur est tapé. Changement : tapez l'immatriculation par-dessus. Les jours passés sont figés chaque matin (changer le camion habituel ne réécrit pas l'historique).
- **FLOTTE** : une ligne par camion. `Activite` (ex : Frigo, Benne, Plateau) sert à suivre la marge par activité ; `Statut` (Disponible, Atelier, Immobilisé…) compte les véhicules indisponibles. `Camion_ID` = l'immatriculation **sans tirets**, comme dans REF_Camions (ex : `GD042ZC`). Mettez `Actif` à `N` pour un camion sorti du parc.
- **SALARIES** : une ligne par salarié, avec les dates de fin de validité.

Dates au format `jj/mm/aaaa`.

### 4. Activer les mails automatiques
**Tableau de bord > 2. Activer les envois automatiques**. Pour tester tout de suite : **Vérifier les alertes et envoyer le mail maintenant**.

### 5. Publier le site
1. Dans l'éditeur Apps Script : **Déployer > Nouveau déploiement**, type **Application Web**.
2. *Exécuter en tant que* : **Moi**. *Qui a accès* : **Moi uniquement**.
3. **Déployer**, puis copiez l'URL (`https://script.google.com/macros/s/…/exec`). Ajoutez-la aux favoris et à l'écran d'accueil du téléphone.

Après une modification du code : **Déployer > Gérer les déploiements > crayon > Version : Nouvelle version**. L'URL ne change pas.

> **Donner accès à un collègue (ex : Adeline)** : partagez le Google Sheet avec son compte Google (droits *Éditeur* si elle doit saisir), puis créez le déploiement avec *Exécuter en tant que* : **Utilisateur qui accède à l'application Web** et *Qui a accès* : **Tous les utilisateurs disposant d'un compte Google**. Seules les personnes qui ont accès au Sheet verront les données.

---

## Mettre à jour le code

Quand une nouvelle version arrive sur GitHub : recollez `Code.gs` et `Dashboard.html` dans Apps Script, enregistrez, puis
1. dans le Google Sheet : **Tableau de bord > 1. Installer / compléter les onglets** (ajoute les nouveaux onglets et colonnes sans rien effacer) ;
2. dans Apps Script : **Déployer > Gérer les déploiements > crayon > Version : Nouvelle version > Déployer**.

---

## Relier le classeur Excel existant (finances)

`maj_mensuelle_flotte.py` continue d'alimenter `TABLEAU_DE_BORD_FLOTTE.xlsx` chaque mois. On ajoute une étape juste après :

```bash
python3 maj_mensuelle_flotte.py --classeur TABLEAU_DE_BORD_FLOTTE.xlsx --sources "TABLEAU DE BORD/SOURCES"
python3 exporter_finances_csv.py --classeur TABLEAU_DE_BORD_FLOTTE.xlsx --sortie "G:/Mon Drive/TABLEAU DE BORD/FINANCES_EXPORT.csv"
```

- Avec **Google Drive pour ordinateur** installé, écrivez le CSV directement dans le dossier `Mon Drive` : il est synchronisé tout seul et importé le lendemain matin. Pour un import immédiat : **Tableau de bord > Importer les finances depuis Drive**.
- Sans Google Drive pour ordinateur, glissez-déposez `FINANCES_EXPORT.csv` dans [drive.google.com](https://drive.google.com).
- Les mois présents dans le CSV remplacent ceux du Google Sheet ; les autres mois sont conservés. On peut donc relancer l'export autant de fois que nécessaire, sans créer de doublons.

---

## Utilisation au quotidien

| Qui | Quand | Quoi |
|---|---|---|
| Chauffeurs / exploitation | chaque jour | *Saisie rapide* : livraisons et pleins (avec le compteur km) |
| Bureau | à chaque événement | mettre à jour FLOTTE / SALARIES (nouvelle date de CT après passage…) et saisir les absences |
| Bureau | chaque mois | script mensuel + export CSV |
| Tout le monde | quand on veut | ouvrir le site. Il se rafraîchit seul toutes les 5 minutes |

## Pistes pour la suite
- Plannings : si les plannings sont déjà dans un tableur, importer directement nombre de livraisons et CA au lieu de les ressaisir.
- Comptabilité : importer la balance ou le grand livre (export CSV du logiciel comptable) pour les charges hors flotte (structure, assurances, frais bancaires…).
- Litrage : importer automatiquement les relevés de carte carburant au lieu de saisir chaque plein.

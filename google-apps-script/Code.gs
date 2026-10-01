/**
 * Tableau de bord Gleyzes / LPB — Google Apps Script
 * ===================================================
 * Script lié au Google Sheet « TABLEAU DE BORD ENTREPRISE ».
 *
 *  - installer()                : crée/complète les onglets et les paramètres (sans rien effacer).
 *  - doGet()                    : sert le site interactif (Dashboard.html).
 *  - getDonnees()               : renvoie toutes les données au site.
 *  - ajouterSaisie(type, objet) : ajoute une livraison / un plein / une absence / des heures / une facture.
 *  - majFacture(numero, action) : marque une facture « payée » ou « relancée » depuis le site.
 *  - verifierAlertes()          : calcule les alertes et envoie le mail (déclencheur quotidien).
 *  - importerFinances()         : importe FINANCES_EXPORT.csv (produit par exporter_finances_csv.py).
 *
 * Principe repris de maj_mensuelle_flotte.py : ne jamais inventer un 0.
 * Une cellule vide reste vide et s'affiche « N/D » sur le site.
 */

// ---------------------------------------------------------------------------
// Structure du classeur
// ---------------------------------------------------------------------------
var FEUILLES = {
  PARAMETRES: ['Paramètre', 'Valeur', 'Description'],
  FINANCES: ['Mois', 'Camion_ID', 'Société', 'CA', 'KM', 'Carburant', 'Peages', 'Salaires',
             'Entretien', 'Charges_Fixes', 'Charges_Mutualisees'],
  FLOTTE: ['Camion_ID', 'Société', 'Marque_Modele', 'Chauffeur_Attitre', 'Prochain_CT',
           'Prochain_Entretien', 'Echeance_Assurance', 'Controle_Tachygraphe', 'Actif', 'Remarques',
           'Activite', 'Statut', 'Indisponible_Jusqu_Au'],
  SALARIES: ['Nom', 'Prénom', 'Société', 'Poste', 'Fin_Validite_Permis', 'Fin_FIMO_FCO',
             'Prochaine_Visite_Medicale', 'Fin_Carte_Conducteur', 'Formation_A_Prevoir',
             'Date_Formation', 'Actif', 'Remarques'],
  ABSENCES: ['Salarié', 'Type', 'Début', 'Fin', 'Remarque'],
  CARBURANT: ['Date', 'Camion_ID', 'Société', 'Litres', 'Montant_TTC', 'KM_Compteur', 'Lieu'],
  LIVRAISONS: ['Date', 'Camion_ID', 'Chauffeur', 'Client', 'Nb_Livraisons', 'CA_HT', 'Remarque', 'Nb_Retards'],
  FACTURES: ['N_Facture', 'Client', 'Date_Facture', 'Echeance', 'Montant_TTC', 'Statut', 'Date_Paiement',
             'Derniere_Relance', 'Remarque'],
  HEURES: ['Date', 'Salarié', 'Heures', 'Remarque'],
  JOURNAL_ALERTES: ['Date_Envoi', 'Clé', 'Niveau', 'Message']
};

// Colonnes de date surveillées : [colonne, libellé affiché dans les alertes]
var ECHEANCES = {
  FLOTTE: [['Prochain_CT', 'Contrôle technique'], ['Prochain_Entretien', 'Entretien'],
           ['Echeance_Assurance', 'Assurance'], ['Controle_Tachygraphe', 'Contrôle chronotachygraphe']],
  SALARIES: [['Fin_Validite_Permis', 'Permis'], ['Fin_FIMO_FCO', 'FIMO / FCO'],
             ['Prochaine_Visite_Medicale', 'Visite médicale'], ['Fin_Carte_Conducteur', 'Carte conducteur'],
             ['Date_Formation', 'Formation']]
};

// Colonnes mises au format jj/mm/aaaa par installer()
var COLONNES_DATE = {
  FLOTTE: ['Prochain_CT', 'Prochain_Entretien', 'Echeance_Assurance', 'Controle_Tachygraphe', 'Indisponible_Jusqu_Au'],
  SALARIES: ['Fin_Validite_Permis', 'Fin_FIMO_FCO', 'Prochaine_Visite_Medicale', 'Fin_Carte_Conducteur', 'Date_Formation'],
  FACTURES: ['Date_Facture', 'Echeance', 'Date_Paiement', 'Derniere_Relance'],
  HEURES: ['Date'], LIVRAISONS: ['Date'], CARBURANT: ['Date'], ABSENCES: ['Début', 'Fin']
};

var PARAMETRES_DEFAUT = [
  ['EMAIL_ALERTES', '', 'Adresse(s) qui reçoivent les alertes, séparées par des virgules'],
  ['JOURS_PREAVIS', 30, 'Une échéance passe « à prévoir » ce nombre de jours avant la date'],
  ['JOURS_URGENT', 7, 'Une échéance passe « urgente » ce nombre de jours avant la date'],
  ['SEUIL_MARGE_PCT', 5, 'Alerte si la marge d\'un camion sur le dernier mois est sous ce % (négatif = urgent)'],
  ['SEUIL_CONSO_L100', 38, 'Alerte si la consommation d\'un camion dépasse ce nombre de L/100 km (90 derniers jours)'],
  ['NOM_FICHIER_FINANCES', 'FINANCES_EXPORT.csv', 'Nom du fichier déposé dans Google Drive par exporter_finances_csv.py'],
  ['RECAP_HEBDO', 'OUI', 'OUI = un mail récapitulatif complet chaque lundi, même sans nouvelle alerte'],
  ['JOURS_RELANCE', 15, 'Une facture échue est « à relancer » si aucune relance depuis ce nombre de jours'],
  ['SEUIL_BAISSE_MARGE_PTS', 5, 'Alerte si la marge d\'une activité perd ce nombre de points par rapport aux 3 mois précédents'],
  ['SEUIL_RETARDS_PCT', 5, 'Alerte si le taux de livraisons en retard du mois dépasse ce % (à partir de 20 livraisons)']
];

// Saisies autorisées depuis le site (rien d'autre ne peut être écrit par le site)
var SAISIES = {
  LIVRAISONS: { obligatoires: ['Date', 'Camion_ID', 'Nb_Livraisons'], nombres: ['Nb_Livraisons', 'CA_HT', 'Nb_Retards'], dates: ['Date'] },
  HEURES: { obligatoires: ['Date', 'Salarié', 'Heures'], nombres: ['Heures'], dates: ['Date'] },
  FACTURES: { obligatoires: ['N_Facture', 'Client', 'Montant_TTC', 'Echeance'], nombres: ['Montant_TTC'], dates: ['Date_Facture', 'Echeance'] },
  CARBURANT: { obligatoires: ['Date', 'Camion_ID', 'Litres'], nombres: ['Litres', 'Montant_TTC', 'KM_Compteur'], dates: ['Date'] },
  ABSENCES: { obligatoires: ['Salarié', 'Type', 'Début'], nombres: [], dates: ['Début', 'Fin'] }
};

var FUSEAU = 'Europe/Paris';

// ---------------------------------------------------------------------------
// Menu et installation
// ---------------------------------------------------------------------------
function onOpen() {
  SpreadsheetApp.getUi().createMenu('Tableau de bord')
    .addItem('1. Installer / compléter les onglets', 'installer')
    .addItem('2. Activer les envois automatiques (chaque matin)', 'installerDeclencheurs')
    .addSeparator()
    .addItem('Importer les finances depuis Drive', 'importerFinancesManuel')
    .addItem('Vérifier les alertes et envoyer le mail maintenant', 'verifierAlertesManuel')
    .addToUi();
}

function classeur_() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  if (ss) return ss;
  var id = PropertiesService.getScriptProperties().getProperty('ID_CLASSEUR');
  if (!id) throw new Error('Classeur introuvable : lancez d\'abord « Installer / compléter les onglets ».');
  return SpreadsheetApp.openById(id);
}

/** Crée les onglets manquants et ajoute les colonnes manquantes, sans jamais effacer de données. */
function installer() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  PropertiesService.getScriptProperties().setProperty('ID_CLASSEUR', ss.getId());
  Object.keys(FEUILLES).forEach(function (nom) {
    var sh = ss.getSheetByName(nom) || ss.insertSheet(nom);
    var attendues = FEUILLES[nom];
    var actuelles = sh.getLastColumn() ? sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0] : [];
    var manquantes = attendues.filter(function (h) { return actuelles.indexOf(h) < 0; });
    if (manquantes.length) {
      var debut = actuelles.filter(String).length + 1;
      sh.getRange(1, debut, 1, manquantes.length).setValues([manquantes]);
    }
    sh.getRange(1, 1, 1, sh.getLastColumn()).setFontWeight('bold').setBackground('#e8eef7');
    sh.setFrozenRows(1);
  });
  var shP = ss.getSheetByName('PARAMETRES');
  var existants = lireTable_(shP).map(function (l) { return l['Paramètre']; });
  PARAMETRES_DEFAUT.forEach(function (p) {
    if (existants.indexOf(p[0]) < 0) shP.appendRow(p);
  });
  if (!lireParametres_()['EMAIL_ALERTES']) {
    var moi = Session.getEffectiveUser().getEmail();
    if (moi) majParametre_('EMAIL_ALERTES', moi);
  }
  // Formats de date lisibles
  Object.keys(COLONNES_DATE).forEach(function (nom) {
    var sh = ss.getSheetByName(nom);
    var entetes = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0];
    COLONNES_DATE[nom].forEach(function (col) {
      var c = entetes.indexOf(col) + 1;
      if (c) sh.getRange(2, c, sh.getMaxRows() - 1, 1).setNumberFormat('dd/mm/yyyy');
    });
  });
  // Listes déroulantes pour limiter les fautes de frappe
  var listes = { FLOTTE: { Statut: ['Disponible', 'Atelier', 'Immobilisé', 'En attente de pièce'] },
                 FACTURES: { Statut: ['En attente', 'Payée', 'Litige'] } };
  Object.keys(listes).forEach(function (nom) {
    var sh = ss.getSheetByName(nom);
    var entetes = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0];
    Object.keys(listes[nom]).forEach(function (col) {
      var c = entetes.indexOf(col) + 1;
      if (c) sh.getRange(2, c, sh.getMaxRows() - 1, 1).setDataValidation(
        SpreadsheetApp.newDataValidation().requireValueInList(listes[nom][col], true).setAllowInvalid(true).build());
    });
  });
  try {
    SpreadsheetApp.getUi().alert('Onglets prêts. Remplissez FLOTTE et SALARIES, puis activez les envois automatiques (menu Tableau de bord > 2).');
  } catch (e) { /* lancé hors interface */ }
}

/** Un déclencheur chaque matin vers 7h : import des finances puis alertes. */
function installerDeclencheurs() {
  ScriptApp.getProjectTriggers().forEach(function (t) {
    if (t.getHandlerFunction() === 'tacheQuotidienne') ScriptApp.deleteTrigger(t);
  });
  ScriptApp.newTrigger('tacheQuotidienne').timeBased().everyDays(1).atHour(7).inTimezone(FUSEAU).create();
  try {
    SpreadsheetApp.getUi().alert('Envois automatiques activés : vérification chaque matin vers 7h.');
  } catch (e) { /* lancé hors interface */ }
}

function tacheQuotidienne() {
  try { importerFinances(false); } catch (e) { console.error('Import finances : ' + e); }
  verifierAlertes(false);
}

// ---------------------------------------------------------------------------
// Site web
// ---------------------------------------------------------------------------
function doGet() {
  return HtmlService.createTemplateFromFile('Dashboard').evaluate()
    .setTitle('Tableau de bord Gleyzes / LPB')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.DEFAULT);
}

/** Toutes les données du site en un seul appel (dates converties en texte AAAA-MM-JJ). */
function getDonnees() {
  var ss = classeur_();
  var d = {};
  ['FINANCES', 'FLOTTE', 'SALARIES', 'ABSENCES', 'CARBURANT', 'LIVRAISONS', 'FACTURES', 'HEURES'].forEach(function (nom) {
    var sh = ss.getSheetByName(nom);
    d[nom] = sh ? lireTable_(sh).map(serialiser_) : [];
  });
  d.FINANCES.forEach(function (l) { l.Mois = moisTexte_(l.Mois); });
  d.FLOTTE = d.FLOTTE.filter(estActif_);
  d.SALARIES = d.SALARIES.filter(estActif_);
  var params = lireParametres_();
  d.parametres = {
    JOURS_PREAVIS: nombre_(params.JOURS_PREAVIS, 30),
    JOURS_URGENT: nombre_(params.JOURS_URGENT, 7),
    SEUIL_MARGE_PCT: nombre_(params.SEUIL_MARGE_PCT, 5),
    SEUIL_CONSO_L100: nombre_(params.SEUIL_CONSO_L100, 38),
    JOURS_RELANCE: nombre_(params.JOURS_RELANCE, 15),
    SEUIL_RETARDS_PCT: nombre_(params.SEUIL_RETARDS_PCT, 5)
  };
  d.alertes = calculerAlertes_(ss, new Date());
  d.genereLe = Utilities.formatDate(new Date(), FUSEAU, "dd/MM/yyyy 'à' HH:mm");
  d.derniereImport = PropertiesService.getScriptProperties().getProperty('DERNIER_IMPORT_TEXTE') || '';
  return d;
}

/** Ajoute une ligne saisie depuis le site. Seuls LIVRAISONS, CARBURANT et ABSENCES sont autorisés. */
function ajouterSaisie(type, objet) {
  var regle = SAISIES[type];
  if (!regle) throw new Error('Saisie non autorisée : ' + type);
  regle.obligatoires.forEach(function (c) {
    if (objet[c] === undefined || objet[c] === null || String(objet[c]).trim() === '') {
      throw new Error('Champ obligatoire manquant : ' + c);
    }
  });
  var ligneObj = {};
  FEUILLES[type].forEach(function (c) {
    var v = objet[c];
    if (v === undefined || v === null || String(v).trim() === '') { ligneObj[c] = ''; return; }
    if (regle.nombres.indexOf(c) >= 0) {
      var n = nombre_(v, null);
      if (n === null || n < 0) throw new Error('Nombre invalide pour ' + c + ' : ' + v);
      ligneObj[c] = n;
    } else if (regle.dates.indexOf(c) >= 0) {
      var dt = dateDepuisTexte_(v);
      if (!dt) throw new Error('Date invalide pour ' + c + ' : ' + v);
      ligneObj[c] = dt;
    } else {
      // Empêche l'injection de formules dans le classeur
      ligneObj[c] = String(v).trim().replace(/^[=+\-@]/, "'$&").slice(0, 200);
    }
  });
  var ss = classeur_();
  if (type === 'LIVRAISONS' && ligneObj.Nb_Retards !== '' && ligneObj.Nb_Retards > ligneObj.Nb_Livraisons) {
    throw new Error('Plus de retards que de livraisons');
  }
  if (type === 'FACTURES') {
    if (!ligneObj.Statut) ligneObj.Statut = 'En attente';
    var existe = lireTable_(ss.getSheetByName('FACTURES')).some(function (l) { return String(l.N_Facture) === String(ligneObj.N_Facture); });
    if (existe) throw new Error('La facture ' + ligneObj.N_Facture + ' existe déjà');
  }
  if (type === 'CARBURANT' && !ligneObj['Société']) {
    var camion = lireTable_(ss.getSheetByName('FLOTTE')).filter(function (l) { return l.Camion_ID === ligneObj.Camion_ID; })[0];
    if (camion) ligneObj['Société'] = camion['Société'];
  }
  var verrou = LockService.getScriptLock();
  verrou.waitLock(10000);
  try {
    var sh = ss.getSheetByName(type);
    var entetes = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0];
    sh.appendRow(entetes.map(function (h) { return ligneObj.hasOwnProperty(h) ? ligneObj[h] : ''; }));
  } finally {
    verrou.releaseLock();
  }
  return true;
}

/** Depuis le site : action 'payee' (Statut + Date_Paiement) ou 'relancee' (Derniere_Relance = aujourd'hui). */
function majFacture(numero, action) {
  if (action !== 'payee' && action !== 'relancee') throw new Error('Action inconnue : ' + action);
  var sh = classeur_().getSheetByName('FACTURES');
  var verrou = LockService.getScriptLock();
  verrou.waitLock(10000);
  try {
    var v = sh.getRange(1, 1, sh.getLastRow(), sh.getLastColumn()).getValues();
    var e = v[0].map(String);
    for (var i = 1; i < v.length; i++) {
      if (String(v[i][e.indexOf('N_Facture')]) !== String(numero)) continue;
      var auj = minuit_(new Date());
      if (action === 'payee') {
        sh.getRange(i + 1, e.indexOf('Statut') + 1).setValue('Payée');
        sh.getRange(i + 1, e.indexOf('Date_Paiement') + 1).setValue(auj);
      } else {
        sh.getRange(i + 1, e.indexOf('Derniere_Relance') + 1).setValue(auj);
      }
      return true;
    }
  } finally {
    verrou.releaseLock();
  }
  throw new Error('Facture introuvable : ' + numero);
}

// ---------------------------------------------------------------------------
// Alertes
// ---------------------------------------------------------------------------
/**
 * Renvoie la liste des alertes, de la plus grave à la moins grave.
 * Niveaux : 'depasse' (date passée / résultat négatif), 'urgent', 'a_prevoir'.
 * Catégories (regroupement du bloc ATTENTION du site) : documents, relances, marge, vehicules,
 * retards, conso, donnees.
 */
function calculerAlertes_(ss, aujourdhui) {
  var p = lireParametres_(ss);
  var preavis = nombre_(p.JOURS_PREAVIS, 30), urgent = nombre_(p.JOURS_URGENT, 7);
  var jour0 = minuit_(aujourdhui);
  var alertes = [];

  // 1. Échéances flotte et salariés
  [['FLOTTE', 'Flotte'], ['SALARIES', 'Salariés']].forEach(function (f) {
    var sh = ss.getSheetByName(f[0]);
    if (!sh) return;
    lireTable_(sh).filter(estActif_).forEach(function (l) {
      var objet = f[0] === 'FLOTTE'
        ? l.Camion_ID + (l['Société'] ? ' (' + l['Société'] + ')' : '')
        : ((l['Prénom'] || '') + ' ' + (l.Nom || '')).trim();
      if (!objet) return;
      ECHEANCES[f[0]].forEach(function (e) {
        var dt = l[e[0]] instanceof Date ? l[e[0]] : dateDepuisTexte_(l[e[0]]);
        if (!dt) return;
        var jours = Math.round((minuit_(dt) - jour0) / 86400000);
        if (jours > preavis) return;
        var niveau = jours < 0 ? 'depasse' : (jours <= urgent ? 'urgent' : 'a_prevoir');
        var quand = jours < 0 ? 'dépassé depuis ' + (-jours) + ' j'
          : (jours === 0 ? 'aujourd\'hui' : 'dans ' + jours + ' j');
        var libelle = e[1] + (e[0] === 'Date_Formation' && l.Formation_A_Prevoir ? ' (' + l.Formation_A_Prevoir + ')' : '');
        alertes.push({
          niveau: niveau, categorie: 'documents', domaine: f[1], objet: objet, sujet: libelle, jours: jours,
          date: Utilities.formatDate(dt, FUSEAU, 'dd/MM/yyyy'),
          message: libelle + ' — ' + objet + ' : ' + Utilities.formatDate(dt, FUSEAU, 'dd/MM/yyyy') + ' (' + quand + ')',
          cle: f[0] + '|' + objet + '|' + e[0] + '|' + Utilities.formatDate(dt, FUSEAU, 'yyyy-MM-dd')
        });
      });
    });
  });

  // 2. Finances : dernier mois importé
  var shF = ss.getSheetByName('FINANCES');
  if (shF) {
    var fin = lireTable_(shF);
    var mois = fin.map(function (l) { return moisTexte_(l.Mois); }).filter(String).sort();
    var dernier = mois[mois.length - 1];
    if (dernier) {
      var seuil = nombre_(p.SEUIL_MARGE_PCT, 5);
      var parCamion = {};
      fin.filter(function (l) { return moisTexte_(l.Mois) === dernier; }).forEach(function (l) {
        var c = parCamion[l.Camion_ID] = parCamion[l.Camion_ID] || { ca: null, charges: null, incomplet: false };
        var ca = nombre_(l.CA, null);
        if (ca !== null) c.ca = (c.ca || 0) + ca;
        ['Carburant', 'Peages', 'Salaires', 'Entretien', 'Charges_Fixes', 'Charges_Mutualisees'].forEach(function (k) {
          var v = nombre_(l[k], null);
          if (v !== null) c.charges = (c.charges || 0) + v;
        });
        if (nombre_(l.Carburant, null) === null || nombre_(l.Salaires, null) === null) c.incomplet = true;
      });
      Object.keys(parCamion).forEach(function (id) {
        var c = parCamion[id];
        if (c.ca === null) {
          alertes.push({ niveau: 'a_prevoir', categorie: 'donnees', domaine: 'Finances', objet: id, sujet: 'CA manquant',
            message: 'CA non renseigné pour ' + id + ' en ' + dernier, cle: 'CA_MANQUANT|' + id + '|' + dernier });
          return;
        }
        if (c.charges === null || c.ca === 0) return;
        var res = c.ca - c.charges, marge = res / c.ca * 100;
        var note = c.incomplet ? ' (charges incomplètes : carburant ou salaire manquant)' : '';
        if (res < 0) {
          alertes.push({ niveau: 'depasse', categorie: 'marge', domaine: 'Finances', objet: id, sujet: 'Résultat négatif',
            message: id + ' perd ' + euros_(-res) + ' en ' + dernier + note, cle: 'RESULTAT|' + id + '|' + dernier });
        } else if (marge < seuil) {
          alertes.push({ niveau: 'urgent', categorie: 'marge', domaine: 'Finances', objet: id, sujet: 'Marge faible',
            message: id + ' : marge de ' + fr1_(marge) + ' % en ' + dernier + note, cle: 'MARGE|' + id + '|' + dernier });
        }
      });
    }
  }

  // 3. Consommation de carburant sur 90 jours (méthode plein à plein)
  var shC = ss.getSheetByName('CARBURANT');
  if (shC) {
    var seuilConso = nombre_(p.SEUIL_CONSO_L100, 38);
    var conso = consommations_(lireTable_(shC), jour0, 90);
    Object.keys(conso).forEach(function (id) {
      if (conso[id] > seuilConso) {
        alertes.push({ niveau: 'urgent', categorie: 'conso', domaine: 'Carburant', objet: id, sujet: 'Consommation élevée',
          message: id + ' consomme ' + fr1_(conso[id]) + ' L/100 km sur 90 jours (seuil ' + seuilConso + ')',
          cle: 'CONSO|' + id + '|' + Utilities.formatDate(jour0, FUSEAU, 'yyyy-ww') });
      }
    });
  }

  var flotte = ss.getSheetByName('FLOTTE') ? lireTable_(ss.getSheetByName('FLOTTE')).filter(estActif_) : [];

  // 4. Véhicules indisponibles (colonne Statut de FLOTTE)
  flotte.forEach(function (l) {
    if (!estIndisponible_(l.Statut)) return;
    var jusqu = l.Indisponible_Jusqu_Au instanceof Date ? l.Indisponible_Jusqu_Au : dateDepuisTexte_(l.Indisponible_Jusqu_Au);
    var objet = l.Camion_ID + (l['Société'] ? ' (' + l['Société'] + ')' : '');
    alertes.push({ niveau: 'urgent', categorie: 'vehicules', domaine: 'Flotte', objet: objet, sujet: 'Indisponible',
      message: objet + ' indisponible : ' + l.Statut + (jusqu ? ' jusqu\'au ' + Utilities.formatDate(jusqu, FUSEAU, 'dd/MM/yyyy') : ''),
      cle: 'INDISPO|' + l.Camion_ID + '|' + l.Statut + '|' + (jusqu ? Utilities.formatDate(jusqu, FUSEAU, 'yyyy-MM-dd') : '') });
  });

  // 5. Clients à relancer : factures échues non payées sans relance récente, regroupées par client
  var shFa = ss.getSheetByName('FACTURES');
  if (shFa) {
    var joursRelance = nombre_(p.JOURS_RELANCE, 15);
    var parClient = {};
    lireTable_(shFa).forEach(function (l) {
      if (!factureEnAttente_(l.Statut)) return;
      var ech = l.Echeance instanceof Date ? l.Echeance : dateDepuisTexte_(l.Echeance);
      if (!ech || minuit_(ech) >= jour0) return;
      var relance = l.Derniere_Relance instanceof Date ? l.Derniere_Relance : dateDepuisTexte_(l.Derniere_Relance);
      if (relance && (jour0 - minuit_(relance)) / 86400000 < joursRelance) return;
      var c = parClient[l.Client] = parClient[l.Client] || { nb: 0, montant: 0, retardMax: 0, relance: null };
      c.nb++;
      c.montant += nombre_(l.Montant_TTC, 0);
      c.retardMax = Math.max(c.retardMax, Math.round((jour0 - minuit_(ech)) / 86400000));
      if (relance && (!c.relance || relance > c.relance)) c.relance = relance;
    });
    Object.keys(parClient).forEach(function (client) {
      var c = parClient[client];
      var depuis = c.relance ? 'dernière relance le ' + Utilities.formatDate(c.relance, FUSEAU, 'dd/MM/yyyy') : 'jamais relancé';
      alertes.push({ niveau: c.retardMax > 45 ? 'depasse' : 'urgent', categorie: 'relances', domaine: 'Factures', objet: client,
        sujet: 'Relance', jours: -c.retardMax,
        message: client + ' : ' + c.nb + ' facture(s) impayée(s), ' + euros_(c.montant) + ', échue(s) depuis ' + c.retardMax + ' j (' + depuis + ')',
        cle: 'RELANCE|' + client + '|' + (c.relance ? Utilities.formatDate(c.relance, FUSEAU, 'yyyy-MM-dd') : 'jamais') });
    });
  }

  // 6. Marge par activité en baisse : dernier mois comparé aux 3 mois précédents
  if (shF) {
    var activiteCamion = {};
    flotte.forEach(function (l) { if (l.Activite) activiteCamion[l.Camion_ID] = String(l.Activite).trim(); });
    alertes = alertes.concat(baissesMarge_(lireTable_(shF), activiteCamion, nombre_(p.SEUIL_BAISSE_MARGE_PTS, 5)));
  }

  // 7. Taux de retards du mois en cours
  var shL = ss.getSheetByName('LIVRAISONS');
  if (shL) {
    var moisCourant = Utilities.formatDate(jour0, FUSEAU, 'yyyy-MM'), nbL = 0, nbR = 0;
    lireTable_(shL).forEach(function (l) {
      var dt = l.Date instanceof Date ? l.Date : dateDepuisTexte_(l.Date);
      if (!dt || Utilities.formatDate(dt, FUSEAU, 'yyyy-MM') !== moisCourant) return;
      nbL += nombre_(l.Nb_Livraisons, 0);
      nbR += nombre_(l.Nb_Retards, 0);
    });
    var seuilRetards = nombre_(p.SEUIL_RETARDS_PCT, 5);
    if (nbL >= 20 && nbR / nbL * 100 > seuilRetards) {
      alertes.push({ niveau: 'urgent', categorie: 'retards', domaine: 'Livraisons', objet: moisCourant, sujet: 'Retards',
        message: fr1_(nbR / nbL * 100) + ' % de livraisons en retard ce mois-ci (' + nbR + ' sur ' + nbL + ', seuil ' + seuilRetards + ' %)',
        cle: 'RETARDS|' + moisCourant });
    }
  }

  var ordre = { depasse: 0, urgent: 1, a_prevoir: 2 };
  alertes.sort(function (a, b) {
    return (ordre[a.niveau] - ordre[b.niveau]) || ((a.jours === undefined ? 999 : a.jours) - (b.jours === undefined ? 999 : b.jours));
  });
  return alertes;
}

/**
 * Compare, par activité (colonne Activite de FLOTTE, à défaut la société), la marge du dernier mois importé
 * à celle des 3 mois précédents cumulés. Seuls les camions dont le CA est connu sont pris en compte.
 */
function baissesMarge_(finances, activiteCamion, seuilPts) {
  var mois = {};
  finances.forEach(function (l) { var m = moisTexte_(l.Mois); if (m) mois[m] = 1; });
  var liste = Object.keys(mois).sort();
  if (liste.length < 2) return [];
  var dernier = liste[liste.length - 1], precedents = liste.slice(-4, -1);
  var agg = {};
  finances.forEach(function (l) {
    var m = moisTexte_(l.Mois), periode = m === dernier ? 'd' : (precedents.indexOf(m) >= 0 ? 'p' : null);
    var ca = nombre_(l.CA, null);
    if (!periode) return;
    var act = activiteCamion[l.Camion_ID] ? 'l\'activité ' + activiteCamion[l.Camion_ID] : 'la société ' + (l['Société'] || '?');
    var a = agg[act] = agg[act] || { d: { ca: 0, ch: 0 }, p: { ca: 0, ch: 0 } };
    if (ca !== null) a[periode].ca += ca;
    ['Carburant', 'Peages', 'Salaires', 'Entretien', 'Charges_Fixes', 'Charges_Mutualisees'].forEach(function (k) {
      a[periode].ch += nombre_(l[k], 0);
    });
  });
  var res = [];
  Object.keys(agg).forEach(function (act) {
    var a = agg[act];
    if (a.d.ca <= 0 || a.p.ca <= 0) return;
    var md = (a.d.ca - a.d.ch) / a.d.ca * 100, mp = (a.p.ca - a.p.ch) / a.p.ca * 100;
    if (mp - md < seuilPts) return;
    res.push({ niveau: 'urgent', categorie: 'marge', domaine: 'Finances', objet: act, sujet: 'Marge en baisse',
      message: 'Marge de ' + act + ' en baisse : ' + fr1_(md) + ' % en ' + dernier + ' contre ' + fr1_(mp)
        + ' % les ' + precedents.length + ' mois précédents',
      cle: 'BAISSE_MARGE|' + act + '|' + dernier });
  });
  return res;
}

function estIndisponible_(statut) {
  var s = String(statut || '').trim().toLowerCase();
  return s !== '' && s.indexOf('dispo') !== 0;
}

function factureEnAttente_(statut) {
  var s = String(statut || '').trim().toLowerCase();
  return s.indexOf('pay') !== 0 && s !== 'réglée' && s !== 'reglee' && s !== 'annulée' && s !== 'annulee';
}

/** L/100 km par camion : litres des pleins après le premier ÷ km parcourus entre premier et dernier plein. */
function consommations_(pleins, jour0, nbJours) {
  var depuis = new Date(jour0.getTime() - nbJours * 86400000);
  var parCamion = {};
  pleins.forEach(function (l) {
    var dt = l.Date instanceof Date ? l.Date : dateDepuisTexte_(l.Date);
    var km = nombre_(l.KM_Compteur, null), litres = nombre_(l.Litres, null);
    if (!dt || dt < depuis || km === null || litres === null || !l.Camion_ID) return;
    (parCamion[l.Camion_ID] = parCamion[l.Camion_ID] || []).push({ km: km, litres: litres });
  });
  var res = {};
  Object.keys(parCamion).forEach(function (id) {
    var p = parCamion[id].sort(function (a, b) { return a.km - b.km; });
    if (p.length < 2) return;
    var distance = p[p.length - 1].km - p[0].km;
    var litres = p.slice(1).reduce(function (s, x) { return s + x.litres; }, 0);
    if (distance > 0) res[id] = litres / distance * 100;
  });
  return res;
}

function verifierAlertesManuel() { verifierAlertes(true); }

/**
 * Envoie un mail si de nouvelles alertes sont apparues (ou chaque lundi un récapitulatif complet).
 * Les alertes déjà envoyées sont mémorisées dans JOURNAL_ALERTES pour ne pas répéter le même mail.
 */
function verifierAlertes(forcer) {
  var ss = classeur_();
  var p = lireParametres_(ss);
  var destinataires = String(p.EMAIL_ALERTES || '').trim();
  var alertes = calculerAlertes_(ss, new Date());
  var shJ = ss.getSheetByName('JOURNAL_ALERTES');
  var dejaEnvoyees = {};
  lireTable_(shJ).forEach(function (l) { dejaEnvoyees[l['Clé'] + '|' + l.Niveau] = true; });
  var nouvelles = alertes.filter(function (a) { return !dejaEnvoyees[a.cle + '|' + a.niveau]; });
  var lundi = new Date().getDay() === 1 && String(p.RECAP_HEBDO || 'OUI').toUpperCase() === 'OUI';

  if (!destinataires) {
    console.warn('EMAIL_ALERTES vide dans PARAMETRES : aucun mail envoyé.');
    if (forcer) SpreadsheetApp.getUi().alert('Renseignez EMAIL_ALERTES dans l\'onglet PARAMETRES.');
    return;
  }
  if (!alertes.length || (!nouvelles.length && !lundi && !forcer)) {
    if (forcer) SpreadsheetApp.getUi().alert('Aucune alerte à signaler aujourd\'hui.');
    return;
  }
  var nbGraves = alertes.filter(function (a) { return a.niveau !== 'a_prevoir'; }).length;
  var sujet = '[Tableau de bord] ' + (nouvelles.length ? nouvelles.length + ' nouvelle(s) alerte(s)' : 'Récapitulatif')
    + ' — ' + nbGraves + ' urgente(s) ou dépassée(s)';
  MailApp.sendEmail({ to: destinataires, subject: sujet, htmlBody: mailHtml_(alertes, nouvelles), name: 'Tableau de bord Gleyzes / LPB' });

  var maintenant = new Date();
  var lignes = nouvelles.map(function (a) { return [maintenant, a.cle, a.niveau, a.message]; });
  if (lignes.length) shJ.getRange(shJ.getLastRow() + 1, 1, lignes.length, 4).setValues(lignes);
  if (forcer) SpreadsheetApp.getUi().alert('Mail envoyé à ' + destinataires + ' (' + alertes.length + ' alerte(s)).');
}

function mailHtml_(alertes, nouvelles) {
  var estNouvelle = {};
  nouvelles.forEach(function (a) { estNouvelle[a.cle + '|' + a.niveau] = true; });
  var styles = {
    depasse: ['#d03b3b', '⛔ Dépassé'], urgent: ['#ec835a', '⚠️ Urgent'], a_prevoir: ['#b07d00', '🕒 À prévoir']
  };
  var lignes = alertes.map(function (a) {
    var s = styles[a.niveau];
    return '<tr><td style="padding:6px 10px;color:' + s[0] + ';font-weight:600;white-space:nowrap">' + s[1] + '</td>'
      + '<td style="padding:6px 10px;color:#52514e">' + echapper_(a.domaine) + '</td>'
      + '<td style="padding:6px 10px">' + echapper_(a.message)
      + (estNouvelle[a.cle + '|' + a.niveau] ? ' <b style="color:#2a78d6">NOUVEAU</b>' : '') + '</td></tr>';
  }).join('');
  var url = '';
  try { url = ScriptApp.getService().getUrl() || ''; } catch (e) { /* pas encore déployé */ }
  return '<div style="font-family:Arial,sans-serif;font-size:14px;color:#0b0b0b">'
    + '<p>Bonjour,</p><p>Voici les points à surveiller au ' + Utilities.formatDate(new Date(), FUSEAU, 'dd/MM/yyyy') + ' :</p>'
    + '<table style="border-collapse:collapse;border:1px solid #e1e0d9">' + lignes + '</table>'
    + (url ? '<p><a href="' + url + '">Ouvrir le tableau de bord</a></p>' : '')
    + '<p style="color:#8a8983;font-size:12px">Mail automatique. Seuils et destinataires : onglet PARAMETRES du Google Sheet.</p></div>';
}

// ---------------------------------------------------------------------------
// Import des finances (CSV produit par exporter_finances_csv.py)
// ---------------------------------------------------------------------------
function importerFinancesManuel() {
  var msg = importerFinances(true);
  SpreadsheetApp.getUi().alert(msg);
}

/**
 * Lit le fichier FINANCES_EXPORT.csv le plus récent dans Google Drive.
 * Les mois présents dans le fichier remplacent les mêmes mois dans FINANCES ; les autres mois sont conservés.
 * Sans « forcer », un fichier déjà importé (même date de modification) est ignoré.
 */
function importerFinances(forcer) {
  var ss = classeur_();
  var p = lireParametres_(ss);
  var nomFichier = String(p.NOM_FICHIER_FINANCES || 'FINANCES_EXPORT.csv');
  var it = DriveApp.getFilesByName(nomFichier), fichier = null;
  while (it.hasNext()) {
    var f = it.next();
    if (!f.isTrashed() && (!fichier || f.getLastUpdated() > fichier.getLastUpdated())) fichier = f;
  }
  if (!fichier) return 'Fichier « ' + nomFichier + ' » introuvable dans Google Drive.';
  var props = PropertiesService.getScriptProperties();
  var marque = String(fichier.getLastUpdated().getTime());
  if (!forcer && props.getProperty('DERNIER_IMPORT') === marque) return 'Déjà importé, rien de nouveau.';

  var texte = fichier.getBlob().getDataAsString('UTF-8').replace(/^﻿/, '');
  var resultat = fusionnerFinances_(ss, Utilities.parseCsv(texte, ';'));
  props.setProperty('DERNIER_IMPORT', marque);
  props.setProperty('DERNIER_IMPORT_TEXTE', Utilities.formatDate(fichier.getLastUpdated(), FUSEAU, 'dd/MM/yyyy HH:mm'));
  return resultat;
}

function fusionnerFinances_(ss, lignesCsv) {
  if (!lignesCsv.length) return 'Fichier vide.';
  var entetes = lignesCsv[0].map(function (h) { return String(h).trim(); });
  var attendues = FEUILLES.FINANCES;
  var manquantes = attendues.filter(function (h) { return entetes.indexOf(h) < 0; });
  if (manquantes.length) throw new Error('Colonnes manquantes dans le CSV : ' + manquantes.join(', '));

  var nouvelles = lignesCsv.slice(1).filter(function (l) { return l.join('').trim() !== ''; }).map(function (l) {
    return attendues.map(function (h) {
      var v = String(l[entetes.indexOf(h)] || '').trim();
      if (h === 'Mois' || h === 'Camion_ID' || h === 'Société') return v;
      return v === '' ? '' : nombre_(v, '');  // vide = donnée manquante, jamais 0
    });
  });
  var moisImportes = {};
  nouvelles.forEach(function (l) { moisImportes[l[0]] = true; });

  var sh = ss.getSheetByName('FINANCES');
  var conservees = lireTable_(sh).filter(function (l) { return !moisImportes[moisTexte_(l.Mois)]; })
    .map(function (l) { return attendues.map(function (h) { return h === 'Mois' ? moisTexte_(l.Mois) : l[h]; }); });
  var toutes = conservees.concat(nouvelles).sort(function (a, b) {
    return a[0] < b[0] ? -1 : a[0] > b[0] ? 1 : (a[1] < b[1] ? -1 : a[1] > b[1] ? 1 : 0);
  });
  if (sh.getLastRow() > 1) sh.getRange(2, 1, sh.getLastRow() - 1, sh.getLastColumn()).clearContent();
  sh.getRange(1, 1, 1, attendues.length).setValues([attendues]);
  if (toutes.length) {
    sh.getRange(2, 1, toutes.length, 1).setNumberFormat('@');  // garde « 2026-08 » en texte
    sh.getRange(2, 1, toutes.length, attendues.length).setValues(toutes);
  }
  return nouvelles.length + ' ligne(s) importée(s) pour ' + Object.keys(moisImportes).sort().join(', ') + '.';
}

// ---------------------------------------------------------------------------
// Utilitaires
// ---------------------------------------------------------------------------
function lireTable_(sh) {
  if (!sh || sh.getLastRow() < 2) return [];
  var valeurs = sh.getRange(1, 1, sh.getLastRow(), sh.getLastColumn()).getValues();
  var entetes = valeurs[0].map(function (h) { return String(h).trim(); });
  return valeurs.slice(1).filter(function (l) { return l.join('') !== ''; }).map(function (l) {
    var o = {};
    entetes.forEach(function (h, i) { if (h) o[h] = l[i]; });
    return o;
  });
}

function lireParametres_(ss) {
  var sh = (ss || classeur_()).getSheetByName('PARAMETRES');
  var p = {};
  lireTable_(sh).forEach(function (l) { p[String(l['Paramètre']).trim()] = l.Valeur; });
  return p;
}

function majParametre_(nom, valeur) {
  var sh = classeur_().getSheetByName('PARAMETRES');
  var v = sh.getRange(1, 1, sh.getLastRow(), 2).getValues();
  for (var i = 1; i < v.length; i++) {
    if (v[i][0] === nom) { sh.getRange(i + 1, 2).setValue(valeur); return; }
  }
}

function estActif_(l) {
  var a = String(l.Actif === undefined ? '' : l.Actif).trim().toUpperCase();
  return a === '' || a === 'O' || a === 'OUI' || a === 'TRUE' || a === 'VRAI';
}

function serialiser_(o) {
  var r = {};
  Object.keys(o).forEach(function (k) {
    var v = o[k];
    r[k] = v instanceof Date ? Utilities.formatDate(v, FUSEAU, 'yyyy-MM-dd') : v;
  });
  return r;
}

/** Accepte un nombre, « 1 234,56 », « 1234.56 € »… Renvoie defaut si illisible. */
function nombre_(v, defaut) {
  if (typeof v === 'number') return isNaN(v) ? defaut : v;
  if (v === null || v === undefined) return defaut;
  var s = String(v).replace(/[\s €%]/g, '');
  if (s === '') return defaut;
  if (s.indexOf(',') >= 0 && s.indexOf('.') >= 0) s = s.replace(/\./g, '');
  s = s.replace(',', '.');
  var n = Number(s);
  return isNaN(n) ? defaut : n;
}

/** Accepte une Date, « 2026-10-01 » ou « 01/10/2026 ». */
function dateDepuisTexte_(v) {
  if (v instanceof Date) return isNaN(v.getTime()) ? null : v;
  var s = String(v || '').trim(), m;
  if ((m = s.match(/^(\d{4})-(\d{2})-(\d{2})/))) return new Date(+m[1], +m[2] - 1, +m[3]);
  if ((m = s.match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})$/))) return new Date(+m[3], +m[2] - 1, +m[1]);
  return null;
}

/** Ramène un mois (Date ou texte) au format AAAA-MM. */
function moisTexte_(v) {
  if (v instanceof Date) return Utilities.formatDate(v, FUSEAU, 'yyyy-MM');
  var s = String(v || '').trim(), m;
  if ((m = s.match(/^(\d{4})-(\d{1,2})/))) return m[1] + '-' + ('0' + m[2]).slice(-2);
  if ((m = s.match(/^(\d{1,2})\/(\d{4})$/))) return m[2] + '-' + ('0' + m[1]).slice(-2);
  return s;
}

function minuit_(d) { return new Date(d.getFullYear(), d.getMonth(), d.getDate()); }

function fr1_(n) { return n.toFixed(1).replace('.', ','); }

function euros_(n) {
  return n.toFixed(0).replace(/\B(?=(\d{3})+(?!\d))/g, ' ') + ' €';
}

function echapper_(s) {
  return String(s).replace(/[&<>"]/g, function (c) { return { '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;' }[c]; });
}

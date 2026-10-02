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
 * Si PARAMETRES > ID_CLASSEUR_ECHEANCES contient le lien du Google Sheet « Échéances flotte », ses onglets
 * « Toutes les échéances », « CONGES… » et « Pense-bête mensuel » sont lus directement (rien n'est recopié).
 * De même pour ID_PLANNING_INTER, ID_PLANNING_CARBURANT (un onglet par mois, une ligne par tournée/livraison)
 * et ID_LITRAGES (un onglet par véhicule, km début / km fin / litres par semaine).
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
           'Activite', 'Statut', 'Indisponible_Jusqu_Au', 'Type'],
  SALARIES: ['Nom', 'Prénom', 'Société', 'Poste', 'Fin_Validite_Permis', 'Fin_FIMO_FCO',
             'Prochaine_Visite_Medicale', 'Fin_Carte_Conducteur', 'Formation_A_Prevoir',
             'Date_Formation', 'Actif', 'Remarques'],
  ABSENCES: ['Salarié', 'Type', 'Début', 'Fin', 'Remarque'],
  CARBURANT: ['Date', 'Camion_ID', 'Société', 'Litres', 'Montant_TTC', 'KM_Compteur', 'Lieu'],
  LIVRAISONS: ['Date', 'Camion_ID', 'Chauffeur', 'Client', 'Nb_Livraisons', 'CA_HT', 'Remarque', 'Nb_Retards'],
  FACTURES: ['N_Facture', 'Client', 'Date_Facture', 'Echeance', 'Montant_TTC', 'Statut', 'Date_Paiement',
             'Derniere_Relance', 'Remarque'],
  HEURES: ['Date', 'Salarié', 'Heures', 'Remarque'],
  CHARGES_MUTUALISEES: ['Poste', 'Société', 'Montant_Mensuel', 'Debut', 'Fin'],
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
  HEURES: ['Date'], LIVRAISONS: ['Date'], CARBURANT: ['Date'], ABSENCES: ['Début', 'Fin'], CHARGES_MUTUALISEES: ['Debut', 'Fin']
};

var PARAMETRES_DEFAUT = [
  ['EMAIL_ALERTES', '', 'Adresse(s) qui reçoivent les alertes, séparées par des virgules'],
  ['ID_CLASSEUR_ECHEANCES', '', 'Lien du Google Sheet « Échéances flotte » : échéances, congés et pense-bête y sont lus directement'],
  ['ID_PLANNING_INTER', '', 'Lien du Google Sheet « Planning Inter (ITM) » : CA et tournées lus directement'],
  ['ID_PLANNING_CARBURANT', '', 'Lien du Google Sheet « Planning Carburant » : CA et livraisons lus directement'],
  ['ID_LITRAGES', '', 'Lien du Google Sheet « Litrages véhicules » : consommation par semaine lue directement'],
  ['SEUIL_HAUSSE_CONSO_PCT', 15, 'Alerte si la consommation récente dépasse de ce % la moyenne habituelle du camion'],
  ['JOURS_MOIS_PRECEDENT', 10, 'Pendant ces premiers jours du mois, les alertes des plannings portent aussi sur le mois précédent (ensuite : mois en cours seulement)'],
  ['NB_SEMAINES_CONSO', 2, 'Nombre de dernières semaines regroupées pour juger la consommation (1 = très réactif mais beaucoup de fausses alertes)'],
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
    .addItem('Remplir la colonne CAMION des plannings', 'preparerCamionsManuel')
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
  console.log('Onglets vérifiés');
  var shP = ss.getSheetByName('PARAMETRES');
  var existants = lireTable_(shP).map(function (l) { return l['Paramètre']; });
  var nouveauxP = PARAMETRES_DEFAUT.filter(function (p) { return existants.indexOf(p[0]) < 0; });
  if (nouveauxP.length) shP.getRange(shP.getLastRow() + 1, 1, nouveauxP.length, 3).setValues(nouveauxP);
  CACHE_PARAMS_ = null;
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
  console.log('Formats et listes déroulantes posés');
  var synchro = '';
  try {
    synchro = synchroniserReferentiels_(ss);
    console.log('Véhicules et chauffeurs synchronisés');
  } catch (e) {
    synchro = '\n\nTableau « Échéances flotte » illisible : ' + e.message;
  }
  // Vérifie chaque lien de fichier et l'indique clairement s'il ne fonctionne pas
  var liensOk = [];
  ['ID_PLANNING_INTER', 'ID_PLANNING_CARBURANT', 'ID_LITRAGES'].forEach(function (k) {
    try { var f = classeurParam_(ss, k); if (f) liensOk.push(f.getName()); }
    catch (e) { synchro += '\n\n⚠ ' + e.message; }
  });
  if (liensOk.length) synchro += '\n\nFichiers reliés : ' + liensOk.join(', ') + '.';
  if (lireParametres_(ss).ID_PLANNING_INTER || lireParametres_(ss).ID_PLANNING_CARBURANT) {
    synchro += '\n\nÉtape suivante : menu Tableau de bord > Remplir la colonne CAMION des plannings.';
  }
  try {
    SpreadsheetApp.getUi().alert('Onglets prêts.' + synchro + '\n\nComplétez FLOTTE (Société, Activite, Statut) et SALARIES (Société), puis activez les envois automatiques (menu Tableau de bord > 2).');
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
  try { preparerCamionsPlannings_(classeur_()); } catch (e) { console.error('Colonne CAMION : ' + e); }
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
  ['FINANCES', 'FLOTTE', 'SALARIES', 'ABSENCES', 'CARBURANT', 'LIVRAISONS', 'FACTURES', 'HEURES', 'CHARGES_MUTUALISEES'].forEach(function (nom) {
    var sh = ss.getSheetByName(nom);
    d[nom] = sh ? lireTable_(sh).map(serialiser_) : [];
  });
  d.FINANCES.forEach(function (l) { l.Mois = moisTexte_(l.Mois); });
  d.FLOTTE = d.FLOTTE.filter(estActif_);
  d.SALARIES = d.SALARIES.filter(estActif_);
  try {
    d.ECHEANCES = echeances_(ss).map(function (e) {
      return { categorie: e.categorie, domaine: e.domaine, nom: e.nom, document: e.document, periodicite: e.periodicite,
        echeance: e.date ? Utilities.formatDate(e.date, FUSEAU, 'yyyy-MM-dd') : '' };
    });
    d.ABSENCES = d.ABSENCES.concat(absencesExternes_(ss).map(serialiser_));
  } catch (e) {
    d.ECHEANCES = d.ECHEANCES || [];
    d.erreurEcheances = 'Tableau « Échéances flotte » illisible : ' + e.message;
  }
  d.erreurs = [];
  try {
    d.PLANNING = plannings_(ss).map(function (l) {
      var o = serialiser_(l);
      o.date = Utilities.formatDate(l.date, FUSEAU, 'yyyy-MM-dd');
      return o;
    });
  } catch (e) { d.PLANNING = []; d.erreurs.push('Plannings illisibles : ' + e.message); }
  try {
    d.LITRAGES = litrages_(ss).map(function (l) {
      var o = serialiser_(l);
      o.fin = l.fin ? Utilities.formatDate(l.fin, FUSEAU, 'yyyy-MM-dd') : '';
      return o;
    });
  } catch (e) { d.LITRAGES = []; d.erreurs.push('Litrages illisibles : ' + e.message); }
  var params = lireParametres_();
  d.parametres = {
    JOURS_PREAVIS: nombre_(params.JOURS_PREAVIS, 30),
    JOURS_URGENT: nombre_(params.JOURS_URGENT, 7),
    SEUIL_MARGE_PCT: nombre_(params.SEUIL_MARGE_PCT, 5),
    SEUIL_CONSO_L100: nombre_(params.SEUIL_CONSO_L100, 38),
    JOURS_RELANCE: nombre_(params.JOURS_RELANCE, 15),
    SEUIL_RETARDS_PCT: nombre_(params.SEUIL_RETARDS_PCT, 5),
    sourceEcheances: !!String(params.ID_CLASSEUR_ECHEANCES || '').trim(),
    sourceInter: !!String(params.ID_PLANNING_INTER || '').trim(),
    sourceCarburant: !!String(params.ID_PLANNING_CARBURANT || '').trim(),
    sourceLitrages: !!String(params.ID_LITRAGES || '').trim(),
    SEUIL_HAUSSE_CONSO_PCT: nombre_(params.SEUIL_HAUSSE_CONSO_PCT, 15),
    NB_SEMAINES_CONSO: Math.max(1, nombre_(params.NB_SEMAINES_CONSO, 2))
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
    var conges = type === 'ABSENCES' ? ongletExterne_(ss, /CONGE/) : null;
    if (conges) {
      // Les absences vont dans l'onglet CONGES du tableau « Échéances flotte », là où elles sont déjà suivies
      var correspondance = { CHAUFFEUR: 'Salarié', SALARIE: 'Salarié', TYPE: 'Type', DEBUT: 'Début', FIN: 'Fin', COMMENTAIRE: 'Remarque', REMARQUE: 'Remarque' };
      var entetesC = conges.getRange(1, 1, 1, conges.getLastColumn()).getValues()[0];
      conges.appendRow(entetesC.map(function (h) { var k = correspondance[cle_(h)]; return k ? ligneObj[k] : ''; }));
      return true;
    }
    var entetes = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0];
    sh.appendRow(entetes.map(function (h) { return ligneObj.hasOwnProperty(h) ? ligneObj[h] : ''; }));
  } finally {
    verrou.releaseLock();
  }
  return true;
}

/**
 * Depuis le site (onglet Charges) : remplace tout le contenu de l'onglet CHARGES_MUTUALISEES par les lignes reçues.
 * Chaque ligne : { Poste, Société (Gleyzes / LPB), Montant_Mensuel, Debut, Fin } ; dates au format AAAA-MM-JJ ou vides.
 */
function enregistrerCharges(lignes) {
  if (!Array.isArray(lignes)) throw new Error('Données invalides');
  var propres = lignes.filter(function (l) {
    return String(l.Poste || '').trim() || nombre_(l.Montant_Mensuel, null) !== null;
  }).map(function (l, i) {
    var poste = String(l.Poste || '').trim().replace(/^[=+\-@]/, "'$&").slice(0, 120);
    if (!poste) throw new Error('Ligne ' + (i + 1) + ' : le poste est vide');
    var soc = cle_(l['Société']);
    soc = soc.indexOf('LPB') >= 0 ? 'LPB' : (soc.indexOf('GLEYZES') >= 0 ? 'GLEYZES' : '');
    if (!soc) throw new Error('Ligne « ' + poste + ' » : choisissez la société (Gleyzes ou LPB)');
    var montant = nombre_(l.Montant_Mensuel, null);
    if (montant === null || montant < 0) throw new Error('Ligne « ' + poste + ' » : montant invalide');
    var date = function (v, nom) {
      if (!String(v || '').trim()) return '';
      var d = dateDepuisTexte_(v);
      if (!d) throw new Error('Ligne « ' + poste + ' » : date de ' + nom + ' invalide');
      return d;
    };
    var debut = date(l.Debut, 'début'), fin = date(l.Fin, 'fin');
    if (debut && fin && fin < debut) throw new Error('Ligne « ' + poste + ' » : la fin est avant le début');
    return [poste, soc, montant, debut, fin];
  });
  var verrou = LockService.getScriptLock();
  verrou.waitLock(10000);
  try {
    var sh = classeur_().getSheetByName('CHARGES_MUTUALISEES');
    if (!sh) throw new Error('Onglet CHARGES_MUTUALISEES introuvable : lancez « Installer / compléter les onglets »');
    var entetes = FEUILLES.CHARGES_MUTUALISEES;
    sh.getRange(1, 1, 1, entetes.length).setValues([entetes]);
    if (sh.getLastRow() > 1) sh.getRange(2, 1, sh.getLastRow() - 1, Math.max(entetes.length, sh.getLastColumn())).clearContent();
    if (propres.length) sh.getRange(2, 1, propres.length, entetes.length).setValues(propres);
  } finally {
    verrou.releaseLock();
  }
  return propres.length;
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

  // 1. Échéances : documents des véhicules, des chauffeurs et de l'entreprise
  var aRenseigner = 0;
  echeances_(ss).forEach(function (e) {
    if (!e.date) { aRenseigner++; return; }
    var jours = Math.round((minuit_(e.date) - jour0) / 86400000);
    if (jours > preavis) return;
    var niveau = jours < 0 ? 'depasse' : (jours <= urgent ? 'urgent' : 'a_prevoir');
    var quand = jours < 0 ? 'dépassé depuis ' + (-jours) + ' j' : (jours === 0 ? 'aujourd\'hui' : 'dans ' + jours + ' j');
    var dateTxt = Utilities.formatDate(e.date, FUSEAU, 'dd/MM/yyyy');
    alertes.push({
      niveau: niveau, categorie: 'documents', domaine: e.domaine, objet: e.nom, sujet: e.document, jours: jours, date: dateTxt,
      message: e.document + ' — ' + e.nom + ' : ' + dateTxt + ' (' + quand + ')',
      cle: 'ECH|' + e.nom + '|' + e.document + '|' + Utilities.formatDate(e.date, FUSEAU, 'yyyy-MM-dd')
    });
  });
  if (aRenseigner) {
    alertes.push({ niveau: 'a_prevoir', categorie: 'arenseigner', domaine: 'Échéances', objet: '', sujet: 'Dates manquantes',
      message: aRenseigner + ' date(s) de contrôle à renseigner dans le tableau des échéances', cle: 'ARENSEIGNER' });
  }

  // 1 bis. Pense-bête mensuel (TVA, péages…) : rappel 7 jours avant la date du mois
  alertes = alertes.concat(taches_(ss, jour0));

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

  // 8. Plannings et litrages (Google Sheets reliés) : une erreur de lecture ne bloque pas les autres alertes
  var debutPlannings = debutPeriodePlannings_(jour0, nombre_(p.JOURS_MOIS_PRECEDENT, 10));
  try { alertes = alertes.concat(alertesPlannings_(plannings_(ss), jour0, debutPlannings)); } catch (e) { console.warn('Plannings : ' + e); }
  try { alertes = alertes.concat(alertesCamionsPlannings_(ss, plannings_(ss), jour0, debutPlannings)); } catch (e) { console.warn('Camions des plannings : ' + e); }
  try {
    alertes = alertes.concat(alertesConsoHebdo_(litrages_(ss), jour0, nombre_(p.SEUIL_CONSO_L100, 38), nombre_(p.SEUIL_HAUSSE_CONSO_PCT, 15),
      Math.max(1, nombre_(p.NB_SEMAINES_CONSO, 2))));
  } catch (e) { console.warn('Litrages : ' + e); }

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

// ---------------------------------------------------------------------------
// Tableau « Échéances flotte » (Google Sheet séparé, lu directement)
// ---------------------------------------------------------------------------
var CACHE_EXTERNE_ = {};

function classeurEcheances_(ss) { return classeurParam_(ss, 'ID_CLASSEUR_ECHEANCES'); }

/** Ouvre le Google Sheet dont le lien (ou l'identifiant) est dans PARAMETRES > nomParam ; null si vide. */
function classeurParam_(ss, nomParam) {
  var v = String(lireParametres_(ss)[nomParam] || '').trim();
  if (!v) return null;
  if (CACHE_EXTERNE_[v]) return CACHE_EXTERNE_[v];
  var m = v.match(/\/d\/([a-zA-Z0-9_-]{20,})/), id = m ? m[1] : (/^[a-zA-Z0-9_-]{25,}$/.test(v) ? v : null);
  if (!id) {
    // Un nom de fichier a été écrit au lieu du lien : on le cherche dans Google Drive (Google Sheets uniquement)
    var it = DriveApp.getFilesByName(v), trouve = null;
    while (it.hasNext()) {
      var f = it.next();
      if (!f.isTrashed() && f.getMimeType() === MimeType.GOOGLE_SHEETS && (!trouve || f.getLastUpdated() > trouve.getLastUpdated())) trouve = f;
    }
    if (!trouve) throw new Error('fichier « ' + v + ' » introuvable dans Google Drive (PARAMETRES > ' + nomParam +
      ') : collez plutôt le lien complet du fichier (https://docs.google.com/spreadsheets/d/…), ou vérifiez qu\'il est bien au format Google Sheets.');
    id = trouve.getId();
  }
  try {
    CACHE_EXTERNE_[v] = SpreadsheetApp.openById(id);
  } catch (e) {
    throw new Error('impossible d\'ouvrir le fichier indiqué dans PARAMETRES > ' + nomParam + ' (' + e.message + ')');
  }
  return CACHE_EXTERNE_[v];
}

/** Premier onglet du tableau « Échéances flotte » dont le nom (sans accents, en majuscules) correspond au motif. */
function ongletExterne_(ss, motif) {
  var ext = classeurEcheances_(ss);
  if (!ext) return null;
  return ext.getSheets().filter(function (sh) { return motif.test(cle_(sh.getName())); })[0] || null;
}

/** « Contrôle / Document » -> « CONTROLE/DOCUMENT » : sert à reconnaître les en-têtes quelle que soit leur écriture. */
function cle_(s) {
  return String(s || '').normalize('NFD').replace(/[̀-ͯ]/g, '').toUpperCase().replace(/[\s_]+/g, ' ').trim()
    .replace(/ ?\/ ?/g, '/');
}

function plaque_(s) { return String(s || '').toUpperCase().replace(/[\s-]/g, ''); }

var CATEGORIES_VEHICULE = /TRACTEUR|REMORQUE|CITERNE|PORTEUR|VEHICULE|TELECHARGEMENT/;

/**
 * Toutes les échéances suivies : colonnes de date de FLOTTE / SALARIES (si remplies)
 * + toutes les lignes du tableau « Échéances flotte » (y compris celles sans date, date = null).
 */
var CACHE_ECHEANCES_ = null;
function echeances_(ss) {
  if (CACHE_ECHEANCES_) return CACHE_ECHEANCES_;
  var res = echeancesLues_(ss);
  return (CACHE_ECHEANCES_ = res);
}

function echeancesLues_(ss) {
  var res = [];
  [['FLOTTE', 'Véhicule', 'Flotte'], ['SALARIES', 'Chauffeur', 'Salariés']].forEach(function (f) {
    var sh = ss.getSheetByName(f[0]);
    if (!sh) return;
    lireTable_(sh).filter(estActif_).forEach(function (l) {
      var nom = f[0] === 'FLOTTE' ? plaque_(l.Camion_ID) : ((l['Prénom'] || '') + ' ' + (l.Nom || '')).trim();
      if (!nom) return;
      ECHEANCES[f[0]].forEach(function (e) {
        var dt = l[e[0]] instanceof Date ? l[e[0]] : dateDepuisTexte_(l[e[0]]);
        if (!dt) return;
        var doc = e[1] + (e[0] === 'Date_Formation' && l.Formation_A_Prevoir ? ' (' + l.Formation_A_Prevoir + ')' : '');
        res.push({ categorie: f[0] === 'FLOTTE' ? (l.Type || f[1]) : f[1], domaine: f[2], nom: nom, document: doc, date: dt, periodicite: null });
      });
    });
  });

  var sh = ongletExterne_(ss, /ECHEANCE/);
  if (!sh || sh.getLastRow() < 2) return res;
  var v = sh.getRange(1, 1, sh.getLastRow(), sh.getLastColumn()).getValues();
  var e = v[0].map(cle_);
  var col = function (motif) { for (var i = 0; i < e.length; i++) if (motif.test(e[i])) return i; return -1; };
  var cCat = col(/^CATEGORIE/), cNom = col(/^NOM/), cDoc = col(/CONTROLE|DOCUMENT/), cPer = col(/^PERIODICITE/),
      cDer = col(/^DERNIER/), cEch = col(/^ECHEANCE/);
  if (cNom < 0 || cDoc < 0 || cEch < 0) throw new Error('colonnes « Nom », « Contrôle / Document » ou « Échéance » introuvables');
  var cat = '', nom = '';
  v.slice(1).forEach(function (l) {
    if (cCat >= 0 && String(l[cCat]).trim()) cat = String(l[cCat]).trim();
    if (String(l[cNom]).trim()) nom = String(l[cNom]).trim();
    var doc = String(l[cDoc] || '').trim();
    if (!doc || !nom) return;
    var date = l[cEch] instanceof Date ? l[cEch] : dateDepuisTexte_(l[cEch]);
    var per = cPer >= 0 ? nombre_(l[cPer], null) : null;
    var dernier = cDer >= 0 ? (l[cDer] instanceof Date ? l[cDer] : dateDepuisTexte_(l[cDer])) : null;
    if (!date && dernier && per) date = new Date(dernier.getFullYear(), dernier.getMonth() + per, dernier.getDate());
    var k = cle_(cat);
    var vehicule = CATEGORIES_VEHICULE.test(k);
    res.push({ categorie: cat || 'Autre', domaine: k === 'CHAUFFEUR' ? 'Salariés' : (vehicule ? 'Flotte' : 'Entreprise'),
      nom: vehicule ? plaque_(nom) : nom, document: doc, date: date, periodicite: per });
  });
  return res;
}

/** Onglet « CONGES - ABSENCES - FORMATION » du tableau « Échéances flotte », au format de l'onglet ABSENCES. */
function absencesExternes_(ss) {
  var sh = ongletExterne_(ss, /CONGE/);
  if (!sh || sh.getLastRow() < 2) return [];
  var v = sh.getRange(1, 1, sh.getLastRow(), sh.getLastColumn()).getValues();
  var e = v[0].map(cle_);
  var i = function (noms) { for (var k = 0; k < noms.length; k++) { var j = e.indexOf(noms[k]); if (j >= 0) return j; } return -1; };
  var cN = i(['CHAUFFEUR', 'SALARIE', 'NOM']), cT = i(['TYPE']), cD = i(['DEBUT']), cF = i(['FIN']), cC = i(['COMMENTAIRE', 'REMARQUE']);
  return v.slice(1).filter(function (l) { return cN >= 0 && String(l[cN]).trim(); }).map(function (l) {
    return { 'Salarié': String(l[cN]).trim(), Type: cT >= 0 ? l[cT] : '', 'Début': cD >= 0 ? l[cD] : '', Fin: cF >= 0 ? l[cF] : '',
      Remarque: cC >= 0 ? l[cC] : '' };
  });
}

/** Pense-bête mensuel : une alerte dans les 7 jours qui précèdent le jour du mois indiqué. */
function taches_(ss, jour0) {
  var sh = ongletExterne_(ss, /PENSE/);
  if (!sh || sh.getLastRow() < 2) return [];
  var v = sh.getRange(1, 1, sh.getLastRow(), sh.getLastColumn()).getValues();
  var res = [];
  v.slice(1).forEach(function (l) {
    var tache = String(l[0] || '').trim(), jour = nombre_(l[1], null), note = String(l[2] || '').trim();
    if (!tache || !jour) return;
    var dueCeMois = new Date(jour0.getFullYear(), jour0.getMonth(), Math.min(jour, new Date(jour0.getFullYear(), jour0.getMonth() + 1, 0).getDate()));
    var due = dueCeMois >= jour0 ? dueCeMois
      : new Date(jour0.getFullYear(), jour0.getMonth() + 1, Math.min(jour, new Date(jour0.getFullYear(), jour0.getMonth() + 2, 0).getDate()));
    var j = Math.round((due - jour0) / 86400000);
    if (j > 7) return;
    var dateTxt = Utilities.formatDate(due, FUSEAU, 'dd/MM/yyyy');
    res.push({ niveau: j <= 2 ? 'urgent' : 'a_prevoir', categorie: 'taches', domaine: 'Pense-bête', objet: tache, sujet: tache, jours: j,
      message: tache + ' : pour le ' + dateTxt + ' (' + (j === 0 ? 'aujourd\'hui' : 'dans ' + j + ' j') + ')' + (note ? ' — ' + note : ''),
      cle: 'TACHE|' + tache + '|' + Utilities.formatDate(due, FUSEAU, 'yyyy-MM-dd') });
  });
  return res;
}

// ---------------------------------------------------------------------------
// Plannings Inter (ITM) et Carburant : deux activités distinctes, jamais mélangées
// ---------------------------------------------------------------------------
var MOIS_ONGLETS = { JANVIER: 1, FEVRIER: 2, MARS: 3, AVRIL: 4, MAI: 5, JUIN: 6, JUILLET: 7, AOUT: 8, SEPTEMBRE: 9,
                     OCTOBRE: 10, NOVEMBRE: 11, DECEMBRE: 12 };
var CACHE_PLANNINGS_ = null, CACHE_LITRAGES_ = null;

/**
 * Une ligne par tournée (Inter) ou par livraison (Carburant), lue dans les onglets mensuels des deux plannings.
 * Les onglets « CA … », TARIFS, COMPARATIF… sont ignorés : le CA est recalculé ligne par ligne.
 */
function plannings_(ss) {
  if (CACHE_PLANNINGS_) return CACHE_PLANNINGS_;
  var res = [];
  [['ID_PLANNING_INTER', 'Inter'], ['ID_PLANNING_CARBURANT', 'Carburant']].forEach(function (src) {
    var ext = classeurParam_(ss, src[0]);
    if (!ext) return;
    ext.getSheets().forEach(function (sh) {
      var nom = cle_(sh.getName());
      if (/^CA\b|TARIF|COMPARATIF|RECAP|CAMIONS/.test(nom)) return;
      if (!Object.keys(MOIS_ONGLETS).some(function (m) { return nom.indexOf(m) >= 0; })) return;
      if (sh.getLastRow() < 2) return;
      var v = sh.getRange(1, 1, sh.getLastRow(), sh.getLastColumn()).getValues();
      // En-tête : la ligne qui contient « SOCIETE… » et « CHAUFFEUR… » (la case DATE est parfois écrasée par un n° de contrat)
      var ligneEntete = -1;
      for (var i = 0; i < Math.min(6, v.length); i++) {
        var h = v[i].map(cle_);
        if (h.some(function (x) { return x.indexOf('SOCIETE') === 0; }) && h.some(function (x) { return x.indexOf('CHAUFFEUR') === 0; })) { ligneEntete = i; break; }
      }
      if (ligneEntete < 0) return;
      var e = v[ligneEntete].map(cle_);
      // Mois et année de l'onglet : corrigent les fautes de frappe sur l'année (« 2025 » ou « 2029 » dans l'onglet JUIN 2026)
      var moisOnglet = MOIS_ONGLETS[Object.keys(MOIS_ONGLETS).filter(function (m) { return nom.indexOf(m) >= 0; })[0]];
      var an = (nom.match(/(20\d\d)/) || [])[1];
      var col = function (test) { for (var k = 0; k < e.length; k++) if (test(e[k])) return k; return -1; };
      var dernier = function (test) { for (var k = e.length - 1; k >= 0; k--) if (test(e[k])) return k; return -1; };
      var c = {
        date: Math.max(0, col(function (h) { return h === 'DATE' || h === 'DATES'; })), ca: e.indexOf('CA'), km: e.indexOf('KM'), litres: e.indexOf('LITRAGE'),
        // Planning Carburant : 1re colonne SOCIETE = société du chauffeur, 2e = société du client (qui facture).
        // Planning Inter : une seule colonne, les deux sont identiques.
        societe: dernier(function (h) { return h.indexOf('SOCIETE') === 0; }),
        societeChauffeur: col(function (h) { return h.indexOf('SOCIETE') === 0; }),
        chauffeur: col(function (h) { return h.indexOf('CHAUFFEUR') === 0; }),
        client: col(function (h) { return h === 'CLIENT'; }),
        lieu: col(function (h) { return h === 'LIEU DE LIVRAISON' || h === 'LIVRAISON'; }),
        contrat: col(function (h) { return h.indexOf('CONTRAT MANQUANT') === 0; }),
        camion: col(function (h) { return /^(CAMION|VEHICULE|IMMAT|TRACTEUR)/.test(h); }),
        attente: col(function (h) { return h.indexOf('HEURES ATTENTE') === 0; })
      };
      var val = function (l, k) { return k >= 0 ? l[k] : ''; };
      if (!an) {   // onglet sans année (« AOUT ») : année la plus fréquente parmi ses dates
        var compte = {};
        v.forEach(function (l) { if (l[c.date] instanceof Date) compte[l[c.date].getFullYear()] = (compte[l[c.date].getFullYear()] || 0) + 1; });
        an = Object.keys(compte).sort(function (a, b) { return compte[b] - compte[a]; })[0];
      }
      v.slice(ligneEntete + 1).forEach(function (l) {
        var dt = l[c.date] instanceof Date ? l[c.date] : dateDepuisTexte_(l[c.date]);
        if (dt && an && dt.getMonth() + 1 === moisOnglet && dt.getFullYear() !== +an) dt = new Date(+an, dt.getMonth(), dt.getDate());
        var chauffeur = String(val(l, c.chauffeur) || '').trim();
        var ca = nombre_(val(l, c.ca), null);
        if (!dt || (!chauffeur && !ca)) return;
        if (!ca && !nombre_(val(l, c.km), null) && !nombre_(val(l, c.litres), null)) return;   // ligne préparée mais pas encore remplie
        var normSoc = function (x) { var k = cle_(x); return k.indexOf('LPB') >= 0 ? 'LPB' : (k.indexOf('GLEYZES') >= 0 ? 'Gleyzes' : k); };
        res.push({
          date: dt, activite: src[1], societe: normSoc(val(l, c.societe)), societeChauffeur: normSoc(val(l, c.societeChauffeur)),
          chauffeur: chauffeur, client: String(val(l, c.client) || '').trim(), lieu: String(val(l, c.lieu) || '').trim(),
          ca: ca, km: nombre_(val(l, c.km), null), litres: nombre_(val(l, c.litres), null),
          attente: nombre_(val(l, c.attente), null), contratManquant: String(val(l, c.contrat) || '').trim(),
          camionSaisi: /^#/.test(String(val(l, c.camion))) ? '' : plaque_(val(l, c.camion))
        });
      });
    });
  });
  attribuerCamions_(ss, res);
  CACHE_PLANNINGS_ = res;
  return res;
}

// ---------------------------------------------------------------------------
// Colonne CAMION des plannings : remplie automatiquement, modifiable à la main
// ---------------------------------------------------------------------------
var SOURCES_PLANNING = [['ID_PLANNING_INTER', 'Inter'], ['ID_PLANNING_CARBURANT', 'Carburant']];

/** Mois / année d'un onglet mensuel (« SEPTEMBRE 2026 », « AOUT ») ; null pour les onglets CA, TARIFS… */
function moisOnglet_(nomOnglet) {
  var nom = cle_(nomOnglet);
  if (/^CA\b|TARIF|COMPARATIF|RECAP|CAMIONS/.test(nom)) return null;
  var m = Object.keys(MOIS_ONGLETS).filter(function (x) { return nom.indexOf(x) >= 0; })[0];
  if (!m) return null;
  var an = (nom.match(/(20\d\d)/) || [])[1];
  return { mois: MOIS_ONGLETS[m], an: an ? +an : null };
}

/** Ligne d'en-tête (0-based) et colonnes utiles d'un onglet mensuel de planning. */
function structurePlanning_(v) {
  for (var i = 0; i < Math.min(6, v.length); i++) {
    var h = v[i].map(cle_);
    var ch = -1, cam = -1, dt = -1;
    h.forEach(function (x, k) {
      if (ch < 0 && x.indexOf('CHAUFFEUR') === 0) ch = k;
      if (cam < 0 && /^(CAMION|VEHICULE|IMMAT|TRACTEUR)/.test(x)) cam = k;
      if (dt < 0 && (x === 'DATE' || x === 'DATES')) dt = k;
    });
    if (ch >= 0 && h.some(function (x) { return x.indexOf('SOCIETE') === 0; })) return { entete: i, chauffeur: ch, camion: cam, date: Math.max(0, dt) };
  }
  return null;
}

/** Immatriculation connue qui ne diffère que d'un caractère (faute de frappe probable), s'il n'y en a qu'une. */
function procheConnue_(p, connues) {
  var c = Object.keys(connues).filter(function (k) { return k.length === p.length && k.split('').filter(function (ch, i) { return ch !== p[i]; }).length === 1; });
  return c.length === 1 ? c[0] : null;
}

function colonneLettre_(n) { var s = ''; while (n > 0) { var r = (n - 1) % 26; s = String.fromCharCode(65 + r) + s; n = Math.floor((n - 1) / 26); } return s; }

/** Onglet CAMIONS d'un planning : nom du chauffeur (normalisé) -> immatriculation. */
function habituelsPlanning_(ext) {
  var sh = ext.getSheetByName('CAMIONS'), res = {};
  if (!sh || sh.getLastRow() < 2) return res;
  sh.getRange(2, 1, sh.getLastRow() - 1, 2).getValues().forEach(function (l) { if (l[0] && l[1]) res[cle_(l[0])] = plaque_(l[1]); });
  return res;
}

function immatsConnues_(ss) {
  var connues = {};
  var shF = ss.getSheetByName('FLOTTE');
  if (shF) lireTable_(shF).forEach(function (l) { if (l.Camion_ID) connues[plaque_(l.Camion_ID)] = 1; });
  try { echeances_(ss).forEach(function (e) { if (e.domaine === 'Flotte') connues[e.nom] = 1; }); } catch (e) { /* tableau non relié */ }
  return connues;
}

function preparerCamionsManuel() {
  var msg;
  try { msg = preparerCamionsPlannings_(classeur_()); }
  catch (e) { SpreadsheetApp.getUi().alert('Problème : ' + e.message); return; }
  SpreadsheetApp.getUi().alert(msg ? msg.trim() : 'Renseignez d\'abord ID_PLANNING_INTER et ID_PLANNING_CARBURANT dans PARAMETRES.');
}

/**
 * Pour chaque planning (Inter, Carburant) :
 *  1. crée / complète l'onglet CAMIONS (chauffeur -> camion habituel), proposé d'après le camion le plus utilisé ;
 *  2. dans les onglets du mois en cours et des mois suivants, ajoute la colonne CAMION (après CHAUFFEUR) si besoin
 *     et écrit le camion habituel dans chaque case vide d'une ligne qui a un chauffeur (simple valeur, pas de
 *     formule : une formule écrite par script échoue en #ERROR! selon la langue du classeur) ;
 *  3. installe un déclencheur « à la modification » sur le planning : dès qu'un chauffeur est tapé, son camion
 *     habituel est écrit tout de suite (surModifPlanning).
 * Une case où un camion a été tapé à la main n'est jamais modifiée.
 */
function preparerCamionsPlannings_(ss) {
  var debutExec = Date.now(), tropLong = function () { return Date.now() - debutExec > 270000; }, interrompu = false;
  var auj = minuit_(new Date()), courant = auj.getFullYear() * 12 + auj.getMonth() + 1;
  var connues = immatsConnues_(ss), msgs = [];
  console.log('Immatriculations connues : ' + Object.keys(connues).length);
  SOURCES_PLANNING.forEach(function (src) {
    if (tropLong()) { interrompu = true; return; }
    var ext = classeurParam_(ss, src[0]);
    if (!ext) return;
    console.log('Planning ' + src[1] + ' ouvert');
    var onglets = ext.getSheets().map(function (sh) { var m = moisOnglet_(sh.getName()); return m ? { sh: sh, mois: m.mois, an: m.an } : null; })
      .filter(Boolean).map(function (o) {
        if (!o.an) o.an = auj.getFullYear();
        o.rang = o.an * 12 + o.mois;
        return o;
      });

    // 1. Onglet CAMIONS, proposé d'après les camions déjà saisis (plus récents d'abord)
    var shC = ext.getSheetByName('CAMIONS');
    if (!shC) {
      shC = ext.insertSheet('CAMIONS');
      shC.getRange(1, 1, 1, 3).setValues([['CHAUFFEUR', 'CAMION HABITUEL', 'REMARQUE']]).setFontWeight('bold');
      shC.setFrozenRows(1);
    }
    var habituels = habituelsPlanning_(ext);
    var recents = {}, vus = {};
    onglets.filter(function (o) { return o.rang >= courant - 3 && o.rang <= courant; }).forEach(function (o) {
      if (o.sh.getLastRow() < 2) return;
      var v = o.sh.getRange(1, 1, o.sh.getLastRow(), o.sh.getLastColumn()).getValues(), st = structurePlanning_(v);
      if (!st) return;
      v.slice(st.entete + 1).forEach(function (l) {
        var ch = cle_(l[st.chauffeur]);
        if (!ch) return;
        vus[ch] = 1;
        var cam = st.camion >= 0 ? plaque_(l[st.camion]) : '';
        if (cam && !connues[cam]) cam = procheConnue_(cam, connues) || '';   // GB042ZC -> GD042ZC
        var dt = l[st.date] instanceof Date ? l[st.date] : dateDepuisTexte_(l[st.date]);
        if (cam && dt) (recents[ch] = recents[ch] || []).push({ dt: dt, cam: cam });
      });
    });
    // Proposition : camion le plus utilisé sur les 10 dernières lignes du chauffeur (suit un changement récent)
    var proposition = function (ch) {
      var l = (recents[ch] || []).sort(function (a, b) { return a.dt - b.dt; }).slice(-10), f = {};
      l.forEach(function (x) { f[x.cam] = (f[x.cam] || 0) + 1; });
      return Object.keys(f).sort(function (a, b) { return f[b] - f[a]; })[0] || '';
    };
    var lignesC = shC.getLastRow() > 1 ? shC.getRange(2, 1, shC.getLastRow() - 1, 3).getValues() : [];
    var dejaListes = lignesC.map(function (l) { return cle_(l[0]); });
    // Chauffeurs déjà listés mais sans camion : on complète la case vide avec la proposition
    var complete = false;
    lignesC.forEach(function (l) {
      if (l[0] && !String(l[1]).trim()) { var pr = proposition(cle_(l[0])); if (pr) { l[1] = pr; l[2] = 'proposé : camion le plus utilisé récemment'; complete = true; } }
    });
    if (complete) shC.getRange(2, 1, lignesC.length, 3).setValues(lignesC);
    var ajouts = Object.keys(vus).filter(function (ch) { return !habituels[ch] && dejaListes.indexOf(ch) < 0; }).sort().map(function (ch) {
      var best = proposition(ch);
      return [ch, best, best ? 'proposé : camion le plus utilisé récemment' : 'à compléter'];
    });
    if (ajouts.length) shC.getRange(shC.getLastRow() + 1, 1, ajouts.length, 3).setValues(ajouts);

    // 2. Colonne CAMION remplie avec le camion habituel (valeurs), mois en cours et suivants (et mois précédent si vide)
    var nbFormules = 0, nbFiges = 0, nbColonnes = 0;
    var hab = habituelsPlanning_(ext);
    onglets.filter(function (o) { return o.rang >= courant - 1; }).forEach(function (o) {
      if (tropLong()) { interrompu = true; return; }
      var sh = o.sh;
      if (sh.getLastRow() < 1) return;
      var v = sh.getRange(1, 1, sh.getLastRow(), sh.getLastColumn()).getValues(), st = structurePlanning_(v);
      if (!st) return;
      if (st.camion < 0) {
        if (o.rang < courant) return;   // on n'ajoute pas de colonne dans un mois terminé
        sh.insertColumnAfter(st.chauffeur + 1);
        sh.getRange(st.entete + 1, st.chauffeur + 2).setValue('CAMION').setFontWeight('bold');
        st.camion = st.chauffeur + 1;
        if (st.date > st.chauffeur) st.date++;
        nbColonnes++;
        v = sh.getRange(1, 1, sh.getLastRow(), sh.getLastColumn()).getValues();
      }
      var n = sh.getLastRow() - st.entete - 1;
      if (n <= 0) return;
      var rng = sh.getRange(st.entete + 2, st.camion + 1, n, 1);
      var formules = rng.getFormulas(), valeurs = rng.getValues();
      // Chaque case : nouvelle valeur, ou null = ne pas toucher
      var actions = valeurs.map(function (x, i) {
        var ligne = v[st.entete + 1 + i] || [], val = x[0] === undefined ? '' : x[0];
        var chauffeur = String(ligne[st.chauffeur] || '').trim(), attendu = chauffeur ? (hab[cle_(chauffeur)] || '') : '';
        if (formules[i][0]) {   // ancienne formule (souvent en #ERROR!) : remplacée par une valeur
          nbFiges++;
          return { x: !/^#/.test(String(val)) && String(val).trim() ? val : attendu };
        }
        if (String(val).trim() !== '' || !attendu) return null;   // saisi à la main, ou pas de chauffeur
        nbFormules++;
        return { x: attendu };
      });
      var i = 0;
      while (i < actions.length) {
        if (!actions[i]) { i++; continue; }
        var j = i;
        while (j < actions.length && actions[j]) j++;
        sh.getRange(st.entete + 2 + i, st.camion + 1, j - i, 1).setValues(actions.slice(i, j).map(function (a) { return [a.x]; }));
        i = j;
      }
      console.log('Planning ' + src[1] + ' : onglet ' + sh.getName() + ' traité');
    });
    // 3. Remplissage instantané à la saisie
    var declencheur = installerDeclencheurPlanning_(ext);
    msgs.push('Planning ' + src[1] + ' : ' + (ajouts.length ? ajouts.length + ' chauffeur(s) ajouté(s) dans l\'onglet CAMIONS, ' : '') +
      (nbColonnes ? nbColonnes + ' colonne(s) CAMION ajoutée(s), ' : '') + nbFormules + ' case(s) CAMION remplie(s)' +
      (nbFiges ? ', ' + nbFiges + ' ancienne(s) formule(s) remplacée(s)' : '') + '. ' + declencheur);
  });
  if (interrompu) msgs.push('Pas tout à fait fini (limite de temps de Google) : relancez « Remplir la colonne CAMION des plannings », il reprendra où il s\'est arrêté.');
  return msgs.length ? '\n\n' + msgs.join('\n') : '';
}

/** Installe (une seule fois) le déclencheur « à la modification » du planning. */
function installerDeclencheurPlanning_(ext) {
  var id = ext.getId();
  var existe = ScriptApp.getProjectTriggers().some(function (t) {
    return t.getHandlerFunction() === 'surModifPlanning' && t.getTriggerSourceId() === id;
  });
  if (existe) return 'Remplissage automatique déjà actif.';
  ScriptApp.newTrigger('surModifPlanning').forSpreadsheet(id).onEdit().create();
  return 'Remplissage automatique activé.';
}

/**
 * Déclencheur : quand un ou plusieurs chauffeurs sont tapés (ou collés) dans un onglet mensuel d'un planning,
 * écrit le camion habituel dans la colonne CAMION si la case est vide, ou si elle contenait le camion habituel
 * de l'ancien chauffeur (changement de chauffeur). Un camion tapé à la main n'est jamais remplacé.
 */
function surModifPlanning(e) {
  try {
    var r = e.range, sh = r.getSheet();
    if (!moisOnglet_(sh.getName())) return;
    var haut = sh.getRange(1, 1, Math.min(6, sh.getLastRow()), sh.getLastColumn()).getValues(), st = structurePlanning_(haut);
    if (!st || st.camion < 0) return;
    var colCh = st.chauffeur + 1;
    if (r.getColumn() > colCh || r.getLastColumn() < colCh) return;   // la modification ne touche pas la colonne CHAUFFEUR
    var premiere = Math.max(r.getRow(), st.entete + 2), derniere = r.getLastRow();
    if (derniere < premiere) return;
    var hab = habituelsPlanning_(e.source);
    var n = derniere - premiere + 1;
    var chauffeurs = sh.getRange(premiere, colCh, n, 1).getValues(), camions = sh.getRange(premiere, st.camion + 1, n, 1).getValues();
    var ancien = n === 1 && e.oldValue !== undefined ? (hab[cle_(e.oldValue)] || '') : null;
    var change = false;
    var sortie = camions.map(function (c, i) {
      var actuel = String(c[0] || '').trim(), nouveau = hab[cle_(chauffeurs[i][0])] || '';
      if (actuel && !/^#/.test(actuel) && !(ancien !== null && actuel === ancien)) return [c[0]];
      if (actuel === nouveau) return [c[0]];
      change = true;
      return [nouveau];
    });
    if (change) sh.getRange(premiere, st.camion + 1, n, 1).setValues(sortie);
  } catch (err) {
    console.error('surModifPlanning : ' + err);
  }
}

/**
 * Camion de chaque ligne de planning : celui de la colonne CAMION (formule ou saisie), sinon le camion habituel du
 * chauffeur (onglet CAMIONS du planning, ou à défaut l'ancien onglet CAMIONS_HABITUELS du tableau de bord).
 * camionSource = 'habituel' si c'est le camion habituel, 'planning' si un autre camion a été tapé (changement).
 */
function attribuerCamions_(ss, lignes) {
  var habituels = {};
  SOURCES_PLANNING.forEach(function (src) {
    try { var ext = classeurParam_(ss, src[1] === 'Inter' ? 'ID_PLANNING_INTER' : 'ID_PLANNING_CARBURANT'); if (ext) habituels[src[1]] = habituelsPlanning_(ext); }
    catch (e) { habituels[src[1]] = {}; }
  });
  var sh = ss.getSheetByName('CAMIONS_HABITUELS');
  var anciens = (sh ? lireTable_(sh) : []).filter(function (r) { return r.Chauffeur && r.Camion_Habituel; });
  lignes.forEach(function (l) {
    var ch = cle_(l.chauffeur), h = habituels[l.activite] || {};
    var hab = h[ch] || (Object.keys(h).filter(function (k) { return memePersonne_(k, ch); }).map(function (k) { return h[k]; })[0]) ||
      (anciens.filter(function (r) { return memePersonne_(r.Chauffeur, ch); }).map(function (r) { return plaque_(r.Camion_Habituel); })[0]) || '';
    if (l.camionSaisi) { l.camion = l.camionSaisi; l.camionSource = l.camionSaisi === hab ? 'habituel' : 'planning'; }
    else { l.camion = hab; l.camionSource = hab ? 'habituel' : ''; }
  });
}

/**
 * Contrôles sur les camions des plannings (mois en cours, et mois précédent en début de mois) :
 *  - immatriculation inconnue (ni dans FLOTTE ni dans le tableau des échéances), avec la correction probable ;
 *  - camion utilisé alors que son contrôle technique était expiré.
 */
function alertesCamionsPlannings_(ss, lignes, jour0, depuis) {
  var res = [];
  var connues = immatsConnues_(ss), ct = {};
  try {
    echeances_(ss).forEach(function (e) { if (e.domaine === 'Flotte' && e.date && /CONTROLE TECHNIQUE/.test(cle_(e.document))) ct[e.nom] = e.date; });
  } catch (e) { /* tableau non relié */ }
  var inconnues = {}, roule = {};
  lignes.forEach(function (l) {
    if (!l.camion || l.date < depuis || l.date > jour0) return;
    if (Object.keys(connues).length && !connues[l.camion]) {
      var x = inconnues[l.camion] = inconnues[l.camion] || { n: 0, act: {}, ch: {} };
      x.n++; x.act[l.activite] = 1; x.ch[cle_(l.chauffeur)] = 1;
    }
    if (ct[l.camion] && l.date > ct[l.camion]) {
      var y = roule[l.camion] = roule[l.camion] || { n: 0, ch: {}, der: l.date };
      y.n++; y.ch[cle_(l.chauffeur)] = 1; if (l.date > y.der) y.der = l.date;
    }
  });
  Object.keys(inconnues).forEach(function (p) {
    var x = inconnues[p];
    var proche = procheConnue_(p, connues);
    res.push({ niveau: 'urgent', categorie: 'donnees', domaine: 'Plannings', objet: p, sujet: 'Immatriculation inconnue',
      message: 'Immatriculation inconnue « ' + p + ' » dans le planning ' + Object.keys(x.act).join(' et ') + ' (' + x.n + ' ligne(s), ' +
        Object.keys(x.ch).join(', ') + ')' + (proche ? ' : faute de frappe pour ' + proche + ' ?' : ''),
      cle: 'IMMAT_INCONNUE|' + p });
  });
  Object.keys(roule).forEach(function (p) {
    var y = roule[p];
    res.push({ niveau: 'depasse', categorie: 'ctroule', domaine: 'Flotte', objet: p, sujet: 'Roule avec CT expiré',
      message: p + ' a roulé ' + y.n + ' fois avec le contrôle technique expiré (depuis le ' + Utilities.formatDate(ct[p], FUSEAU, 'dd/MM/yyyy') +
        ', dernière fois le ' + Utilities.formatDate(y.der, FUSEAU, 'dd/MM/yyyy') + ', ' + Object.keys(y.ch).join(', ') + ')',
      cle: 'CT_ROULE|' + p + '|' + Utilities.formatDate(y.der, FUSEAU, 'yyyy-MM-dd') });
  });
  return res;
}

/**
 * Début de la période surveillée dans les plannings : le 1er du mois en cours, ou le 1er du mois précédent
 * pendant les premiers jours du mois (le 2 octobre on regarde encore septembre, plus août).
 */
function debutPeriodePlannings_(jour0, joursMoisPrecedent) {
  return jour0.getDate() <= joursMoisPrecedent
    ? new Date(jour0.getFullYear(), jour0.getMonth() - 1, 1)
    : new Date(jour0.getFullYear(), jour0.getMonth(), 1);
}

/** Contrats manquants (Inter) et livraisons sans prix (Carburant) des 60 derniers jours. */
function alertesPlannings_(lignes, jour0, depuis) {
  var res = [], sansPrix = {};
  lignes.forEach(function (l) {
    if (l.date < depuis || l.date > jour0) return;
    var dateTxt = Utilities.formatDate(l.date, FUSEAU, 'dd/MM/yyyy');
    if (l.activite === 'Inter' && l.contratManquant) {
      res.push({ niveau: 'a_prevoir', categorie: 'contrats', domaine: 'Inter', objet: l.chauffeur, sujet: 'Contrat manquant',
        message: 'Contrat manquant (Inter) : ' + dateTxt + ', ' + l.chauffeur + ', ' + l.lieu + ' — ' + l.contratManquant,
        cle: 'CONTRAT|' + Utilities.formatDate(l.date, FUSEAU, 'yyyy-MM-dd') + '|' + l.chauffeur + '|' + l.contratManquant });
    }
    if (l.activite === 'Carburant' && l.litres > 0 && !l.ca) {
      var k = l.client + ' ' + l.lieu;
      (sansPrix[k] = sansPrix[k] || []).push(dateTxt);
    }
  });
  var sansCamion = {};
  lignes.forEach(function (l) { if (l.date >= depuis && l.date <= jour0 && !l.camion && l.chauffeur) sansCamion[cle_(l.chauffeur)] = 1; });
  if (Object.keys(sansCamion).length) {
    res.push({ niveau: 'a_prevoir', categorie: 'donnees', domaine: 'Plannings', objet: '', sujet: 'Camion habituel manquant',
      message: Object.keys(sansCamion).length + ' chauffeur(s) du planning sans camion (' + Object.keys(sansCamion).sort().join(', ') + ') : à compléter dans l\'onglet CAMIONS du planning',
      cle: 'SANSCAMION|' + Object.keys(sansCamion).sort().join(',') });
  }
  Object.keys(sansPrix).forEach(function (k) {
    res.push({ niveau: 'a_prevoir', categorie: 'donnees', domaine: 'Carburant', objet: k, sujet: 'Prix manquant',
      message: 'Livraison(s) carburant sans prix pour ' + k + ' (' + sansPrix[k].join(', ') + ') : prix à saisir dans l\'onglet TARIFS',
      cle: 'SANSPRIX|' + k + '|' + sansPrix[k].length });
  });
  return res;
}

// ---------------------------------------------------------------------------
// Litrages véhicules : un onglet par camion, des blocs mensuels « Km début / Km fin / Litrage » par semaine
// ---------------------------------------------------------------------------
/**
 * Une ligne par semaine et par véhicule : { immat, semaine, fin, approx, km, litres, conso, valide, raison }.
 * Les semaines incohérentes (km de fin < km de début, plus de 5 000 km, conso hors 10-70 L/100…) sont gardées
 * mais marquées valide = false, pour ne jamais déclencher d'alerte sur une faute de frappe.
 */
function litrages_(ss) {
  if (CACHE_LITRAGES_) return CACHE_LITRAGES_;
  var ext = classeurParam_(ss, 'ID_LITRAGES'), res = [];
  if (!ext) return (CACHE_LITRAGES_ = res);
  ext.getSheets().forEach(function (sh) {
    var immat = plaque_(sh.getName());
    if (!/^[A-Z]{2}\d{3}[A-Z]{2}$/.test(immat) || sh.getLastRow() < 3) return;
    var v = sh.getRange(1, 1, sh.getLastRow(), sh.getLastColumn()).getValues();
    for (var r = 0; r < v.length; r++) {
      for (var c = 0; c < v[r].length; c++) {
        if (cle_(v[r][c]) !== 'KM DEBUT') continue;
        // Mois du bloc : la case au-dessus de « Km début » (une date, ou un texte « 01/03/2026-31/03/2026 »)
        var haut = r > 0 ? v[r - 1][c] : '', mois = null;
        if (haut instanceof Date) mois = new Date(haut.getFullYear(), haut.getMonth(), 1);
        else { var m = String(haut).match(/(\d{1,2})\/(\d{1,2})\/(\d{4})/); if (m) mois = new Date(+m[3], +m[2] - 1, 1); }
        if (!mois) continue;
        var finMois = new Date(mois.getFullYear(), mois.getMonth() + 1, 0);
        for (var i = 1; i <= 6 && r + i < v.length; i++) {
          var l = v[r + i], libelle = c > 0 ? String(l[c - 1] || '').trim() : '';
          if (/TOTAL/.test(cle_(libelle)) || cle_(l[c]) === 'KM DEBUT') break;
          var kmD = nombre_(l[c], null), kmF = nombre_(l[c + 1], null), litres = nombre_(l[c + 2], null);
          if (kmF === null && !litres) continue;   // ligne vide ou sous-total
          // Fin de semaine : « du 12 au 17/01 » ; sinon estimée d'après le rang de la ligne dans le mois
          var fin = null, approx = false, mm = libelle.match(/au\s*(\d{1,2})\s*\/\s*(\d{1,2})/i);
          if (mm) fin = new Date(mois.getFullYear(), +mm[2] - 1, +mm[1]);
          if (!fin || isNaN(fin.getTime())) { fin = new Date(Math.min(finMois.getTime(), new Date(mois.getFullYear(), mois.getMonth(), 7 * i).getTime())); approx = true; }
          var km = kmD !== null && kmF !== null ? kmF - kmD : null;
          var conso = km > 0 && litres > 0 ? litres / km * 100 : null;
          var raison = km === null ? 'km manquant' : km <= 0 ? 'km de fin inférieur au km de début' : km > 5000 ? 'plus de 5 000 km (faute de frappe ?)'
            : !(litres > 0) ? 'litres manquants' : (conso < 10 || conso > 70) ? 'consommation impossible (' + conso.toFixed(0) + ' L/100)' : '';
          res.push({ immat: immat, semaine: libelle || ('semaine ' + i + ' de ' + Utilities.formatDate(mois, FUSEAU, 'MM/yyyy')),
            fin: fin, approx: approx, km: km, litres: litres, conso: raison ? null : conso, valide: !raison, raison: raison });
        }
      }
    }
  });
  res.sort(function (a, b) { return a.immat < b.immat ? -1 : a.immat > b.immat ? 1 : a.fin - b.fin; });
  return (CACHE_LITRAGES_ = res);
}

/**
 * Consommation des nb dernières semaines saisies de chaque camion (regroupées : un plein qui tombe à cheval sur
 * deux semaines fausse une semaine isolée), comparée au seuil et à la moyenne des 8 semaines d'avant.
 */
function alertesConsoHebdo_(semaines, jour0, seuil, haussePct, nb) {
  var parImmat = {}, res = [];
  semaines.forEach(function (s) { if (s.valide) (parImmat[s.immat] = parImmat[s.immat] || []).push(s); });
  var ratio = function (l) { return l.reduce(function (t, x) { return t + x.litres; }, 0) / l.reduce(function (t, x) { return t + x.km; }, 0) * 100; };
  Object.keys(parImmat).forEach(function (immat) {
    var l = parImmat[immat], der = l[l.length - 1];
    if (l.length < nb || (jour0 - der.fin) / 86400000 > 45) return;   // pas de saisie récente : rien à signaler
    var recent = l.slice(l.length - nb), prec = l.slice(Math.max(0, l.length - nb - 8), l.length - nb);
    var conso = ratio(recent), moy = prec.length >= 3 ? ratio(prec) : null;
    var hausse = moy ? (conso - moy) / moy * 100 : null;
    if (conso <= seuil && !(hausse !== null && hausse > haussePct)) return;
    var periode = recent.map(function (x) { return x.semaine; }).join(' + ');
    res.push({ niveau: 'urgent', categorie: 'conso', domaine: 'Carburant', objet: immat, sujet: 'Consommation récente',
      message: immat + ' consomme ' + fr1_(conso) + ' L/100 km sur ' + (nb === 1 ? 'la semaine ' : 'les ' + nb + ' dernières semaines ') + '(' + periode + ')' +
        (moy ? ' contre ' + fr1_(moy) + ' habituellement (' + (hausse >= 0 ? '+' : '') + fr1_(hausse) + ' %)' : '') +
        (conso > seuil ? ', au-dessus du seuil de ' + seuil : ''),
      cle: 'CONSO_SEM|' + immat + '|' + Utilities.formatDate(der.fin, FUSEAU, 'yyyy-MM-dd') });
  });
  return res;
}

/** Deux noms désignent la même personne si les mots du plus court sont tous dans le plus long (ordre et accents ignorés). */
function memePersonne_(a, b) {
  var m = function (s) { return cle_(s).split(/[\s-]+/).filter(String); };
  var x = m(a), y = m(b);
  if (!x.length || !y.length) return false;
  var court = x.length <= y.length ? x : y, long = x.length <= y.length ? y : x;
  return court.every(function (w) { return long.indexOf(w) >= 0; });
}

/** Ajoute à FLOTTE et SALARIES les véhicules et chauffeurs du tableau « Échéances flotte » qui n'y sont pas encore. */
function synchroniserReferentiels_(ss) {
  if (!classeurEcheances_(ss)) return '';
  var shF = ss.getSheetByName('FLOTTE'), shS = ss.getSheetByName('SALARIES');
  var flotte = lireTable_(shF).map(function (l) { return plaque_(l.Camion_ID); });
  var salaries = lireTable_(shS).map(function (l) { return ((l['Prénom'] || '') + ' ' + (l.Nom || '')).trim(); });
  var eF = shF.getRange(1, 1, 1, shF.getLastColumn()).getValues()[0], eS = shS.getRange(1, 1, 1, shS.getLastColumn()).getValues()[0];
  var nbV = 0, nbC = 0, vus = {}, ajoutsF = [], ajoutsS = [];
  echeances_(ss).forEach(function (e) {
    if (vus[e.nom]) return;
    vus[e.nom] = 1;
    var k = cle_(e.categorie);
    if (e.domaine === 'Flotte' && k !== 'TELECHARGEMENT' && flotte.indexOf(e.nom) < 0) {
      var lv = { Camion_ID: e.nom, Type: e.categorie, Actif: 'O', Statut: 'Disponible' };
      ajoutsF.push(eF.map(function (h) { return lv[h] || ''; }));
      flotte.push(e.nom); nbV++;
    } else if (k === 'CHAUFFEUR' && !salaries.some(function (n) { return memePersonne_(n, e.nom); })) {
      var ls = { Nom: e.nom, Poste: 'Chauffeur', Actif: 'O' };
      ajoutsS.push(eS.map(function (h) { return ls[h] || ''; }));
      salaries.push(e.nom); nbC++;
    }
  });
  // Une seule écriture par onglet (bien plus rapide que des ajouts ligne par ligne)
  if (ajoutsF.length) shF.getRange(shF.getLastRow() + 1, 1, ajoutsF.length, eF.length).setValues(ajoutsF);
  if (ajoutsS.length) shS.getRange(shS.getLastRow() + 1, 1, ajoutsS.length, eS.length).setValues(ajoutsS);
  return '\n\nTableau « Échéances flotte » relié : ' + nbV + ' véhicule(s) et ' + nbC + ' chauffeur(s) ajouté(s) dans FLOTTE et SALARIES.';
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

var CACHE_PARAMS_ = null;
function lireParametres_(ss) {
  if (CACHE_PARAMS_) return CACHE_PARAMS_;
  var sh = (ss || classeur_()).getSheetByName('PARAMETRES');
  var p = {};
  lireTable_(sh).forEach(function (l) { p[String(l['Paramètre']).trim()] = l.Valeur; });
  return (CACHE_PARAMS_ = p);
}

function majParametre_(nom, valeur) {
  var sh = classeur_().getSheetByName('PARAMETRES');
  var v = sh.getRange(1, 1, sh.getLastRow(), 2).getValues();
  for (var i = 1; i < v.length; i++) {
    if (v[i][0] === nom) { sh.getRange(i + 1, 2).setValue(valeur); CACHE_PARAMS_ = null; return; }
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

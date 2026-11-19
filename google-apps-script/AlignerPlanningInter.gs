/**
 * À coller dans le PLANNING INTER (Extensions > Apps Script), puis lancer « alignerNovembreDecembre ».
 * Rend les onglets NOVEMBRE et DECEMBRE identiques à OCTOBRE : mêmes colonnes, mêmes formules, même mise en forme
 * (couleurs, retours à la ligne, largeurs). Les saisies sont reprises colonne par colonne (un KM tapé est gardé).
 * Avant de modifier un onglet, une sauvegarde cachée « SAUV <MOIS> » est créée.
 */
var MODELE = 'OCTOBRE 2026';
var CIBLES = ['NOVEMBRE 2026', 'DECEMBRE 2026'];

function alignerNovembreDecembre() {
  var ss = SpreadsheetApp.getActiveSpreadsheet(), modele = ss.getSheetByName(MODELE);
  if (!modele) throw new Error('Onglet « ' + MODELE + ' » introuvable');
  var lc = modele.getLastColumn(), hM = modele.getRange(1, 1, 1, lc).getValues()[0];
  while (lc > 1 && !String(hM[lc - 1]).trim()) lc--;
  hM = hM.slice(0, lc);
  var col = function (re) { for (var i = 0; i < hM.length; i++) if (re.test(cle(hM[i]))) return i; return -1; };
  var cLiv = col(/^LIVRAISON/), cKm = col(/^KM$/);
  // formules du modèle : ligne 2 pour les colonnes calculées ; pour KM, la première ligne qui a une formule
  var fM = modele.getRange(2, 1, Math.min(Math.max(1, modele.getLastRow() - 1), 400), lc).getFormulas();
  var colsFormule = [], ligneKm = -1, finFormules = 2;
  for (var c = 0; c < lc; c++) {
    if (c === cKm) { for (var r = 0; r < fM.length; r++) if (fM[r][c]) { ligneKm = r + 2; break; } continue; }
    if (fM[0][c]) colsFormule.push(c);
  }
  for (r = fM.length - 1; r >= 0; r--) if (fM[r].some(String)) { finFormules = r + 2; break; }
  var formuleKm = ligneKm > 0 ? fM[ligneKm - 2][cKm] : '';
  var caModele = ss.getSheetByName('CA ' + MODELE), compte = [];

  CIBLES.forEach(function (nom, k) {
    var T = ss.getSheetByName(nom); if (!T) { compte.push(nom + ' : introuvable'); return; }
    // sauvegarde cachée de l'onglet tel qu'il est
    var sauv = T.copyTo(ss).setName('SAUV ' + nom + ' ' + Utilities.formatDate(new Date(), 'Europe/Paris', 'dd-MM HH-mm')); sauv.hideSheet();
    var lr = Math.max(1, T.getLastRow()), lcT = Math.max(1, T.getLastColumn());
    var anc = T.getRange(1, 1, lr, lcT).getValues(), ancF = T.getRange(1, 1, lr, lcT).getFormulas(), hA = anc[0];
    // colonnes : même titre (« N CONTRAT » = « N° CONTRAT »)
    var idx = {}; hA.forEach(function (h, i) { var x = cleCol(h); if (x && idx[x] === undefined) idx[x] = i; });
    var pris = {}, map = hM.map(function (h) { var i = idx[cleCol(h)]; if (i !== undefined) pris[i] = 1; return i === undefined ? -1 : i; });
    var extras = []; hA.forEach(function (h, i) { if (!pris[i] && String(h || '').trim() && !/^COLUMN \d+$/.test(cle(h))) extras.push(i); });
    var derniere = 1; for (r = anc.length - 1; r >= 1; r--) if (String(anc[r][0]).trim() || String(anc[r][2]).trim()) { derniere = r + 1; break; }
    var nData = derniere - 1, nF = Math.max(finFormules - 1, nData), largeur = lc + extras.length;
    if (T.getMaxColumns() < largeur) T.insertColumnsAfter(T.getMaxColumns(), largeur - T.getMaxColumns());
    if (T.getMaxRows() < nF + 1) T.insertRowsAfter(T.getMaxRows(), nF + 1 - T.getMaxRows());
    // 1. tout effacer, puis mise en forme du modèle
    T.getRange(1, 1, T.getMaxRows(), T.getMaxColumns()).clear(); T.setConditionalFormatRules([]);
    var zone = modele.getRange(1, 1, Math.min(modele.getMaxRows(), T.getMaxRows()), lc), dest = T.getRange(1, 1);
    zone.copyTo(dest, SpreadsheetApp.CopyPasteType.PASTE_FORMAT, false);
    zone.copyTo(dest, SpreadsheetApp.CopyPasteType.PASTE_DATA_VALIDATION, false);
    zone.copyTo(dest, SpreadsheetApp.CopyPasteType.PASTE_CONDITIONAL_FORMATTING, false);
    zone.copyTo(dest, SpreadsheetApp.CopyPasteType.PASTE_COLUMN_WIDTHS, false);
    T.setRowHeight(1, modele.getRowHeight(1));
    T.setFrozenRows(modele.getFrozenRows()); T.setFrozenColumns(modele.getFrozenColumns());
    // 2. titres
    T.getRange(1, 1, 1, lc).setValues([hM]);
    if (extras.length) T.getRange(1, lc + 1, 1, extras.length).setValues([extras.map(function (i) { return hA[i]; })]);
    // 3. formules du modèle sur toutes les lignes
    colsFormule.forEach(function (c) { modele.getRange(2, c + 1).copyTo(T.getRange(2, c + 1, nF, 1)); });
    if (ligneKm > 0) modele.getRange(ligneKm, cKm + 1).copyTo(T.getRange(2, cKm + 1, nF, 1));
    // 4. saisies reprises (valeurs)
    if (nData > 0) {
      map.forEach(function (i, c) {
        if (i < 0 || colsFormule.indexOf(c) >= 0) return;
        if (c === cKm && ligneKm > 0) {
          for (var r = 1; r <= nData; r++) { var v = anc[r][i]; if (!ancF[r][i] && v !== '' && v !== null && !isNaN(Number(v))) T.getRange(r + 1, c + 1).setValue(Number(v)); }
          return;
        }
        T.getRange(2, c + 1, nData, 1).setValues(anc.slice(1, derniere).map(function (l) { return [l[i]]; }));
      });
      extras.forEach(function (i, k2) { T.getRange(2, lc + k2 + 1, nData, 1).setValues(anc.slice(1, derniere).map(function (l) { return [l[i]]; })); });
    }
    // 5. la formule KM va chercher le mois précédent de CET onglet (octobre pour novembre, novembre pour décembre)
    var prec = k === 0 ? MODELE : CIBLES[k - 1], n = 0, re = /'[^']+'!\$[A-Z]+\$\d+:\$[A-Z]+\$\d+/g, m;
    while (formuleKm && (m = re.exec(formuleKm))) {
      var lettre = lettreCol((n === 0 ? cLiv : cKm) + 1); n++;
      T.createTextFinder(m[0]).matchFormulaText(true).replaceAllWith("'" + prec + "'!$" + lettre + '$2:$' + lettre + '$250');
    }
    // 6. onglet « CA <MOIS> » refait comme « CA OCTOBRE »
    var ca = ss.getSheetByName('CA ' + nom);
    if (caModele && ca) {
      ca.getRange(1, 1, ca.getMaxRows(), ca.getMaxColumns()).clear(); ca.setConditionalFormatRules([]);
      var zc = caModele.getRange(1, 1, Math.min(caModele.getMaxRows(), ca.getMaxRows()), Math.min(caModele.getMaxColumns(), ca.getMaxColumns()));
      zc.copyTo(ca.getRange(1, 1)); zc.copyTo(ca.getRange(1, 1), SpreadsheetApp.CopyPasteType.PASTE_COLUMN_WIDTHS, false);
      ca.createTextFinder("'" + MODELE + "'!").matchFormulaText(true).replaceAllWith("'" + nom + "'!");
      var v = ca.getDataRange().getValues(), f = ca.getDataRange().getFormulas();
      v.forEach(function (l, r) { l.forEach(function (x, c) { if (!f[r][c] && typeof x === 'string' && x.indexOf(MODELE) >= 0) ca.getRange(r + 1, c + 1).setValue(x.split(MODELE).join(nom)); }); });
    }
    compte.push(nom + ' : ' + nData + ' lignes reprises' + (extras.length ? ' (colonnes en plus gardées à droite : ' + extras.map(function (i) { return hA[i]; }).join(', ') + ')' : ''));
  });
  SpreadsheetApp.flush();
  try { SpreadsheetApp.getUi().alert('Terminé :\n' + compte.join('\n')); } catch (e) { Logger.log(compte.join('\n')); }
}

function cle(s) { return String(s || '').normalize('NFD').replace(/[̀-ͯ]/g, '').toUpperCase().replace(/\s+/g, ' ').trim(); }
function cleCol(s) { return cle(s).replace(/[^A-Z0-9]/g, ''); }
function lettreCol(n) { var s = ''; while (n > 0) { var r = (n - 1) % 26; s = String.fromCharCode(65 + r) + s; n = Math.floor((n - 1) / 26); } return s; }

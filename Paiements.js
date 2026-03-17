function getSemestreEnCours() {
  const today = new Date();
  const annee = today.getFullYear();
  const mois  = today.getMonth();
  return mois < 6
    ? { numero: 1, debut: new Date(annee, 0, 1), fin: new Date(annee, 5, 30),  label: "S1 " + annee }
    : { numero: 2, debut: new Date(annee, 6, 1), fin: new Date(annee, 11, 31), label: "S2 " + annee };
}

function calculerStatut(attendu, paye) {
  if (attendu === 0)   return "";
  if (paye >= attendu) return "✅ Payé";
  if (paye === 0)      return "⚠️ Non payé";
  return "⚠️ Montant insuffisant";
}

function calculerStatutCotisation(montantPaye) {
  if (montantPaye === 0)                               return "⚠️ Non payé";
  if (montantPaye < CONFIG.MONTANT_MINIMUM_COTISATION) return "⚠️ Montant insuffisant";
  return "✅ Payé";
}

// ============================================================
// MISE À JOUR SUIVI PARRAINAGES
// ============================================================

function mergeParrainages() {
  const ss       = SpreadsheetApp.getActiveSpreadsheet();
  const semestre = getSemestreEnCours();
  const props    = PropertiesService.getScriptProperties();

  const parrainsMap = {};
  ss.getSheetByName(CONFIG.SHEET_PARRAIN).getDataRange().getValues().slice(1).forEach(row => {
    const id = row[CONFIG.COLS.PARRAIN.SPONSOR_ID - 1];
    if (id) parrainsMap[id] = {
      nom:   row[CONFIG.COLS.PARRAIN.SPONSOR_NAME - 1],
      email: row[CONFIG.COLS.PARRAIN.ADRESSE_MAIL - 1]
    };
  });

  const parrainsOngoing = {};
  ss.getSheetByName(CONFIG.SHEET_PARRAINAGE).getDataRange().getValues().slice(1).forEach(row => {
    const id     = row[CONFIG.COLS.PARRAINAGE.SPONSOR_ID - 1];
    const statut = row[CONFIG.COLS.PARRAINAGE.STATUS - 1];
    if (id && statut === "Ongoing") parrainsOngoing[id] = (parrainsOngoing[id] || 0) + 1;
  });

  // Fusion des paiements S1 et S2 — chaque onglet est lu séparément
  // puis fusionné pour obtenir l'index complet du semestre en cours
  const indexPaiements = {
    ...indexerPaiementsHelloAsso(CONFIG.SHEET_HELLO_ASSO_1),  // lit "helloasso-s1"
    ...indexerPaiementsHelloAsso(CONFIG.SHEET_HELLO_ASSO_2),  // lit "helloasso-s2"
     ...indexerPaiementsManuels("Parrainage") 
  };

  const sheetSuivi = ss.getSheetByName(CONFIG.SHEET_SUIVI);
  const suiviData  = sheetSuivi.getDataRange().getValues();
  const suiviMap   = {};
  const surplusMap = {};

  suiviData.slice(1).forEach((r, i) => {
    const id     = r[CONFIG.COLS.SUIVI.SPONSOR_ID - 1];
    const semLbl = r[CONFIG.COLS.SUIVI.SEMESTRE   - 1];
    suiviMap[`${id}-${semLbl}`] = i + 2;

    const surplus = parseFloat(r[CONFIG.COLS.SUIVI.SURPLUS_PAIEMENT - 1]) || 0;
    if (surplus > 0) surplusMap[id] = surplus;
  });


  Object.entries(parrainsOngoing).forEach(([sponsorId, nb]) => {
  const parrain = parrainsMap[sponsorId];
  if (!parrain) return;

  // Montant attendu pour le semestre courant
  const montantAttendu = nb * CONFIG.MONTANT_PAR_PARRAINAGE;

  // Paiement le plus récent
  const paiement = indexPaiements[parrain.email.toLowerCase()];
  let montantPaye = paiement?.montant || 0;
  const datePaiement = paiement?.date || "";

  // Calcul du surplus pour le semestre courant uniquement
  let nouveauSurplus = 0;
  if (montantPaye > montantAttendu) {
    nouveauSurplus = montantPaye - montantAttendu;
  }

  // Calcul du statut et lien relance
  const statut = calculerStatut(montantAttendu, montantPaye);
  const c = CONFIG.COLS.SUIVI;
  const ligne = suiviMap[`${sponsorId}-${semestre.label}`];

  const tokenRelance = Utilities.getUuid();
  props.setProperty(`TOKEN_RELANCE_${sponsorId}`, tokenRelance);
  const lienRelance = statut.includes("⚠️")
    ? `=HYPERLINK("${CONFIG.WEBAPP_URL}?action=relancer&token=${tokenRelance}&sponsor=${sponsorId}&ligne=${ligne || ''}","Relancer")`
    : "";

  // Mise à jour ou création
  if (ligne) {
    sheetSuivi.getRange(ligne, c.MONTANT_ATTENDU).setValue(montantAttendu);
    sheetSuivi.getRange(ligne, c.MONTANT_PAYE).setValue(montantPaye);
    sheetSuivi.getRange(ligne, c.STATUT).setValue(statut);
    sheetSuivi.getRange(ligne, c.DATE_PAIEMENT).setValue(datePaiement);
    sheetSuivi.getRange(ligne, c.HYPERLINK_RELANCE).setFormula(lienRelance);
    sheetSuivi.getRange(ligne, c.SURPLUS_PAIEMENT).setValue(nouveauSurplus);
  } else {
    sheetSuivi.appendRow([
      semestre.label.split(" ")[1], semestre.label, new Date(),
      parrain.nom, parrain.email, sponsorId,
      montantAttendu, montantPaye, nouveauSurplus,
      statut, datePaiement, "", lienRelance
    ]);
  }
});


  Logger.log("✅ Suivi_Parrainages mis à jour.");
}


// ============================================================
// MISE À JOUR SUIVI COTISATIONS
// ============================================================

function mergeCotisations() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const annee = new Date().getFullYear().toString();

  const sheetMerged = ss.getSheetByName(CONFIG.SHEET_COTISATIONS_MERGED);
  if (!sheetMerged) {
    Logger.log(`❌ Onglet "${CONFIG.SHEET_COTISATIONS_MERGED}" introuvable`);
    return;
  }

  const lastRow = sheetMerged.getLastRow();
  if (lastRow > 1) {
    sheetMerged.getRange(2, 1, lastRow - 1, sheetMerged.getLastColumn()).clearContent();
  }

  const cm = CONFIG.COLS.COTISATIONS_MERGED;

  // Source 1 — HelloAsso cotisation
  const sheetHA = ss.getSheetByName(CONFIG.SHEET_HELLO_ASSO_COTISATION);
  if (sheetHA) {
    const lignes = [];
    sheetHA.getDataRange().getValues().slice(1).forEach(row => {
      const ligne = new Array(Object.keys(CONFIG.COLS.COTISATIONS_MERGED).length).fill("");
      ligne[cm.DATE    - 1] = row[CONFIG.COLS.HELLO_ASSO.DATE    - 1];
      ligne[cm.PRENOM  - 1] = row[CONFIG.COLS.HELLO_ASSO.PRENOM  - 1];
      ligne[cm.NOM     - 1] = row[CONFIG.COLS.HELLO_ASSO.NOM     - 1];
      ligne[cm.EMAIL   - 1] = row[CONFIG.COLS.HELLO_ASSO.EMAIL   - 1];
      ligne[cm.MONTANT - 1] = row[CONFIG.COLS.HELLO_ASSO.MONTANT - 1];
      ligne[cm.SOURCE  - 1] = "HelloAsso - " + annee;
      lignes.push(ligne);
    });
    if (lignes.length > 0)
      sheetMerged.getRange(sheetMerged.getLastRow() + 1, 1, lignes.length, lignes[0].length)
        .setValues(lignes);
    Logger.log(`✅ ${lignes.length} lignes HelloAsso cotisation copiées`);
  }

  // 🔑 Correspondance "NOM PRENOM" → { email, nom, prenom } depuis le sheet Membres
  const sheetMembres = ss.getSheetByName(CONFIG.SHEET_MEMBRES);
  const cm2 = CONFIG.COLS.MEMBRES;
  const infoParNomComplet = {};
  if (sheetMembres) {
    sheetMembres.getDataRange().getValues().slice(1).forEach(r => {
      const nom    = String(r[cm2.MEMBRE_NAME      - 1] || '').trim();
      const prenom = String(r[cm2.MEMBRE_FIRSTNAME - 1] || '').trim();
      const email  = String(r[cm2.ADRESSE_MAIL     - 1] || '').trim().toLowerCase();
      if (nom && prenom) {
        infoParNomComplet[`${nom} ${prenom}`] = { nom, prenom, email };
      }
    });
  }

  // Source 2 — Paiements manuels filtrés sur Cotisation annuelle
  const sheetManuels = ss.getSheetByName(CONFIG.SHEET_PAIEMENTS_MANUELS);
  if (sheetManuels) {
    const c      = CONFIG.COLS.PAIEMENT_MANUEL;
    const lignes = [];
    sheetManuels.getDataRange().getValues().slice(1).forEach(row => {
      if (row[c.TYPE - 1]?.toString().trim() !== "Cotisation annuelle") return;

      const nomMembre = String(row[c.MEMBRE - 1] || '').trim();
      const info      = infoParNomComplet[nomMembre];

      // Nom/prénom depuis le sheet Membres si trouvé, sinon split basique
      const nom    = info ? info.nom    : nomMembre.split(' ')[0] || nomMembre;
      const prenom = info ? info.prenom : nomMembre.split(' ').slice(1).join(' ') || '';
      const email  = row[c.EMAIL - 1] || (info ? info.email : '') || '';

      const ligne = new Array(Object.keys(CONFIG.COLS.COTISATIONS_MERGED).length).fill("");
      ligne[cm.DATE    - 1] = row[c.DATE_COTISATION    - 1] || row[c.HORODATAGE - 1];
      ligne[cm.PRENOM  - 1] = prenom;
      ligne[cm.NOM     - 1] = nom;
      ligne[cm.EMAIL   - 1] = email;
      ligne[cm.MONTANT - 1] = parseFloat(String(row[c.MONTANT_COTISATION - 1]).replace(',', '.')) || 0;
      ligne[cm.SOURCE  - 1] = (row[c.MOYEN_COTISATION  - 1] || "Manuel") + " - " + annee;
      lignes.push(ligne);
    });
    if (lignes.length > 0)
      sheetMerged.getRange(sheetMerged.getLastRow() + 1, 1, lignes.length, lignes[0].length)
        .setValues(lignes);
    Logger.log(`✅ ${lignes.length} paiements manuels cotisation copiés`);
  }

  Logger.log("✅ mergeCotisations terminé");
}
// ============================================================
// Indexer Paiements HelloAsso, Manuels
// ============================================================

function indexerPaiementsHelloAsso(nomOnglet) {
  // nomOnglet détermine QUEL onglet on lit
  // CONFIG.COLS.HELLO_ASSO détermine QUELLES colonnes on lit (même structure pour les 3 onglets)
  const index = {};
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(nomOnglet);
  if (!sheet) return index;

  sheet.getDataRange().getValues().slice(1).forEach(row => {
    const email   = row[CONFIG.COLS.HELLO_ASSO.EMAIL   - 1].toString().toLowerCase().trim();
    const montant = parseFloat(row[CONFIG.COLS.HELLO_ASSO.MONTANT - 1]) || 0;
    const date    = new Date(row[CONFIG.COLS.HELLO_ASSO.DATE - 1]);
    if (!index[email] || date > index[email].date)
      index[email] = { montant, date };
  });
  return index;
}

function indexerPaiementsManuels(filtreType = null) {
  const index = {};
  const sheet = SpreadsheetApp.getActiveSpreadsheet()
    .getSheetByName(CONFIG.SHEET_PAIEMENTS_MANUELS);
  if (!sheet) return index;

  const c = CONFIG.COLS.PAIEMENT_MANUEL;
  sheet.getDataRange().getValues().slice(1).forEach(row => {
    const type    = row[c.TYPE - 1]?.toString().trim();
    const email   = row[c.EMAIL - 1]?.toString().toLowerCase().trim();

    // Lecture des bonnes colonnes selon la section empruntée
    const montant = type === "Parrainage"
      ? parseFloat(row[c.MONTANT_PARRAINAGE - 1]) || 0
      : parseFloat(row[c.MONTANT_COTISATION - 1]) || 0;
    const date    = type === "Parrainage"
      ? new Date(row[c.DATE_PARRAINAGE - 1] || row[c.HORODATAGE - 1])
      : new Date(row[c.DATE_COTISATION - 1] || row[c.HORODATAGE - 1]);

    if (filtreType && type !== filtreType) return;
    if (!email || montant === 0) return;
    if (!index[email] || date > index[email].date)
      index[email] = { montant, date, source: "manuel" };
  });

  return index;
}

/*------------------------------------------------
/* Process ajout paiement manuel form
/*---------------------------------------------------
*/
function processPaiementManuel(newRow) {
  const sheet    = newRow.sheet;
  const rowIndex = newRow.row;
  const lastCol  = sheet.getLastColumn();

  // Relire la ligne après déplacement
  const rowData = sheet.getRange(rowIndex, 1, 1, lastCol).getValues()[0];

  const c    = CONFIG.COLS.PAIEMENT_MANUEL;
  const type = rowData[c.TYPE - 1]?.toString().trim();

  const nom     = type === "Parrainage"
    ? rowData[c.PARRAIN            - 1]
    : rowData[c.MEMBRE             - 1];
  const montant = type === "Parrainage"
    ? rowData[c.MONTANT_PARRAINAGE - 1]
    : rowData[c.MONTANT_COTISATION - 1];
  const date    = type === "Parrainage"
    ? rowData[c.DATE_PARRAINAGE    - 1]
    : rowData[c.DATE_COTISATION    - 1];
  const moyen   = type === "Parrainage"
    ? rowData[c.MOYEN_PARRAINAGE   - 1]
    : rowData[c.MOYEN_COTISATION   - 1];

  Logger.log(`✅ Paiement manuel — Type: ${type} | Nom: ${nom} | Montant: ${montant} | Date: ${date} | Moyen: ${moyen}`);

  const email = _getEmailDepuisNom(nom, type);
  if (!email) {
    Logger.log(`❌ Email introuvable pour : ${nom} (${type})`);
    return;
  }

  sheet.getRange(rowIndex, c.EMAIL).setValue(email);
  Logger.log(`✅ Email résolu et écrit : ${email}`);

  if (type === "Parrainage") mergeParrainages();
  if (type === "Cotisation") mergeCotisations();
}

function _getEmailDepuisNom(valeurDropdown, type) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  if (type === "Parrainage") {
    // Extraire le SPONSOR_ID entre parenthèses
    const match = valeurDropdown.match(/\(([^)]+)\)$/);
    if (!match) {
      Logger.log(`❌ Impossible d'extraire l'ID depuis : ${valeurDropdown}`);
      return null;
    }
    const sponsorId = match[1].trim();
    const rows = ss.getSheetByName(CONFIG.SHEET_PARRAIN)
      .getDataRange().getValues().slice(1);
    const row = rows.find(r => r[CONFIG.COLS.PARRAIN.SPONSOR_ID - 1].toString().trim() === sponsorId);
    if (!row) {
      Logger.log(`❌ Parrain introuvable pour ID : ${sponsorId}`);
      return null;
    }
    return row[CONFIG.COLS.PARRAIN.ADRESSE_MAIL - 1].toString().toLowerCase().trim();
  }

    if (type === "Cotisation") {
    const rows = ss.getSheetByName(CONFIG.SHEET_MEMBRES)
      .getDataRange().getValues().slice(1);
    const row = rows.find(r => {
      const nomComplet = `${r[CONFIG.COLS.MEMBRES.MEMBRE_NAME - 1]} ${r[CONFIG.COLS.MEMBRES.MEMBRE_FIRSTNAME - 1]}`.trim();
      return nomComplet === valeurDropdown;
    });
    if (!row) {
      Logger.log(`❌ Membre introuvable pour : ${valeurDropdown}`);
      return null;
    }
    return row[CONFIG.COLS.MEMBRES.ADRESSE_MAIL - 1].toString().toLowerCase().trim();
  }
}



// ============================================================
// RECHERCHE ET EXPORT MANUEL
// ============================================================

function forcerRechercheEtExport() {
  const token     = getAccessToken();
  const semestre  = getSemestreEnCours();
  const campagnes = getSlugsCampagneSemestre(token);
  const ui        = SpreadsheetApp.getUi();

  if (campagnes.length === 0) {
    ui.alert(
      `⚠️ Aucune campagne trouvée pour ${semestre.label}.\n\n`
      + `Vérifiez que la campagne a bien été créée sur HelloAsso `
      + `avec un titre contenant "Parrainage", "S${semestre.numero}" et "${semestre.label.split(" ")[1]}".`
    );
    return;
  }

  const nomOnglet = semestre.numero === 1 ? CONFIG.SHEET_HELLO_ASSO_1 : CONFIG.SHEET_HELLO_ASSO_2;
  campagnes.forEach(c => {
    mettreAJourConfiguration(c.slug, c.url);
    _exportSlugToSheet(c.slug, nomOnglet, token);
  });
  mergeParrainages();

  ui.alert(
    `✅ ${campagnes.length} campagne(s) synchronisée(s) pour ${semestre.label}.\n\n`
    + campagnes.map(c => `• ${c.slug}`).join("\n")
  );
}



function getParrainsActifsIds() {
  const ids = new Set();
  SpreadsheetApp.getActiveSpreadsheet()
    .getSheetByName(CONFIG.SHEET_PARRAINAGE)
    .getDataRange().getValues().slice(1)
    .forEach(row => {
      const id     = row[CONFIG.COLS.PARRAINAGE.SPONSOR_ID - 1];
      const statut = row[CONFIG.COLS.PARRAINAGE.STATUS - 1];
      if (id && statut === "Ongoing") ids.add(id);
    });
  return ids;
}






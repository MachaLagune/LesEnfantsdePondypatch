/**
 * ===============================================
 * TRAITEMENT D'UNE NOUVELLE SOUMISSION DE PARRAINAGE
 * ===============================================
 */
function processSponsorshipSubmission() {
  try {
    Logger.log("=== Début processSponsorshipSubmission ===");

    // 🔹 Petit délai pour laisser le Form écrire la ligne
    Utilities.sleep(2000);

    // Accès à la feuille Parrainage
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(CONFIG.SHEET_PARRAINAGE);
    if (!sheet) {
      Logger.log(`Feuille '${CONFIG.SHEET_PARRAINAGE}' introuvable !`);
      return;
    }

    // Dernière ligne
    const actualRow = sheet.getLastRow();

    // ==================================================
    // 1️⃣ Séparer Kid Name / Kid_ID
    // ==================================================
    const sourceColKid = CONFIG.COLS.PARRAINAGE.KID_NAME;
    const idColKid     = CONFIG.COLS.PARRAINAGE.KID_ID;


    const cellKid = sheet.getRange(actualRow, sourceColKid);
    const valueKid = cellKid.getValue();

    if (typeof valueKid === "string") {
      const match = valueKid.match(/^(.+?)\s*\(([^)]+)\)$/);
      if (match) {
        cellKid.setValue(match[1].trim());
        sheet.getRange(actualRow, idColKid).setValue(match[2].trim());
        Logger.log(`✅ Sponsor séparé : ${match[1]} / ${match[2]}`);
      }
    }

    // ==================================================
    // 1️⃣ Séparer Sponsor Name / Sponsor_ID
    // ==================================================
    const sourceCol = CONFIG.COLS.PARRAINAGE.SPONSOR_NAME;
    const idCol     = CONFIG.COLS.PARRAINAGE.SPONSOR_ID;

    const cell = sheet.getRange(actualRow, sourceCol);
    const value = cell.getValue();

    if (typeof value === "string") {
      const match = value.match(/^(.+?)\s*\(([^)]+)\)$/);
      if (match) {
        cell.setValue(match[1].trim());
        sheet.getRange(actualRow, idCol).setValue(match[2].trim());
        Logger.log(`✅ Sponsor séparé : ${match[1]} / ${match[2]}`);
      }
    }

    // ==================================================
    // 2️⃣ Assigner Sponsorship_ID si vide
    // ==================================================
    const spipCol = CONFIG.COLS.PARRAINAGE.SPONSORSHIP_ID;
    const spipCell = sheet.getRange(actualRow, spipCol);

    if (!spipCell.getValue()) {
      const allIds = sheet.getRange(2, spipCol, sheet.getLastRow() - 1).getValues();
      let maxId = 0;
      allIds.forEach(r => {
        if (r[0]) {
          const num = parseInt(r[0].replace("SPIP-", ""), 10);
          if (!isNaN(num) && num > maxId) maxId = num;
        }
      });
      maxId++;
      spipCell.setValue("SPIP-" + Utilities.formatString("%04d", maxId));
      Logger.log(`✅ Sponsorship_ID assigné : ${spipCell.getValue()}`);
    }

    // ==================================================
    // 3️⃣ Statut par défaut "Ongoing"
    // ==================================================
    const statusCol = CONFIG.COLS.PARRAINAGE.STATUS;
    const statusCell = sheet.getRange(actualRow, statusCol);
    if (!statusCell.getValue()) {
      statusCell.setValue("Ongoing");
      Logger.log("✅ Statut par défaut 'Ongoing' appliqué");
    }

    Logger.log(`✅ Traitement terminé pour la ligne ${actualRow}`);

  } catch (err) {
    Logger.log("❌ Erreur processSponsorshipSubmission : " + err);
    throw err;
  }
}




/**
 * ===============================================
 * Traitement d'une nouvelle soumission d fin de parrainage
 * ===============================================
 */
function markParrainageFinished() {
  Utilities.sleep(2000);
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  const finSheet = ss.getSheetByName(CONFIG.SHEET_FIN_PARRAINAGE);
  const parrainageSheet = ss.getSheetByName(CONFIG.SHEET_PARRAINAGE);

  if (!finSheet || !parrainageSheet) {
    Logger.log("❌ Feuilles manquantes");
    return;
  }

  const lastRow = finSheet.getLastRow();
  if (lastRow < 2) { // pas de données
    Logger.log("⚠️ Pas de ligne à traiter dans Fin_Parrainage");
    return;
  }

  const finCols = CONFIG.COLS.FIN;
  const parCols = CONFIG.COLS.PARRAINAGE;

  const kidNameWithID = finSheet.getRange(lastRow, finCols.KID_NAME).getValue();
  if (!kidNameWithID) {
    Logger.log("❌ Kid Name vide à la dernière ligne");
    return;
  }

  // 🔹 Extraire Kid_ID
  const match = kidNameWithID.match(/\((KID-\d+)\)$/);
  if (!match) {
    Logger.log("❌ Impossible d'extraire le Kid_ID depuis : " + kidNameWithID);
    return;
  }
  const kidID = match[1];

  // 📝 Écrire Kid_ID dans Fin_Parrainage si vide
  const currentKidID = finSheet.getRange(lastRow, finCols.KID_ID).getValue();
  if (!currentKidID) {
    finSheet.getRange(lastRow, finCols.KID_ID).setValue(kidID);
    Logger.log(`📝 Kid_ID ${kidID} écrit dans Fin_Parrainage (ligne ${lastRow})`);
  }

  // 🔎 Chercher le dernier parrainage actif dans Parrainage
  const parData = parrainageSheet.getDataRange().getValues();
  let lastRowIndex = -1;
  let sponsorshipId = null;
  for (let i = parData.length - 1; i >= 1; i--) {
    if (parData[i][parCols.KID_ID - 1] === kidID && parData[i][parCols.STATUS - 1] !== "Finished") {
      lastRowIndex = i;
      sponsorshipId = parData[i][parCols.SPONSORSHIP_ID - 1];
      break;
    }
  }

  if (lastRowIndex === -1) {
    Logger.log(`⚠️ Aucun parrainage actif trouvé pour ${kidID}`);
    return;
  }

  // 📝 Écrire Sponsorship_ID dans Fin_Parrainage si vide
  const currentSponsorship = finSheet.getRange(lastRow, finCols.SPONSORSHIP_ID).getValue();
  if (!currentSponsorship && sponsorshipId) {
    finSheet.getRange(lastRow, finCols.SPONSORSHIP_ID).setValue(sponsorshipId);
    Logger.log(`📝 Sponsorship_ID ${sponsorshipId} écrit dans Fin_Parrainage (ligne ${lastRow})`);
  }

  // ⚡️ Mettre à jour le statut dans Parrainage, la date de fin et la raison de fin

  // Récupérer la date de fin et la raison depuis Fin_Parrainage
  const endDate = finSheet.getRange(lastRow, finCols.DATE_FIN_PARRAINAGE).getValue();
  const reason  = finSheet.getRange(lastRow, finCols.RAISON_ARRET_PARRAINAGE).getValue();

  // Écrire ces valeurs dans Parrainage
  parrainageSheet.getRange(lastRowIndex + 1, parCols.SPONSORSHIP_END_DATE).setValue(endDate);
  parrainageSheet.getRange(lastRowIndex + 1, parCols.REASON_STOPPING).setValue(reason);
  parrainageSheet.getRange(lastRowIndex + 1, parCols.STATUS).setValue("Finished");


  Logger.log(`✅ Parrainage ${sponsorshipId} clôturé pour ${kidID} (ligne ${lastRow})`);
}

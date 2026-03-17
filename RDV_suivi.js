/**
 * ===============================================
 * TRAITEMENT RDV_SUIVI 
 * ===============================================
 */
function processRdvSuivi() {
  try {
    Logger.log("=== Début processRdvSuivi ===");

    // 🔹 Pause pour laisser Google Forms écrire la ligne
    Utilities.sleep(2000);

    // ----------------------------
    // 1️⃣ Accès aux feuilles
    // ----------------------------
    const ss = SpreadsheetApp.getActiveSpreadsheet();

    const parrainageSheet = ss.getSheetByName(CONFIG.SHEET_PARRAINAGE);
    if (!parrainageSheet) {
      Logger.log("❌ Feuille " + CONFIG.SHEET_PARRAINAGE + " introuvable");
      return;
    }

    const rdvSheet = ss.getSheetByName(CONFIG.SHEET_RDV);
    if (!rdvSheet) {
      Logger.log("❌ Feuille " + CONFIG.SHEET_RDV + " introuvable");
      return;
    }

    // ----------------------------
    // 2️⃣ Dernière ligne du Rdv
    // ----------------------------
    const actualRow = rdvSheet.getLastRow();
    Logger.log("🔹 Traitement de la ligne " + actualRow);

    // ----------------------------
    // 3️⃣ Récupérer Kid Name + Kid_ID de la dernière ligne
    // ----------------------------
    const kidNameWithID = rdvSheet.getRange(actualRow, CONFIG.COLS.RDV.KID_NAME).getValue();
    if (!kidNameWithID) {
      Logger.log("❌ Kid Name vide");
      return;
    }

    const match = kidNameWithID.match(/^(.+?)\s*\((KID-\d+)\)$/);
    if (!match) {
      Logger.log("❌ Impossible d'extraire Kid Name et Kid_ID : " + kidNameWithID);
      return;
    }

    const kidName = match[1].trim();
    const kidID = match[2];
    Logger.log(`✅ Kid identifié : ${kidName} (${kidID})`);

    // ----------------------------
    // 4️⃣ Récupérer le Sponsor_ID le plus récent
    // ----------------------------
    const parrainageData = parrainageSheet.getDataRange().getValues();
    const COL_KID_ID = CONFIG.COLS.PARRAINAGE.KID_ID;
    const COL_SPONSOR_ID = CONFIG.COLS.PARRAINAGE.SPONSOR_ID;
    const COL_SPONSOR_NAME = CONFIG.COLS.PARRAINAGE.SPONSOR_NAME;

    let sponsorID = "";
    let sponsorName = "";

    for (let i = parrainageData.length - 1; i >= 1; i--) {
      if (parrainageData[i][COL_KID_ID - 1] === kidID) {
        sponsorID   = parrainageData[i][COL_SPONSOR_ID - 1];
        sponsorName = parrainageData[i][COL_SPONSOR_NAME - 1];
        break;
      }
    }

    if (!sponsorID) {
      Logger.log("⚠️ Aucun Sponsor_ID trouvé pour " + kidID);
    } else {
      Logger.log(`✅ Sponsor_ID récupéré : ${sponsorID}`);
    }

    // ----------------------------
    // 5️⃣ Écriture dans Rdv_suivi
    // ----------------------------
    rdvSheet.getRange(actualRow, CONFIG.COLS.RDV.KID_NAME).setValue(kidName);
    rdvSheet.getRange(actualRow, CONFIG.COLS.RDV.KID_ID).setValue(kidID);
    rdvSheet.getRange(actualRow, CONFIG.COLS.RDV.SPONSOR_NAME).setValue(sponsorName);
    rdvSheet.getRange(actualRow, CONFIG.COLS.RDV.SPONSOR_ID).setValue(sponsorID);

    Logger.log("✅ processRdvSuivi terminé avec succès");

  } catch (err) {
    Logger.log("❌ Erreur processRdvSuivi : " + err);
  }
}



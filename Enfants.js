/**
 * ===============================================
 * MAIN — Mise à jour complète des enfants
 * ===============================================
 */
function mainUpdateKids() {
  try {
    Logger.log("=== Début mise à jour complète des enfants ===");

    // 1️⃣ Assigne les KID_ID et calcule les âges
    assignKidIDandCalculateAge();

    SpreadsheetApp.flush(); // ← force l'écriture avant la lecture

    // 2️⃣ Met à jour les Forms avec les enfants
    UpdateKidsName();

    Logger.log("✅ Mise à jour complète terminée !");
  } catch (err) {
    Logger.log("❌ Erreur dans mainUpdateKids : " + err);
  }
}


/**
 * ===============================================
 * ASSIGNE KID_ID ET CALCULE AGE
 * ===============================================
 */
function assignKidIDandCalculateAge() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(CONFIG.SHEET_ENFANT);
  if (!sheet) return;

  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return; // pas de données

  // Colonnes via CONFIG
  const colKidID = CONFIG.COLS.ENFANT.KID_ID;
  const colAge = CONFIG.COLS.ENFANT.AGE;
  const colBirthDate = CONFIG.COLS.ENFANT.DATE_BIRTH;
  const colName = CONFIG.COLS.ENFANT.KID_NAME;

  if (!colKidID || !colAge || !colBirthDate || !colName) {
    Logger.log("❌ Colonnes KID_ID, AGE, DATE_OF_BIRTH ou NAME non définies dans CONFIG");
    return;
  }

  // Assigner KID_ID manquants et recalculer âge uniquement si nécessaire

  const idRange = sheet.getRange(2, colKidID, lastRow - 1);
  const ids = idRange.getValues();
  const birthDates = sheet.getRange(2, colBirthDate, lastRow - 1).getValues();
  const ages = sheet.getRange(2, colAge, lastRow - 1).getValues();

  let maxId = 0;
  // Trouver le max KID_ID existant
  ids.forEach(r => {
    if (r[0]) {
      const num = parseInt(r[0].toString().replace("KID-", ""), 10);
      if (!isNaN(num) && num > maxId) maxId = num;
    }
  });

  const today = new Date();

  for (let i = 0; i < ids.length; i++) {
    const kidName = sheet.getRange(i + 2, colName).getValue();
    const kidID = ids[i][0];
    const birthDate = birthDates[i][0];

    // Seulement si Nom enfant existe et KID_ID est vide
    if (kidName && !kidID) {
      // Assigner KID_ID
      maxId++;
      ids[i][0] = "KID-" + Utilities.formatString("%04d", maxId);
      Logger.log(`✅ Assigné KID_ID=${ids[i][0]} à la ligne ${i + 2}`);

      // Recalculer l'âge si Date de Naissance valide
      if (birthDate instanceof Date && !isNaN(birthDate)) {
        let age = today.getFullYear() - birthDate.getFullYear();
        const m = today.getMonth() - birthDate.getMonth();
        const d = today.getDate() - birthDate.getDate();
        if (m < 0 || (m === 0 && d < 0)) age--;
        ages[i][0] = age;
        Logger.log(`Ligne ${i + 2} : Age recalculé = ${age}`);
      } else {
        ages[i][0] = "";
        Logger.log(`Ligne ${i + 2} : Date de naissance invalide`);
      }
    }
  }

  // Écriture en une seule fois
  idRange.setValues(ids);
  sheet.getRange(2, colAge, lastRow - 1).setValues(ages);

  Logger.log("✅ Mise à jour des KID_ID et âges des nouvelles lignes terminée");
}


/**
 * ===============================================
 * MISE À JOUR DES GOOGLE FORMS
 * ===============================================
 */
function UpdateKidsName() {
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);

  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(CONFIG.SHEET_ENFANT);
    if (!sheet) return;

    // Colonnes via CONFIG (converties en index 0-based pour getValues())
    const colName  = CONFIG.COLS.ENFANT.KID_NAME - 1;
    const colKidID = CONFIG.COLS.ENFANT.KID_ID   - 1;

    const data = sheet.getDataRange().getValues();


    let kidsForm = data
      .slice(1)
      .map(r => {
        const nom = r[colName]  ? r[colName].toString().trim()  : null;
        const id  = r[colKidID] ? r[colKidID].toString().trim() : null;

        if (!nom || !id) {
          if (nom || id) Logger.log(`⚠️ Ligne ignorée — nom: "${nom}", id: "${id}"`);
          return null;
        }

        return `${nom} (${id})`;
      })
      .filter(Boolean);

    // Dédoublonner et mettre les derniers ajouts en haut
    kidsForm = dedupeKeepLast_(kidsForm);
    kidsForm = kidsForm.sort((a, b) => a.localeCompare(b, 'fr', { sensitivity: 'base' }));

    if (kidsForm.length === 0) {
      Logger.log("⚠️ Aucun enfant valide trouvé — mise à jour annulée");
      return;
    }

    // Mise à jour des Forms via IDs dans CONFIG
    const forms = [
      ["Parrainage",  CONFIG.FORM_PARRAINAGE],
      ["RDV Suivi",   CONFIG.FORM_RDV_SUIVI],
      ["Fin",         CONFIG.FORM_FIN_PARRAINAGE]
    ].filter(([_, id]) => id && id.toString().trim() !== "");

    if (forms.length === 0) {
      throw new Error("Aucun Google Form valide configuré");
    }

    forms.forEach(([name, id]) => {
      Logger.log(`➡️ Mise à jour du form : ${name}`);
      const form = FormApp.openById(id);
      updateDropdown_(form, "Nom enfant", kidsForm);
    });

    Logger.log(`✅ Mise à jour Forms terminée : ${kidsForm.length} enfants`);

  } catch (err) {
    Logger.log("❌ Erreur UpdateKidsName : " + err);
    throw err;
  } finally {
    lock.releaseLock();
  }
}

/**
 * ===============================================
 * CHERCHER ENFANT
 * ===============================================
 */
/**
 * Cherche un enfant dans "Fiche_Enfant_Form" par son ID.
 *
 * @param {string} kidId - ID de l'enfant.
 * @return {{kidId: string, age: number, gender: string}|null} Données de l'enfant ou null si non trouvé.
 */


function chercherEnfant(kidId) {
  const ss    = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(CONFIG.SHEET_ENFANT);
  if (!sheet) {
    Logger.log("❌ Feuille Fiche_Enfant_Form introuvable");
    return null;
  }

  const colId     = CONFIG.COLS.ENFANT.KID_ID   - 1;
  const colAge    = CONFIG.COLS.ENFANT.AGE       - 1;
  const colGender = CONFIG.COLS.ENFANT.GENDER    - 1;

  const data = sheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][colId] || '').trim() === String(kidId || '').trim()) {
      return {
        kidId:  data[i][colId],
        age:    data[i][colAge],
        gender: data[i][colGender]
      };
    }
  }

  Logger.log(`⚠️ Enfant introuvable pour kidId : "${kidId}"`);
  return null;
}



/**
 * ===============================================
 * majTous les suivis et parrainages dans fiche enfant
 * ===============================================
 */

function majTousLesSuivisEtParrainage() {
  Logger.log("=== Début recalcul des âges ===");

  const ss           = SpreadsheetApp.getActiveSpreadsheet();
  const sheetEnfants = ss.getSheetByName(CONFIG.SHEET_ENFANT);
  if (!sheetEnfants) throw new Error("Feuille enfants introuvable");

  const COLS_ENFANTS = CONFIG.COLS.ENFANT;
  const today        = new Date();
  const startRow     = 2;
  const numRows      = sheetEnfants.getLastRow() - startRow + 1;
  if (numRows < 1) return;

  const data = sheetEnfants.getRange(startRow, 1, numRows, sheetEnfants.getLastColumn()).getValues();

  const ages = data.map(row => {
    const birthDate = row[COLS_ENFANTS.DATE_BIRTH - 1];
    if (birthDate instanceof Date && !isNaN(birthDate)) {
      let age = today.getFullYear() - birthDate.getFullYear();
      const m = today.getMonth() - birthDate.getMonth();
      const d = today.getDate() - birthDate.getDate();
      if (m < 0 || (m === 0 && d < 0)) age--;
      return [age];
    }
    return [""];
  });

  sheetEnfants.getRange(startRow, COLS_ENFANTS.AGE, ages.length, 1).setValues(ages);
  Logger.log(`✅ Âges recalculés pour ${ages.length} enfants`);
}
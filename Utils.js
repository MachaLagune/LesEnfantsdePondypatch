/**
 * Vérifie si le formulaire déclencheur est bien celui attendu
 * @param {GoogleAppsScript.Events.FormsOnFormSubmit} e 
 * @param {string} allowedFormId 
 * @return {boolean}
 */
 
function isAllowedForm_(e, allowedFormId) {
  if (!e || !e.source) return false;
  return e.source.getId() === allowedFormId;
}

/**
 * Supprime les doublons d'une liste en gardant la dernière occurrence
 * @param {Array} list 
 * @return {Array}
 */
function dedupeKeepLast_(list) {
  const seen = new Set();
  return list
    .slice()
    .reverse()
    .filter(v => {
      if (seen.has(v)) return false;
      seen.add(v);
      return true;
    })
    .reverse();
}

/**
 * Déplace les éléments récents en premier
 * @param {Array} list 
 * @param {number} nb - Nombre d'éléments récents à déplacer
 * @return {Array}
 */
function moveRecentFirst_(list, nb) {
  if (!Array.isArray(list) || list.length === 0) return [];
  const last = list.slice(-nb).reverse();
  const older = list.slice(0, -nb);
  return [...last, ...older];
}

/**
 * Met à jour ou crée une liste déroulante dans un Google Form
 */
function updateDropdown_(form, title, values) {
  if (!values || values.length === 0) return;
  const items = form.getItems(FormApp.ItemType.LIST);
  let question = items.find(i => i.getTitle() === title);
  if (!question) question = form.addListItem().setTitle(title);
  else question = question.asListItem();
  question.setChoiceValues(values);
}


/**
 * Récupère un Blob d'une image depuis un lien Google Drive
 * @param {string} url - URL du fichier Google Drive
 * @return {Blob|null}
 */
function getImageBlobFromUrl(url) {
  const match = url.match(/[-\w]{25,}/);
  if (!match) return null;
  try {
    return DriveApp.getFileById(match[0]).getBlob();
  } catch (e) {
    Logger.log("⚠️ Impossible de récupérer l'image depuis l'URL : " + url);
    return null;
  }
}


/**
 * Ajoute des paragraphes vides dans un document pour créer des espaces
 * @param {Body} body - Body d'un DocumentApp
 * @param {number} n - Nombre de paragraphes vides à ajouter
 */
function addSpacer(body, n = 1) {
  for (let i = 0; i < n; i++) body.appendParagraph("");
}


/**
 * Extrait l'ID d'un document Google (Docs, Sheets, Drive) à partir de son URL.
 * @param {string} url - L'URL complète du document Google.
 * @return {string|null} L'ID du document ou null si l'URL est invalide.
 */
function extractDocId(url) {
  if (!url) return null;
  const match = url.match(/[-\w]{25,}/);
  return match ? match[0] : null;
}



/**
 * Déplace une ligne à la fin d'une sheet
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet 
 * @param {number} rowIndex 
 * @param {number} timestampCol - numéro de colonne de l'horodatage
 */

function moveRowToEnd(sheet, rowIndex,timestampCol) {
  const lastRow = sheet.getLastRow();
  Logger.log(`moveRowToEnd — rowIndex: ${rowIndex}, lastRow: ${lastRow}`);
  
  if (rowIndex >= lastRow) {
    Logger.log("⚠️ Ligne déjà en dernière position, rien à faire");
    return;
  }

  const range = sheet.getRange(rowIndex, 1, 1, sheet.getLastColumn());
  const values = range.getValues();
  const formats = range.getNumberFormats();
  const backgrounds = range.getBackgrounds();
  
  sheet.insertRowAfter(lastRow);
  const targetRange = sheet.getRange(lastRow + 1, 1, 1, sheet.getLastColumn());
  targetRange.setValues(values);
  targetRange.setBackgrounds(backgrounds);

  sheet.deleteRow(rowIndex);
}


function checkAndMoveNewRow() {
  Logger.log("🔵 checkAndMoveNewRow démarrée");
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  const sheetsToCheck = [
    {
      name: CONFIG.SHEET_ENFANT,
      cols: CONFIG.COLS.ENFANT,
      check: (row, c) => row[c.KID_NAME - 1] && !row[c.KID_ID - 1]
    },
    {
      name: CONFIG.SHEET_PARRAIN,
      cols: CONFIG.COLS.PARRAIN,
      check: (row, c) => row[c.SPONSOR_NAME - 1] && !row[c.SPONSOR_ID - 1]
    },
    {
      name: CONFIG.SHEET_PARRAINAGE,
      cols: CONFIG.COLS.PARRAINAGE,
      check: (row, c) => row[c.KID_NAME - 1] && !row[c.SPONSORSHIP_ID - 1]
    },
    {
      name: CONFIG.SHEET_RDV,
      cols: CONFIG.COLS.RDV,
      check: (row, c) => row[c.KID_NAME - 1] && !row[c.KID_ID - 1]
    },
    {
      name: CONFIG.SHEET_FIN_PARRAINAGE,
      cols: CONFIG.COLS.FIN,
      check: (row, c) => row[c.KID_NAME - 1] && !row[c.KID_ID - 1]
    },

     {
      name: CONFIG.SHEET_VALIDATION_CR,
      cols: CONFIG.COLS.VALIDATION,
      check: (row, c) => row[c.LIEN_CR - 1] && !row[c.CR_ENVOYE - 1]
    },

    { 
      name: CONFIG.SHEET_MEMBRES,
      cols: CONFIG.COLS.MEMBRES,
      check: null, // pas de check classique
      useLatestTimestamp: true,
      timestampCol: CONFIG.COLS.MEMBRES.HORODATEUR
    },

    
    {name: CONFIG.SHEET_PAIEMENTS_MANUELS,
      cols: CONFIG.COLS.PAIEMENT_MANUEL,
      check: null,
      useLatestTimestamp: true,
      timestampCol: CONFIG.COLS.PAIEMENT_MANUEL.HORODATEUR  // a un type mais pas encore d'email résolu
    }

  ];

  for (const s of sheetsToCheck) {
    const sheet = ss.getSheetByName(s.name);
    Logger.log(`🔍 Scan sheet: ${s.name} → trouvée: ${!!sheet}`);
    if (!sheet) continue;

    const lastRow = sheet.getLastRow();
    Logger.log(`   lastRow: ${lastRow}`);
    if (lastRow < 2) continue;

    const lastCol = sheet.getLastColumn();
    const data = sheet.getRange(2, 1, lastRow - 1, lastCol).getValues();
    Logger.log(`   lignes à scanner: ${data.length}`);

    let targetIndex;

    if (s.useLatestTimestamp) {
      const tsCol = s.timestampCol - 1;
      Logger.log(`   mode timestamp — tsCol: ${tsCol}`);
      let latestTime = 0;
      let latestIndex = -1;

      for (let i = 0; i < data.length; i++) {
        const ts = data[i][tsCol];
        Logger.log(`   ligne ${i + 2} → timestamp: ${ts} (type: ${typeof ts}, isDate: ${ts instanceof Date})`);
        if (ts instanceof Date && ts.getTime() > latestTime) {
          latestTime = ts.getTime();
          latestIndex = i;
          Logger.log(`   ✅ Nouveau candidat ligne ${i + 2} : ${ts}`);
        }
      }

      if (latestIndex === -1) {
        Logger.log(`   ⚠️ Aucun timestamp valide trouvé dans ${s.name}`);
        continue;
      }
      targetIndex = latestIndex;

    } else {
      targetIndex = -1;
      for (let i = data.length - 1; i >= 0; i--) {
        if (s.check(data[i], s.cols)) {
          targetIndex = i;
          break;
        }
      }
      if (targetIndex === -1) {
        Logger.log(`   ⚠️ Aucune ligne valide trouvée dans ${s.name}`);
        continue;
      }
    }

    const rowIndex = targetIndex + 2;
    Logger.log(`✅ Ligne à traiter détectée dans ${s.name} (ligne ${rowIndex})`);
    moveRowToEnd(sheet, rowIndex, s.cols.HORODATEUR);
    return { sheet, row: sheet.getLastRow() };
  }

  Logger.log("⚠️ Aucune nouvelle ligne détectée");
  return null;
}
























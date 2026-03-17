/**
 * ===============================================
 * MAIN — Mise à jour complète des sponsors
 * ===============================================
 */
function mainUpdateSponsors() {
  try {
    Logger.log("=== Début mise à jour complète des sponsors ===");

    // 🔹 Pause pour laisser Google Forms écrire la ligne
    Utilities.sleep(2000);

    // 1️⃣ Assigner les SPONSOR_ID
    assignSponsorIDs();

    // 2️⃣ Mettre à jour le Form
    UpdateSponsors();

    Logger.log("✅ Mise à jour complète Sponsors terminée !");
  } catch (err) {
    Logger.log("❌ Erreur dans mainUpdateSponsors : " + err);
  }
}


/**
 * ===============================================
 * ASSIGNE SPONSOR_ID
 * ===============================================
 */
function assignSponsorIDs() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(CONFIG.SHEET_PARRAIN);
  if (!sheet) {
    Logger.log(`Feuille '${CONFIG.SHEET_PARRAIN}' introuvable !`);
    return;
  }

  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return;

  const colId   = CONFIG.COLS.PARRAIN.SPONSOR_ID;
  const idRange = sheet.getRange(2, colId, lastRow - 1, 1);
  const ids     = idRange.getValues();

  let maxId = 0;
  ids.forEach(r => {
    if (r[0]) {
      const num = parseInt(String(r[0]).replace("SP-", ""), 10);
      if (!isNaN(num) && num > maxId) maxId = num;
    }
  });

  for (let i = 0; i < ids.length; i++) {
    if (!ids[i][0]) {
      maxId++;
      ids[i][0] = "SP-" + Utilities.formatString("%04d", maxId);
      Logger.log(`Assigné SPONSOR_ID=${ids[i][0]} à la ligne ${i + 2}`);
    }
  }

  idRange.setValues(ids);
  Logger.log("✅ Assignation des SPONSOR_ID terminée");
}

/**
 * ===============================================
 * MISE À JOUR DU GOOGLE FORM
 * ===============================================
 */
function UpdateSponsors() {
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);

  try {
    Utilities.sleep(2500);

    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(CONFIG.SHEET_PARRAIN);
    if (!sheet) {
      Logger.log(`Feuille introuvable : ${CONFIG.SHEET_PARRAIN}`);
      return;
    }

    const colName = CONFIG.COLS.PARRAIN.SPONSOR_NAME - 1;
    const colId   = CONFIG.COLS.PARRAIN.SPONSOR_ID   - 1;

    const data = sheet.getDataRange().getValues();

    let sponsorsList = data.slice(1)
      .filter(r => r[colName] && r[colId])
      .map(r => `${r[colName].toString().trim()} (${r[colId]})`);

    sponsorsList = dedupeKeepLast_(sponsorsList);
    sponsorsList = sponsorsList.sort((a, b) => a.localeCompare(b, 'fr', { sensitivity: 'base' }));
    

    const form = FormApp.openById(CONFIG.FORM_PARRAINAGE);
    updateDropdown_(form, "Nom du parrain", sponsorsList);
    Logger.log(`✅ Form parrainage mis à jour : ${sponsorsList.length} parrains`);

    try {
      const formPaiement = FormApp.openById(CONFIG.FORM_PAIEMENTS_MANUELS);
      updateDropdown_(formPaiement, "Nom du parrain", sponsorsList);
      Logger.log(`✅ Form paiements manuels — dropdown parrains mis à jour : ${sponsorsList.length} entrées`);

      const membersList = getMembersList();
      updateDropdown_(formPaiement, "Nom du membre", membersList);
      Logger.log(`✅ Form paiements manuels — dropdown membres mis à jour : ${membersList.length} entrées`);
    } catch (err) {
      Logger.log(`❌ Erreur mise à jour Form paiements manuels : ${err}`);
    }

  } catch (err) {
    Logger.log("❌ Erreur UpdateSponsors : " + err);
    throw err;
  } finally {
    lock.releaseLock();
  }
}


/**
 * ===============================================
 * Charcher parrain par ID dans "Fiche_Parrain_Form"
 * ===============================================
 */
/**
 * Cherche un parrain dans "Fiche_Parrain_Form" par son ID.
 *
 * @param {string} sponsorId - ID du parrain.
 * @return {{sponsorId: string, email: string}|null} Données du parrain ou null si non trouvé.
 */


function chercherParrain(sponsorId) {
  const ss    = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(CONFIG.SHEET_PARRAIN);
  if (!sheet) {
    Logger.log("❌ Feuille Fiche_Parrain_Form introuvable");
    return null;
  }

  const colId    = CONFIG.COLS.PARRAIN.SPONSOR_ID   - 1;
  const colEmail = CONFIG.COLS.PARRAIN.ADRESSE_MAIL - 1;
  const colName  = CONFIG.COLS.PARRAIN.SPONSOR_NAME - 1;

  const data = sheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][colId] || '').trim() === String(sponsorId || '').trim()) {
      return {
        sponsorId: data[i][colId],
        name:      data[i][colName],
        email:     data[i][colEmail]
      };
    }
  }

  Logger.log(`⚠️ Parrain introuvable pour sponsorId : "${sponsorId}"`);
  return null;
}


// ============================================================
// ENVOI D'EMAILS — PARRAINAGES - COTISATIONS
// ============================================================

function envoyerLienPaiementAuxParrains() {
  const ss       = SpreadsheetApp.getActiveSpreadsheet();
  const semestre = getSemestreEnCours();
  const props    = PropertiesService.getScriptProperties();
  const cleEnvoi = `ENVOI_PARRAINS_${semestre.label}`;

  if (props.getProperty(cleEnvoi)) {
    Logger.log(`ℹ️ Emails déjà envoyés pour ${semestre.label}`);
    try { SpreadsheetApp.getUi().alert(`ℹ️ Les emails ont déjà été envoyés pour ${semestre.label}.`); } catch(e) {}
    return;
  }

  const actifs       = getParrainsActifsIds();
  const lienCampagne = getLienCampagne();
  let nbEnvoyes      = 0;

  ss.getSheetByName(CONFIG.SHEET_PARRAIN).getDataRange().getValues().slice(1).forEach(row => {
    const id    = row[CONFIG.COLS.PARRAIN.SPONSOR_ID   - 1];
    const email = row[CONFIG.COLS.PARRAIN.ADRESSE_MAIL - 1];
    const nom   = row[CONFIG.COLS.PARRAIN.SPONSOR_NAME - 1];
    if (!id || !actifs.has(id) || !email) return;

    GmailApp.sendEmail(
    CONFIG.MODE_TEST ? CONFIG.EMAIL_TEST : email,
  `${CONFIG.MODE_TEST ? "TEST - " : ""} [Parrainages] Campagne ${semestre.label} ouverte`,
  "",
  { htmlBody: getTemplateAppelCotisation(nom, semestre, lienCampagne), name: CONFIG.ASSO_NOM }
  );
    nbEnvoyes++;
    pauseAntiSpam(nbEnvoyes);
  });

  if (!CONFIG.MODE_TEST) props.setProperty(cleEnvoi, new Date().toISOString());
  Logger.log(`✅ Emails parrainage envoyés à ${nbEnvoyes} parrain(s).`);
  try { SpreadsheetApp.getUi().alert(`✅ Emails envoyés à ${nbEnvoyes} parrain(s).`); } catch(e) {}
}



function envoyerRelanceIndividuelle(sponsorId, ligne) {
  const ss          = SpreadsheetApp.getActiveSpreadsheet();
  const semestre    = getSemestreEnCours();
  const sheetSuivi  = ss.getSheetByName(CONFIG.SHEET_SUIVI);
  const row         = sheetSuivi.getRange(ligne, 1, 1, sheetSuivi.getLastColumn()).getValues()[0];
  const nom         = row[CONFIG.COLS.SUIVI.SPONSOR_NAME        - 1];
  const email       = String(row[CONFIG.COLS.SUIVI.ADRESSE_MAIL - 1] || '').trim().toLowerCase();
  const dernierRappel = row[CONFIG.COLS.SUIVI.DATE_DERNIER_RAPPEL - 1];

  if (dernierRappel && joursDepuis(dernierRappel) < CONFIG.DELAI_RELANCE_JOURS) {
    Logger.log(`⚠️ Relance trop récente pour ${nom}`);
    return false;
  }

  GmailApp.sendEmail(
    CONFIG.MODE_TEST ? CONFIG.EMAIL_TEST : email,
    `${CONFIG.MODE_TEST ? "TEST - " : ""} [Rappel] Parrainage ${semestre.label} - ${CONFIG.ASSO_NOM}`,
    "",
    { htmlBody: getTemplateRelance(nom, semestre, getLienCampagne()), name: CONFIG.ASSO_NOM }
  );

  sheetSuivi.getRange(ligne, CONFIG.COLS.SUIVI.DATE_DERNIER_RAPPEL).setValue(new Date());
  
  // ✅ Synchroniser dans ScriptProperties pour le dashboard
  PropertiesService.getScriptProperties().setProperty(`RELANCE_PARRAIN_${email}`, new Date().toISOString());
  
  Logger.log(`✅ Relance envoyée à ${nom}`);
  return true;
}



function relancerTousLesEligibles() {
  const ss       = SpreadsheetApp.getActiveSpreadsheet();
  const semestre = getSemestreEnCours();
  const data     = ss.getSheetByName(CONFIG.SHEET_SUIVI).getDataRange().getValues();
  let nbRelances = 0;

  data.slice(1).forEach((row, i) => {
    if (row[CONFIG.COLS.SUIVI.SEMESTRE - 1] !== semestre.label) return;
    if (!row[CONFIG.COLS.SUIVI.STATUT  - 1]?.includes("⚠️"))   return;

    const dernierRappel = row[CONFIG.COLS.SUIVI.DATE_DERNIER_RAPPEL - 1];
    if (!dernierRappel || joursDepuis(dernierRappel) >= CONFIG.DELAI_RELANCE_JOURS) {
      envoyerRelanceIndividuelle(row[CONFIG.COLS.SUIVI.SPONSOR_ID - 1], i + 2);
      nbRelances++;
      pauseAntiSpam(nbRelances, 20);
    }
  });

  return { count: nbRelances };
}


function lancerRelancesTousEligibles() {
  const nb = relancerTousLesEligibles();
  SpreadsheetApp.getUi().alert(`✅ ${nb} relance(s) envoyée(s).`);
}




function joursDepuis(date) {
  return Math.floor((new Date() - new Date(date)) / 86400000);
}







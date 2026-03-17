
function getMembersList() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet()
    .getSheetByName(CONFIG.SHEET_MEMBRES);
  if (!sheet) return [];

  return sheet.getDataRange().getValues().slice(1)
    .map(row => {
      const prenom = row[CONFIG.COLS.MEMBRES.MEMBRE_NAME     - 1]?.toString().trim() || "";
      const nom    = row[CONFIG.COLS.MEMBRES.MEMBRE_FIRSTNAME  - 1]?.toString().trim() || "";
      return `${prenom} ${nom}`.trim();
    })
    .filter(Boolean)
    .sort();
}


function UpdateMembresName() {
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);

  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(CONFIG.SHEET_MEMBRES);
    if (!sheet) return;

    const colName      = CONFIG.COLS.MEMBRES.MEMBRE_NAME      - 1;
    const colFirstname = CONFIG.COLS.MEMBRES.MEMBRE_FIRSTNAME - 1;

    const data = sheet.getDataRange().getValues();

    let membresForm = data
      .slice(1)
      .map(r => {
        const nom    = r[colName]      ? r[colName].toString().trim()      : null;
        const prenom = r[colFirstname] ? r[colFirstname].toString().trim() : null;

        if (!nom || !prenom) {
          if (nom || prenom) Logger.log(`⚠️ Ligne ignorée — nom: "${nom}", prénom: "${prenom}"`);
          return null;
        }

        return `${nom} ${prenom}`;
      })
      .filter(Boolean);

    membresForm = dedupeKeepLast_(membresForm);
    membresForm = moveRecentFirst_(membresForm, 10);

    if (membresForm.length === 0) {
      Logger.log("⚠️ Aucun membre valide trouvé — mise à jour annulée");
      return;
    }

    const form = FormApp.openById(CONFIG.FORM_PAIEMENTS_MANUELS);
    updateDropdown_(form, "Nom du membre", membresForm);

    Logger.log(`✅ Dropdown membres mis à jour : ${membresForm.length} membres`);

  } catch (err) {
    Logger.log("❌ Erreur UpdateMembresName : " + err);
    throw err;
  } finally {
    lock.releaseLock();
  }
}


// ============================================================
// ENVOI D'EMAILS — COTISATIONS
// ============================================================

function envoyerAppelCotisationAuxMembres() {
  const ss       = SpreadsheetApp.getActiveSpreadsheet();
  const annee    = new Date().getFullYear().toString();
  const props    = PropertiesService.getScriptProperties();
  const cleEnvoi = `ENVOI_COTISATION_${annee}`;

  if (props.getProperty(cleEnvoi)) {
    Logger.log(`ℹ️ Cotisation déjà envoyée pour ${annee}`);
    try { SpreadsheetApp.getUi().alert(`ℹ️ Les emails ont déjà été envoyés pour ${annee}.`); } catch(e) {}
    return;
  }

  const lienCampagne = getLienCotisation();
  let nbEnvoyes = 0;

  ss.getSheetByName(CONFIG.SHEET_MEMBRES).getDataRange().getValues().slice(1).forEach(row => {
    const nom   = row[CONFIG.COLS.MEMBRES.MEMBRE_NAME  - 1];
    const email = row[CONFIG.COLS.MEMBRES.ADRESSE_MAIL - 1];
    if (!email) return;

    GmailApp.sendEmail(
    CONFIG.MODE_TEST ? CONFIG.EMAIL_TEST : email,
  `${CONFIG.MODE_TEST ? "TEST — " : ""} [Cotisation] Adhésion ${annee} — ${CONFIG.ASSO_NOM}`,
  "",
  { htmlBody: getTemplateAppelCotisationAnnuelle(nom, annee, lienCampagne), name: CONFIG.ASSO_NOM }
    );
    nbEnvoyes++;
    pauseAntiSpam(nbEnvoyes);
  });

  if (!CONFIG.MODE_TEST) props.setProperty(cleEnvoi, new Date().toISOString());
  Logger.log(`✅ Appel cotisation envoyé à ${nbEnvoyes} membre(s).`);
  try { SpreadsheetApp.getUi().alert(`✅ Appel cotisation envoyé à ${nbEnvoyes} membre(s).`); } catch(e) {}
}





function relancerMembreIndividuel(email) {
  const ss    = SpreadsheetApp.getActiveSpreadsheet();
  const annee = new Date().getFullYear().toString();
  const props = PropertiesService.getScriptProperties();

  const cle          = `RELANCE_MEMBRE_${email}_${annee}`;
  const dernierEnvoi = props.getProperty(cle);

  if (dernierEnvoi) {
    const jours = (new Date() - new Date(dernierEnvoi)) / (1000 * 60 * 60 * 24);
    if (jours < CONFIG.DELAI_RELANCE_MEMBRES_JOURS) {
      Logger.log(`⏭️ Relance trop récente pour ${email} (${Math.round(jours)}j)`);
      return `Relance ignorée - envoyée il y a ${Math.round(jours)} jours`;
    }
  }

  // Chercher le nom du membre
  const membres = getSheetRows(ss, CONFIG.SHEET_MEMBRES);
  const membre  = membres.find(r =>
    String(r[CONFIG.COLS.MEMBRES.ADRESSE_MAIL] || '').toLowerCase().trim() === email.toLowerCase().trim()
  );

  if (!membre) {
    Logger.log(`❌ Membre introuvable pour email : ${email}`);
    return "Membre introuvable";
  }

  const nom          = String(membre[CONFIG.COLS.MEMBRES.MEMBRE_NAME] || '').trim();
  const lienCotisation = getLienCotisation();

  GmailApp.sendEmail(
    CONFIG.MODE_TEST ? CONFIG.EMAIL_TEST : email,
    `${CONFIG.MODE_TEST ? "TEST - " : ""}[Rappel] Cotisation ${annee} — ${CONFIG.ASSO_NOM}`,
    "",
    { htmlBody: getTemplateRelanceMembre(nom, annee, lienCotisation), name: CONFIG.ASSO_NOM }
  );

  props.setProperty(cle, new Date().toISOString());
  Logger.log(`✅ Relance individuelle membre envoyée à ${nom} (${email})`);
  return `Relance envoyée à ${nom}`;
}


function relancerMembresEligibles() {
  const ss    = SpreadsheetApp.getActiveSpreadsheet();
  const annee = new Date().getFullYear().toString();
  const props = PropertiesService.getScriptProperties();

  const cotisationsPayees = new Set(
    getSheetRows(ss, CONFIG.SHEET_COTISATIONS_MERGED)
      .filter(r => {
        const match = String(r[CONFIG.COLS.COTISATIONS_MERGED.SOURCE] || '').match(/(\d{4})/);
        return match && match[1] === annee;
      })
      .map(r => String(r[CONFIG.COLS.COTISATIONS_MERGED.EMAIL] || '').toLowerCase().trim())
  );

  const membres = getSheetRows(ss, CONFIG.SHEET_MEMBRES);
  let nb = 0;

  membres.forEach(row => {
    const email = String(row[CONFIG.COLS.MEMBRES.ADRESSE_MAIL] || '').toLowerCase().trim();
    if (!email || cotisationsPayees.has(email)) return;

    const cle          = `RELANCE_MEMBRE_${email}_${annee}`;
    const dernierEnvoi = props.getProperty(cle);

    if (dernierEnvoi) {
      const jours = (new Date() - new Date(dernierEnvoi)) / (1000 * 60 * 60 * 24);
      if (jours < CONFIG.DELAI_RELANCE_MEMBRES_JOURS) return;
    }

    const result = relancerMembreIndividuel(email);
    if (result.startsWith('Relance envoyée')) nb++;
  });

  return { count: nb };
}

function pauseAntiSpam(compteur, seuil = 50) {
  if (compteur % seuil === 0) Utilities.sleep(1000);
}





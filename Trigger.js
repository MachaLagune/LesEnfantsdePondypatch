function processNewSubmission() {
  Utilities.sleep(2000);
  const newRow = checkAndMoveNewRow();
  if (!newRow) return;

  const sheetName = newRow.sheet.getName();

  switch(sheetName) {
    case CONFIG.SHEET_ENFANT:
      mainUpdateKids();
      break;
    case CONFIG.SHEET_PARRAIN:
      mainUpdateSponsors();
      break;
    case CONFIG.SHEET_PARRAINAGE:
      processSponsorshipSubmission();
      break;
    case CONFIG.SHEET_RDV:
      Logger.log('▶ Début processRdvSuivi');
      processRdvSuivi();
      Logger.log('✅ processRdvSuivi terminé');
      SpreadsheetApp.flush();
      Utilities.sleep(1000);
      Logger.log('▶ Début genererCRPourCetteLigne');
      genererCRPourCetteLigne();
      Logger.log('✅ genererCRPourCetteLigne terminé');
      break;
     case CONFIG.SHEET_FIN_PARRAINAGE:
      markParrainageFinished();
      break;
    case CONFIG.SHEET_VALIDATION_CR:
      envoi_CR_Parrain();
      break;
    case CONFIG.SHEET_MEMBRES:
      UpdateMembresName();
      break;
    case CONFIG.SHEET_PAIEMENTS_MANUELS:
      processPaiementManuel(newRow);
      break;
  }
}

// ============================================================
// POINT D'ENTRÉE QUOTIDIEN
// ============================================================

function MajOngletSuivi() {
  // Forcer un token frais à chaque exécution
  const props = PropertiesService.getScriptProperties();
  props.deleteProperty("HA_ACCESS_TOKEN");
  props.deleteProperty("HA_TOKEN_EXPIRY");

  _resetLogoCache();
  exportToSheets();
  mergeParrainages();
  exportCotisationToSheets();
  mergeCotisations();
  majTousLesSuivisEtParrainage();
}

// ============================================================
// DÉCLENCHEURS
// ============================================================

function declencheurValidationSemestrielle() {
  const mois     = new Date().getMonth();
  const annee    = new Date().getFullYear();
  const semestre = getSemestreEnCours();
  if (mois !== 0 && mois !== 6) return;

  const blocsHTML = [];

  blocsHTML.push(`
    <tr><td style="padding:0 40px 30px;">
      <div style="border:1px solid #ddd;border-radius:8px;overflow:hidden;">
        <div style="background:#3A7CA5;padding:15px 20px;">
          <p style="margin:0;color:#fff;font-weight:bold;font-size:15px;">📌 Campagne Parrainage</p>
        </div>
        <div style="padding:20px;">
          <table width="100%">
            <tr>
              <td width="40%" style="color:#666;font-size:14px;padding:8px 0;">Titre à saisir :</td>
              <td style="font-size:14px;padding:8px 0;">
                <strong style="background:#FFF8DC;padding:4px 10px;border-radius:4px;font-family:monospace;">
                  Parrainages S${semestre.numero} ${annee}
                </strong>
              </td>
            </tr>
            <tr>
              <td width="40%" style="color:#666;font-size:14px;padding:8px 0;">Slug généré :</td>
              <td style="font-size:14px;padding:8px 0;">
                <span style="background:#f0f0f0;padding:4px 10px;border-radius:4px;font-family:monospace;color:#555;">
                  parrainages-s${semestre.numero}-${annee}
                </span>
              </td>
            </tr>
            <tr>
              <td width="40%" style="color:#666;font-size:14px;padding:8px 0;">Concernés :</td>
              <td style="font-size:14px;padding:8px 0;"><strong>Tous les parrains actifs</strong></td>
            </tr>
          </table>
        </div>
      </div>
    </td></tr>`);

  if (mois === 0) {
    blocsHTML.push(`
    <tr><td style="padding:0 40px 30px;">
      <div style="border:1px solid #ddd;border-radius:8px;overflow:hidden;">
        <div style="background:#3A7CA5;padding:15px 20px;">
          <p style="margin:0;color:#fff;font-weight:bold;font-size:15px;">📌 Campagne Cotisation Annuelle</p>
        </div>
        <div style="padding:20px;">
          <table width="100%">
            <tr>
              <td width="40%" style="color:#666;font-size:14px;padding:8px 0;">Titre à saisir :</td>
              <td style="font-size:14px;padding:8px 0;">
                <strong style="background:#FFF8DC;padding:4px 10px;border-radius:4px;font-family:monospace;">
                  Adhésion ${annee}
                </strong>
              </td>
            </tr>
            <tr>
              <td width="40%" style="color:#666;font-size:14px;padding:8px 0;">Slug généré :</td>
              <td style="font-size:14px;padding:8px 0;">
                <span style="background:#f0f0f0;padding:4px 10px;border-radius:4px;font-family:monospace;color:#555;">
                  adhesion-${annee}
                </span>
              </td>
            </tr>
            <tr>
              <td width="40%" style="color:#666;font-size:14px;padding:8px 0;">Concernés :</td>
              <td style="font-size:14px;padding:8px 0;"><strong>Tous les membres actifs</strong></td>
            </tr>
            <tr>
              <td width="40%" style="color:#666;font-size:14px;padding:8px 0;">Montant minimum :</td>
              <td style="font-size:14px;padding:8px 0;"><strong>${CONFIG.MONTANT_MINIMUM_COTISATION}€ (participation libre)</strong></td>
            </tr>
          </table>
        </div>
      </div>
    </td></tr>`);
  }

  const logo     = getLogoBase64();
  const logoHtml = logo
    ? `<img src="data:${logo.mimeType};base64,${logo.base64}" alt="${CONFIG.ASSO_NOM}"
           style="max-height:50px;max-width:160px;margin-bottom:10px;display:block;margin-left:auto;margin-right:auto;">`
    : "";

  const htmlEmail = `<!DOCTYPE html>
<html lang="fr"><head><meta charset="UTF-8"></head>
<body style="margin:0;padding:0;background:#f4f4f4;font-family:Arial,sans-serif;">
  <table width="100%" cellpadding="0" cellspacing="0" style="background:#f4f4f4;padding:30px 0;">
    <tr><td align="center">
      <table width="640" cellpadding="0" cellspacing="0" style="background:#fff;border-radius:8px;overflow:hidden;box-shadow:0 2px 8px rgba(0,0,0,0.1);">
        <tr>
          <td style="background:#3A7CA5;padding:30px;text-align:center;">
            ${logoHtml}
            <h1 style="color:#fff;margin:0;font-size:24px;">${CONFIG.ASSO_NOM}</h1>
            <p style="color:#FFD966;margin:8px 0 0;font-size:14px;">[Action requise] Créer les campagnes – ${semestre.label}</p>
          </td>
        </tr>
        <tr>
          <td style="padding:30px 40px 20px;">
            <p style="color:#333;font-size:15px;line-height:1.7;margin:0 0 10px;">Bonjour,</p>
            <p style="color:#555;font-size:15px;line-height:1.7;margin:0 0 20px;">
              Le semestre <strong>${semestre.label}</strong> vient de débuter. Merci de créer
              ${mois === 0 ? "les deux campagnes suivantes" : "la campagne suivante"} sur HelloAsso
              en respectant bien les titres indiqués.
            </p>
          </td>
        </tr>
        ${blocsHTML.join("\n")}
        <tr>
          <td style="padding:0 40px 30px;">
            <div style="background:#E8F4FD;border-left:4px solid #3A7CA5;padding:15px 20px;border-radius:0 6px 6px 0;">
              <p style="margin:0;color:#555;font-size:14px;line-height:1.7;">
                ✅ <strong>Aucune autre action requise.</strong> Une fois la campagne créée,
                les emails partent automatiquement. Vous recevrez une confirmation.
              </p>
            </div>
          </td>
        </tr>
        <tr>
          <td style="background:#3A7CA5;padding:20px 40px;text-align:center;">
            <p style="color:#fff;font-size:13px;margin:0;">${CONFIG.ASSO_NOM}</p>
            <p style="margin:5px 0 0;">
              <a href="mailto:${CONFIG.ASSO_EMAIL}" style="color:#FFD966;font-size:13px;text-decoration:none;">
                ${CONFIG.ASSO_EMAIL}
              </a>
            </p>
          </td>
        </tr>
      </table>
    </td></tr>
  </table>
</body></html>`;

  const destinataires = CONFIG.MODE_TEST ? [CONFIG.EMAIL_TEST] : CONFIG.EMAILS_DIRECTION;
  destinataires.forEach(email =>
    GmailApp.sendEmail(email,
      `${CONFIG.MODE_TEST ? "TEST — " : ""}[Action requise] Créer les campagnes HelloAsso – ${semestre.label}`,
      "",
      { htmlBody: htmlEmail, name: CONFIG.ASSO_NOM }
    )
  );
  Logger.log(`✅ Rappel semestriel envoyé pour ${semestre.label}`);
}





// ============================================================
// Boutons d'action manuelle sur le sheet Tableau_Complet
// ============================================================


function onOpen() {
  const ui = SpreadsheetApp.getUi();

  ui.createMenu("📋 CR Enfants")
    .addItem("Générer les CR manquants", "genererCRPourCetteLigne")
    .addItem("📨 Envoyer le CR de cette ligne", "ouvrirSidebarEnvoi")
    .addToUi();

  ui.createMenu("🤖 Parrainages")
    .addItem("🔄Synchroniser maintenant", "MajOngletSuivi")
    .addSeparator()
    .addItem("🔍 Forcer recherche campagne", "forcerRechercheEtExport")
    .addSeparator()
    .addItem("📧 Envoyer appel cotisation parrains", "envoyerLienPaiementAuxParrains")
    .addItem("📧 Envoyer appel cotisation membres", "envoyerAppelCotisationAuxMembres")
    .addSeparator()
    .addItem("🔔 Relancer tous les parrains éligibles", "relancerTousLesEligibles")
    .addItem("🔔 Relancer les membres en retard", "relancerMembresEligibles")
    .addSeparator()
    .addItem("✉️ Tester les templates emails", "testerTemplatesEmails")
    .addToUi();
}

// ============================================================
// Mise à jour onglets diagnostic et non reconcilés
// ============================================================
function majOngletDiagnostic() {

  // 1️⃣ Non-réconciliés + Diagnostic (tout en un)
  remplirDiagnostic();

}



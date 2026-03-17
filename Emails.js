// ============================================================
// LOGO — CHARGEMENT DEPUIS DRIVE AVEC CACHE
// ============================================================

let _logoCache = null;

function getLogoBase64() {
  if (_logoCache) return _logoCache;
  try {
    const file = DriveApp.getFileById(CONFIG.LOGO_ID);
    const blob = file.getBlob();
    _logoCache = {
      base64:   Utilities.base64Encode(blob.getBytes()),
      mimeType: blob.getContentType()
    };
    Logger.log("✅ Logo chargé depuis Drive.");
    return _logoCache;
  } catch (err) {
    Logger.log("⚠️ Logo inaccessible : " + err.message);
    return null;
  }
}
function _resetLogoCache() { _logoCache = null; }


// ============================================================
// TEMPLATES EMAILS HTML
// ============================================================

function _buildEmailHtml({ sousTitre, corps, lienCampagne, boutonLabel, note }) {
  const logo     = getLogoBase64();
  const logoHtml = logo
    ? `<img src="data:${logo.mimeType};base64,${logo.base64}"
           alt="${CONFIG.ASSO_NOM}"
           style="max-height:60px;max-width:200px;margin-bottom:12px;display:block;margin-left:auto;margin-right:auto;">`
    : "";

  return `<!DOCTYPE html>
<html lang="fr"><head><meta charset="UTF-8"></head>
<body style="margin:0;padding:0;background:#f4f4f4;font-family:Arial,sans-serif;">
  <table width="100%" cellpadding="0" cellspacing="0" style="background:#f4f4f4;padding:30px 0;">
    <tr><td align="center">
      <table width="600" cellpadding="0" cellspacing="0" style="background:#fff;border-radius:8px;overflow:hidden;box-shadow:0 2px 8px rgba(0,0,0,0.1);">
        <tr>
          <td style="background:#3A7CA5;padding:30px;text-align:center;">
            ${logoHtml}
            <h1 style="color:#fff;margin:0;font-size:24px;letter-spacing:1px;">${CONFIG.ASSO_NOM}</h1>
            <p style="color:#FFD966;margin:8px 0 0;font-size:14px;">${sousTitre}</p>
          </td>
        </tr>
        <tr>
          <td style="padding:35px 40px;">
            ${corps}
            <table width="100%" cellpadding="0" cellspacing="0">
              <tr><td align="center" style="padding:10px 0 30px;">
                <a href="${lienCampagne}"
                   style="background:#FFD966;color:#3A7CA5;padding:14px 35px;border-radius:25px;text-decoration:none;font-weight:bold;font-size:15px;display:inline-block;">
                  ${boutonLabel}
                </a>
              </td></tr>
            </table>
            <p style="color:#888;font-size:13px;font-style:italic;margin:0;">${note}</p>
          </td>
        </tr>
        <tr>
          <td style="background:#3A7CA5;padding:20px 40px;text-align:center;">
            <p style="color:#fff;font-size:13px;margin:0;">${CONFIG.ASSO_NOM}</p>
            <p style="margin:5px 0 0;">
              <a href="mailto:${CONFIG.ASSO_EMAIL}"
                 style="color:#FFD966;font-size:13px;text-decoration:none;">${CONFIG.ASSO_EMAIL}</a>
            </p>
          </td>
        </tr>
      </table>
    </td></tr>
  </table>
</body></html>`;
}

function getTemplateAppelCotisation(sponsorName, semestre, lienCampagne) {
  return _buildEmailHtml({
    sousTitre: `Appel de cotisation - ${semestre.label}`,
    lienCampagne, boutonLabel: "Régler ma cotisation",
    note: "Si vous avez déjà effectué votre règlement, merci d'ignorer ce message. Votre statut sera mis à jour automatiquement.",
    corps: `
      <p style="color:#333;font-size:16px;margin:0 0 15px;">Cher parrain, chère marraine,</p>
      <p style="color:#555;font-size:15px;line-height:1.7;margin:0 0 15px;">Nous sommes ravis de vous retrouver pour ce nouveau semestre et vous remercions chaleureusement pour votre fidélité et votre générosité.</p>
      <p style="color:#555;font-size:15px;line-height:1.7;margin:0 0 15px;">Votre parrainage est essentiel : il soutient directement un enfant pour sa scolarité, l'achat de fournitures et de vêtements.</p>
      <p style="color:#555;font-size:15px;line-height:1.7;margin:0 0 25px;">Pour renouveler votre parrainage pour le <strong>${semestre.label}</strong>, cliquez ci-dessous :</p>`
  });
}

function getTemplateRelance(sponsorName, semestre, lienCampagne) {
  return _buildEmailHtml({
    sousTitre: `Rappel de cotisation - ${semestre.label}`,
    lienCampagne, boutonLabel: "Régler ma cotisation",
    note: "Si vous avez déjà effectué votre règlement, merci d'ignorer ce message. Votre statut sera mis à jour automatiquement sous 24h.",
    corps: `
      <p style="color:#333;font-size:16px;margin:0 0 15px;">Cher parrain, chère marraine,</p>
      <p style="color:#555;font-size:15px;line-height:1.7;margin:0 0 15px;">Nous espérons que vous allez bien ! Nous vous contactons car nous n'avons pas encore reçu votre cotisation pour le <strong>${semestre.label}</strong>.</p>
      <p style="color:#555;font-size:15px;line-height:1.7;margin:0 0 15px;">Votre soutien est précieux : chaque cotisation contribue directement à la scolarité et aux fournitures d'un enfant que vous parrainez.</p>
      <p style="color:#555;font-size:15px;line-height:1.7;margin:0 0 25px;">Si vous souhaitez continuer l'aventure, réglez votre cotisation en cliquant ci-dessous :</p>`
  });
}

function getTemplateAppelCotisationAnnuelle(membreName, annee, lienCampagne) {
  return _buildEmailHtml({
    sousTitre: `Cotisation annuelle ${annee}`,
    lienCampagne, boutonLabel: "Régler ma cotisation",
    note: `Participation libre à partir de ${CONFIG.MONTANT_MINIMUM_COTISATION}€. Si vous avez déjà réglé, merci d'ignorer ce message.`,
    corps: `
      <p style="color:#333;font-size:16px;margin:0 0 15px;">Cher(e) membre,</p>
      <p style="color:#555;font-size:15px;line-height:1.7;margin:0 0 15px;">Nous vous souhaitons une belle année ${annee} et vous remercions pour votre soutien continu à l'association.</p>
      <p style="color:#555;font-size:15px;line-height:1.7;margin:0 0 15px;">Il est temps de renouveler votre adhésion pour ${annee}. Votre cotisation, libre à partir de <strong>${CONFIG.MONTANT_MINIMUM_COTISATION}€</strong>, contribue directement au fonctionnement de l'association.</p>
      <p style="color:#555;font-size:15px;line-height:1.7;margin:0 0 25px;">Pour régler votre cotisation, cliquez ci-dessous :</p>`
  });
}






function getTemplateRelanceMembre(membreName, annee, lienCampagne) {
  return _buildEmailHtml({
    sousTitre: `Rappel cotisation annuelle ${annee}`,
    lienCampagne,
    boutonLabel: "Régler ma cotisation",
    note: `Participation libre à partir de ${CONFIG.MONTANT_MINIMUM_COTISATION}€. Si vous avez déjà réglé, merci d'ignorer ce message.`,
    corps: `
      <p style="color:#333;font-size:16px;margin:0 0 15px;">Cher(e) membre,</p>
      <p style="color:#555;font-size:15px;line-height:1.7;margin:0 0 15px;">
        Nous espérons que vous allez bien ! Nous vous contactons car nous n'avons pas encore reçu votre cotisation pour <strong>${annee}</strong>.
      </p>
      <p style="color:#555;font-size:15px;line-height:1.7;margin:0 0 15px;">
        Votre adhésion contribue directement au fonctionnement de l'association et au soutien des enfants que nous accompagnons.
      </p>
      <p style="color:#555;font-size:15px;line-height:1.7;margin:0 0 25px;">
        Pour régler votre cotisation, cliquez ci-dessous :
      </p>`
  });
}







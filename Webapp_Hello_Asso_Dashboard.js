// ============================================================
// Code.gs — Routeur principal (doGet + doPost)
// NE CONTIENT AUCUNE LOGIQUE MÉTIER — uniquement du routage
// ============================================================

function doGet(e) {
  const { action, token, sponsor, ligne,email } = e.parameter;
  const props = PropertiesService.getScriptProperties();

  const page = (titre, couleur, message, avecBouton = false) =>
    HtmlService.createHtmlOutput(`
      <html><body style="background:#f4f4f4;font-family:Arial;padding:50px;text-align:center;">
        <div style="background:white;padding:40px;border-radius:8px;max-width:500px;margin:auto;">
          <h1 style="color:${couleur};">${titre}</h1><p>${message}</p>
          ${avecBouton ? `<button onclick="window.close()" style="background:#5A8BA8;color:white;padding:10px 20px;border:none;border-radius:4px;cursor:pointer;">Fermer</button>` : ""}
        </div>
      </body></html>`);

  try {
    switch (action) {

      case "accueil":
      case undefined:
        return HtmlService.createHtmlOutputFromFile('accueil')
          .setTitle('Les Enfants de Pondypatch')
          .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);

      case "dashboard_page":
        return HtmlService.createHtmlOutputFromFile('dashboard')
          .setTitle('Dashboard — Les Enfants de Pondypatch')
          .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);

      case "guide":
        return HtmlService.createHtmlOutputFromFile('guide_installation')
          .setTitle('Guide d\'installation — Les Enfants de Pondypatch')
          .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);

      case "timeline":
        return HtmlService.createHtmlOutputFromFile('timeline')
          .setTitle('Timeline — Les Enfants de Pondypatch')
          .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);

      case "formulaires":
        return HtmlService.createHtmlOutputFromFile('formulaires')
          .setTitle('Formulaires — Les Enfants de Pondypatch')
          .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);

      case "process":
        return HtmlService.createHtmlOutputFromFile('process')
          .setTitle('Process Parrainage — Les Enfants de Pondypatch')
          .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);

      case "relancer": {
        const tokenAttendu = props.getProperty(`TOKEN_RELANCE_${sponsor}`);
        if (token !== tokenAttendu)
          return page("❌ Lien invalide ou expiré", "#E74C3C", "Ce lien n'est plus valide.");
        props.setProperty(`TOKEN_RELANCE_${sponsor}`, Utilities.getUuid());
        const ok = envoyerRelanceIndividuelle(sponsor, ligne);
        return ok
          ? page("✅ Relance envoyée !", "#2ECC71", "L'email a été envoyé au parrain.", true)
          : page("⚠️ Relance non envoyée", "#FFA500", "Ce parrain a déjà été relancé récemment.", true);
      }

      case "relancer_tous": {
        const nb = relancerTousLesEligibles();
        return page("✅ Relances envoyées !", "#2ECC71", `${nb} email(s) envoyés.`, true);
      }

      case "relancer_membre": {
        const email = (e.parameter.email || '').toLowerCase().trim(); // ← ajouter
        const token = e.parameter.token;
        Logger.log(`🔍 email reçu : "${email}"`);
        Logger.log(`🔍 token reçu : "${token}"`);
        const tokenAttendu = props.getProperty(`TOKEN_RELANCE_MEMBRE_${email}`);
        Logger.log(`🔍 token attendu : "${tokenAttendu}"`);
        if (token !== tokenAttendu)
          return page("❌ Lien invalide ou expiré", "#E74C3C", "Ce lien n'est plus valide.");
        props.setProperty(`TOKEN_RELANCE_MEMBRE_${email}`, Utilities.getUuid());
        const ok = relancerMembreIndividuel(email);
        return ok
          ? page("✅ Relance envoyée !", "#2ECC71", "L'email a été envoyé au membre.", true)
          : page("⚠️ Relance non envoyée", "#FFA500", "Ce membre a déjà été relancé récemment.", true);
      }

      case "relancer_tous_membres": {
        if (props.getProperty("TOKEN_RELANCE_MEMBRES_TOUS") !== token)
          return page("❌ Lien invalide ou expiré", "#E74C3C", "Ce lien n'est plus valide.");
        props.setProperty("TOKEN_RELANCE_MEMBRES_TOUS", Utilities.getUuid());
        const nb = relancerMembresEligibles();
        return page("✅ Relances envoyées !", "#2ECC71", `${nb} membre(s) relancé(s).`, true);
      }

      default:
        return page("❌ Action non reconnue", "#E74C3C", "");  // ← default en dernier
    }

  } catch (err) {
    Logger.log("Erreur doGet: " + err);
    return page("❌ Une erreur est survenue", "#E74C3C", err.toString());
  }
}



function doPost(e) {
  // → HelloAsso.gs — inchangé
  return traiterWebhookHelloAsso(e);
}
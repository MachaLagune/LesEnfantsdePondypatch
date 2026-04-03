// ============================================================
// AUTHENTIFICATION HELLOASSO
// ============================================================

function getAccessToken() {
  const props       = PropertiesService.getScriptProperties();
  const tokenCache  = props.getProperty("HA_ACCESS_TOKEN");
  const tokenExpiry = props.getProperty("HA_TOKEN_EXPIRY");

  if (tokenCache && tokenExpiry && new Date().getTime() < parseInt(tokenExpiry) - 60000) {
    Logger.log("ℹ️ Token réutilisé depuis le cache.");
    return tokenCache;
  }

  const maxRetries = 3;
  let lastError = null;

  for (let attempt = 1; attempt <= maxRetries; attempt++) {
    const response = UrlFetchApp.fetch("https://api.helloasso.com/oauth2/token", {
      method: "post",
      contentType: "application/x-www-form-urlencoded",
      payload: {
        grant_type:    "client_credentials",
        client_id:     CONFIG.CLIENT_ID,
        client_secret: CONFIG.CLIENT_SECRET
      },
      muteHttpExceptions: true
    });

    const statusCode = response.getResponseCode();
    const rawText    = response.getContentText();

    // Vérifier que c'est bien du JSON avant de parser
    if (statusCode === 429 || rawText.includes("error code:")) {
      lastError = `HTTP ${statusCode} — Cloudflare/rate limit: ${rawText}`;
      Logger.log(`⚠️ Tentative ${attempt}/${maxRetries} échouée : ${lastError}`);
      if (attempt < maxRetries) Utilities.sleep(attempt * 3000); // 3s, 6s
      continue;
    }

    if (statusCode !== 200) {
      throw new Error(`❌ Erreur HTTP ${statusCode} : ${rawText}`);
    }

    let data;
    try {
      data = JSON.parse(rawText);
    } catch (e) {
      throw new Error(`❌ Réponse non-JSON (${statusCode}) : ${rawText}`);
    }

    if (!data.access_token) {
      throw new Error(`❌ Token absent dans la réponse : ${rawText}`);
    }

    const expiry = new Date().getTime() + (data.expires_in || 1800) * 1000;
    props.setProperty("HA_ACCESS_TOKEN", data.access_token);
    props.setProperty("HA_TOKEN_EXPIRY", expiry.toString());

    Logger.log("✅ Nouveau token HelloAsso obtenu.");
    return data.access_token;
  }

  throw new Error(`❌ Impossible d'obtenir un token après ${maxRetries} tentatives. Dernière erreur : ${lastError}`);
}


// ============================================================
// EXPORT HELLOASSO → GOOGLE SHEETS
// ============================================================

function getSlugsCampagneSemestre(token) {
  const semestre     = getSemestreEnCours();
  const annee        = semestre.label.split(" ")[1];
  const num          = semestre.numero.toString();
  const slugsTrouves = [];
  let pageIndex = 1, continuer = true;

  while (continuer) {
    const url = `https://api.helloasso.com/v5/organizations/${CONFIG.ORG_SLUG}/forms`
              + `?pageIndex=${pageIndex}&pageSize=100`;

    const response = UrlFetchApp.fetch(url, {
      method: "get",
      headers: { Authorization: "Bearer " + token },
      muteHttpExceptions: true
    });

    const code = response.getResponseCode();
    if (code !== 200) { Logger.log(`⚠️ Erreur HTTP ${code}`); break; }

    const data = JSON.parse(response.getContentText());
    if (!data.data || data.data.length === 0) { continuer = false; break; }

    data.data.forEach(form => {
      const slug = (form.formSlug || "").toLowerCase();
      const estParrainage = slug.includes("parrainage")
        && slug.includes(annee)
        && (slug.includes(`s${num}`) || slug.includes(`s-${num}`) || slug.includes(`semestre-${num}`));
      if (estParrainage) {
        slugsTrouves.push({ slug, url: form.url || "" });
        Logger.log(`✅ Campagne parrainage trouvée : ${slug}`);
      }
    });

    const pagination = data.pagination;
    if (!pagination || pageIndex >= pagination.totalPages) continuer = false;
    else { pageIndex++; Utilities.sleep(200); }
  }

  if (slugsTrouves.length === 0)
    Logger.log(`⚠️ Aucune campagne parrainage trouvée pour ${semestre.label}`);

  return slugsTrouves;
}

function getSlugCotisationAnnee(token) {
  const annee    = new Date().getFullYear().toString();
  const url      = `https://api.helloasso.com/v5/organizations/${CONFIG.ORG_SLUG}/forms?pageSize=100`;
  const response = UrlFetchApp.fetch(url, {
    method: "get",
    headers: { Authorization: "Bearer " + token },
    muteHttpExceptions: true
  });

  if (response.getResponseCode() !== 200) return null;
  const data = JSON.parse(response.getContentText());
  if (!data.data) return null;

  const form = data.data.find(f => {
    const slug = (f.formSlug || "").toLowerCase();
    return (slug.includes("adhesion") || slug.includes("cotisation")) && slug.includes(annee);
  });

  if (form) {
    Logger.log(`✅ Campagne cotisation trouvée : ${form.formSlug}`);
    return { slug: form.formSlug, url: form.url || "" };
  }

  Logger.log(`⚠️ Aucune campagne cotisation trouvée pour ${annee}`);
  return null;
}

function _exportSlugToSheet(slug, nomOnglet, token) {
  const ss    = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(nomOnglet) || ss.insertSheet(nomOnglet);

  if (sheet.getLastRow() === 0)
    sheet.appendRow(["Date", "Prénom", "Nom", "Email", "Montant (€)", "Statut", "Slug"]);

  // Indexer les paiements déjà présents pour éviter les doublons
  // Clé : email + jour (10 premiers caractères de la date) + montant
  const existants = new Set();
  sheet.getDataRange().getValues().slice(1).forEach(row => {
    const email   = row[CONFIG.COLS.HELLO_ASSO.EMAIL   - 1].toString().toLowerCase().trim();
    const date    = row[CONFIG.COLS.HELLO_ASSO.DATE    - 1].toString().substring(0, 10);
    const montant = parseFloat(row[CONFIG.COLS.HELLO_ASSO.MONTANT - 1]).toFixed(0);
    existants.add(`${email}_${date}_${montant}`);
  });

  Logger.log(`🔍 ${existants.size} paiement(s) déjà présent(s) dans "${nomOnglet}"`);

  let pageIndex = 1, nbAjoutes = 0, continuer = true;

  while (continuer) {
    const url = `https://api.helloasso.com/v5/organizations/${CONFIG.ORG_SLUG}`
              + `/forms/${CONFIG.TYPE_FORMULAIRE}/${slug}/payments`
              + `?pageIndex=${pageIndex}&pageSize=100`;

    const response = UrlFetchApp.fetch(url, {
      method: "get",
      headers: { Authorization: "Bearer " + token },
      muteHttpExceptions: true
    });

    const code = response.getResponseCode();
    if (code !== 200) { Logger.log(`⚠️ Erreur HTTP ${code} pour ${slug}`); break; }

    const data = JSON.parse(response.getContentText());
    if (!data.data || data.data.length === 0) { continuer = false; break; }

    data.data.forEach(p => {
      const email   = (p.payer?.email || "").toLowerCase().trim();
      const jour    = (p.date || "").toString().substring(0, 10);
      const montant = (p.amount / 100).toFixed(0);
      const cle     = `${email}_${jour}_${montant}`;

      if (!existants.has(cle)) {
        sheet.appendRow([
          p.date,
          p.payer?.firstName || "",
          p.payer?.lastName  || "",
          p.payer?.email     || "",
          montant,
          p.state,
          slug
        ]);
        existants.add(cle);
        nbAjoutes++;
        Logger.log(`➕ Nouveau paiement ajouté : ${email} | ${jour} | ${montant}€`);
      } else {
        Logger.log(`⏭️ Doublon ignoré : ${email} | ${jour} | ${montant}€`);
      }
    });

    const pagination = data.pagination;
    if (!pagination || pageIndex >= pagination.totalPages) continuer = false;
    else { pageIndex++; Utilities.sleep(200); }
  }

  Logger.log(`✅ ${nbAjoutes} nouveau(x) paiement(s) ajouté(s) dans "${nomOnglet}" pour ${slug}`);
}

function exportToSheets() {
  const token     = getAccessToken();
  const semestre  = getSemestreEnCours();
  const campagnes = getSlugsCampagneSemestre(token);

  if (campagnes.length === 0) {
    Logger.log("⚠️ Aucune campagne parrainage à exporter.");
    return;
  }

  // Quel onglet selon le semestre en cours
  const nomOnglet = semestre.numero === 1
    ? CONFIG.SHEET_HELLO_ASSO_1   // → "helloasso-s1"
    : CONFIG.SHEET_HELLO_ASSO_2;  // → "helloasso-s2"

  campagnes.forEach(c => {
    mettreAJourConfiguration(c.slug, c.url);
    _exportSlugToSheet(c.slug, nomOnglet, token);
  });
}

function exportCotisationToSheets() {
  const token    = getAccessToken();
  const campagne = getSlugCotisationAnnee(token);

  if (!campagne) {
    Logger.log("⚠️ Aucune campagne cotisation à exporter.");
    return;
  }

  PropertiesService.getScriptProperties()
    .setProperty("LIEN_HELLOASSO_COTISATION", campagne.url);

  // Toujours dans l'onglet "helloasso-cotisation"
  _exportSlugToSheet(campagne.slug, CONFIG.SHEET_HELLO_ASSO_COTISATION, token);
}


// ============================================================
// WEBHOOK HELLOASSO (doPost)
// ============================================================

function traiterWebhookHelloAsso(e) {
  try {
    const payload = JSON.parse(e.postData.contents);
    Logger.log("📨 Webhook reçu : " + JSON.stringify(payload));

    switch (payload.eventType) {

      case "Form": {
        const slug  = (payload.data?.formSlug || "").toLowerCase();
        const annee = new Date().getFullYear();
        const sem   = getSemestreEnCours();
        const num   = sem.numero.toString();

        const estParrainage = slug.includes("parrainage")
          && slug.includes(annee)
          && (slug.includes(`s${num}`) || slug.includes(`s-${num}`) || slug.includes(`semestre-${num}`));
        const estCotisation = (slug.includes("adhesion") || slug.includes("cotisation"))
          && slug.includes(annee);

        Logger.log(`🔍 estParrainage : ${estParrainage}`);
        Logger.log(`🔍 estCotisation : ${estCotisation}`);

        if (estParrainage)      _traiterWebhookParrainage(slug, payload,false);
        else if (estCotisation) _traiterWebhookCotisation(slug, payload,false);
        else Logger.log(`ℹ️ Slug "${slug}" non reconnu`);
        break;
      }

      case "Order": {
        if (payload.data?.state !== "Authorized") break;

        const formSlug = (payload.data?.order?.formSlug || payload.data?.formSlug || "").toLowerCase();
        Logger.log(`🔍 formSlug détecté : "${formSlug}"`);

        const annee    = new Date().getFullYear().toString();
        Logger.log(`🔍 annee : "${annee}"`);

        const sem      = getSemestreEnCours();
        const num      = sem.numero.toString();

        const estParrainage = formSlug.includes("parrainage")
          && formSlug.includes(annee)
          && (formSlug.includes(`s${num}`) || formSlug.includes(`s-${num}`) || formSlug.includes(`semestre-${num}`));
        const estCotisation = (formSlug.includes("adhesion") || formSlug.includes("cotisation"))
          && formSlug.includes(annee);
        Logger.log(`🔍 estCotisation : ${estCotisation}`);

        if (estParrainage)      _traiterPaiementParrainage(payload);
        else if (estCotisation) _traiterPaiementCotisation(payload);
        break;
      }

      default:
        Logger.log(`ℹ️ Événement ignoré : ${payload.eventType}`);
    }

    return ContentService.createTextOutput("ok");

  } catch (err) {
    Logger.log("❌ Erreur doPost : " + err.toString());
    return ContentService.createTextOutput("error: " + err.toString());
  }
}

function _traiterWebhookParrainage(slug, payload, modeTest = true) {
  const semestre = getSemestreEnCours();
  const props    = PropertiesService.getScriptProperties();

  if (!modeTest && props.getProperty(`ENVOI_PARRAINS_${semestre.label}`)) return;

  mettreAJourConfiguration(slug, payload.data?.url || "","LIEN_HELLOASSO_ACTUEL");

  if (modeTest) {
    GmailApp.sendEmail(
      CONFIG.MODE_TEST ? CONFIG.EMAIL_TEST : CONFIG.EMAIL_DIRIGEANT,
      `TEST - Campagne parrainage ${semestre.label}`,
      `Ceci est un test.\nLe lien détecté est : ${payload.data?.url}`
    );
  } else {
    envoyerLienPaiementAuxParrains();

    GmailApp.sendEmail(
      CONFIG.MODE_TEST ? CONFIG.EMAIL_TEST : CONFIG.EMAIL_DIRIGEANT,
      `✅ Campagne parrainage ${semestre.label} détectée — emails envoyés`,
      `La campagne "${slug}" a été détectée.\nEmails envoyés aux parrains actifs.\n\n${CONFIG.ASSO_NOM} 🤖`
    );
  }
}

function _traiterWebhookCotisation(slug, payload, modeTest = true) {
  const annee = new Date().getFullYear().toString();
  const props = PropertiesService.getScriptProperties();

  if (!modeTest && props.getProperty(`ENVOI_COTISATION_${annee}`)) return;

  props.setProperty("LIEN_HELLOASSO_COTISATION", payload.data?.url || "");

  if (modeTest) {
    GmailApp.sendEmail(
      CONFIG.MODE_TEST ? CONFIG.EMAIL_TEST : CONFIG.EMAIL_DIRIGEANT,
      `TEST - Campagne cotisation ${annee}`,
      `Ceci est un test.\nLien détecté : ${payload.data?.url}`
    );
  } else {
    // PROD : envoi réel aux membres
    envoyerAppelCotisationAuxMembres();

    GmailApp.sendEmail(
      CONFIG.MODE_TEST ? CONFIG.EMAIL_TEST : CONFIG.EMAIL_DIRIGEANT,
      `✅ Campagne cotisation ${annee} détectée — emails envoyés`,
      `La campagne "${slug}" a été détectée.\nEmails envoyés à tous les membres.\n\n${CONFIG.ASSO_NOM} 🤖`
    );
  }
}

function _traiterPaiementParrainage(payload) {
  const email   = (payload.data?.payer?.email || "").toLowerCase().trim();
  const montant = ((payload.data?.amount || payload.data?.items?.[0]?.amount || 0)) / 100;
  const date    = new Date(payload.data?.date || new Date());
  const slug    = (payload.data?.order?.formSlug || payload.data?.formSlug || "").toLowerCase();
  if (!email || montant === 0) return;

  const semestre  = getSemestreEnCours();
  const nomOnglet = semestre.numero === 1 ? CONFIG.SHEET_HELLO_ASSO_1 : CONFIG.SHEET_HELLO_ASSO_2;
  _ajouterOuMettreAJourPaiementSheet(nomOnglet, email, montant, date,
    payload.data?.payer?.firstName || "",
    payload.data?.payer?.lastName  || "",
    slug
  );
  mergeParrainages();
}

function _traiterPaiementCotisation(payload) {
  Logger.log("🔍 _traiterPaiementCotisation appelée");
  const email   = (payload.data?.payer?.email || "").toLowerCase().trim();
  const montant = ((payload.data?.amount || payload.data?.items?.[0]?.amount || 0)) / 100;
  const date    = new Date(payload.data?.date || new Date());
  const slug    = (payload.data?.order?.formSlug || payload.data?.formSlug || "").toLowerCase();
  if (!email || montant === 0) return;

  _ajouterOuMettreAJourPaiementSheet(
    CONFIG.SHEET_HELLO_ASSO_COTISATION,
    email,
    montant,
    date,
    payload.data?.payer?.firstName || "",
    payload.data?.payer?.lastName  || "",
    slug
  );

}
function _ajouterOuMettreAJourPaiementSheet(nomOnglet, email, montant, date, prenom, nom,slug) {
  const ss    = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(nomOnglet) || ss.insertSheet(nomOnglet);
  const data  = sheet.getDataRange().getValues();
  const c     = CONFIG.COLS.HELLO_ASSO;

  let ligneExistante = null;
  data.slice(1).forEach((row, i) => {
    if (row[c.EMAIL - 1].toString().toLowerCase().trim() === email)
      ligneExistante = i + 2;
  });

  if (ligneExistante) {
    sheet.getRange(ligneExistante, c.DATE).setValue(date);
    sheet.getRange(ligneExistante, c.MONTANT).setValue(montant.toFixed(0));
  } else {
    const ligne = ["", "", "", "", "", "Authorized", ""];
    ligne[c.DATE    - 1] = date;
    ligne[c.PRENOM  - 1] = prenom;
    ligne[c.NOM     - 1] = nom;
    ligne[c.EMAIL   - 1] = email;
    ligne[c.MONTANT - 1] = montant.toFixed(0);
    ligne[c.SLUG    - 1] = slug;
    sheet.appendRow(ligne);
  }
}




function getLienCampagne() {
  const lien = PropertiesService.getScriptProperties().getProperty("LIEN_HELLOASSO_ACTUEL")
    || CONFIG.LIEN_HELLOASSO_ACTUEL;
  Logger.log("🔗 Lien campagne : " + lien);
  return lien;
}

function getLienCotisation() {
  return PropertiesService.getScriptProperties().getProperty("LIEN_HELLOASSO_COTISATION")
    || CONFIG.LIEN_HELLOASSO_ACTUEL;
}

function mettreAJourConfiguration(nouveauSlug, nouveauLien, cle = "LIEN_HELLOASSO_ACTUEL") {
  PropertiesService.getScriptProperties().setProperty(cle, nouveauLien);
  Logger.log(`✅ Configuration mise à jour : ${nouveauSlug} — ${nouveauLien}`);
}


// ════════════════════════════════════════════════════
//  UTILITAIRE : lire un onglet → tableau d'objets
// ════════════════════════════════════════════════════

function getSheetRowsByName(ss, sheetName) {
  const sheet = ss.getSheetByName(sheetName);
  if (!sheet) throw new Error("Onglet introuvable : " + sheetName);
  const data = sheet.getDataRange().getValues();
  const headers = data[0];
  return data.slice(1).map(row => {
    const obj = {};
    headers.forEach((h, i) => { obj[h] = row[i]; });
    return obj;
  });
}

// ════════════════════════════════════════════════════
//  UTILITAIRE : nom du header à partir de l'index (1-based)
// ════════════════════════════════════════════════════

function getHeaderName(ss, sheetName, colIndex) {
  const sheet = ss.getSheetByName(sheetName);
  if (!sheet) throw new Error("Onglet introuvable : " + sheetName);
  return sheet.getRange(1, colIndex).getValue();
}

// ════════════════════════════════════════════════════
//  NORMALISATION
// ════════════════════════════════════════════════════

function normaliser(str) {
  if (!str) str = '';
  return String(str)
    .toLowerCase()
    .normalize('NFD').replace(/[\u0300-\u036f]/g, '')
    .replace(/[-']/g, ' ')
    .replace(/\s+/g, ' ')
    .trim();
}

// ════════════════════════════════════════════════════
//  COMPARAISON DE NOMS SOUPLE
//  Gère : "Dupont Jean-Pierre", "Dupont J", "Dupont"
//  dans les deux sens
// ════════════════════════════════════════════════════

function nomCorrespondent(nomA, nomB) {
  const motsA = normaliser(nomA).split(' ').filter(m => m.length > 0);
  const motsB = normaliser(nomB).split(' ').filter(m => m.length > 0);

  if (motsA.length === 0 || motsB.length === 0) return false;

  // Essaie toutes les combinaisons possibles de "nom de famille"
  // ex: "Le Dilly Violaine" → teste "le", "le dilly", "le dilly violaine"
  for (let nbMotsNomA = 1; nbMotsNomA < motsA.length; nbMotsNomA++) {
    for (let nbMotsNomB = 1; nbMotsNomB < motsB.length; nbMotsNomB++) {

      const nomFamilleA = motsA.slice(0, nbMotsNomA).join(' ');
      const nomFamilleB = motsB.slice(0, nbMotsNomB).join(' ');

      if (nomFamilleA.length < 3) continue;
      if (nomFamilleA !== nomFamilleB) continue;

      // Noms de famille identiques — vérifie le prénom
      const prenomA = motsA.slice(nbMotsNomA).join(' ');
      const prenomB = motsB.slice(nbMotsNomB).join(' ');

      // Pas de prénom d'un côté → nom seul suffit
      if (prenomA === '' || prenomB === '') return true;

      // Prénom complet identique
      if (prenomA === prenomB) return true;

      // Initiales identiques
      if (prenomA[0] === prenomB[0]) return true;

      // Un prénom commence par l'initiale de l'autre
      if (prenomA.startsWith(prenomB[0]) || prenomB.startsWith(prenomA[0])) return true;
    }
  }

  // Cas où un des deux n'a qu'un seul mot — compare directement
  const nomSeulA = motsA.join(' ');
  const nomSeulB = motsB.join(' ');
  if (motsA.length === 1 && nomSeulA.length >= 3 && nomSeulB.startsWith(nomSeulA)) return true;
  if (motsB.length === 1 && nomSeulB.length >= 3 && nomSeulA.startsWith(nomSeulB)) return true;

  return false;
}

// ════════════════════════════════════════════════════
//  CORRESPONDANCE SOUPLE
//  Pour cotisations (nom+prénom séparés) vs membres
// ════════════════════════════════════════════════════

function trouverMembrePourCotisation(c, membres, cotisEmail, cotisNom, cotisPrenom, membreEmail, membreNom, membrePrenom) {
  const emailCible  = normaliser(c[cotisEmail] || '');
  const nomComplet  = ((c[cotisNom] || '') + ' ' + (c[cotisPrenom] || '')).trim();

  return membres.find(m => {
    // 1. Email exact
    if (emailCible !== '' && normaliser(m[membreEmail]) === emailCible) return true;

    // 2. Nom + prénom combinés vs nom + prénom combinés
    const nomMembre = ((m[membreNom] || '') + ' ' + (m[membrePrenom] || '')).trim();
    return nomCorrespondent(nomComplet, nomMembre);

  }) || null;
}

// ════════════════════════════════════════════════════
//  CORRESPONDANCE SOUPLE
//  Pour suivi_parrainages (SPONSOR_NAME) vs parrains (SPONSOR_NAME)
//  Les deux ont nom+prénom ou nom+initiale dans une seule colonne
// ════════════════════════════════════════════════════

function trouverParrainPourSuivi(s, parrains, suiviEmail, suiviNom, parrainEmail, parrainNom) {
  const emailCible = normaliser(s[suiviEmail] || '');
  const nomCible   = s[suiviNom] || '';

  Logger.log("=== Recherche pour : '" + nomCible + "' / '" + emailCible + "'");

  return parrains.find(p => {
    Logger.log("  → Compare avec parrain : '" + p[parrainNom] + "' / '" + p[parrainEmail] + "'");
    Logger.log("     nomCorrespondent : " + nomCorrespondent(nomCible, p[parrainNom] || ''));

    if (emailCible !== '' && normaliser(p[parrainEmail]) === emailCible) return true;
    return nomCorrespondent(nomCible, p[parrainNom] || '');
  }) || null;
}
// ════════════════════════════════════════════════════
//  DIAGNOSTIC PRINCIPAL
// ════════════════════════════════════════════════════

function buildDiagnostic() {
  Logger.log("buildDiagnostic() exécutée");
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  const parrains = getSheetRowsByName(ss, CONFIG.SHEET_PARRAIN);
  const membres  = getSheetRowsByName(ss, CONFIG.SHEET_MEMBRES);
  const cotis    = getSheetRowsByName(ss, CONFIG.SHEET_COTISATIONS_MERGED);
  const suivi    = getSheetRowsByName(ss, CONFIG.SHEET_SUIVI);

  // Noms réels des colonnes
  const cotisEmail   = getHeaderName(ss, CONFIG.SHEET_COTISATIONS_MERGED, CONFIG.COLS.COTISATIONS_MERGED.EMAIL);
  const cotisNom     = getHeaderName(ss, CONFIG.SHEET_COTISATIONS_MERGED, CONFIG.COLS.COTISATIONS_MERGED.NOM);
  const cotisPrenom  = getHeaderName(ss, CONFIG.SHEET_COTISATIONS_MERGED, CONFIG.COLS.COTISATIONS_MERGED.PRENOM);
  const cotisMontant = getHeaderName(ss, CONFIG.SHEET_COTISATIONS_MERGED, CONFIG.COLS.COTISATIONS_MERGED.MONTANT);

  const membreEmail  = getHeaderName(ss, CONFIG.SHEET_MEMBRES, CONFIG.COLS.MEMBRES.ADRESSE_MAIL);
  const membreNom    = getHeaderName(ss, CONFIG.SHEET_MEMBRES, CONFIG.COLS.MEMBRES.MEMBRE_NAME);
  const membrePrenom = getHeaderName(ss, CONFIG.SHEET_MEMBRES, CONFIG.COLS.MEMBRES.MEMBRE_FIRSTNAME);

  const parrainEmail = getHeaderName(ss, CONFIG.SHEET_PARRAIN, CONFIG.COLS.PARRAIN.ADRESSE_MAIL);
  const parrainNom   = getHeaderName(ss, CONFIG.SHEET_PARRAIN, CONFIG.COLS.PARRAIN.SPONSOR_NAME);

  const suiviEmail   = getHeaderName(ss, CONFIG.SHEET_SUIVI, CONFIG.COLS.SUIVI.ADRESSE_MAIL);
  const suiviNom     = getHeaderName(ss, CONFIG.SHEET_SUIVI, CONFIG.COLS.SUIVI.SPONSOR_NAME);

  const lignes = [];

  // ── 1. cotisations-merged vs membres ──
  cotis.forEach(c => {
    const nomComplet = ((c[cotisNom] || '') + ' ' + (c[cotisPrenom] || '')).trim();

    const trouveMembre = trouverMembrePourCotisation(
      c, membres,
      cotisEmail, cotisNom, cotisPrenom,
      membreEmail, membreNom, membrePrenom
    );

    if (!trouveMembre) {
      lignes.push([
        new Date(),
        '⚠️ Cotisation sans membre',
        nomComplet,
        c[cotisEmail] || '',
        'Absent de la liste membres'
      ]);
    } else if (normaliser(trouveMembre[membreEmail]) !== normaliser(c[cotisEmail] || '')) {
      lignes.push([
        new Date(),
        '⚠️ Email différent (cotisation/membre)',
        nomComplet,
        c[cotisEmail] || '',
        'Email dans membres : ' + trouveMembre[membreEmail]
      ]);
    }
  });

  // ── 2. suivi_parrainages vs parrains ──
  suivi.forEach(s => {
    const trouveParrain = trouverParrainPourSuivi(
      s, parrains,
      suiviEmail, suiviNom,
      parrainEmail, parrainNom
    );

    if (!trouveParrain) {
      lignes.push([
        new Date(),
        '⚠️ Parrainage sans parrain',
        s[suiviNom]   || '',
        s[suiviEmail] || '',
        'Absent de la liste parrains'
      ]);
    } else if (normaliser(trouveParrain[parrainEmail]) !== normaliser(s[suiviEmail] || '')) {
      lignes.push([
        new Date(),
        '⚠️ Email différent (parrainage/parrain)',
        s[suiviNom]   || '',
        s[suiviEmail] || '',
        'Email dans parrains : ' + trouveParrain[parrainEmail]
      ]);
    }
  });

  // ── Écriture dans l'onglet diagnostic ──
  let sheet = ss.getSheetByName(CONFIG.SHEET_DIAGNOSTIC);
  if (!sheet) {
    sheet = ss.insertSheet(CONFIG.SHEET_DIAGNOSTIC);
  } else {
    sheet.clearContents();
  }

  sheet.getRange(1, 1, 1, 5)
    .setValues([['Date vérif', 'Problème', 'Nom', 'Email', 'Détail']])
    .setFontWeight('bold');

  if (lignes.length > 0) {
    sheet.getRange(2, 1, lignes.length, 5).setValues(lignes);
    Logger.log(lignes.length + " problème(s) détecté(s)");
  } else {
    sheet.appendRow(['', '✅ Aucun problème détecté', '', '', '']);
    Logger.log("Aucun problème détecté");
  }

  return lignes.length;
}



// ════════════════════════════════════════════════════
//  POINT D'ENTRÉE MANUEL
// ════════════════════════════════════════════════════

function remplirDiagnostic() {
  const nb = buildDiagnostic();
  Logger.log("Diagnostic terminé : " + nb + " problème(s)");
}

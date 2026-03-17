// ============================================================
// TESTS
// ============================================================

function testerConnexionHelloAsso() {
  try {
    const token = getAccessToken();
    Logger.log("✅ Connexion OK — token : " + token.substring(0, 20) + "...");
  } catch (err) {
    Logger.log("❌ Erreur : " + err.toString());
  }
}

function testerLogo() {
  try {
    const file = DriveApp.getFileById(CONFIG.LOGO_ID);
    Logger.log("✅ Fichier trouvé : " + file.getName());
    Logger.log("Type MIME : " + file.getBlob().getContentType());
  } catch (err) {
    Logger.log("❌ Erreur : " + err.message);
  }
}

function testerRechercheCampagnes() {
  const token      = getAccessToken();
  const campagnes  = getSlugsCampagneSemestre(token);
  const cotisation = getSlugCotisationAnnee(token);

  Logger.log(`📋 Campagnes parrainage : ${campagnes.length}`);
  campagnes.forEach(c => Logger.log(`  → ${c.slug} | ${c.url}`));

  Logger.log(cotisation
    ? `📋 Cotisation : ${cotisation.slug}`
    : "⚠️ Aucune campagne cotisation trouvée");
}

function testerTemplatesEmails() {
  const semestre  = getSemestreEnCours();
  const annee     = new Date().getFullYear().toString();
  const lienTest  = "https://www.helloasso.com/test";
  const emailTest = Session.getActiveUser().getEmail();

  GmailApp.sendEmail(emailTest, "TEST - Appel parrainage", "", {
    htmlBody: getTemplateAppelCotisation("Jean Dupont", semestre, lienTest),
    name: CONFIG.ASSO_NOM
  });
  GmailApp.sendEmail(emailTest, "TEST - Relance parrainage", "", {
    htmlBody: getTemplateRelance("Jean Dupont", semestre, lienTest),
    name: CONFIG.ASSO_NOM
  });
  GmailApp.sendEmail(emailTest, "TEST - Appel cotisation annuelle", "", {
    htmlBody: getTemplateAppelCotisationAnnuelle("Marie Martin", annee, lienTest),
    name: CONFIG.ASSO_NOM
  });

  Logger.log(`✅ 3 emails de test envoyés à ${emailTest}`);
}

function testerSynchronisationComplete() {
  Logger.log("🔄 Début synchronisation...");
  MajOngletSuivi();
  Logger.log("✅ Synchronisation terminée.");
}




function debugExportCotisation() {
  const token    = getAccessToken();
  const campagne = getSlugCotisationAnnee(token);
  Logger.log("Campagne trouvée : " + JSON.stringify(campagne));
  
  if (campagne) {
    exportCotisationToSheets();
    Logger.log("Export terminé — vérifie l'onglet helloasso-cotisation");
  }
}

function testerTemplatesEmails() {
  const semestre  = getSemestreEnCours();
  const annee     = new Date().getFullYear().toString();
  const lienTest  = "https://www.helloasso.com/test";
  const emailTest = Session.getActiveUser().getEmail();

  GmailApp.sendEmail(emailTest, "TEST - Appel parrainage", "", {
    htmlBody: getTemplateAppelCotisation("Jean Dupont", semestre, lienTest),
    name: CONFIG.ASSO_NOM
  });
  GmailApp.sendEmail(emailTest, "TEST - Relance parrainage", "", {
    htmlBody: getTemplateRelance("Jean Dupont", semestre, lienTest),
    name: CONFIG.ASSO_NOM
  });
  GmailApp.sendEmail(emailTest, "TEST - Appel cotisation annuelle", "", {
    htmlBody: getTemplateAppelCotisationAnnuelle("Marie Martin", annee, lienTest),
    name: CONFIG.ASSO_NOM
  });

  GmailApp.sendEmail(emailTest, "TEST - Relance membre", "", {
  htmlBody: getTemplateRelanceMembre("Marie Martin", annee, lienTest),
  name: CONFIG.ASSO_NOM
});

  Logger.log(`✅ 3 emails de test envoyés à ${emailTest}`);
}

// ══════════════════════════════════════════════════════════════
// TEST merge des cotisations

// ══════════════════════════════════════════════════════════════

function testerMergeCotisations() {
  Logger.log("🔄 Début test mergeCotisations...");
  mergeCotisations();
  Logger.log("✅ Test mergeCotisations terminé.");
}

// ══════════════════════════════════════════════════════════════
// TEST —connexion Dashboard

// ══════════════════════════════════════════════════════════════
function testerConnexion() {
  const result = buildDashboardData();
  Logger.log('✅ Parrains : '    + result.parrains);
  Logger.log('✅ Enfants : '     + result.enfants);
  Logger.log('✅ Parrainages : ' + result.parrainages);
  Logger.log('✅ Ongoing : '     + result.ongoing);
  Logger.log('✅ Durée moy : '   + result.duree_moy + ' ans');
  Logger.log('✅ Années : '      + result.annees.join(', '));
  Logger.log(JSON.stringify(result.enfants_data.all.sexe));
}

// ============================================================
//  Tests_Webhook.gs — Simulation des webhooks HelloAsso
//  À coller dans Tests.gs ou dans un fichier dédié
//  Nécessite MODE_TEST: true dans Config.gs
// ============================================================

// ------------------------------------------------------------
// 1. NOUVELLE CAMPAGNE PARRAINAGE
//    Simule HelloAsso qui détecte une nouvelle collecte S1/S2
//    → doit déclencher l'envoi des emails aux parrains actifs
// ------------------------------------------------------------
function testWebhookNouvelleCampagneParrainage() {
  const fakeEvent = {
    postData: {
      contents: JSON.stringify({
        eventType: "Form",
        data: {
          formType:         "CrowdFunding",
          formSlug:         "parrainage-2026-semestre-1",
          organizationSlug: "les-enfants-de-pondypatch",
          title:            "Parrainage 2026 — Semestre 1",
          url:              "https://www.helloasso.com/associations/les-enfants-de-pondypatch/adhesions/parrainage-2026-semestre-1"
        }
      })
    }
  };

  console.log("▶ Test : nouvelle campagne parrainage S1");
  traiterWebhookHelloAsso(fakeEvent);
  console.log("✅ Terminé — vérifie EMAIL_TEST pour l'email aux parrains + confirmation direction");
}


// ------------------------------------------------------------
// 2. PAIEMENT PARRAINAGE REÇU
//    Simule un parrain qui règle son semestre via HelloAsso
//    → doit mettre à jour le Sheet + envoyer confirmation
// ------------------------------------------------------------
function testWebhookPaiementParrainage() {
  const fakeEvent = {
    postData: {
      contents: JSON.stringify({
        eventType: "Order",
        data: {
          state: "Authorized",
          formType:         "CrowdFunding",
          formSlug:         "parrainage-2026-semestre-1",
          organizationSlug: "les-enfants-de-pondypatch",
          payer: {
            email:     "parrain.test@example.com",
            firstName: "Jean",
            lastName:  "Dupont"
          },
          items: [
            {
              amount:    15000,   // en centimes → 150,00 €
              type:      "Payment",
              state:     "Processed"
            }
          ],
          date: new Date().toISOString()
        }
      })
    }
  };

  console.log("▶ Test : paiement parrainage reçu");
  traiterWebhookHelloAsso(fakeEvent);
  console.log("✅ Terminé — vérifie le Sheet Paiements + EMAIL_TEST");
}


// ------------------------------------------------------------
// 3. NOUVELLE CAMPAGNE COTISATION ANNUELLE
//    Simule HelloAsso qui détecte la nouvelle campagne adhésion
//    → doit déclencher l'envoi des emails à tous les membres
// ------------------------------------------------------------
function testWebhookNouvelleCampagneCotisation() {
  const fakeEvent = {
    postData: {
      contents: JSON.stringify({
        eventType: "Form",
        data: {
          formType:         "Membership",
          formSlug:         "adhesion-edp-2026",
          organizationSlug: "les-enfants-de-pondypatch",
          title:            "Adhésion EDP 2026",
          url:              "https://www.helloasso.com/associations/les-enfants-de-pondypatch/adhesions/adhesion-edp-2026"
        }
      })
    }
  };

  console.log("▶ Test : nouvelle campagne cotisation annuelle");
  traiterWebhookHelloAsso(fakeEvent);
  console.log("✅ Terminé — vérifie EMAIL_TEST pour l'email aux membres + confirmation direction");
}


// ------------------------------------------------------------
// 4. PAIEMENT COTISATION REÇU
//    Simule un membre qui règle son adhésion annuelle
//    → doit mettre à jour le Sheet Membres + envoyer confirmation
// ------------------------------------------------------------
function testWebhookPaiementCotisation() {
  const fakeEvent = {
    postData: {
      contents: JSON.stringify({
        eventType: "Order",
        data: {
          state: "Authorized",
          formType:         "Membership",
          formSlug:         "adhesion-edp-2026",
          organizationSlug: "les-enfants-de-pondypatch",
          payer: {
            email:     "membre.test@example.com",
            firstName: "Marie",
            lastName:  "Martin"
          },
          items: [
            {
              amount:    1000,   // en centimes → 10,00 € (montant libre minimum)
              type:      "Payment",
              state:     "Processed"
            }
          ],
          date: new Date().toISOString()
        }
      })
    }
  };

  console.log("▶ Test : paiement cotisation reçu");
  traiterWebhookHelloAsso(fakeEvent);
  console.log("✅ Terminé — vérifie le Sheet Membres + EMAIL_TEST");
}


// ------------------------------------------------------------
// TOUT TESTER D'UN COUP (optionnel)
// Lance les 4 scénarios en séquence avec pause entre chaque
// ------------------------------------------------------------
function testTousLesWebhooks() {
  console.log("=== Début des tests webhook ===\n");

  testWebhookNouvelleCampagneParrainage();
  Utilities.sleep(2000);

  testWebhookPaiementParrainage();
  Utilities.sleep(2000);

  testWebhookNouvelleCampagneCotisation();
  Utilities.sleep(2000);

  testWebhookPaiementCotisation();

  console.log("\n=== Tous les tests terminés ===");
  console.log("👉 Vérifie EMAIL_TEST et les Journaux d'exécution");
}


// ============================================================
//  Tests_Relances.gs — Simulation des emails de relance
//  À coller dans Tests.gs ou dans un fichier dédié
//  Nécessite MODE_TEST: true dans Config.gs
// ============================================================


// ------------------------------------------------------------
// 1. RÉCAPITULATIF MENSUEL IMPAYÉS PARRAINS
//    Simule le déclencheur du 1er du mois
//    → envoie à EMAIL_TEST le récap avec les boutons Relancer
// ------------------------------------------------------------
function testRecapMensuelRelances() {
  console.log("▶ Test : récapitulatif mensuel impayés parrains");
  envoyerRecapMensuelRelances();
  console.log("✅ Terminé — vérifie EMAIL_TEST pour le récap avec boutons Relancer");
}


// ------------------------------------------------------------
// 2. RELANCE INDIVIDUELLE PARRAIN
//    Simule un clic sur "Relancer" depuis l'email récap
//    → envoie à EMAIL_TEST l'email de relance du parrain
//
//    ⚠️ Remplace SPONSOR_ID et LIGNE par des valeurs réelles
//    depuis ton onglet Suivi_Parrainages
// ------------------------------------------------------------
function testRelanceIndividuelleParrain() {
  const SPONSOR_ID = "SP-0186";  // ← remplace par un vrai SPONSOR_ID de ton Sheet
  const LIGNE      = 247;        // ← remplace par le numéro de ligne dans Suivi_Parrainages

  console.log(`▶ Test : relance individuelle parrain ${SPONSOR_ID} (ligne ${LIGNE})`);

  // Désactive temporairement le délai anti-spam pour le test
  const props = PropertiesService.getScriptProperties();
  const cleDelai = `RELANCE_${getSemestreEnCours().label}_${SPONSOR_ID}`;
  const valeurOriginale = props.getProperty(cleDelai);
  props.deleteProperty(cleDelai);  // supprime la date de dernière relance

  const resultat = envoyerRelanceIndividuelle(SPONSOR_ID, LIGNE);

  // Restaure la valeur originale si elle existait
  if (valeurOriginale) props.setProperty(cleDelai, valeurOriginale);

  console.log(`✅ Terminé — résultat : ${resultat}`);
  console.log("👉 Vérifie EMAIL_TEST pour l'email de relance parrain");
}


// ------------------------------------------------------------
// 3. RELANCER TOUS LES PARRAINS ÉLIGIBLES
//    Simule le bouton "Relancer tous" depuis l'email récap
//    → envoie à EMAIL_TEST un email par parrain éligible
// ------------------------------------------------------------
function testRelancerTousLesEligibles() {
  console.log("▶ Test : relancer tous les parrains éligibles");
  const nb = relancerTousLesEligibles();
  console.log(`✅ Terminé — ${nb} relance(s) envoyée(s) vers EMAIL_TEST`);
}


// ------------------------------------------------------------
// 4. RELANCE INDIVIDUELLE MEMBRE
//    Simule un clic sur "Relancer" depuis l'email récap membres
//    → envoie à EMAIL_TEST l'email de relance du membre
//
//    ⚠️ Remplace EMAIL_MEMBRE par un vrai email de ton Sheet Membres
// ------------------------------------------------------------
function testRelanceIndividuelleMembre() {
  
  console.log(`▶ Test : relance individuelle membre ${EMAIL_MEMBRE}`);

  // Désactive temporairement le délai anti-spam pour le test
  const annee = new Date().getFullYear().toString();
  const props = PropertiesService.getScriptProperties();
  const cle   = `RELANCE_MEMBRE_${EMAIL_TEST}_${annee}`;
  const valeurOriginale = props.getProperty(cle);
  props.deleteProperty(cle);  // supprime la date de dernière relance

  const resultat = relancerMembreIndividuel(EMAIL_TEST);

  // Restaure la valeur originale si elle existait
  if (valeurOriginale) props.setProperty(cle, valeurOriginale);

  console.log(`✅ Terminé — résultat : ${resultat}`);
  console.log("👉 Vérifie EMAIL_TEST pour l'email de relance membre");
}


// ------------------------------------------------------------
// 5. RÉCAPITULATIF RELANCES MEMBRES (juillet)
//    Simule le déclencheur du 1er juillet
//    → envoie à EMAIL_TEST le récap membres avec boutons Relancer
// ------------------------------------------------------------
function testRecapRelancesMembres() {
  console.log("▶ Test : récapitulatif relances membres");
  envoyerRecapMensuelRelancesMembres();
  console.log("✅ Terminé — vérifie EMAIL_TEST pour le récap membres");
  Logger.log("🔗 WEBAPP_URL utilisée : " + CONFIG.WEBAPP_URL);
}


// ------------------------------------------------------------
// TOUT TESTER D'UN COUP (optionnel)
// ⚠️ Peut envoyer beaucoup d'emails vers EMAIL_TEST
// ------------------------------------------------------------
function testToutesLesRelances() {
  console.log("=== Début des tests relances ===\n");

  testRecapMensuelRelances();
  Utilities.sleep(2000);

  testRelancerTousLesEligibles();
  Utilities.sleep(2000);

  testRelanceIndividuelleParrain();
  Utilities.sleep(2000);

  testRelanceIndividuelleMembre();
  Utilities.sleep(2000);

  testRecapRelancesMembres();

  console.log("\n=== Tous les tests relances terminés ===");
  console.log("👉 Vérifie EMAIL_TEST et les Journaux d'exécution");
}

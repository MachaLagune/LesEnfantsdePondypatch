
// ══════════════════════════════════════════════════════════════
// CONSTRUCTION DES DONNÉES DASHBOARD
// ══════════════════════════════════════════════════════════════
function buildDashboardData() {
  const ss = SpreadsheetApp.openById(CONFIG.spreadsheetId);

  const parrains    = getSheetRows(ss, CONFIG.SHEET_PARRAIN);
  const parrainages = getSheetRows(ss, CONFIG.SHEET_PARRAINAGE);
  const enfants     = getSheetRows(ss, CONFIG.SHEET_ENFANT);
  const rdvs        = getSheetRows(ss, CONFIG.SHEET_RDV);

  // ── GLOBALE ──
  const C_P  = CONFIG.COLS.PARRAIN;
  const C_PA = CONFIG.COLS.PARRAINAGE;
  const C_E  = CONFIG.COLS.ENFANT;
  const C_R  = CONFIG.COLS.RDV;

  const totalParrains    = countDistinct(parrains,    C_P.SPONSOR_ID);
  const totalParrainages = parrainages.length;
  const totalEnfants     = countDistinct(enfants, C_E.KID_ID);
  const ongoing          = parrainages.filter(r => r[C_PA.STATUS] === 'Ongoing').length;
  const agesMoyens = calculerAgesMoyens(parrainages, enfants, C_PA, C_E);
  const dureeMoy   = (agesMoyens.age_moy_fin - agesMoyens.age_moy_debut) || calculerDureeMoyenne(parrainages, C_PA);


  // ── PAR ANNÉE ──
  const annees    = extraireAnnees(parrainages, C_PA.SPONSORSHIP_START_DATE);
  const par_annee = {};
  annees.forEach(a => {
    par_annee[a] = calculerParAnnee(parrainages, rdvs, a, C_PA, C_R);
  });

  // ── PROFILS ENFANTS FILTRÉS ──
  const enfants_data = {
  all:      calculerProfilEnfants(enfants, rdvs, parrainages, null,       C_E, C_R, C_PA),
  Ongoing:  calculerProfilEnfants(enfants, rdvs, parrainages, 'Ongoing',  C_E, C_R, C_PA),
  Finished: calculerProfilEnfants(enfants, rdvs, parrainages, 'Finished', C_E, C_R, C_PA),
};

  return {
    parrains:      totalParrains,
    enfants:       totalEnfants,
    parrainages:   totalParrainages,
    ongoing:       ongoing,
    duree_moy:     dureeMoy,
    age_moy_debut: agesMoyens.age_moy_debut,
    age_moy_fin:   agesMoyens.age_moy_fin,
    annees:        annees,
    par_annee:     par_annee,
    enfants_data:  enfants_data,
    lastUpdate:    Utilities.formatDate(new Date(), 'Europe/Paris', 'dd/MM/yyyy')
  };
}

// ══════════════════════════════════════════════════════════════
// CALCULS
// ══════════════════════════════════════════════════════════════

function calculerDureeMoyenne(parrainages, C) {
  const termines = parrainages.filter(r => r[C.SPONSORSHIP_START_DATE] && r[C.SPONSORSHIP_END_DATE]);
  if (!termines.length) return 0;
  const total = termines.reduce((sum, r) => {
    const debut = new Date(r[C.SPONSORSHIP_START_DATE]);
    const fin   = new Date(r[C.SPONSORSHIP_END_DATE]);
    return sum + (fin - debut) / (1000 * 60 * 60 * 24 * 365.25);
  }, 0);
  return Math.round(total / termines.length);
}

function calculerAgesMoyens(parrainages, enfants, C_PA, C_E) {
  // Dictionnaire kid_id → date de naissance
  const naissances = {};
  enfants.forEach(e => {
    const id  = e[C_E.KID_ID];
    const dob = e[C_E.DATE_BIRTH];
    if (id && dob) naissances[id] = new Date(dob);
  });

  // ✅ Uniquement les parrainages terminés (comme calculerDureeMoyenne)
  const termines = parrainages.filter(r =>
    r[C_PA.SPONSORSHIP_START_DATE] && r[C_PA.SPONSORSHIP_END_DATE]
  );

  const agesDebut = [];
  const agesFin   = [];

  termines.forEach(p => {
    const id    = p[C_PA.KID_ID];
    const debut = p[C_PA.SPONSORSHIP_START_DATE];
    const fin   = p[C_PA.SPONSORSHIP_END_DATE];
    const dob   = naissances[id];
    if (!dob) return;

    const ageDebut = (new Date(debut) - dob) / (1000 * 60 * 60 * 24 * 365.25);
    const ageFin   = (new Date(fin)   - dob) / (1000 * 60 * 60 * 24 * 365.25);

    if (ageDebut > 0) agesDebut.push(ageDebut);
    if (ageFin   > 0) agesFin.push(ageFin);
  });

  const moy = arr => arr.length
    ? Math.round(arr.reduce((a, b) => a + b, 0) / arr.length)
    : 0;

  return {
    age_moy_debut: moy(agesDebut),
    age_moy_fin:   moy(agesFin)
  };
}

function extraireAnnees(parrainages, col) {
  const annees = new Set();
  parrainages.forEach(r => {
    const d = r[col];
    if (d) annees.add(new Date(d).getFullYear());
  });
  return [...annees].filter(a => !isNaN(a)).sort();
}

function calculerParAnnee(parrainages, rdvs, annee, C_PA, C_R) {

  // Nouveaux : démarrent cette année
  const nouveaux = parrainages.filter(r => {
    const d = r[C_PA.SPONSORSHIP_START_DATE];
    return d && new Date(d).getFullYear() === annee;
  }).length;

  // Terminés : se terminent cette année
  const termines = parrainages.filter(r => {
    const d = r[C_PA.SPONSORSHIP_END_DATE];
    return d && new Date(d).getFullYear() === annee;
  }).length;

  // En cours : démarrés AVANT l'année ET terminés APRÈS l'année (pas pendant)
  const encours = parrainages.filter(r => {
    const deb = r[C_PA.SPONSORSHIP_START_DATE];
    const fin = r[C_PA.SPONSORSHIP_END_DATE];
    if (!deb) return false;
    const debAnnee = new Date(deb).getFullYear();
    const finAnnee = fin ? new Date(fin).getFullYear() : 9999;
    return debAnnee < annee && finAnnee > annee;
  }).length;

  // Entretiens : RDV réalisés cette année
  const entretiens = rdvs.filter(r => {
    const d = r[C_R.DATE];
    return d && new Date(d).getFullYear() === annee;
  }).length;

  return { nouveaux, termines, encours, entretiens };
}

function calculerProfilEnfants(enfants, rdvs, parrainages, filtre, C_E, C_R, C_PA) {

  // ── Source : parrainages filtrés ──
  const rows = parrainages.filter(p =>
    filtre ? p[C_PA.STATUS] === filtre : true
  );

  // ── Jointure kidId → enfant (pour sexe et âge uniquement) ──
  const enfantsParKid = {};
  enfants.forEach(e => {
    const kidId = String(e[C_E.KID_ID] || '').trim();
    if (kidId) enfantsParKid[kidId] = e;
  });

  // ── Jointure kidId → dernier RDV (pour niveau, domaine, statut CR) ──
  const dernierRDVParKid = {};
  rdvs.forEach(r => {
    const kidId = String(r[C_R.KID_ID] || '').trim();
    const date  = r[C_R.DATE] instanceof Date ? r[C_R.DATE].getTime() : 0;
    if (!kidId) return;
    if (!dernierRDVParKid[kidId] || date > dernierRDVParKid[kidId].date) {
      dernierRDVParKid[kidId] = {
        date,
        niveau:  String(r[C_R.NIVEAU]    || '').trim(),
        domaine: String(r[C_R.DOMAINE]   || '').trim(),
        statut:  String(r[C_R.STATUT_CR] || '').trim(),
      };
    }
  });

  // ── Sexe — depuis fiche enfant ──
  const sexe = { F: 0, M: 0 };
  rows.forEach(p => {
    const kidId  = String(p[C_PA.KID_ID] || '').trim();
    const enfant = enfantsParKid[kidId];
    if (!enfant) return;
    const s = String(enfant[C_E.GENDER] || '').trim().toUpperCase();
    if (['F','FILLE','FEMALE'].includes(s))         sexe.F++;
    else if (['M','GARÇON','MALE','H'].includes(s)) sexe.M++;
  });

  // ── Domaines — depuis dernier RDV ──
  const domaines = {};
  rows.forEach(p => {
    const kidId = String(p[C_PA.KID_ID] || '').trim();
    const entry = dernierRDVParKid[kidId];
    const d     = (entry && entry.domaine) ? entry.domaine : 'Non renseigné';
    domaines[d] = (domaines[d] || 0) + 1;
  });

  // ── Niveaux — depuis dernier RDV ──
  const niveaux = {};
  rows.forEach(p => {
    const kidId = String(p[C_PA.KID_ID] || '').trim();
    const entry = dernierRDVParKid[kidId];
    const n     = (entry && entry.niveau) ? entry.niveau : 'Non renseigné';
    niveaux[n] = (niveaux[n] || 0) + 1;
  });

  // ── Ages — depuis fiche enfant ──
  const ages = [];
  rows.forEach(p => {
    const kidId  = String(p[C_PA.KID_ID] || '').trim();
    const enfant = enfantsParKid[kidId];
    if (!enfant) return;
    const age = parseInt(enfant[C_E.AGE]);
    if (!isNaN(age) && age > 0) ages.push(age);
  });
  const ageMoy = ages.length ? Math.round(ages.reduce((a,b) => a+b,0) / ages.length) : 0;
  const ageMin = ages.length ? Math.min(...ages) : 0;
  const ageMax = ages.length ? Math.max(...ages) : 0;

  // ── Statut CR — depuis dernier RDV ──
  const cr = {};
  rows.forEach(p => {
    const kidId  = String(p[C_PA.KID_ID] || '').trim();
    const entry  = dernierRDVParKid[kidId];
    const statut = (entry && entry.statut) ? entry.statut : 'Non renseigné';
    cr[statut] = (cr[statut] || 0) + 1;
  });

  return {
    total:    rows.length,
    sexe,
    domaines: trierEtLimiter(domaines, 6),
    niveaux:  trierParNiveau(niveaux),
    ages:     { moy: ageMoy, min: ageMin, max: ageMax },
    cr
  };
}

// ══════════════════════════════════════════════════════════════
// UTILITAIRES
// ══════════════════════════════════════════════════════════════

function getSheetRows(ss, sheetName) {
  const sheet = ss.getSheetByName(sheetName);
  if (!sheet) { Logger.log('Onglet introuvable : ' + sheetName); return []; }
  const data = sheet.getDataRange().getValues();
  if (data.length < 2) return [];
  return data.slice(1).map(row => {
    const proxy = {};
    row.forEach((val, i) => { proxy[i + 1] = val; });
    return proxy;
  });
}

function countDistinct(rows, col) {
  return new Set(
    rows.map(r => r[col]).filter(v => v !== '' && v !== null && v !== undefined)
  ).size;
}

function trierEtLimiter(obj, n) {
  return Object.fromEntries(
    Object.entries(obj).sort((a,b) => b[1]-a[1]).slice(0, n)
  );
}

const ORDRE_NIVEAUX = [
  '1st grade','2nd grade','3rd grade','4th grade','5th grade','6th grade',
  '7th grade','8th grade','9th grade','10th grade','11th grade','12th grade',
  '1ère année','2ème année','3ème année','4ème année',
  'Bachelor','Master 1','Master 2','PhD','Autre','Non renseigné'
];
function trierParNiveau(obj) {
  const sorted = {};
  ORDRE_NIVEAUX.forEach(k => { if (obj[k] !== undefined) sorted[k] = obj[k]; });
  Object.entries(obj).forEach(([k,v]) => { if (!sorted[k]) sorted[k] = v; });
  return sorted;
}




function buildMembresData() {
  const ss          = SpreadsheetApp.openById(CONFIG.spreadsheetId);
  const membres     = getSheetRows(ss, CONFIG.SHEET_MEMBRES);
  const cotisations = getSheetRows(ss, CONFIG.SHEET_COTISATIONS_MERGED);
  const C_M  = CONFIG.COLS.MEMBRES;
  const C_C  = CONFIG.COLS.COTISATIONS_MERGED;
  const props = PropertiesService.getScriptProperties();

  const anneeEnCours = new Date().getFullYear();

  const extraireAnnee = source => {
    const match = String(source || '').match(/(\d{4})/);
    return match ? parseInt(match[1]) : null;
  };

  const payesIds = new Set(
    cotisations
      .filter(r => extraireAnnee(r[C_C.SOURCE]) === anneeEnCours)
      .map(r => {
        const nom   = String(r[C_C.NOM]   || '').trim().toLowerCase();
        const email = String(r[C_C.EMAIL] || '').trim().toLowerCase();
        return `${nom}|${email}`;
      })
  );

  const liste = membres.map(r => {
    const email = String(r[C_M.ADRESSE_MAIL] || '').trim().toLowerCase();
    const nom   = String(r[C_M.MEMBRE_NAME]  || '').trim().toLowerCase();
    const cotisationOk = payesIds.has(`${nom}|${email}`);

    let urlRelance = null;
    let eligible   = false;

    if (!cotisationOk && email) {
      // ✅ Réutiliser le token existant plutôt que d'en générer un nouveau
      const tokenCle = `TOKEN_RELANCE_MEMBRE_${email}`;
      let token = props.getProperty(tokenCle);
      if (!token) {
        token = Utilities.getUuid();
        props.setProperty(tokenCle, token);
      }
      urlRelance = `${CONFIG.WEBAPP_URL}?action=relancer_membre&email=${encodeURIComponent(email)}&token=${token}`;

      // ✅ Éligibilité basée sur la vraie date de dernière relance
      const cleRelance    = `RELANCE_MEMBRE_${email}_${anneeEnCours}`;
      const derniereRelance = props.getProperty(cleRelance);
      const jours = derniereRelance
        ? (new Date() - new Date(derniereRelance)) / (1000 * 60 * 60 * 24)
        : null;
      eligible = jours === null || jours >= CONFIG.DELAI_RELANCE_MEMBRES_JOURS;
    }

    return {
      nom:          String(r[C_M.MEMBRE_NAME]      || '').trim(),
      prenom:       String(r[C_M.MEMBRE_FIRSTNAME] || '').trim(),
      email,
      cotisationOk,
      urlRelance,
      eligible,
    };
  });

  const actifs   = liste.filter(m => m.cotisationOk).length;
  const enRetard = liste.filter(m => !m.cotisationOk);

  const totauxParAnnee = {};
  cotisations.forEach(r => {
    const annee   = extraireAnnee(r[C_C.SOURCE]);
    const montant = parseFloat(r[C_C.MONTANT]) || 0;
    if (annee) totauxParAnnee[annee] = (totauxParAnnee[annee] || 0) + montant;
  });

  return {
    total:         membres.length,
    actifs,
    enRetard,
    liste,
    totauxParAnnee,
    anneeEnCours,
    totalCollecte: totauxParAnnee[anneeEnCours] || 0,
  };
}

// ══════════════════════════════════════════════════════════════
// COTISATIONS PARRAINAGES — à ajouter dans Dashboard.gs
// ══════════════════════════════════════════════════════════════

function buildCotisationsData() {
  const ss    = SpreadsheetApp.openById(CONFIG.spreadsheetId);
  const suivi = getSheetRows(ss, CONFIG.SHEET_SUIVI);
  const props = PropertiesService.getScriptProperties();
  const C     = CONFIG.COLS.SUIVI;

  const extraireAnnee = sem => {
    const match = String(sem || '').match(/(\d{4})/);
    return match ? parseInt(match[1]) : null;
  };

  const anneeEnCours = new Date().getFullYear();
  const PAYE = '✅ Payé';

  // ── TOTAUX PAR ANNÉE ──
  const totauxParAnnee    = {};
  const attenduParAnnee   = {};
  const cotisantsParAnnee = {};

  suivi.forEach(r => {
    const annee = extraireAnnee(r[C.SEMESTRE]);
    if (!annee) return;
    const montantPaye    = parseFloat(r[C.MONTANT_PAYE])    || 0;
    const montantAttendu = parseFloat(r[C.MONTANT_ATTENDU]) || 0;
    totauxParAnnee[annee]  = (totauxParAnnee[annee]  || 0) + montantPaye;
    attenduParAnnee[annee] = (attenduParAnnee[annee] || 0) + montantAttendu;
    if (r[C.STATUT] === PAYE) {
      cotisantsParAnnee[annee] = (cotisantsParAnnee[annee] || 0) + 1;
    }
  });

  // ── SEMESTRES DE L'ANNÉE EN COURS ──
  const s1Key    = 'S1 ' + anneeEnCours;
  const s2Key    = 'S2 ' + anneeEnCours;
  const lignesS1 = suivi.filter(r => String(r[C.SEMESTRE] || '').includes(s1Key));
  const lignesS2 = suivi.filter(r => String(r[C.SEMESTRE] || '').includes(s2Key));

  const statsSemestre = (lignes) => ({
    total:          lignes.length,
    payes:          lignes.filter(r => r[C.STATUT] === PAYE).length,
    montantPaye:    lignes.reduce((s, r) => s + (parseFloat(r[C.MONTANT_PAYE])    || 0), 0),
    montantAttendu: lignes.reduce((s, r) => s + (parseFloat(r[C.MONTANT_ATTENDU]) || 0), 0),
  });

  // ── EN RETARD (année en cours, non payé) ──
  const enRetard = suivi
    .filter(r => {
      const annee = extraireAnnee(r[C.SEMESTRE]);
      return annee === anneeEnCours && r[C.STATUT] !== PAYE;
    })
    .map(r => ({
      nom:              String(r[C.SPONSOR_NAME]      || '').trim(),
      email:            String(r[C.ADRESSE_MAIL]      || '').trim(),
      semestre:         String(r[C.SEMESTRE]           || '').trim(),
      montantAttendu:   parseFloat(r[C.MONTANT_ATTENDU]) || 0,
      dateDernierRappel: r[C.DATE_DERNIER_RAPPEL]
        ? Utilities.formatDate(new Date(r[C.DATE_DERNIER_RAPPEL]), 'Europe/Paris', 'dd/MM/yyyy')
        : null,
      peutRelancer: !r[C.DATE_DERNIER_RAPPEL] ||
        (new Date() - new Date(r[C.DATE_DERNIER_RAPPEL])) / (1000 * 60 * 60 * 24) >= CONFIG.DELAI_RELANCE_JOURS
    }));

  // ── HISTORIQUE PAR PARRAIN ──
  const parParrain = {};
  suivi.forEach((r, i) => {
    const id    = String(r[C.SPONSOR_ID]   || '').trim();
    const nom   = String(r[C.SPONSOR_NAME] || '').trim();
    const email = String(r[C.ADRESSE_MAIL] || '').trim().toLowerCase();
    if (!id) return;
    if (!parParrain[id]) {
      parParrain[id] = { nom, email, lignes: [] };
    }
    parParrain[id].lignes.push({
      semestre:       String(r[C.SEMESTRE]        || '').trim(),
      statut:         String(r[C.STATUT]          || '').trim(),
      montantAttendu: parseFloat(r[C.MONTANT_ATTENDU]) || 0,
      montantPaye:    parseFloat(r[C.MONTANT_PAYE])    || 0,
      datePaiement:   r[C.DATE_PAIEMENT]
        ? Utilities.formatDate(new Date(r[C.DATE_PAIEMENT]), 'Europe/Paris', 'dd/MM/yyyy')
        : '—',
    });
  });

  // ── TRIER + CALCULER aJour, hyperlinkRelance, eligible ──
  Object.entries(parParrain).forEach(([sponsorId, p]) => {
    p.lignes.sort((a, b) => a.semestre.localeCompare(b.semestre));

    p.aJour = p.lignes
      .filter(l => extraireAnnee(l.semestre) === anneeEnCours)
      .every(l => l.statut === PAYE);

    if (!p.aJour) {
      // ✅ Trouver la ligne de la première cotisation en retard de l'année en cours
      const ligneEnRetard = suivi.findIndex(r =>
        String(r[C.SPONSOR_ID] || '').trim() === sponsorId &&
        extraireAnnee(r[C.SEMESTRE]) === anneeEnCours &&
        String(r[C.STATUT] || '').trim() !== PAYE
      );
      const ligneNum = ligneEnRetard !== -1 ? ligneEnRetard + 2 : null;

      // ✅ Réutiliser le token existant
      const tokenCle = `TOKEN_RELANCE_${sponsorId}`;
      let token = props.getProperty(tokenCle);
      if (!token) {
        token = Utilities.getUuid();
        props.setProperty(tokenCle, token);
      }
      p.hyperlinkRelance = ligneNum
        ? `${CONFIG.WEBAPP_URL}?action=relancer&sponsor=${encodeURIComponent(sponsorId)}&ligne=${ligneNum}&token=${token}`
        : null;
    } else {
      p.hyperlinkRelance = null;
    }

    // ✅ Éligibilité via ScriptProperties
    const cleRelance      = `RELANCE_PARRAIN_${p.email}`;
    const derniereRelance = props.getProperty(cleRelance);
    const jours = derniereRelance
      ? (new Date() - new Date(derniereRelance)) / (1000 * 60 * 60 * 24)
      : null;
    p.eligible = jours === null || jours >= CONFIG.DELAI_RELANCE_JOURS;
  });

  Logger.log('parParrain count : ' + Object.keys(parParrain).length);
  Logger.log('enRetard count : ' + enRetard.length);
  Logger.log('s1 : ' + JSON.stringify(statsSemestre(lignesS1)));

  return {
    anneeEnCours,
    totauxParAnnee,
    attenduParAnnee,
    cotisantsParAnnee,
    s1: statsSemestre(lignesS1),
    s2: statsSemestre(lignesS2),
    enRetard,
    parParrain: Object.values(parParrain).sort((a, b) => a.nom.localeCompare(b.nom)),
  };
}

function getCRsData() {
  const ss          = SpreadsheetApp.openById(CONFIG.spreadsheetId);
  const rdvs        = getSheetRows(ss, CONFIG.SHEET_RDV);
  const parrainages = getSheetRows(ss, CONFIG.SHEET_PARRAINAGE);
  const parrains    = getSheetRows(ss, CONFIG.SHEET_PARRAIN);

  const C_R  = CONFIG.COLS.RDV;
  const C_PA = CONFIG.COLS.PARRAINAGE;
  const C_P  = CONFIG.COLS.PARRAIN;

  // ✅ Sorti de la boucle
  const timezone = Session.getScriptTimeZone();

  const nomParrain = {};
  parrains.forEach(p => {
    nomParrain[p[C_P.SPONSOR_ID]] = String(p[C_P.SPONSOR_NAME] || '').trim();
  });

  const kidToSponsor = {};
  parrainages.forEach(p => {
    const kidId     = p[C_PA.KID_ID];
    const sponsorId = p[C_PA.SPONSOR_ID];
    if (!kidToSponsor[kidId] || p[C_PA.STATUS] === 'Ongoing') {
      kidToSponsor[kidId] = sponsorId;
    }
  });

  const ongoingKidIds = new Set(
    parrainages
      .filter(p => p[C_PA.STATUS] === 'Ongoing')
      .map(p => p[C_PA.KID_ID])
  );

  const seen = {};

  rdvs.forEach((row, i) => {
    const kidId   = row[C_R.KID_ID];
    const kidName = row[C_R.KID_NAME];
    if (!kidId || !kidName) return;

    const rowNumber    = i + 2;
    const rdvDate      = row[C_R.DATE] instanceof Date ? row[C_R.DATE] : null;
    const rdvTimestamp = rdvDate ? rdvDate.getTime() : 0;
    const existing     = seen[kidId];

    if (!existing || rdvTimestamp > (existing._timestamp || 0)) {
      const crUrl     = String(row[C_R.DERNIER_CR] || '').trim() || null;
      const sponsorId = kidToSponsor[kidId] || row[C_R.SPONSOR_ID];

      seen[kidId] = {
        rowNumber,
        kidId,
        kidName:                 String(kidName).trim(),
        parrainName:             nomParrain[sponsorId] || '—',
        dernierEntretienAffiche: rdvDate
          ? Utilities.formatDate(rdvDate, timezone, 'dd/MM/yyyy')  // ✅ variable réutilisée
          : null,
        crUrl,
        crStatut:   String(row[C_R.STATUT_CR] || '').trim() || null,
        validationUrl: (crUrl && String(row[C_R.STATUT_CR] || '').trim() !== 'CR envoyé')
          ? `https://docs.google.com/forms/d/e/${CONFIG.formId}/viewform?usp=pp_url&${CONFIG.entryIdLienCR}=${encodeURIComponent(crUrl)}`
          : null,
        _timestamp: rdvTimestamp,
      };
    }
  });

  return Object.values(seen)
    .filter(r => ongoingKidIds.has(r.kidId))
    .sort((a, b) => a.kidName.localeCompare(b.kidName))
    .map(({ _timestamp, kidId, ...rest }) => rest);
}

function buildDiagnosticData() {
  const ss       = SpreadsheetApp.getActiveSpreadsheet();
  const parrains = getSheetRows(ss, CONFIG.SHEET_PARRAINS);
  const membres  = getSheetRows(ss, CONFIG.SHEET_MEMBRES);
  const cotis    = getSheetRows(ss, CONFIG.SHEET_COTISATIONS_MERGED);
  const suivi    = getSheetRows(ss, CONFIG.SHEET_SUIVI);

  const emailsCotis = new Set(cotis.map(r => normaliser(r[C_C.EMAIL])));
  const emailsSuivi = new Set(suivi.map(r => normaliser(r[C_S.EMAIL_PARRAIN])));

  const alertes = [];

  parrains.forEach(p => {
    const nom = p[C_P.NOM] + ' ' + p[C_P.PRENOM];
    const email = p[C_P.EMAIL];
    if (!emailsCotis.has(normaliser(email)))
      alertes.push({ type: 'parrain', probleme: 'Sans cotisation', nom, email });
    if (!emailsSuivi.has(normaliser(email)))
      alertes.push({ type: 'parrain', probleme: 'Sans parrainage', nom, email });
  });

  membres.forEach(m => {
    const nom = m[C_M.NOM] + ' ' + m[C_M.PRENOM];
    const email = m[C_M.EMAIL];
    if (!emailsCotis.has(normaliser(email)))
      alertes.push({ type: 'membre', probleme: 'Sans cotisation', nom, email });
  });

  return {
    alertes,
    total: alertes.length,
    ok: alertes.length === 0
  };
}



function serveDashboardData() {
  try {
    const base        = buildDashboardData();
    const membres     = buildMembresData();
    const cotisations = buildCotisationsData();
    const diagnostic  = buildDiagnosticData(); // 

    let crs = [];
    try {
      crs = getCRsData();
    } catch(e) {
      Logger.log('❌ Erreur getCRsData : ' + e.toString());
    }

    // ✅ Retour explicite
    return {
      parrains:      base.parrains,
      enfants:       base.enfants,
      parrainages:   base.parrainages,
      ongoing:       base.ongoing,
      duree_moy:     base.duree_moy,
      age_moy_debut: base.age_moy_debut,
      age_moy_fin:   base.age_moy_fin,
      annees:        base.annees,
      par_annee:     base.par_annee,
      enfants_data:  base.enfants_data,
      lastUpdate:    base.lastUpdate,
      membres:       membres,
      cotisations:   cotisations,
      crs:           crs,
      diagnostic:    diagnostic 
    };

  } catch (e) {
    Logger.log('❌ Erreur globale : ' + e.toString());
    return { error: e.toString() };
  }
}




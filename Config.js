/**
 * ==================================================
 * CONFIGURATION GLOBALE
 * ==================================================
 */
const CONFIG = {
  // Feuilles / Spreadsheet
  spreadsheetId: '1Ft7B67LbMBp80hjjdcI_KXCmuz5AoRwLdwLE2UXzJF4',
  SHEET_ENFANT          : 'Fiche_Enfant_Form',
  SHEET_PARRAIN         : 'Fiche_Parrain_Form',
  SHEET_RDV             : 'Rdv_suivi',
  SHEET_PARRAINAGE      : 'Fiche_Parrainage_Form',
  SHEET_FIN_PARRAINAGE  : 'Fin_Parrainage_Form',
  SHEET_VALIDATION_CR   : 'Validation_CR_Form',
  
  // Paiements manuels
  SHEET_PAIEMENTS_MANUELS : "paiements-manuels",
  SHEET_COTISATIONS_MERGED: "cotisations-merged",
  

  // Parrainages hello asso
  SHEET_HELLO_ASSO_1: "helloasso-s1",
  SHEET_HELLO_ASSO_2: "helloasso-s2",
  SHEET_SUIVI:        "suivi_parrainages",

  // Cotisations hello asso
  SHEET_MEMBRES:               "Membres",
  SHEET_HELLO_ASSO_COTISATION: "helloasso-cotisation",

  // DIAGNOSTIC DIFFERENCE HELLO ASSO ET FICHIERS MEMEBRES/PARRAINS
  SHEET_DIAGNOSTIC:      'diagnostic',

  // Formulaires

  FORM_ENFANT: '1EFfX_YdS_LTkVyQOVEw5YtyDqoKF2YUS5qjotNepznc',// Nouvel Enfant
  FORM_PARRAIN : '1Y6ok1UqCBPcQmfcQ56S1ErWJrMmIVTZNUtwpgD4hIHM',// Nouveau Parrain
  FORM_PARRAINAGE: '17dFr6URnFbInFKxOzMR6CqBgKI2GyFWGTqkJMz3SlQk',// Nouveau Parrainage
  FORM_RDV_SUIVI: '1I4-zBU3RdXsgVDfwPxm4Di7ZNKt4P4F2CBpiMjgYlAw',// Nouvel Entretien
  FORM_FIN_PARRAINAGE: '1JfguhPFqoGFCy27uMtPkjrbDGKTqEnDjszcZvUSPuKw', // Fin de Parrainage
  FORM_VALIDATION_CR: '1TYRtILDX3_mwnXkNt42paU86mPeePO2fXTehT2wVcoU',  // Validation CR
  FORM_MEMBRES : '1rxYZl8NswOlM-f_IMNjT-bPbPpIPfo5AIfJ52B-hrhQ', // Nouveau membre
  FORM_PAIEMENTS_MANUELS  : '1Ci611Y0fB4uckmeIB7Vdorfpw0C97ZN9SRk4M464U4s', // Nouveau paiement manuel


// Colonnes (index 1-based)
  COLS: {
    // Fiche Enfant
    ENFANT: {
      TIMESTAMP: 1,
      KID_NAME: 2,
      DATE_BIRTH : 3,
      GENDER : 4,
      SCHOOL: 5,
      KID_ID: 6,
      AGE: 7,
    },

    // Fiche PARRAIN
    PARRAIN: {
      TIMESTAMP: 1,
      SPONSOR_NAME: 2,
      ADRESSE_MAIL: 3,
      SPONSOR_ID: 4
    },
    // Fiche Parrainage
    PARRAINAGE: {
      TIMESTAMP: 1,
      KID_NAME : 2,
      SPONSOR_NAME: 3,
      SPONSORSHIP_START_DATE : 4,
      STATUS: 5,
      SPONSORSHIP_END_DATE : 6,
      REASON_STOPPING : 7,
      SPONSORSHIP_ID : 8,   
      KID_ID: 9,
      SPONSOR_ID: 10,
      COTISATIONS : 11
     
    },

    // Rdv Suivi
    RDV: {
      TIMESTAMP: 1,
      KID_NAME : 2,
      SPONSOR_NAME: 3,
      DATE: 4,
      VU_PAR : 5,
      GENERAL_CONDITION: 6,
      ENGLISH_SKILLS : 7,
      NIVEAU: 8,
      DOMAINE: 9,
      STUDY_PLAN : 10,
      FAMILY : 11,
      OTHER : 12,
      PHOTO: 13,
      KID_ID: 14,
      SPONSOR_ID: 15,
      DERNIER_CR: 16,
      STATUT_CR: 17,
      LIEN_ENVOI: 18
    },

    // Fin Parrainage Form
    FIN: {
      TIMESTAMP:1,
      KID_NAME: 2,
      KID_ID: 5,
      DATE_FIN_PARRAINAGE: 3,
      RAISON_ARRET_PARRAINAGE: 4,
      SPONSORSHIP_ID: 6
    },

    //Validation CR
    VALIDATION : {
      TIMESTAMP: 1,
      LIEN_CR : 2,
      CR_FINALISE: 3,
      CR_ENVOYE : 4
    },

    // Structure commune aux 3 onglets helloasso-s1, helloasso-s2, helloasso-cotisation
    HELLO_ASSO: {
      DATE:    1,
      PRENOM:  2,
      NOM:     3,
      EMAIL:   4,
      MONTANT: 5,
      STATUT:  6,
      SLUG:    7
    },

    // Suivi Cotisations parrainages
    SUIVI: {
      ANNEE: 1, SEMESTRE: 2, HORODATEUR: 3,
      SPONSOR_NAME: 4, ADRESSE_MAIL: 5, SPONSOR_ID: 6,
      MONTANT_ATTENDU: 7, MONTANT_PAYE: 8, SURPLUS_PAIEMENT: 9,
      STATUT: 10, DATE_PAIEMENT: 11, DATE_DERNIER_RAPPEL: 12, HYPERLINK_RELANCE: 13
    },

    MEMBRES:    { HORODATEUR: 1, MEMBRE_NAME: 2,MEMBRE_FIRSTNAME: 3, ADRESSE_MAIL: 4 },

    PAIEMENT_MANUEL: {
    HORODATAGE:          1,  // toujours col A
    TYPE:                2,  // toujours col B
    // Section Parrainage
    PARRAIN:             3,
    MONTANT_PARRAINAGE:  4,
    DATE_PARRAINAGE:     5,
    MOYEN_PARRAINAGE:    6,
    NOTES_PARRAINAGE:    7,
    // Section Cotisation
    MEMBRE:              8,
    MONTANT_COTISATION:  9,
    DATE_COTISATION:     10,
    MOYEN_COTISATION:    11,
    NOTES_COTISATION:    12,
    // Ajouté par le script
    EMAIL:               13
    

    },

    COTISATIONS_MERGED: {
    DATE:    1,
    PRENOM:  2,
    NOM:     3,
    EMAIL:   4,
    MONTANT: 5,
    SOURCE:  6
    },


    C_DIAG : {
    DATE_VERIF: 0,
    PROBLEME:   1,
    NOM:        2,
    EMAIL:      3,
    DETAIL:     4,
    },
  },


  // Paramètres métier
  MONTANT_PAR_PARRAINAGE: 125, // Par parrainage Ongoing
  MONTANT_MINIMUM_COTISATION: 12,
  DELAI_RELANCE_JOURS: 30,
  DELAI_RELANCE_MEMBRES_JOURS: 180, // 6 mois
  LIEN_BASE_RENVOI: "https://www.helloasso.com/",    

  // Contacts
  EMAIL_DIRIGEANT: "lesenfantsdepondypatch@gmail.com",
  EMAILS_DIRECTION: ["lesenfantsdepondypatch@gmail.com"],

  // Web App Hello asso
  WEBAPP_URL: "https://script.google.com/macros/s/AKfycbwPUTZF7Iv-nOOXZVd4TJCtC-VmTiBz2u-OCHyXFDjiFICR9-vqJpDp4IfQ43a2RlqzfQ/exec",



// Drive / Docs
  driveFolderId: '1R2v6FUHCJr_0sKv89vd3QWYnBEq0_sD8',
  LOGO_ID: '1SJd5jAG1dc5mkYTM4NcR7iyL6iIVMyNR',
 

// Formulaire CR
  formId: '1FAIpQLSeXvkoNJFrN8HaA3_uwp0AShPb2Wrkyte7WAC_1t-7EXYS3zg',
  entryIdLienCR: 'entry.856165431',


// Identité association
  ASSO_NOM:   "Les Enfants de Pondypatch",
  ASSO_EMAIL: "lesenfantsdepondypatch@gmail.com",


  // HelloAsso
  CLIENT_ID: "1895823f4c214e43873a15d25df79037",
  CLIENT_SECRET: "VLAy82R3AQ/YfwN51veJz4hWambgz2RR",
  ORG_SLUG: "les-enfants-de-pondypatch",
  TYPE_FORMULAIRE: "Membership",// Membership, CrowdFunding, Event, Donation...
  LIEN_HELLOASSO_ACTUEL:     "",  // lien parrainage — mis à jour par le webhook
  LIEN_HELLOASSO_COTISATION: "",  // lien adhésion — mis à jour par le webhook


  MODE_TEST: true,
  EMAIL_TEST: 'machalagune@gmail.com', // reçoit tous les emails en mode test

    // Styles
  JAUNE_FONCE: '#F9A825',

    // Couleurs Sheets
  COULEUR_PAYE            : "#d9ead3", // vert clair
  COULEUR_NON_PAYE        : "#f4cccc", // rouge clair
  COULEUR_INSUFFISANT     : "#fce5cd", // orange clair

}

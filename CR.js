/**
 * ===============================================
 * Génération CR
 * ===============================================
 */

function genererCRPourCetteLigne() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(CONFIG.SHEET_RDV);
  if (!sheet) return;

  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return;

  const row = sheet.getRange(lastRow, 1, 1, sheet.getLastColumn()).getValues()[0];

  const kidId    = row[CONFIG.COLS.RDV.KID_ID - 1];
  const sponsorId = row[CONFIG.COLS.RDV.SPONSOR_ID - 1];
  const crUrl    = row[CONFIG.COLS.RDV.DERNIER_CR - 1];

  Logger.log(`🔍 Ligne ${lastRow} — kidId: ${kidId}, sponsorId: ${sponsorId}, crUrl: ${crUrl}`);

  if (!crUrl && kidId && sponsorId) {
    try {
      genererCRProtege(lastRow, row);
    } catch(e) {
      Logger.log(`❌ Erreur ligne ${lastRow} : ${e.message}`);
    }
  } else {
    Logger.log(`🚫 Ligne ${lastRow} ignorée : CR déjà généré ou données manquantes`);
  }
}

function genererCRProtege(rowNumber, rdvRow) {
  const kidName = rdvRow[CONFIG.COLS.RDV.KID_NAME - 1];
  const kidId = rdvRow[CONFIG.COLS.RDV.KID_ID - 1];
  const sponsorId = rdvRow[CONFIG.COLS.RDV.SPONSOR_ID - 1];

  const enfantData = chercherEnfant(kidId);
  const parrainData = chercherParrain(sponsorId);
  if (!enfantData || !parrainData) return;

  const timezone = Session.getScriptTimeZone();
  const rdvDate = rdvRow[CONFIG.COLS.RDV.DATE - 1] instanceof Date 
    ? Utilities.formatDate(rdvRow[CONFIG.COLS.RDV.DATE - 1], timezone, "dd/MM/yyyy") 
    : "10/02/2023";

  const Pronoun = enfantData.gender === 'M' ? 'Il' : 'Elle';
  const Vu = enfantData.gender === 'M' ? 'vu' : 'vue';

  // --- Création du document ---
  const doc = DocumentApp.create(`CR - ${kidName} - ${rdvDate.replace(/\//g, '-')}`);
  const docUrl = doc.getUrl();
  const file = DriveApp.getFileById(doc.getId());
  DriveApp.getFolderById(CONFIG.driveFolderId).addFile(file);
  DriveApp.getRootFolder().removeFile(file);

  const body = doc.getBody();
  body.clear();
  body.setMarginTop(30);

  // --- 🟦 HEADER ---
  let header = doc.getHeader() || doc.addHeader();
  header.clear();
  
  try {
    const logoBlob = DriveApp.getFileById(CONFIG.LOGO_ID).getAs('image/jpeg');
    const logoPara = header.appendParagraph("");
    const logo = logoPara.appendInlineImage(logoBlob);
    logo.setWidth(75).setHeight(75);
    logoPara.setAlignment(DocumentApp.HorizontalAlignment.CENTER);
    logoPara.setSpacingAfter(0);
  } catch(e) { 
    Logger.log("⚠️ Logo inaccessible : " + e.message);
  }

  const t1 = header.appendParagraph(`Compte rendu de l’entretien – ${kidName}`);
  t1.setHeading(DocumentApp.ParagraphHeading.TITLE);
  t1.setAlignment(DocumentApp.HorizontalAlignment.CENTER);
  t1.setSpacingAfter(0);
  t1.editAsText().setBold(true);

  const t2 = header.appendParagraph(`Date de l’entretien : ${rdvDate}`);
  t2.setHeading(DocumentApp.ParagraphHeading.SUBTITLE);
  t2.setAlignment(DocumentApp.HorizontalAlignment.CENTER);
  t2.setSpacingAfter(0);
  t2.editAsText().setForegroundColor('#808080');

  // --- 🟨 SECTION PROFIL ---
  const profilTable = body.appendTable([["Profil"]]);
  profilTable.setBorderWidth(1).setBorderColor('#DAA520');
  const pCell = profilTable.getCell(0, 0);
  pCell.setBackgroundColor(CONFIG.JAUNE_FONCE).setPaddingTop(0).setPaddingBottom(0);
  
  const pPara = pCell.getChild(0).asParagraph();
  pPara.setHeading(DocumentApp.ParagraphHeading.HEADING2);
  pPara.setSpacingBefore(0).setSpacingAfter(0);
  pPara.editAsText().setBold(true);

  const infoTable = body.appendTable([["", ""]]);
  infoTable.setBorderWidth(0).setColumnWidth(1, 130);
  
  const infoCell = infoTable.getCell(0, 0);
  const namePara = infoCell.getChild(0).asParagraph();
  namePara.setText(`Nom : ${kidName}`);
  namePara.setSpacingBefore(5).setSpacingAfter(1);
  namePara.editAsText().setBold(true);

  infoCell.appendParagraph(`Âge : ${enfantData.age} ans`).setSpacingAfter(1);
  infoCell.appendParagraph(`Classe : ${rdvRow[CONFIG.COLS.RDV.NIVEAU - 1]}`).setSpacingAfter(1);
  infoCell.appendParagraph(`Domaine : ${rdvRow[CONFIG.COLS.RDV.DOMAINE - 1]}`).setSpacingAfter(1);

  if (rdvRow[CONFIG.COLS.RDV.PHOTO - 1]) {
    try {
      const photoBlob = getImageBlobFromUrl(rdvRow[CONFIG.COLS.RDV.PHOTO - 1]);
      if (photoBlob) {
        const photoPara = infoTable.getCell(0, 1).getChild(0).asParagraph();
        const img = photoPara.appendInlineImage(photoBlob);
        img.setWidth(120).setHeight(120); 
        photoPara.setAlignment(DocumentApp.HorizontalAlignment.RIGHT);
      }
    } catch (e) { Logger.log("⚠️ Photo ignorée pour " + kidName); }
  }

  // --- 🟨 SECTION RÉSUMÉ ---
  const resumeTable = body.appendTable([["Résumé de l’entretien"]]);
  resumeTable.setBorderWidth(1).setBorderColor('#DAA520');
  const rCell = resumeTable.getCell(0, 0);
  rCell.setBackgroundColor(CONFIG.JAUNE_FONCE).setPaddingTop(0).setPaddingBottom(0);
  
  const rPara = rCell.getChild(0).asParagraph();
  rPara.setHeading(DocumentApp.ParagraphHeading.HEADING2);
  rPara.setSpacingBefore(0).setSpacingAfter(0);
  rPara.editAsText().setBold(true);

  const r1 = body.appendParagraph(`${kidName} a été ${Vu} par ${rdvRow[CONFIG.COLS.RDV.VU_PAR - 1]} le ${rdvDate}.`);
  r1.setSpacingBefore(6).setSpacingAfter(2);

  body.appendParagraph(`${Pronoun} a ${enfantData.age} ans.`).setSpacingAfter(2);
  body.appendParagraph(`${Pronoun} est actuellement en classe de ${rdvRow[CONFIG.COLS.RDV.NIVEAU - 1]} et étudie ${rdvRow[CONFIG.COLS.RDV.DOMAINE - 1]}.`).setSpacingAfter(2);

  if (rdvRow[CONFIG.COLS.RDV.ENGLISH_SKILLS - 1]) { body.appendParagraph(`Niveau d'anglais : ${rdvRow[CONFIG.COLS.RDV.ENGLISH_SKILLS - 1]}`).setSpacingAfter(2); }
  if (rdvRow[CONFIG.COLS.RDV.GENERAL_CONDITION - 1]) {
    body.appendParagraph(`Condition générale : ${rdvRow[CONFIG.COLS.RDV.GENERAL_CONDITION - 1]}`).setSpacingAfter(2);
  }
  if (rdvRow[CONFIG.COLS.RDV.FAMILY - 1]) { body.appendParagraph(`Concernant sa famille, ${rdvRow[CONFIG.COLS.RDV.FAMILY - 1]}`).setSpacingAfter(2); }
  if (rdvRow[CONFIG.COLS.RDV.OTHER - 1]) { body.appendParagraph(`Par ailleurs, ${rdvRow[CONFIG.COLS.RDV.OTHER - 1]}`).setSpacingAfter(2); }

  const msgText = `Votre présence et votre soutien sont précieux pour ${kidName}.\nMerci pour votre engagement à nos côtés.`;
  const msgPara = body.appendParagraph(msgText);
  msgPara.setAlignment(DocumentApp.HorizontalAlignment.LEFT);
  msgPara.setSpacingBefore(30);
  msgPara.editAsText().setItalic(true);

  let footer = doc.getFooter() || doc.addFooter();
  footer.clear();
  const fPara = footer.appendParagraph("Les Enfants de Pondypatch");
  fPara.setAlignment(DocumentApp.HorizontalAlignment.CENTER);
  fPara.setSpacingAfter(0);
  fPara.editAsText().setBold(true);
  footer.appendParagraph("lesenfantsdepondypatch@gmail.com").setAlignment(DocumentApp.HorizontalAlignment.CENTER);

  doc.saveAndClose();
  
  // Update Spreadsheet
  const sheetRef = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Rdv_suivi");
  sheetRef.getRange(rowNumber, CONFIG.COLS.RDV.DERNIER_CR).setValue(doc.getUrl());
  sheetRef.getRange(rowNumber, CONFIG.COLS.RDV.STATUT_CR).setValue('CR généré');
  ecrireLienEnvoiDansSheet(sheetRef, rowNumber, kidName);

  // Lien prérempli formulaire
  const formUrl = `https://docs.google.com/forms/d/e/${CONFIG.formId}/viewform?usp=pp_url&${CONFIG.entryIdLienCR}=${encodeURIComponent(docUrl)}`;


  Logger.log('✅ CR généré pour ' + kidName);
}

// ============================================================
// SIDEBAR — Ouvrir, vérifier et envoyer le CR au parrain
// ============================================================

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu("📋 CR Enfants")
    .addItem("Générer les CR manquants", "genererCRPourCetteLigne")
    .addItem("📨 Envoyer le CR de cette ligne", "ouvrirSidebarDepuisSelection")
    .addToUi();
}

// Appelé par le lien dans la colonne — ouvre la sidebar pour la ligne donnée
function ouvrirSidebarEnvoi(rowNumber) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CONFIG.SHEET_RDV);
  const row = sheet.getRange(rowNumber, 1, 1, sheet.getLastColumn()).getValues()[0];

  const kidName   = row[CONFIG.COLS.RDV.KID_NAME   - 1];
  const kidId     = row[CONFIG.COLS.RDV.KID_ID     - 1];
  const sponsorId = row[CONFIG.COLS.RDV.SPONSOR_ID - 1];
  const crUrl     = row[CONFIG.COLS.RDV.DERNIER_CR - 1];

  if (!crUrl) {
    SpreadsheetApp.getUi().alert("⚠️ Aucun CR trouvé sur cette ligne.");
    return;
  }

  const parrainData = chercherParrain(sponsorId);
  const parrainEmail = parrainData ? parrainData.email : "";
  const parrainNom   = parrainData ? parrainData.name  : "Parrain/Marraine";

  const defaultSubject = `Compte rendu de l'entretien – ${kidName}`;
  const defaultBody = `Bonjour ${parrainNom},\n\nVeuillez trouver ci-dessous le lien vers le compte rendu de l'entretien de ${kidName}.\n\n📄 Lire le CR : ${crUrl}\n\nMerci pour votre soutien et votre engagement à nos côtés.\n\nCordialement,\nLes Enfants de Pondypatch`;

  const html = HtmlService.createHtmlOutput(`
    <!DOCTYPE html>
    <html>
    <head>
      <style>
        * { box-sizing: border-box; font-family: Arial, sans-serif; }
        body { margin: 0; padding: 16px; background: #f9f9f9; }
        h2 { font-size: 15px; color: #333; margin-bottom: 4px; }
        .subtitle { font-size: 12px; color: #888; margin-bottom: 16px; }

        .card {
          background: white; border-radius: 8px; padding: 14px;
          margin-bottom: 12px; border: 1px solid #e0e0e0;
        }
        label { font-size: 12px; font-weight: bold; color: #555; display: block; margin-bottom: 4px; }
        input, textarea {
          width: 100%; padding: 8px; border: 1px solid #ccc;
          border-radius: 6px; font-size: 13px; color: #333;
        }
        textarea { resize: vertical; min-height: 130px; }
        input:focus, textarea:focus { border-color: #4285f4; outline: none; }

        .btn-doc {
          display: block; width: 100%; padding: 10px;
          background: #4285f4; color: white; border: none;
          border-radius: 6px; font-size: 13px; font-weight: bold;
          text-align: center; text-decoration: none;
          margin-bottom: 12px; cursor: pointer;
        }
        .btn-send {
          display: block; width: 100%; padding: 12px;
          background: #34a853; color: white; border: none;
          border-radius: 6px; font-size: 14px; font-weight: bold;
          cursor: pointer; margin-top: 4px;
        }
        .btn-send:hover { background: #2d9247; }
        .btn-send:disabled { background: #aaa; cursor: not-allowed; }
        .status { margin-top: 10px; font-size: 13px; text-align: center; }
        .ok  { color: #34a853; font-weight: bold; }
        .err { color: #e53935; font-weight: bold; }
      </style>
    </head>
    <body>
      <h2>📋 CR de ${kidName}</h2>
      <p class="subtitle">Ligne ${rowNumber}</p>

      <a class="btn-doc" href="${crUrl}" target="_blank">📝 Ouvrir et modifier le CR</a>

      <div class="card">
        <label>Destinataire</label>
        <input id="email" type="email" value="${parrainEmail}" />
      </div>

      <div class="card">
        <label>Objet</label>
        <input id="subject" type="text" value="${defaultSubject}" />
        <label style="margin-top:10px;">Message</label>
        <textarea id="body">${defaultBody}</textarea>
      </div>

      <button class="btn-send" id="sendBtn" onclick="envoyer()">✉️ Envoyer au parrain</button>
      <div class="status" id="status"></div>

      <script>
        function envoyer() {
          const btn = document.getElementById('sendBtn');
          const status = document.getElementById('status');
          btn.disabled = true;
          btn.textContent = "Envoi en cours…";
          status.innerHTML = "";

          const email   = document.getElementById('email').value.trim();
          const subject = document.getElementById('subject').value.trim();
          const body    = document.getElementById('body').value.trim();

          if (!email) {
            status.innerHTML = '<span class="err">⚠️ Email manquant</span>';
            btn.disabled = false;
            btn.textContent = "✉️ Envoyer au parrain";
            return;
          }

          google.script.run
            .withSuccessHandler(function() {
              btn.textContent = "✅ Envoyé !";
              status.innerHTML = '<span class="ok">Email envoyé avec succès.</span>';
            })
            .withFailureHandler(function(err) {
              btn.disabled = false;
              btn.textContent = "✉️ Envoyer au parrain";
              status.innerHTML = '<span class="err">❌ Erreur : ' + err.message + '</span>';
            })
            .envoyerCRParrain(${rowNumber}, email, subject, body);
        }
      </script>
    </body>
    </html>
  `)
  .setTitle(`Envoyer CR — ${kidName}`)
  .setWidth(340);

  SpreadsheetApp.getUi().showSidebar(html);
}

// Fonction appelée par le bouton HTML — envoie l'email au parrain
function envoyerCRParrain(rowNumber, email, subject, bodyText) {
  MailApp.sendEmail({
    to: CONFIG.MODE_TEST ? CONFIG.EMAIL_TEST : email,
    subject: `${CONFIG.MODE_TEST ? "TEST — " : ""}${subject}`,
    body: bodyText
  });

  // Mise à jour du statut dans le sheet
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CONFIG.SHEET_RDV);
  sheet.getRange(rowNumber, CONFIG.COLS.RDV.STATUT_CR).setValue("CR envoyé");
}

// ============================================================
// À appeler dans genererCRProtege() — insère le lien dans la colonne
// ============================================================

  // Formule qui ouvre la sidebar via un lien custom menu
  // Comme on ne peut pas appeler directement une fonction depuis une cellule,
  // on met un =HYPERLINK vers un Apps Script Web App OU on utilise une note + un raccourci.
  // Solution simple et fiable : on met le numéro de ligne dans une cellule dédiée
  // et l'utilisateur clique sur "Envoyer ce CR" dans le menu après avoir sélectionné la ligne.

  // On génère un lien affichant le nom, qui au clic sélectionne la ligne (visuellement)

function ecrireLienEnvoiDansSheet(sheetRef, rowNumber, kidName) {
  sheetRef.getRange(rowNumber, CONFIG.COLS.RDV.LIEN_ENVOI)
    .setValue("📨 Envoyer")
    .setNote(`Row:${rowNumber}`)
    .setFontColor("#1a73e8")
    .setFontWeight("bold");
}


/**
 * ===============================================
 * Validation CR
 * ===============================================
 */



function envoi_CR_Parrain() {
  Utilities.sleep(2000);
  try {
    Logger.log("=== Début Envoi_CR_Parrain ===");

    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const validationSheet = ss.getSheetByName(CONFIG.SHEET_VALIDATION_CR);
    const rdvSheet        = ss.getSheetByName(CONFIG.SHEET_RDV);

    if (!validationSheet || !rdvSheet) {
      Logger.log("❌ Feuilles manquantes");
      return;
    }

    const lastRow = validationSheet.getLastRow();
    for (let row = 2; row <= lastRow; row++) { // commence à 2 si headers
      const data = validationSheet.getRange(row, 1, 1, validationSheet.getLastColumn()).getValues()[0];

      // 🔹 Nouvelle ligne ? => colonne "CR envoyé?" vide
      const crStatutCell = data[CONFIG.COLS.VALIDATION.CR_ENVOYE - 1];
      if (crStatutCell) continue; // déjà traité


      const crUrl = data[CONFIG.COLS.VALIDATION.LIEN_CR - 1];
      Logger.log(`URL brute : "${crUrl}"`)
      const crId  = extractDocId(crUrl);
      Logger.log(`ID extrait : "${crId}"`);
      if (!crId) continue;
   
     
      // 🔹 Trouver ligne correspondante dans Rdv_suivi
      const rdvData = rdvSheet.getDataRange().getValues();
      let targetRow = -1;
      for (let i = 1; i < rdvData.length; i++) {
        const rowCrId = extractDocId(rdvData[i][CONFIG.COLS.RDV.DERNIER_CR - 1]);
        if (rowCrId === crId) {
          targetRow = i + 1;
          break;
        }
      }
      if (targetRow === -1) continue;

      // 🔹 Vérifier anti double envoi
      const statutActuel = rdvData[targetRow - 1][CONFIG.COLS.RDV.STATUT_CR - 1];
      if (statutActuel === "CR envoyé") continue;

      // 🔹 Infos enfant / parrain
      const sponsorId = rdvData[targetRow - 1][CONFIG.COLS.RDV.SPONSOR_ID - 1];
      const kidName   = rdvData[targetRow - 1][CONFIG.COLS.RDV.KID_NAME - 1];
      const parrain   = chercherParrain(sponsorId);
      if (!parrain) continue;

      // 🔹 Générer PDF et envoyer email
    
      const pdfBlob = DriveApp.getFileById(crId)
        .getAs("application/pdf")
        .setName(`CR - ${kidName}.pdf`);

      MailApp.sendEmail({
        to: CONFIG.MODE_TEST ? CONFIG.EMAIL_TEST : parrain.email,
        subject: `${CONFIG.MODE_TEST ? " TEST — " : ""}Compte-rendu de ${kidName}`,
        htmlBody: `<p>Bonjour,</p><p>Veuillez trouver ci-joint le compte-rendu concernant <strong>${kidName}</strong>.</p><p>Cordialement</p>`,
        attachments: [pdfBlob]
      });

      // 🔹 Marquer CR comme envoyé dans Validation_CR et Rdv_suivi
      validationSheet.getRange(row, CONFIG.COLS.VALIDATION.CR_ENVOYE).setValue("Oui");
      rdvSheet.getRange(targetRow, CONFIG.COLS.RDV.STATUT_CR).setValue("CR envoyé");

      Logger.log(`✅ CR envoyé à ${parrain.email} pour ${kidName} (ligne ${targetRow})`);
    }

  } catch (err) {
    Logger.log("❌ Erreur Envoi_CR_Parrain : " + err);
  }
}




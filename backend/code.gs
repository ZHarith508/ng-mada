// ============================================================
//  SUIVI APPORTS ASSOCIÉS — Backend
//  VERSION FINALE AVEC SYNCHRONISATION COMPLÈTE
// ============================================================

const CONFIG = {
  SHEET_MEMBRE: 'MEMBRE',
  SHEET_APPORTS: 'Apports des Associés',
  SHEET_FACTURE: 'FACTURE',
  SHEET_FINANCE: 'Gestion Financière',
  SHEET_DASHBOARD: 'Tableau de bord',
  MEMBRE_ID_PREFIX: 'M',
  MEMBRE_ID_PAD: 4,
  STATUT_NOUVEAU: 'Nouveau - à vérifier',
  STATUT_VALIDE: 'Associé validé',
  STATUT_PARTIEL: 'Partiellement versé',
  FONCTION_DEFAUT: 'Membre simple',
  SEUIL_VALIDATION_DEFAUT: 1000000,
  DRIVE_FOLDER_PHOTOS: '11I5ajvNiu0qRgtCChHtIWxEU122bFmrU',
  DRIVE_FOLDER_PHOTOS_PROFIL: '11I5ajvNiu0qRgtCChHtIWxEU122bFmrU',
  DRIVE_FOLDER_AFFICHES_PHOTOS: '1-1aYDgllz0sCWShD4w1a0V1fW7i7y9t3',
  DRIVE_FOLDER_AFFICHES_PDF: '17Rf2ir8kliEcPGLLtieaSMBeEAatJ8sQ',
  DRIVE_FOLDER_JUSTIFICATIFS_DEBIT: '1NRz3Ap6QV-LrbnGnURYEN5iKQP2Aon6V',
  DRIVE_FOLDER_JUSTIFICATIFS_CREDIT: '1jbWcBfMtc_lS4T9kn37EpEwqe2SUr7Eb',
  SHEET_UTILISATEURS: 'Utilisateures',
  SESSION_DUREE_SEC: 6 * 60 * 60,
};

const FACTURE_SHEET_NAME = CONFIG.SHEET_FACTURE;
const COL = {
  INVOICE: 1, TIMESTAMP: 2, EMAIL: 3, SCORE: 4, NOM: 5, CONTACT: 6, PERIODE: 7,
  MONTANT: 8, MODE_PAIEMENT: 9, REF_PAIEMENT: 10, DATE_PAIEMENT: 11, RECU_PAR: 12,
  LINK_QR: 13, QR_CODE: 14, MERGED_DOC_ID: 15, MERGED_DOC_URL: 16,
  LINK_MERGED_DOC: 17, MERGE_STATUS: 18, VALIDATION: 19
};
const VALIDATION_VALUE = 'VALIDER';
const TEMPLATE_ID = '1pvlcZuUmn0vn_zkSLdPKNvadYUD4BQZfFaX7m-wpsTI';
const OUTPUT_FOLDER_ID = '1NRz3Ap6QV-LrbnGnURYEN5iKQP2Aon6V';
const ADMIN_EMAIL = 'tovoherinirinaeliceronaldo@gmail.com';
const TO_EMAIL = 'zzaidharith@gmail.com';

// ============================================================
//  SECRETS VIA PROPERTYSERVICE
// ============================================================
function verifierPin(pin) {
  const adminPin = PropertiesService.getScriptProperties().getProperty('PIN_ADMIN');
  return String(pin || '').trim() === String(adminPin);
}
function verifierPinOuErreur(pin) {
  if (!verifierPin(pin)) throw new Error('PIN incorrect. Action refusée.');
}
function hasherMotDePasse(motDePasse) {
  const pepper = PropertiesService.getScriptProperties().getProperty('PASSWORD_PEPPER') || '';
  const contenu = pepper + '::' + String(motDePasse || '');
  const octets = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, contenu, Utilities.Charset.UTF_8);
  return octets.map(function (o) { return ('0' + (o & 0xFF).toString(16)).slice(-2); }).join('');
}

// ============================================================
//  AUTHENTIFICATION UTILISATEUR
// ============================================================
function getUtilisateursSheet_() {
  return SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CONFIG.SHEET_UTILISATEURS);
}
function inscrireUtilisateur(nom, email, motDePasse) {
  try {
    nom = String(nom || '').trim();
    const emailNorm = normalizeEmail(email);
    motDePasse = String(motDePasse || '');
    if (!nom || !emailNorm || !motDePasse) return { success: false, message: 'Merci de remplir tous les champs.' };
    if (motDePasse.length < 6) return { success: false, message: 'Le mot de passe doit contenir au moins 6 caractères.' };
    if (!/^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(emailNorm)) return { success: false, message: 'Adresse e-mail invalide.' };
    const sheet = getUtilisateursSheet_();
    if (!sheet) return { success: false, message: 'Onglet "Utilisateures" introuvable.' };
    const header = getHeaderMap(sheet);
    const headerRow = header.__headerRow || 1;
    const lastRow = sheet.getLastRow();
    let dernierNumero = 0;
    if (lastRow > headerRow) {
      const data = sheet.getRange(headerRow + 1, 1, lastRow - headerRow, sheet.getLastColumn()).getValues();
      for (let i = 0; i < data.length; i++) {
        const emailExistant = normalizeEmail(getCell(data[i], header, ['email']));
        if (emailExistant && emailExistant === emailNorm) return { success: false, message: 'Un compte existe déjà avec cette adresse e-mail.' };
        const n = parseInt(getCell(data[i], header, ['n°', 'numero', 'n']), 10);
        if (!isNaN(n) && n > dernierNumero) dernierNumero = n;
      }
    }
    const ligne = new Array(sheet.getLastColumn()).fill('');
    setCell(ligne, header, ['n°', 'numero', 'n'], dernierNumero + 1);
    setCell(ligne, header, ['nom'], nom);
    setCell(ligne, header, ['email'], emailNorm);
    setCell(ligne, header, ['mot de passe (hash)', 'mot de passe', 'hash'], hasherMotDePasse(motDePasse));
    setCell(ligne, header, ['date de création', 'date de creation'], new Date());
    sheet.appendRow(ligne);
    return { success: true };
  } catch (err) {
    logErreur('inscrireUtilisateur', err);
    return { success: false, message: err.message };
  }
}
function connecterUtilisateur(email, motDePasse) {
  try {
    const emailNorm = normalizeEmail(email);
    if (!emailNorm || !motDePasse) return { success: false, message: 'Merci de saisir votre e-mail et votre mot de passe.' };
    const sheet = getUtilisateursSheet_();
    if (!sheet) return { success: false, message: 'Onglet "Utilisateures" introuvable.' };
    const header = getHeaderMap(sheet);
    const headerRow = header.__headerRow || 1;
    const lastRow = sheet.getLastRow();
    if (lastRow <= headerRow) return { success: false, message: 'E-mail ou mot de passe incorrect.' };
    const data = sheet.getRange(headerRow + 1, 1, lastRow - headerRow, sheet.getLastColumn()).getValues();
    const hashSaisi = hasherMotDePasse(motDePasse);
    for (let i = 0; i < data.length; i++) {
      const emailLigne = normalizeEmail(getCell(data[i], header, ['email']));
      if (emailLigne !== emailNorm) continue;
      const hashStocke = String(getCell(data[i], header, ['mot de passe (hash)', 'mot de passe', 'hash']) || '');
      if (hashStocke !== hashSaisi) return { success: false, message: 'E-mail ou mot de passe incorrect.' };
      const nom = getCell(data[i], header, ['nom']);
      const token = Utilities.getUuid();
      CacheService.getScriptCache().put(
        'session_' + token,
        JSON.stringify({ nom: nom, email: emailNorm }),
        CONFIG.SESSION_DUREE_SEC
      );
      setCellDansLigne(sheet, i + headerRow + 1, header, ['date et heure connecté', 'date et heure connecte'], new Date());
      return { success: true, token: token, nom: nom, email: emailNorm };
    }
    return { success: false, message: 'E-mail ou mot de passe incorrect.' };
  } catch (err) {
    logErreur('connecterUtilisateur', err);
    return { success: false, message: err.message };
  }
}
function verifierSession(token) {
  try {
    if (!token) return { success: false };
    const brut = CacheService.getScriptCache().get('session_' + token);
    if (!brut) return { success: false };
    const infos = JSON.parse(brut);
    return { success: true, nom: infos.nom, email: infos.email };
  } catch (err) { return { success: false }; }
}
function deconnecterUtilisateur(token) {
  try { if (token) CacheService.getScriptCache().remove('session_' + token); return { success: true }; } catch (err) { return { success: true }; }
}

// ============================================================
//  ROUTAGE PRINCIPAL (doGet)
// ============================================================
function doGet(e) {
  const params = (e && e.parameter) || {};
  if (params.verify) {
    return pageVerificationMembre(params.verify);
  }
  if (params.id) {
    return HtmlService.createHtmlOutputFromFile('PhotoUpload')
      .setTitle('Ma photo — NG-MADA')
      .addMetaTag('viewport', 'width=device-width, initial-scale=1');
  }
  if (params.reset) {
    const token = params.reset;
    const cached = CacheService.getScriptCache().get('reset_' + token);
    if (!cached) {
      return HtmlService.createHtmlOutput(
        '<html><head><meta charset="UTF-8"><meta name="viewport" content="width=device-width, initial-scale=1">' +
        '<style>body{font-family:Arial,sans-serif;background:#0f172a;color:#f1f5f9;text-align:center;padding:40px 20px;}' +
        'h1{color:#f87171;font-size:24px;}.box{background:#1e293b;border-radius:16px;padding:30px;max-width:400px;margin:0 auto;border:1px solid #334155;}' +
        'a{color:#facc15;text-decoration:none;}.btn{display:inline-block;margin-top:16px;padding:10px 24px;background:#facc15;color:#0f172a;border-radius:8px;text-decoration:none;font-weight:700;}' +
        '</style></head><body>' +
        '<div class="box"><h1>❌ Lien expiré ou invalide</h1>' +
        '<p>Le lien de réinitialisation a expiré ou n\'est plus valide.</p>' +
        '<p>Vous pouvez faire une nouvelle demande depuis la page de connexion.</p>' +
        '<a href="' + ScriptApp.getService().getUrl() + '" class="btn">🔐 Retourner à la connexion</a>' +
        '</div></body></html>'
      ).setTitle('Lien expiré — NG-MADA');
    }
    const template = HtmlService.createTemplateFromFile('ResetPassword');
    template.token = token;
    return template.evaluate()
      .setTitle('Nouveau mot de passe — NG-MADA')
      .addMetaTag('viewport', 'width=device-width, initial-scale=1')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
  }
  if (params.session) {
    const session = verifierSession(params.session);
    if (session && session.success) {
      return HtmlService.createHtmlOutputFromFile('Dashboard')
        .setTitle('Suivi Apports Associés — NG-MADA')
        .addMetaTag('viewport', 'width=device-width, initial-scale=1');
    }
  }
  return HtmlService.createHtmlOutputFromFile('Dashboard')
    .setTitle('Suivi Apports Associés — NG-MADA')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

// ============================================================
//  RÉINITIALISATION MOT DE PASSE
// ============================================================
function demanderReinitialisationMotDePasse(email) {
  try {
    const emailNorm = normalizeEmail(email);
    if (!emailNorm) return { success: false, message: 'Adresse e-mail invalide.' };
    const sheet = getUtilisateursSheet_();
    if (!sheet) return { success: false, message: 'Onglet "Utilisateures" introuvable.' };
    const header = getHeaderMap(sheet);
    const headerRow = header.__headerRow || 1;
    const lastRow = sheet.getLastRow();
    if (lastRow <= headerRow) return { success: false, message: 'Aucun compte associé à cette adresse e-mail.' };
    const data = sheet.getRange(headerRow + 1, 1, lastRow - headerRow, sheet.getLastColumn()).getValues();
    let ligneTrouvee = -1, nom = '';
    for (let i = 0; i < data.length; i++) {
      const emailLigne = normalizeEmail(getCell(data[i], header, ['email']));
      if (emailLigne === emailNorm) { ligneTrouvee = i + headerRow + 1; nom = getCell(data[i], header, ['nom']); break; }
    }
    if (ligneTrouvee === -1) return { success: false, message: 'Aucun compte associé à cette adresse e-mail.' };
    const token = Utilities.getUuid();
    const expires = new Date(); expires.setHours(expires.getHours() + 1);
    CacheService.getScriptCache().put('reset_' + token, JSON.stringify({ email: emailNorm, expires: expires.getTime() }), 3600);
    const url = ScriptApp.getService().getUrl();
    const resetLink = url + '?reset=' + encodeURIComponent(token);
    const htmlBody = `
      <div style="font-family:Arial,sans-serif;max-width:500px;margin:0 auto;padding:20px;background:#0f172a;border-radius:16px;color:#f1f5f9;border:1px solid rgba(250,204,21,0.2);">
        <h2 style="color:#facc15;text-align:center;">🔑 Réinitialisation du mot de passe</h2>
        <p>Bonjour <strong>${nom}</strong>,</p>
        <p>Vous avez demandé la réinitialisation de votre mot de passe pour votre compte <strong>NG-MADA</strong>.</p>
        <p style="text-align:center;margin:24px 0;">
          <a href="${resetLink}" style="background:#facc15;color:#0f172a;padding:12px 32px;border-radius:10px;text-decoration:none;font-weight:700;">🔄 Réinitialiser mon mot de passe</a>
        </p>
        <p style="font-size:12px;color:#94a3b8;">Ce lien est valable <strong>1 heure</strong>. Si vous n'êtes pas à l'origine de cette demande, ignorez cet email.</p>
        <hr style="border-color:#334155;margin:16px 0;">
        <p style="font-size:11px;color:#475569;text-align:center;">NEXT GENERATION MADA — "NG-MADA"</p>
      </div>
    `;
    MailApp.sendEmail({ to: emailNorm, subject: '🔑 Réinitialisation de votre mot de passe NG-MADA', htmlBody: htmlBody, name: 'NG-MADA' });
    return { success: true };
  } catch (err) {
    logErreur('demanderReinitialisationMotDePasse', err);
    return { success: false, message: err.message };
  }
}
function verifierTokenReinitialisation(token) {
  try {
    if (!token) return { success: false, message: 'Token manquant.' };
    const cached = CacheService.getScriptCache().get('reset_' + token);
    if (!cached) return { success: false, message: 'Token invalide ou expiré.' };
    const data = JSON.parse(cached);
    return { success: true, email: data.email };
  } catch (err) { return { success: false, message: err.message }; }
}
function reinitialiserMotDePasse(token, nouveauMotDePasse) {
  try {
    if (!token) return { success: false, message: 'Token manquant.' };
    if (!nouveauMotDePasse || nouveauMotDePasse.length < 6) return { success: false, message: 'Le mot de passe doit contenir au moins 6 caractères.' };
    const cached = CacheService.getScriptCache().get('reset_' + token);
    if (!cached) return { success: false, message: 'Lien expiré ou invalide.' };
    const data = JSON.parse(cached);
    const emailNorm = data.email;
    CacheService.getScriptCache().remove('reset_' + token);
    const sheet = getUtilisateursSheet_();
    const header = getHeaderMap(sheet);
    const headerRow = header.__headerRow || 1;
    const lastRow = sheet.getLastRow();
    const dataRows = sheet.getRange(headerRow + 1, 1, lastRow - headerRow, sheet.getLastColumn()).getValues();
    const colEmail = colIndexOf(header, ['email']);
    const colHash = colIndexOf(header, ['mot de passe (hash)', 'mot de passe', 'hash']);
    for (let i = 0; i < dataRows.length; i++) {
      const emailLigne = normalizeEmail(dataRows[i][colEmail - 1]);
      if (emailLigne === emailNorm) {
        const row = i + headerRow + 1;
        const nouveauHash = hasherMotDePasse(nouveauMotDePasse);
        sheet.getRange(row, colHash).setValue(nouveauHash);
        return { success: true };
      }
    }
    return { success: false, message: 'Utilisateur introuvable.' };
  } catch (err) {
    logErreur('reinitialiserMotDePasse', err);
    return { success: false, message: err.message };
  }
}

// ============================================================
//  UTILITAIRES D'ÉCRITURE OPTIMISÉE
// ============================================================
function ecrireLigne(sheet, row, header, valeurs) {
  const lastCol = sheet.getLastColumn();
  const range = sheet.getRange(row, 1, 1, lastCol);
  const rowData = range.getValues()[0];
  for (const [clef, valeur] of Object.entries(valeurs)) {
    const col = colIndexOf(header, [clef]);
    if (col) rowData[col - 1] = valeur;
  }
  range.setValues([rowData]);
}

// ============================================================
//  FACTURE (Apports) — génération PDF + doublon
// ============================================================
function onFormSubmit(e) { try { notifierNouveauPaiementRecu(e); } catch (err) { logErreur('onFormSubmit', err); } }
function notifierNouveauPaiementRecu(e) {
  if (!e || !e.range) return;
  const sheet = e.range.getSheet();
  if (sheet.getName() !== FACTURE_SHEET_NAME) return;
  if (!ADMIN_EMAIL) return;
  const row = e.range.getRow();
  const rowValues = sheet.getRange(row, 1, 1, sheet.getLastColumn()).getValues()[0];
  const nom = rowValues[COL.NOM - 1];
  const contact = rowValues[COL.CONTACT - 1];
  const montant = rowValues[COL.MONTANT - 1];
  const modePaie = rowValues[COL.MODE_PAIEMENT - 1];
  const refPaie = rowValues[COL.REF_PAIEMENT - 1];
  const datePaie = rowValues[COL.DATE_PAIEMENT - 1];
  const recuPar = rowValues[COL.RECU_PAR - 1];
  ecrireLienQR(sheet, row, '');
  const sujet = '💰 Nouveau paiement à valider — ' + (nom || 'Sans nom');
  const corps = 'Un nouveau paiement vient d\'arriver via le formulaire (ligne ' + row + '), en attente de validation.\n\n' +
    'Nom : ' + (nom || '—') + '\nContact : ' + (contact || '—') + '\nMontant : ' + (montant || '—') +
    '\nMode de paiement : ' + (modePaie || '—') + '\nRéférence : ' + (refPaie || '—') +
    '\nDate : ' + (datePaie || '—') + '\nReçu par : ' + (recuPar || '—') + '\n\nOuvrez le Dashboard pour valider.';
  MailApp.sendEmail(ADMIN_EMAIL, sujet, corps);
}

// ============================================================
//  GESTION DES NUMÉROS DE FACTURE
// ============================================================
function genererNumeroFacture() {
  var scriptProps = PropertiesService.getScriptProperties();
  var now = new Date();
  var annee = now.getFullYear();
  var dernierNumero = 0;
  var dernier = scriptProps.getProperty('DERNIER_NUMERO_FACTURE');
  if (dernier) {
    var parties = dernier.split('-');
    if (parties.length === 3) {
      var anneeStockee = parseInt(parties[1]);
      var seq = parseInt(parties[2]);
      if (anneeStockee < annee) {
        dernierNumero = 0;
      } else {
        dernierNumero = seq;
      }
    }
  }
  var nouveauNumero = dernierNumero + 1;
  var numFacture = 'NG-' + annee + '-' + String(nouveauNumero).padStart(5, '0');
  scriptProps.setProperty('DERNIER_NUMERO_FACTURE', numFacture);
  Logger.log('📄 Nouveau numéro de facture généré : ' + numFacture);
  return numFacture;
}

function getProchainNumeroFacture() {
  var scriptProps = PropertiesService.getScriptProperties();
  var now = new Date();
  var annee = now.getFullYear();
  var dernierNumero = 0;
  var dernier = scriptProps.getProperty('DERNIER_NUMERO_FACTURE');
  if (dernier) {
    var parties = dernier.split('-');
    if (parties.length === 3) {
      var anneeStockee = parseInt(parties[1]);
      var seq = parseInt(parties[2]);
      if (anneeStockee < annee) {
        dernierNumero = 0;
      } else {
        dernierNumero = seq;
      }
    }
  }
  return 'NG-' + annee + '-' + String(dernierNumero + 1).padStart(5, '0');
}

function verifierNumeroFactureExistant(financeSheet, numFacture) {
  if (!financeSheet || !numFacture) return false;
  const header = getHeaderMap(financeSheet);
  const headerRow = header.__headerRow || 1;
  const lastRow = financeSheet.getLastRow();
  if (lastRow <= headerRow) return false;
  const colNumFacture = colIndexOf(header, ['n° facture', 'numero facture']);
  if (!colNumFacture) return false;
  const data = financeSheet.getRange(headerRow + 1, colNumFacture, lastRow - headerRow, 1).getValues();
  for (let i = 0; i < data.length; i++) {
    const num = String(data[i][0] || '').trim();
    if (num === String(numFacture).trim()) {
      return true;
    }
  }
  return false;
}

// ============================================================
//  VALIDATION FACTURE - CORRIGÉE
// ============================================================
function validerFactureWeb(pin, row, numFacture) {
  try {
    verifierPinOuErreur(pin);
    row = parseInt(row, 10);
    if (!row || row < 2) return { success: false, message: 'Ligne invalide.' };
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(FACTURE_SHEET_NAME);
    if (!sheet) return { success: false, message: 'Onglet FACTURE introuvable.' };
    const validationActuelle = String(sheet.getRange(row, COL.VALIDATION).getValue()).trim().toUpperCase();
    if (validationActuelle.indexOf('DOUBLON') !== -1) {
      return { success: false, message: 'Cette ligne est marquée comme doublon.' };
    }
    if (sheet.getRange(row, COL.INVOICE).getValue() !== '') {
      return { success: false, message: 'Cette facture a déjà été traitée.' };
    }
    var numeroFinal = numFacture || genererNumeroFacture();
    Logger.log('📄 Validation facture ligne ' + row + ' → ' + numeroFinal);
    const financeSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CONFIG.SHEET_FINANCE);
    var numeroExistant = verifierNumeroFactureExistant(financeSheet, numeroFinal);
    if (numeroExistant) {
      const header = getHeaderMap(sheet);
      ecrireLigne(sheet, row, header, {
        'validation': 'DOUBLON - N° ' + numeroFinal + ' existe déjà dans la caisse'
      });
      return { 
        success: false, 
        message: '⚠️ Doublon détecté : le numéro ' + numeroFinal + ' existe déjà dans la caisse.',
        doublon: true
      };
    }
    const rowValues = sheet.getRange(row, 1, 1, sheet.getLastColumn()).getValues()[0];
    const nom = rowValues[COL.NOM - 1];
    const contact = rowValues[COL.CONTACT - 1];
    const montant = rowValues[COL.MONTANT - 1];
    const refPaiement = rowValues[COL.REF_PAIEMENT - 1];
    const texteQr = [numeroFinal, nom || '', contact || '', montant || '', refPaiement || '']
      .filter(function (v) { return v !== '' && v !== null && v !== undefined; })
      .join(', ');
    const qrUrl = 'https://quickchart.io/qr?text=' + encodeURIComponent(texteQr) + '&size=300';
    const header = getHeaderMap(sheet);
    ecrireLigne(sheet, row, header, {
      'n° facture': numeroFinal,
      'validation': VALIDATION_VALUE,
      'lien qr': qrUrl
    });
    SpreadsheetApp.flush();
    generateAndSendInvoice(sheet, row, numeroFinal);
    try {
      traiterLigneFactureApport(sheet, row, numeroFinal);
    } catch (syncErr) {
      Logger.log('ERREUR SYNCHRO : ' + syncErr.message);
      ecrireStatutSync(sheet, row, '⚠️ Erreur synchronisation Apports : ' + syncErr.message);
    }
    ngmadaSynchroniserTableauDeBord();
    return { 
      success: true, 
      numFacture: numeroFinal,
      message: 'Facture validée avec succès'
    };
  } catch (err) {
    logErreur('validerFactureWeb', err);
    return { success: false, message: err.message };
  }
}

function ecrireLienQR(sheet, row, invoiceNumber) {
  const rowValues = sheet.getRange(row, 1, 1, sheet.getLastColumn()).getValues()[0];
  const nom = rowValues[COL.NOM - 1];
  const contact = rowValues[COL.CONTACT - 1];
  const montant = rowValues[COL.MONTANT - 1];
  const refPaiement = rowValues[COL.REF_PAIEMENT - 1];
  const texteQr = [invoiceNumber || '', nom, contact, montant, refPaiement]
    .filter(function (v) { return v !== '' && v !== null && v !== undefined; })
    .join(', ');
  const url = 'https://quickchart.io/qr?text=' + encodeURIComponent(texteQr) + '&size=300';
  sheet.getRange(row, COL.LINK_QR).setValue(url);
  return url;
}

function detecterDoublonFacture(sheet, row, rowValues) {
  const nom = normaliserNomComplet(rowValues[COL.NOM - 1]);
  const contact = normalizePhone(rowValues[COL.CONTACT - 1]);
  const montant = toNombre(rowValues[COL.MONTANT - 1]);
  const ref = String(rowValues[COL.REF_PAIEMENT - 1] || '').trim().toUpperCase();
  if (!nom || !montant) return null;
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return null;
  const data = sheet.getRange(2, 1, lastRow - 1, sheet.getLastColumn()).getValues();
  for (let i = 0; i < data.length; i++) {
    const r = i + 2;
    if (r === row) continue;
    const inv = String(data[i][COL.INVOICE - 1] || '').trim();
    if (!inv) continue;
    const nom2 = normaliserNomComplet(data[i][COL.NOM - 1]);
    const contact2 = normalizePhone(data[i][COL.CONTACT - 1]);
    const montant2 = toNombre(data[i][COL.MONTANT - 1]);
    const ref2 = String(data[i][COL.REF_PAIEMENT - 1] || '').trim().toUpperCase();
    const memeRef = ref && ref2 && ref !== '0000' && ref !== '00000' && ref === ref2;
    const memePersonneMontant = nom === nom2 && contact === contact2 && montant === montant2;
    if (memeRef || memePersonneMontant) return { numFacture: inv };
  }
  return null;
}

function generateAndSendInvoice(sheet, row, invoiceNumber) {
  const rowData = sheet.getRange(row, 1, 1, sheet.getLastColumn()).getValues()[0];
  const data = {
    invoice: invoiceNumber,
    horodateur: rowData[COL.TIMESTAMP - 1],
    email: rowData[COL.EMAIL - 1],
    nom: rowData[COL.NOM - 1],
    contact: rowData[COL.CONTACT - 1],
    periode: rowData[COL.PERIODE - 1],
    montant: rowData[COL.MONTANT - 1],
    modePaiement: rowData[COL.MODE_PAIEMENT - 1],
    refPaiement: rowData[COL.REF_PAIEMENT - 1],
    datePaiement: rowData[COL.DATE_PAIEMENT - 1],
    recuPar: rowData[COL.RECU_PAR - 1],
    qrUrl: rowData[COL.LINK_QR - 1]
  };
  const templateFile = DriveApp.getFileById(TEMPLATE_ID);
  const outputFolder = DriveApp.getFolderById(OUTPUT_FOLDER_ID);
  const docName = 'NG FACT ' + invoiceNumber;
  const copiedFile = templateFile.makeCopy(docName, outputFolder);
  const doc = DocumentApp.openById(copiedFile.getId());
  const body = doc.getBody();
  const formattedDate = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'dd/MM/yyyy');
  body.replaceText('<<Horodateur>>', formatValue(data.horodateur));
  body.replaceText('<<N° Facture>>', data.invoice);
  body.replaceText('<<NOM>>', formatValue(data.nom));
  body.replaceText('<<Adresse e-mail>>', formatValue(data.email));
  body.replaceText('<<CONTACT>>', formatValue(data.contact));
  body.replaceText('<<PERIODE>>', formatValue(data.periode));
  body.replaceText('<<MONTANT>>', formatMontantAr(data.montant));
  body.replaceText('<<Mode de paiement>>', formatValue(data.modePaiement));
  body.replaceText('<<REF DE PAIEMENT:>>', formatValue(data.refPaiement));
  body.replaceText('<<DATE DE PAIEMENT>>', formatValue(data.datePaiement));
  body.replaceText('<<Reçu par>>', formatValue(data.recuPar));
  body.replaceText('<<TODAY>>', formattedDate);
  if (data.qrUrl) {
    try {
      const qrBlob = UrlFetchApp.fetch(data.qrUrl).getBlob();
      insertImageAtPlaceholder(body, '<<LINK QR CODE>>', qrBlob);
    } catch (imgErr) { body.replaceText('<<LINK QR CODE>>', '[QR code indisponible]'); }
  } else { body.replaceText('<<LINK QR CODE>>', ''); }
  doc.saveAndClose();
  const pdfBlob = DriveApp.getFileById(copiedFile.getId()).getAs(MimeType.PDF);
  pdfBlob.setName(docName + '.pdf');
  const pdfFile = outputFolder.createFile(pdfBlob);
  const recipients = [TO_EMAIL];
  if (data.email) recipients.push(data.email);
  const subject = 'FACTURE ' + invoiceNumber + ' ' + formattedDate;
  const emailBody = 'ASALAM ALAIKUM!\nIty ilay FACTURENAO.\n\nMankasitraka.\n\nCordialement,\n\nNG-MADA';
  GmailApp.sendEmail(recipients.join(','), subject, emailBody, { cc: ADMIN_EMAIL, attachments: [pdfFile.getAs(MimeType.PDF)], name: 'NG-MADA' });
  sheet.getRange(row, COL.MERGED_DOC_ID).setValue(copiedFile.getId());
  sheet.getRange(row, COL.MERGED_DOC_URL).setValue(pdfFile.getUrl());
  sheet.getRange(row, COL.LINK_MERGED_DOC).setValue(docName);
  sheet.getRange(row, COL.MERGE_STATUS).setValue('Document créé et fusionné avec succès ; PDF créé ; Email envoyé à : ' + recipients.join(',') + ' ; ' + new Date().toLocaleString('fr-FR'));
}

function insertImageAtPlaceholder(body, placeholder, imageBlob) {
  const searchResult = body.findText(placeholder);
  if (!searchResult) return;
  const element = searchResult.getElement();
  const parent = element.getParent();
  const type = parent.getType();
  element.asText().setText('');
  const IMAGE_SIZE = 100;
  let insertedImage = null;
  if (type === DocumentApp.ElementType.PARAGRAPH) insertedImage = parent.asParagraph().insertInlineImage(0, imageBlob);
  else if (type === DocumentApp.ElementType.LIST_ITEM) insertedImage = parent.asListItem().insertInlineImage(0, imageBlob);
  else if (type === DocumentApp.ElementType.TABLE_CELL) insertedImage = parent.asTableCell().insertParagraph(0, '').insertInlineImage(0, imageBlob);
  if (insertedImage) { insertedImage.setWidth(IMAGE_SIZE); insertedImage.setHeight(IMAGE_SIZE); }
}

function formatValue(value) { if (value === null || value === undefined || value === '') return ''; if (value instanceof Date) return Utilities.formatDate(value, Session.getScriptTimeZone(), 'dd/MM/yyyy'); return String(value); }
function formatMontantAr(value) { const n = toNombre(value); return n.toLocaleString('fr-FR') + ' Ar'; }

// ============================================================
//  SYNCHRONISATION FACTURE -> APPORTS / MEMBRE / FINANCE
// ============================================================
function traiterLigneFactureApport(factureSheet, row, numFacture) {
  const ss = factureSheet.getParent();
  const rowValues = factureSheet.getRange(row, 1, 1, factureSheet.getLastColumn()).getValues()[0];
  const nom = rowValues[COL.NOM - 1];
  const contact = rowValues[COL.CONTACT - 1];
  const email = rowValues[COL.EMAIL - 1];
  const montant = toNombre(rowValues[COL.MONTANT - 1]);
  const modePaie = rowValues[COL.MODE_PAIEMENT - 1];
  const datePaie = rowValues[COL.DATE_PAIEMENT - 1] || new Date();
  const recuPar = rowValues[COL.RECU_PAR - 1];
  if (!nom && !contact) return;
  const membresSheet = ss.getSheetByName(CONFIG.SHEET_MEMBRE);
  const apportsSheet = ss.getSheetByName(CONFIG.SHEET_APPORTS);
  const financeSheet = ss.getSheetByName(CONFIG.SHEET_FINANCE);
  if (verifierNumeroFactureExistant(financeSheet, numFacture)) {
    ecrireStatutSync(factureSheet, row, '⚠️ Doublon détecté - N° ' + numFacture + ' existe déjà dans la caisse');
    return;
  }
  const index = buildMembresIndexApport(membresSheet);
  const nomNorm = normaliserNomComplet(nom);
  const correspondances = index.byNomComplet.get(nomNorm) || [];
  let membre = null, nouveauMembre = false, homonyme = false;
  if (correspondances.length === 1) {
    membre = correspondances[0];
  } else if (correspondances.length > 1) {
    homonyme = true;
  }
  if (!membre) {
    membre = creerNouveauMembreApport(membresSheet, index, { nom: nom, contact: contact, email: email }, homonyme);
    nouveauMembre = true;
  }
  const statutFinal = mettreAJourLigneApport(apportsSheet, membre, {
    montant: montant, contact: contact, modePaie: modePaie, datePaie: datePaie, numFacture: numFacture
  });
  if (!verifierNumeroFactureExistant(financeSheet, numFacture)) {
    ecrireLigneGestionFinanciere(ss, {
      date: datePaie, nom: membre.nomComplet, idMembre: membre.id, modePaie: modePaie,
      montant: montant, numFacture: numFacture, recuPar: recuPar
    });
  } else {
    ecrireStatutSync(factureSheet, row, '⚠️ Doublon - N° ' + numFacture + ' déjà dans la caisse');
  }
  if (statutFinal === CONFIG.STATUT_VALIDE) {
    mettreAJourStatutMembre(membresSheet, membre.id, CONFIG.STATUT_VALIDE);
  }
  let statutMsg;
  if (homonyme) {
    notifierHomonymeApport(nom, membre.id, correspondances, numFacture);
    statutMsg = '⚠️ Synchronisé — homonyme détecté, nouveau membre ' + membre.id + ' créé par précaution';
  } else if (nouveauMembre) {
    if (ADMIN_EMAIL) {
      MailApp.sendEmail(ADMIN_EMAIL, 'Nouveau membre créé automatiquement - Suivi Apports',
        'Un nouveau membre a été créé depuis la facture ' + numFacture + ' : ' + membre.nomComplet +
        '\nID : ' + membre.id + '\nÀ vérifier dans l\'onglet MEMBRE.');
    }
    statutMsg = 'Synchronisé — nouveau membre ' + membre.id + ' créé + apport enregistré';
  } else {
    statutMsg = 'Synchronisé — apport ajouté pour ' + membre.id + ' (cumul mis à jour' +
      (statutFinal === CONFIG.STATUT_VALIDE ? ', statut : Associé validé' : '') + ')';
  }
  ecrireStatutSync(factureSheet, row, statutMsg);
}

function buildMembresIndexApport(membresSheet) {
  const header = getHeaderMap(membresSheet);
  const headerRow = header.__headerRow || 1;
  const lastRow = membresSheet.getLastRow();
  const byNomComplet = new Map(), byId = new Map();
  let maxIdNum = 0;
  if (lastRow > headerRow) {
    const data = membresSheet.getRange(headerRow + 1, 1, lastRow - headerRow, membresSheet.getLastColumn()).getValues();
    data.forEach(function (r) {
      const id = getCell(r, header, ['id membre', 'id']);
      const nomComplet = String(getCell(r, header, ['nom & prénom', 'nom et prenom', 'nom']) || '').trim();
      const obj = { id: id, nomComplet: nomComplet };
      if (id) byId.set(id, obj);
      const nomNorm = normaliserNomComplet(nomComplet);
      if (nomNorm) {
        if (!byNomComplet.has(nomNorm)) byNomComplet.set(nomNorm, []);
        byNomComplet.get(nomNorm).push(obj);
      }
      const num = parseInt(String(id).replace(/[^0-9]/g, ''), 10);
      if (!isNaN(num) && num > maxIdNum) maxIdNum = num;
    });
  }
  return { byNomComplet: byNomComplet, byId: byId, maxIdNum: maxIdNum, header: header };
}

function creerNouveauMembreApport(membresSheet, index, infos, homonyme) {
  const header = index.header;
  const nextId = CONFIG.MEMBRE_ID_PREFIX + String(index.maxIdNum + 1).padStart(CONFIG.MEMBRE_ID_PAD, '0');
  const ligne = new Array(membresSheet.getLastColumn()).fill('');
  setCell(ligne, header, ['id membre', 'id'], nextId);
  setCell(ligne, header, ['nom & prénom', 'nom et prenom', 'nom'], infos.nom || '');
  setCell(ligne, header, ['téléphone', 'telephone'], infos.contact || '');
  setCell(ligne, header, ['adresse e-mail', 'email', 'e-mail'], infos.email || '');
  setCell(ligne, header, ['fonction'], CONFIG.FONCTION_DEFAUT);
  setCell(ligne, header, ["date d'adhésion", 'date adhesion'], new Date());
  setCell(ligne, header, ['statut'], homonyme ? 'Nouveau - ⚠️ homonyme à vérifier' : CONFIG.STATUT_NOUVEAU);
  membresSheet.appendRow(ligne);
  const obj = { id: nextId, nomComplet: (infos.nom || '').trim() };
  const nomNorm = normaliserNomComplet(infos.nom);
  if (nomNorm) {
    if (!index.byNomComplet.has(nomNorm)) index.byNomComplet.set(nomNorm, []);
    index.byNomComplet.get(nomNorm).push(obj);
  }
  index.byId.set(nextId, obj);
  index.maxIdNum += 1;
  return obj;
}

function mettreAJourLigneApport(apportsSheet, membre, infos) {
  const header = getHeaderMap(apportsSheet);
  const headerRow = header.__headerRow || 1;
  const lastRow = apportsSheet.getLastRow();
  const colIdMembre = colIndexOf(header, ['id membre', 'id']);
  let numeroLigne = -1;
  if (lastRow > headerRow) {
    const ids = apportsSheet.getRange(headerRow + 1, colIdMembre, lastRow - headerRow, 1).getValues();
    for (let i = 0; i < ids.length; i++) {
      if (String(ids[i][0]) === String(membre.id)) { numeroLigne = i + headerRow + 1; break; }
    }
  }
  if (numeroLigne === -1) {
    let dernierNumero = 0;
    if (lastRow > headerRow) {
      const numeros = apportsSheet.getRange(headerRow + 1, colIndexOf(header, ['n°', 'numero', 'n']), lastRow - headerRow, 1).getValues();
      numeros.forEach(function (r) { const n = parseInt(r[0], 10); if (!isNaN(n) && n > dernierNumero) dernierNumero = n; });
    }
    const ligne = new Array(apportsSheet.getLastColumn()).fill('');
    setCell(ligne, header, ['n°', 'numero', 'n'], dernierNumero + 1);
    setCell(ligne, header, ['id membre', 'id'], membre.id);
    setCell(ligne, header, ['nom & prénom', 'nom'], membre.nomComplet);
    setCell(ligne, header, ['téléphone', 'telephone'], infos.contact || '');
    setCell(ligne, header, ["date d'engagement"], new Date());
    apportsSheet.appendRow(ligne);
    numeroLigne = apportsSheet.getLastRow();
  }
  const rowActuelle = apportsSheet.getRange(numeroLigne, 1, 1, apportsSheet.getLastColumn()).getValues()[0];
  const apportPrevuActuel = toNombre(getCell(rowActuelle, header, ['apport prévu', 'apport prevu']));
  const montantVerseActuel = toNombre(getCell(rowActuelle, header, ['montant versé', 'montant verse']));
  const facturesActuelles = String(getCell(rowActuelle, header, ['n° facture', 'numero facture']) || '').trim();
  const nouveauMontantVerse = montantVerseActuel + infos.montant;
  const seuil = apportPrevuActuel > 0 ? apportPrevuActuel : CONFIG.SEUIL_VALIDATION_DEFAUT;
  const reste = Math.max(seuil - nouveauMontantVerse, 0);
  const pourcentage = seuil > 0 ? Math.min(nouveauMontantVerse / seuil, 1) : 0;
  const statut = nouveauMontantVerse >= seuil ? CONFIG.STATUT_VALIDE : CONFIG.STATUT_PARTIEL;
  const nouvellesFactures = infos.numFacture + (facturesActuelles ? '\n' + facturesActuelles : '');
  const valeurs = {
    'montant versé': nouveauMontantVerse,
    'reste à verser': reste,
    '% versé': pourcentage,
    'date de versement final': infos.datePaie,
    'mode de paiement': infos.modePaie,
    'statut': statut,
    'n° facture': nouvellesFactures
  };
  if (!apportPrevuActuel) valeurs['apport prévu'] = seuil;
  ecrireLigne(apportsSheet, numeroLigne, header, valeurs);
  return statut;
}

function setCellDansLigne(sheet, numeroLigne, header, keywords, valeur) {
  const col = colIndexOf(header, keywords);
  if (col) sheet.getRange(numeroLigne, col).setValue(valeur);
}

function mettreAJourStatutMembre(membresSheet, idMembre, statut) {
  const header = getHeaderMap(membresSheet);
  const headerRow = header.__headerRow || 1;
  const lastRow = membresSheet.getLastRow();
  if (lastRow <= headerRow) return;
  const colId = colIndexOf(header, ['id membre', 'id']);
  const ids = membresSheet.getRange(headerRow + 1, colId, lastRow - headerRow, 1).getValues();
  for (let i = 0; i < ids.length; i++) {
    if (String(ids[i][0]) === String(idMembre)) {
      setCellDansLigne(membresSheet, i + headerRow + 1, header, ['statut'], statut);
      return;
    }
  }
}

function notifierHomonymeApport(nom, nouvelId, correspondances, numFacture) {
  if (!ADMIN_EMAIL) return;
  const idsExistants = correspondances.map(function (m) { return m.id; }).join(', ');
  MailApp.sendEmail(ADMIN_EMAIL, 'Homonyme détecté - vérification requise - Suivi Apports',
    'HOMONYME : le nom "' + nom + '" (facture ' + numFacture + ') correspond à plusieurs membres ' +
    'existants (' + idsExistants + '). Un nouveau membre (' + nouvelId + ') a été créé par précaution. ' +
    'Merci de vérifier s\'il s\'agit d\'un doublon dans l\'onglet MEMBRE.');
}

function ecrireStatutSync(factureSheet, row, message) {
  const cell = factureSheet.getRange(row, COL.MERGE_STATUS);
  const existant = cell.getValue();
  const separateur = existant ? ' ; ' : '';
  cell.setValue(existant + separateur + message + ' (' + new Date().toLocaleString('fr-FR') + ')');
}

function ecrireLigneGestionFinanciere(ss, infos) {
  const sheet = ss.getSheetByName(CONFIG.SHEET_FINANCE);
  if (!sheet) return;
  if (infos.numFacture && verifierNumeroFactureExistant(sheet, infos.numFacture)) {
    Logger.log('⚠️ Doublon détecté : N° ' + infos.numFacture + ' existe déjà dans la caisse. Ignoré.');
    return;
  }
  const header = getHeaderMap(sheet);
  const headerRow = header.__headerRow || 1;
  const lastRow = sheet.getLastRow();
  const colNum = colIndexOf(header, ['n°', 'numero', 'n']);
  let dernierNumero = 0;
  if (lastRow > headerRow) {
    const numeros = sheet.getRange(headerRow + 1, colNum, lastRow - headerRow, 1).getValues();
    numeros.forEach(function (r) { const n = parseInt(r[0], 10); if (!isNaN(n) && n > dernierNumero) dernierNumero = n; });
  }
  const ligne = new Array(sheet.getLastColumn()).fill('');
  setCell(ligne, header, ['n°', 'numero', 'n'], dernierNumero + 1);
  setCell(ligne, header, ['date'], infos.date || new Date());
  setCell(ligne, header, ['type'], 'Entrée');
  setCell(ligne, header, ['catégorie', 'categorie'], 'Apport associé');
  setCell(ligne, header, ['description'], infos.idMembre + ' - ' + infos.nom + ' - Apport en capital - ' + infos.numFacture);
  setCell(ligne, header, ['nom'], infos.nom);
  setCell(ligne, header, ['mode de paiement'], infos.modePaie || 'Non spécifié');
  setCell(ligne, header, ['entré', 'entrée', 'entree'], infos.montant || 0);
  setCell(ligne, header, ['sortie'], 0);
  setCell(ligne, header, ['n° facture', 'numero facture'], infos.numFacture || '');
  setCell(ligne, header, ['enrégistré par', 'enregistré par', 'reçu par', 'recu par'], infos.recuPar || 'Admin');
  sheet.appendRow(ligne);
  SpreadsheetApp.flush();
}

// ============================================================
//  CAISSE (Gestion Financière)
// ============================================================
function genererNumeroFinance() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CONFIG.SHEET_FINANCE);
  const header = getHeaderMap(sheet);
  const headerRow = header.__headerRow || 1;
  const lastRow = sheet.getLastRow();
  const year = new Date().getFullYear();
  const prefix = 'GF-' + year + '-';
  let dernierNumero = 0;
  if (lastRow > headerRow) {
    const col = colIndexOf(header, ['n° facture', 'numero facture']);
    const values = sheet.getRange(headerRow + 1, col, lastRow - headerRow, 1).getValues();
    values.forEach(function (r) {
      const val = String(r[0] || '').trim();
      if (val.indexOf(prefix) === 0) {
        const num = parseInt(val.substring(prefix.length), 10);
        if (!isNaN(num) && num > dernierNumero) dernierNumero = num;
      }
    });
  }
  return prefix + String(dernierNumero + 1).padStart(5, '0');
}

function genererPdfMouvementFinance(infos, numFacture) {
  const estDebit = infos.type === 'Entrée';
  const folderId = estDebit ? CONFIG.DRIVE_FOLDER_JUSTIFICATIFS_DEBIT : CONFIG.DRIVE_FOLDER_JUSTIFICATIFS_CREDIT;
  const folder = DriveApp.getFolderById(folderId);
  const doc = DocumentApp.create('Justificatif ' + numFacture);
  const body = doc.getBody();
  body.setMarginTop(28).setMarginBottom(28).setMarginLeft(40).setMarginRight(40);
  try {
    const logoUrl = 'https://drive.google.com/thumbnail?id=1tNXUPl0gC6aAeaJ5C4evF4zMaZJNOBgb&sz=w400';
    const logoBlob = UrlFetchApp.fetch(logoUrl).getBlob();
    const logoPara = body.appendParagraph('');
    logoPara.setAlignment(DocumentApp.HorizontalAlignment.CENTER);
    const logoImg = logoPara.insertInlineImage(0, logoBlob);
    logoImg.setWidth(70);
    logoImg.setHeight(70);
  } catch (imgErr) { Logger.log('Logo indisponible : ' + imgErr.message); }
  const titre = body.appendParagraph('NEXT GENERATION MADA — "NG-MADA"');
  titre.setAlignment(DocumentApp.HorizontalAlignment.CENTER).setHeading(DocumentApp.ParagraphHeading.HEADING2);
  const sousTitre = body.appendParagraph(estDebit ? 'Reçu — Débit (entrée en caisse)' : 'Reçu — Crédit (sortie de caisse)');
  sousTitre.setAlignment(DocumentApp.HorizontalAlignment.CENTER).setHeading(DocumentApp.ParagraphHeading.HEADING3);
  body.appendParagraph(' ');
  const table = body.appendTable();
  table.setBorderWidth(0);
  function ajouterLigneTableau(table, label, valeur) {
    const row = table.appendTableRow();
    const cell1 = row.appendTableCell(label);
    cell1.setBackgroundColor('#f1f5f9');
    cell1.setWidth(150);
    row.appendTableCell(valeur);
  }
  ajouterLigneTableau(table, 'N°', numFacture);
  const dateFormatee = Utilities.formatDate(new Date(infos.date), Session.getScriptTimeZone(), 'dd/MM/yyyy');
  ajouterLigneTableau(table, 'Date', dateFormatee);
  ajouterLigneTableau(table, 'Type', estDebit ? 'Débit (Entrée)' : 'Crédit (Sortie)');
  ajouterLigneTableau(table, 'Description', infos.description || '—');
  const modePaiement = infos.modePaiement || 'Non spécifié';
  ajouterLigneTableau(table, 'Mode de paiement', modePaiement);
  const quantite = parseInt(infos.quantite) || 1;
  ajouterLigneTableau(table, 'Quantité', String(quantite));
  const prixUnitaire = toNombre(infos.prixUnitaire);
  const prixUnitaireFormate = prixUnitaire > 0 ? prixUnitaire.toLocaleString('fr-FR') + ' Ar' : '—';
  ajouterLigneTableau(table, 'Prix unitaire', prixUnitaireFormate);
  const montant = toNombre(infos.montant);
  const rowTotal = table.appendTableRow();
  const cellLabelTotal = rowTotal.appendTableCell('Montant total');
  cellLabelTotal.setBackgroundColor('#facc15');
  cellLabelTotal.setBold(true);
  const cellValeurTotal = rowTotal.appendTableCell(montant.toLocaleString('fr-FR') + ' Ar');
  cellValeurTotal.setBackgroundColor('#facc15');
  cellValeurTotal.setBold(true);
  ajouterLigneTableau(table, 'Reçu par', infos.recuPar || '—');
  body.appendParagraph(' ');
  try {
    const texteQr = [numFacture, infos.description || '—', 'Mode: ' + modePaiement, 'Qté: ' + quantite,
      'Prix unit: ' + (prixUnitaire > 0 ? prixUnitaire + ' Ar' : '—'), 'Total: ' + montant + ' Ar',
      'Reçu par: ' + (infos.recuPar || '—')
    ].filter(Boolean).join(', ');
    const qrUrl = 'https://quickchart.io/qr?text=' + encodeURIComponent(texteQr) + '&size=300';
    const qrBlob = UrlFetchApp.fetch(qrUrl).getBlob();
    const qrPara = body.appendParagraph('');
    qrPara.setAlignment(DocumentApp.HorizontalAlignment.CENTER);
    const qrImg = qrPara.insertInlineImage(0, qrBlob);
    qrImg.setWidth(100);
    qrImg.setHeight(100);
    const qrLegende = body.appendParagraph('Scanner pour voir le résumé du mouvement');
    qrLegende.setAlignment(DocumentApp.HorizontalAlignment.CENTER);
    qrLegende.setFontSize(8);
    qrLegende.setForegroundColor('#94a3b8');
  } catch (qrErr) { Logger.log('QR indisponible : ' + qrErr.message); }
  body.appendParagraph(' ');
  const pied = body.appendParagraph('Document généré automatiquement le ' + new Date().toLocaleString('fr-FR'));
  pied.setAlignment(DocumentApp.HorizontalAlignment.CENTER);
  pied.setFontSize(8);
  pied.setForegroundColor('#94a3b8');
  doc.saveAndClose();
  const docFile = DriveApp.getFileById(doc.getId());
  const pdfBlob = docFile.getAs(MimeType.PDF);
  pdfBlob.setName('Justificatif ' + numFacture + '.pdf');
  const pdfFile = folder.createFile(pdfBlob);
  docFile.setTrashed(true);
  return pdfFile.getUrl();
}

function previsualiserMouvementFinance(pin, infos) {
  try {
    if (!verifierPin(pin)) return { success: false, message: 'PIN incorrect' };
    if (!infos || !infos.date || !infos.description || !infos.type || !infos.montant || !infos.recuPar) {
      return { success: false, message: 'Merci de remplir tous les champs (y compris "Reçu par").' };
    }
    const montant = toNombre(infos.montant);
    if (!montant || montant <= 0) return { success: false, message: 'Montant invalide.' };
    const modePaiement = infos.modePaiement || 'Non spécifié';
    const quantite = parseInt(infos.quantite) || 1;
    const prixUnitaire = toNombre(infos.prixUnitaire) || 0;
    const montantCalcule = quantite * prixUnitaire;
    if (montantCalcule > 0 && Math.abs(montantCalcule - montant) > 0.01) {
      return { success: false, message: '⚠️ Le montant total ne correspond pas à quantité × prix unitaire.' };
    }
    const numFacture = genererNumeroFinance();
    const dossierCible = infos.type === 'Entrée' ? 'Débit (entrée)' : 'Crédit (sortie)';
    return { success: true, numFacture: numFacture, dossierCible: dossierCible, modePaiement: modePaiement, quantite: quantite, prixUnitaire: prixUnitaire, montant: montant };
  } catch (err) {
    logErreur('previsualiserMouvementFinance', err);
    return { success: false, message: err.message };
  }
}

function ajouterMouvementFinance(pin, infos) {
  try {
    if (!verifierPin(pin)) return { success: false, message: 'PIN incorrect' };
    if (!infos || !infos.date || !infos.description || !infos.type || !infos.montant || !infos.recuPar) {
      return { success: false, message: 'Merci de remplir tous les champs (y compris "Reçu par").' };
    }
    const montant = toNombre(infos.montant);
    if (!montant || montant <= 0) return { success: false, message: 'Montant invalide.' };
    const modePaiement = infos.modePaiement || 'Non spécifié';
    const quantite = parseInt(infos.quantite) || 1;
    const prixUnitaire = toNombre(infos.prixUnitaire) || 0;
    const numFacture = genererNumeroFinance();
    const pdfUrl = genererPdfMouvementFinance(infos, numFacture);
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CONFIG.SHEET_FINANCE);
    const header = getHeaderMap(sheet);
    const headerRow = header.__headerRow || 1;
    const lastRow = sheet.getLastRow();
    const colNum = colIndexOf(header, ['n°', 'numero', 'n']);
    let dernierNumero = 0;
    if (lastRow > headerRow) {
      const numeros = sheet.getRange(headerRow + 1, colNum, lastRow - headerRow, 1).getValues();
      numeros.forEach(function (r) { const n = parseInt(r[0], 10); if (!isNaN(n) && n > dernierNumero) dernierNumero = n; });
    }
    const estEntree = infos.type === 'Entrée';
    const ligne = new Array(sheet.getLastColumn()).fill('');
    setCell(ligne, header, ['n°', 'numero', 'n'], dernierNumero + 1);
    setCell(ligne, header, ['date'], new Date(infos.date));
    setCell(ligne, header, ['type'], infos.type);
    setCell(ligne, header, ['catégorie', 'categorie'], infos.categorie || 'Divers');
    setCell(ligne, header, ['description'], infos.description);
    setCell(ligne, header, ['mode de paiement'], modePaiement);
    setCell(ligne, header, ['quantité', 'quantite'], quantite);
    setCell(ligne, header, ['prix unitaire', 'prix_unitaire'], prixUnitaire);
    setCell(ligne, header, ['entré', 'entrée', 'entree'], estEntree ? montant : 0);
    setCell(ligne, header, ['sortie'], estEntree ? 0 : montant);
    setCell(ligne, header, ['n° facture', 'numero facture'], numFacture);
    setCell(ligne, header, ['enrégistré par', 'enregistré par', 'reçu par', 'recu par'], infos.recuPar);
    setCell(ligne, header, ['lien justificatif', 'justificatif'], pdfUrl);
    sheet.appendRow(ligne);
    SpreadsheetApp.flush();
    if (ADMIN_EMAIL) {
      const sujet = (estEntree ? '⬆️ Débit' : '⬇️ Crédit') + ' enregistré — ' + numFacture;
      const corps = 'Un mouvement vient d\'être enregistré dans la Caisse.\n\n' +
        'N° : ' + numFacture + '\nType : ' + (estEntree ? 'Débit (Entrée)' : 'Crédit (Sortie)') +
        '\nDate : ' + Utilities.formatDate(new Date(infos.date), Session.getScriptTimeZone(), 'dd/MM/yyyy') +
        '\nDescription : ' + infos.description + '\nMode : ' + modePaiement + '\nQuantité : ' + quantite +
        '\nPrix unitaire : ' + prixUnitaire.toLocaleString('fr-FR') + ' Ar\nMontant total : ' + montant.toLocaleString('fr-FR') + ' Ar' +
        '\nReçu par : ' + infos.recuPar + '\n\nJustificatif : ' + pdfUrl;
      MailApp.sendEmail(ADMIN_EMAIL, sujet, corps);
    }
    ngmadaSynchroniserTableauDeBord();
    return { success: true, numFacture: numFacture, pdfUrl: pdfUrl };
  } catch (err) {
    logErreur('ajouterMouvementFinance', err);
    return { success: false, message: err.message };
  }
}

function getFinanceData() {
  try {
    SpreadsheetApp.flush();
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CONFIG.SHEET_FINANCE);
    const header = getHeaderMap(sheet);
    const headerRow = header.__headerRow || 1;
    const lastRow = sheet.getLastRow();
    if (lastRow <= headerRow) return { items: [], totalEntree: 0, totalSortie: 0 };
    const data = sheet.getRange(headerRow + 1, 1, lastRow - headerRow, sheet.getLastColumn()).getValues();
    let totalEntree = 0, totalSortie = 0;
    const items = [];
    data.forEach(function (r, i) {
      if (r.join('').toString().trim() === '') return;
      const entree = toNombre(getCell(r, header, ['entré', 'entrée', 'entree']));
      const sortie = toNombre(getCell(r, header, ['sortie']));
      totalEntree += entree;
      totalSortie += sortie;
      const dateCell = getCell(r, header, ['date']);
      let dateFormatee = '';
      if (dateCell instanceof Date) {
        dateFormatee = Utilities.formatDate(dateCell, Session.getScriptTimeZone(), 'dd/MM/yyyy');
      } else if (dateCell) {
        try {
          const d = new Date(dateCell);
          if (!isNaN(d)) {
            dateFormatee = Utilities.formatDate(d, Session.getScriptTimeZone(), 'dd/MM/yyyy');
          }
        } catch(e) {}
      }
      items.push({
        row: i + headerRow + 1,
        date: dateFormatee || String(dateCell || ''),
        type: entree > 0 ? 'Entrée' : 'Sortie',
        categorie: getCell(r, header, ['catégorie', 'categorie']),
        description: getCell(r, header, ['description']),
        nom: getCell(r, header, ['nom']),
        modePaiement: getCell(r, header, ['mode de paiement']),
        entree: entree,
        sortie: sortie,
        numFacture: getCell(r, header, ['n° facture', 'numero facture']),
        justificatif: getCell(r, header, ['lien justificatif', 'justificatif']),
        enregistrePar: getCell(r, header, ['enrégistré par', 'enregistré par', 'reçu par', 'recu par'])
      });
    });
    items.reverse();
    return { items: items, totalEntree: totalEntree, totalSortie: totalSortie };
  } catch (err) {
    logErreur('getFinanceData', err);
    return { __error: err.message };
  }
}

function extraireIdDriveDepuisUrl(url) { if (!url) return null; const m = String(url).match(/[-\w]{25,}/); return m ? m[0] : null; }

function trouverJustificatifParNumero(numFacture) {
  if (!numFacture) return null;
  const nomFichier = 'Justificatif ' + numFacture + '.pdf';
  const dossiers = [CONFIG.DRIVE_FOLDER_JUSTIFICATIFS_DEBIT, CONFIG.DRIVE_FOLDER_JUSTIFICATIFS_CREDIT];
  for (let i = 0; i < dossiers.length; i++) {
    try {
      const folder = DriveApp.getFolderById(dossiers[i]);
      const it = folder.getFilesByName(nomFichier);
      if (it.hasNext()) return it.next();
    } catch (e) { Logger.log('trouverJustificatifParNumero : dossier inaccessible (' + dossiers[i] + ') : ' + e.message); }
  }
  return null;
}

function trouverFactureParNumero_(numFacture) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(FACTURE_SHEET_NAME);
  const lastRow = sheet.getLastRow();
  if (lastRow < 2 || !numFacture) return null;
  const data = sheet.getRange(2, 1, lastRow - 1, sheet.getLastColumn()).getValues();
  for (let i = 0; i < data.length; i++) {
    if (String(data[i][COL.INVOICE - 1]).trim() === String(numFacture).trim()) {
      return { row: i + 2, mergedDocId: data[i][COL.MERGED_DOC_ID - 1] };
    }
  }
  return null;
}

function envoyerJustificatifParEmail(pin, row) {
  try {
    if (!verifierPin(pin)) return { success: false, message: 'PIN incorrect' };
    const destinataire = ADMIN_EMAIL;
    if (!destinataire) return { success: false, message: 'Aucune adresse email configurée.' };
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CONFIG.SHEET_FINANCE);
    const header = getHeaderMap(sheet);
    const rowValues = sheet.getRange(row, 1, 1, sheet.getLastColumn()).getValues()[0];
    const lien = getCell(rowValues, header, ['lien justificatif', 'justificatif']);
    const numFacture = getCell(rowValues, header, ['n° facture', 'numero facture']);
    const estMouvementManuel = String(numFacture || '').indexOf('GF-') === 0;
    let pdfBlob = null;
    let nomFichier = 'Justificatif ' + numFacture + '.pdf';
    if (!estMouvementManuel) {
      const facture = trouverFactureParNumero_(numFacture);
      if (facture && facture.mergedDocId) {
        try { pdfBlob = DriveApp.getFileById(facture.mergedDocId).getAs(MimeType.PDF); nomFichier = 'NG FACT ' + numFacture + '.pdf'; } catch (e) { pdfBlob = null; }
      }
    } else {
      if (lien) {
        const fileId = extraireIdDriveDepuisUrl(lien);
        if (fileId) { try { pdfBlob = DriveApp.getFileById(fileId).getBlob(); } catch (e) { pdfBlob = null; } }
      }
      if (!pdfBlob) {
        const fichier = trouverJustificatifParNumero(numFacture);
        if (fichier) pdfBlob = fichier.getBlob();
      }
    }
    if (!pdfBlob) {
      const message = estMouvementManuel
        ? 'PDF introuvable dans les dossiers Drive pour ' + numFacture + '.'
        : 'Facture ' + (numFacture || '—') + ' introuvable (ou pas encore générée).';
      return { success: false, message: message };
    }
    pdfBlob.setName(nomFichier);
    MailApp.sendEmail(destinataire, (estMouvementManuel ? 'Justificatif ' : 'Facture ') + numFacture + ' - Suivi Apports Associés',
      'Voici le document demandé, en pièce jointe.', { attachments: [pdfBlob] });
    return { success: true, email: destinataire };
  } catch (err) {
    logErreur('envoyerJustificatifParEmail', err);
    return { success: false, message: err.message };
  }
}

// ============================================================
//  CARTE MEMBRE — envoi par email au membre
// ============================================================
function envoyerCarteMembreParEmail(pin, idMembre) {
  try {
    verifierPinOuErreur(pin);
    if (!idMembre) throw new Error('ID membre manquant.');
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CONFIG.SHEET_MEMBRE);
    const header = getHeaderMap(sheet);
    const headerRow = header.__headerRow || 1;
    const lastRow = sheet.getLastRow();
    if (lastRow <= headerRow) throw new Error('Membre introuvable.');
    const data = sheet.getRange(headerRow + 1, 1, lastRow - headerRow, sheet.getLastColumn()).getValues();
    let membre = null;
    for (let i = 0; i < data.length; i++) {
      if (String(getCell(data[i], header, ['id membre', 'id'])) === String(idMembre)) {
        membre = {
          id: idMembre,
          nom: getCell(data[i], header, ['nom & prénom', 'nom']),
          fonction: getCell(data[i], header, ['fonction']),
          statut: getCell(data[i], header, ['statut']),
          photo: getCell(data[i], header, ['photo']),
          email: getCell(data[i], header, ['adresse e-mail', 'email', 'e-mail']),
          telephone: getCell(data[i], header, ['téléphone', 'telephone']),
          dateAdhesion: getCell(data[i], header, ["date d'adhésion"])
        };
        break;
      }
    }
    if (!membre) throw new Error('Membre introuvable.');
    const emailMembre = String(membre.email || '').trim();
    if (!emailMembre || !emailMembre.includes('@')) return { success: false, message: 'Adresse e-mail invalide pour ce membre.' };
    const pdfResult = genererCarteMembrePDF(membre);
    if (!pdfResult.success) return { success: false, message: 'Erreur PDF : ' + pdfResult.message };
    const pdfBlob = pdfResult.pdfBlob;
    const url = ScriptApp.getService().getUrl();
    const verifyUrl = url ? (url + '?verify=' + encodeURIComponent(membre.id)) : '';
    const textBody = 'Bonjour ' + membre.nom + ',\n\nVotre carte de membre NG-MADA est disponible en pièce jointe.\n' +
      'Vous pouvez aussi la consulter en ligne : ' + verifyUrl + '\n\nCordialement,\nNG-MADA\nContact : ' + ADMIN_EMAIL;
    const htmlBody = '<div style="font-family:Arial,sans-serif;padding:20px;">' +
      '<h2>Votre carte de membre NG-MADA</h2>' +
      '<p>Bonjour <strong>' + membre.nom + '</strong>,</p>' +
      '<p>Votre carte de membre est jointe à cet email au format PDF.</p>' +
      '<p>Consultez-la en ligne : <a href="' + verifyUrl + '">' + verifyUrl + '</a></p>' +
      '<p>Cordialement,<br>NG-MADA</p></div>';
    MailApp.sendEmail({ to: emailMembre, subject: 'Votre carte de membre NG-MADA – ' + membre.nom,
      body: textBody, htmlBody: htmlBody, replyTo: ADMIN_EMAIL, name: 'NG-MADA', attachments: [pdfBlob] });
    if (ADMIN_EMAIL && ADMIN_EMAIL !== emailMembre) {
      MailApp.sendEmail({ to: ADMIN_EMAIL, subject: '📨 Carte membre PDF envoyée à ' + membre.nom + ' (' + emailMembre + ')',
        body: 'Une carte de membre au format PDF a été envoyée à ' + membre.nom + ' (' + emailMembre + ') le ' + new Date().toLocaleString('fr-FR') + '.' });
    }
    return { success: true, email: emailMembre };
  } catch (err) {
    logErreur('envoyerCarteMembreParEmail', err);
    return { success: false, message: err.message };
  }
}

function genererCarteMembrePDF(membre) {
  try {
    if (!membre || !membre.id) throw new Error('Données du membre incomplètes.');
    const templateFile = DriveApp.getFileById('14beWdRuhWFJAtLdLZN42yBNyrXrv_bUEyWGbu5Tj9M4');
    const outputFolder = DriveApp.getFolderById('1YuWNNE5hrKJeKKk6eY_9EBsx98h1GL4T');
    const docName = 'Carte_Membre_NG-MADA_' + membre.id;
    const copiedFile = templateFile.makeCopy(docName, outputFolder);
    const doc = DocumentApp.openById(copiedFile.getId());
    const body = doc.getBody();
    const dateAdhesionTxt = formatDateAffichage(membre.dateAdhesion);
    const statutTxt = membre.statut || 'Nouveau - à vérifier';
    const telephoneTxt = membre.telephone || '—';
    body.replaceText('<<NOM>>', membre.nom || '');
    body.replaceText('<<ID>>', membre.id);
    body.replaceText('<<FONCTION>>', membre.fonction || 'Membre simple');
    body.replaceText('<<TELEPHONE>>', telephoneTxt);
    body.replaceText('<<DATE_ADHESION>>', dateAdhesionTxt || '—');
    body.replaceText('<<STATUT>>', statutTxt);
    if (membre.photo) {
      try {
        const photoResponse = UrlFetchApp.fetch(membre.photo);
        const photoBlob = photoResponse.getBlob();
        const photoPlaceholder = body.findText('<<PHOTO>>');
        if (photoPlaceholder) {
          const element = photoPlaceholder.getElement();
          const parent = element.getParent();
          element.asText().setText('');
          let img;
          if (parent.getType() === DocumentApp.ElementType.PARAGRAPH) {
            img = parent.asParagraph().insertInlineImage(0, photoBlob);
            parent.asParagraph().setAlignment(DocumentApp.HorizontalAlignment.CENTER);
          } else if (parent.getType() === DocumentApp.ElementType.TABLE_CELL) {
            const cell = parent.asTableCell();
            const p = cell.insertParagraph(0, '');
            img = p.insertInlineImage(0, photoBlob);
            p.setAlignment(DocumentApp.HorizontalAlignment.CENTER);
          }
          if (img) { img.setWidth(60); img.setHeight(60); }
        }
      } catch (e) { body.replaceText('<<PHOTO>>', '[Photo indisponible]'); }
    } else { body.replaceText('<<PHOTO>>', ''); }
    const url = ScriptApp.getService().getUrl();
    const verifyUrl = url ? (url + '?verify=' + encodeURIComponent(membre.id)) : '';
    if (verifyUrl) {
      try {
        const qrApiUrl = 'https://api.qrserver.com/v1/create-qr-code/?size=140x140&data=' + encodeURIComponent(verifyUrl);
        const qrResponse = UrlFetchApp.fetch(qrApiUrl);
        const qrBlob = qrResponse.getBlob();
        const qrPlaceholder = body.findText('<<QR>>');
        if (qrPlaceholder) {
          const element = qrPlaceholder.getElement();
          const parent = element.getParent();
          element.asText().setText('');
          let img;
          if (parent.getType() === DocumentApp.ElementType.PARAGRAPH) {
            img = parent.asParagraph().insertInlineImage(0, qrBlob);
            parent.asParagraph().setAlignment(DocumentApp.HorizontalAlignment.CENTER);
          } else if (parent.getType() === DocumentApp.ElementType.TABLE_CELL) {
            const cell = parent.asTableCell();
            const p = cell.insertParagraph(0, '');
            img = p.insertInlineImage(0, qrBlob);
            p.setAlignment(DocumentApp.HorizontalAlignment.CENTER);
          }
          if (img) { img.setWidth(40); img.setHeight(40); }
        }
      } catch (e) { body.replaceText('<<QR>>', '[QR indisponible]'); }
    } else { body.replaceText('<<QR>>', ''); }
    doc.saveAndClose();
    const pdfBlob = DriveApp.getFileById(copiedFile.getId()).getAs(MimeType.PDF);
    pdfBlob.setName(docName + '.pdf');
    copiedFile.setTrashed(true);
    return { success: true, pdfBlob: pdfBlob };
  } catch (err) {
    logErreur('genererCarteMembrePDF', err);
    return { success: false, message: err.message };
  }
}

// ============================================================
//  DASHBOARD — données
// ============================================================
function getDashboardData() {
  try {
    SpreadsheetApp.flush();
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const apports = lireFeuilleObjets(ss.getSheetByName(CONFIG.SHEET_APPORTS));
    const finance = lireFeuilleObjets(ss.getSheetByName(CONFIG.SHEET_FINANCE));
    const dashboardSheet = ss.getSheetByName(CONFIG.SHEET_DASHBOARD);
    let totalVerse = 0, totalReste = 0, associesValides = 0, enAttente = 0;
    apports.forEach(function (a) {
      totalVerse += toNombre(getVal(a, ['montant versé', 'montant verse']));
      totalReste += toNombre(getVal(a, ['reste à verser', 'reste a verser']));
      const statut = String(getVal(a, ['statut']) || '');
      if (statut.toLowerCase().indexOf('valid') !== -1) associesValides++;
      else enAttente++;
    });
    let entreesCaisse = 0, sortiesCaisse = 0;
    finance.forEach(function (m) {
      entreesCaisse += toNombre(getVal(m, ['entré', 'entrée', 'entree']));
      sortiesCaisse += toNombre(getVal(m, ['sortie']));
    });
    let joursRestants = null;
    const dateLimite = trouverDateLimite(dashboardSheet);
    if (dateLimite) {
      const diff = Math.ceil((dateLimite - new Date()) / 86400000);
      joursRestants = diff > 0 ? diff : 0;
    }
    return { totalVerse: totalVerse, totalReste: totalReste, associesValides: associesValides, enAttente: enAttente,
      nombreApports: apports.length, soldeCaisse: entreesCaisse - sortiesCaisse, joursRestants: joursRestants };
  } catch (err) {
    logErreur('getDashboardData', err);
    return { __error: err.message };
  }
}

function getApportsData() {
  try {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CONFIG.SHEET_APPORTS);
    return lireFeuilleObjets(sheet).map(function (a) {
      return {
        idMembre: getVal(a, ['id membre', 'id']),
        nom: getVal(a, ['nom & prénom', 'nom']),
        apportPrevu: toNombre(getVal(a, ['apport prévu', 'apport prevu'])),
        montantVerse: toNombre(getVal(a, ['montant versé', 'montant verse'])),
        resteAVerser: toNombre(getVal(a, ['reste à verser', 'reste a verser'])),
        pourcentage: toNombre(getVal(a, ['% versé', '% verse'])),
        statut: getVal(a, ['statut']),
        modePaiement: getVal(a, ['mode de paiement']),
        dateVersementFinal: formatDateAffichage(getVal(a, ['date de versement final'])),
        factures: String(getVal(a, ['n° facture', 'numero facture']) || '').split('\n').filter(Boolean)
      };
    });
  } catch (err) {
    logErreur('getApportsData', err);
    return { __error: err.message };
  }
}

function getFacturesData() {
  try {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CONFIG.SHEET_FACTURE);
    const header = getHeaderMap(sheet);
    const headerRow = header.__headerRow || 1;
    const lastRow = sheet.getLastRow();
    if (lastRow <= headerRow) return [];
    const data = sheet.getRange(headerRow + 1, 1, lastRow - headerRow, sheet.getLastColumn()).getValues();
    const list = [];
    data.forEach(function (r, i) {
      if (r.join('').toString().trim() === '') return;
      const ligne = { __row: r, __header: header };
      const nom = getVal(ligne, ['nom']);
      if (!nom) return;
      list.push({
        row: i + headerRow + 1,
        numFacture: getVal(ligne, ['n° facture', 'numero facture']),
        nom: nom,
        montant: toNombre(getVal(ligne, ['montant'])),
        modePaiement: getVal(ligne, ['mode de paiement']),
        datePaiement: formatDateAffichage(getVal(ligne, ['date de paiement'])),
        validation: getVal(ligne, ['validation']),
        mergedDocUrl: getVal(ligne, ['lien pdf', 'pdf url', 'merged doc url'])
      });
    });
    list.reverse();
    return list;
  } catch (err) {
    logErreur('getFacturesData', err);
    return { __error: err.message };
  }
}

function getApportsMensuels() {
  try {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CONFIG.SHEET_APPORTS);
    const data = lireFeuilleObjets(sheet);
    const moisMap = {};
    data.forEach(function (a) {
      const date = getVal(a, ['date d’engagement', 'date']);
      const d = date instanceof Date ? date : new Date(date);
      if (!isNaN(d)) {
        const key = Utilities.formatDate(d, Session.getScriptTimeZone(), 'MM/yyyy');
        const montant = toNombre(getVal(a, ['montant versé', 'montant verse']));
        moisMap[key] = (moisMap[key] || 0) + montant;
      }
    });
    const labels = Object.keys(moisMap).sort();
    const dataArray = labels.map(function (l) { return moisMap[l]; });
    return { labels: labels, data: dataArray };
  } catch (err) {
    return { __error: err.message };
  }
}

// ============================================================
//  MEMBRES
// ============================================================
function getMembres() {
  try {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CONFIG.SHEET_MEMBRE);
    return lireFeuilleObjets(sheet).map(function (m) {
      return {
        id: getVal(m, ['id membre', 'id']),
        nom: getVal(m, ['nom & prénom', 'nom']),
        telephone: getVal(m, ['téléphone', 'telephone']),
        email: getVal(m, ['adresse e-mail', 'email', 'e-mail']),
        fonction: getVal(m, ['fonction']),
        dateAdhesion: formatDateAffichage(getVal(m, ["date d'adhésion"])),
        statut: getVal(m, ['statut']),
        photo: getVal(m, ['photo']),
        dateNaissance: formatDateAffichage(getVal(m, ['date de naissance'])),
        adressePostale: getVal(m, ['adresse postale'])
      };
    });
  } catch (err) {
    logErreur('getMembres', err);
    return { __error: err.message };
  }
}

function getMembreLienPhoto(idMembre) {
  try { const url = ScriptApp.getService().getUrl(); return url ? url + '?id=' + encodeURIComponent(idMembre) : ''; } catch (err) { return ''; }
}

function getMembrePublic(idMembre) {
  try {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CONFIG.SHEET_MEMBRE);
    const header = getHeaderMap(sheet);
    const headerRow = header.__headerRow || 1;
    const lastRow = sheet.getLastRow();
    if (lastRow <= headerRow) return null;
    const data = sheet.getRange(headerRow + 1, 1, lastRow - headerRow, sheet.getLastColumn()).getValues();
    for (let i = 0; i < data.length; i++) {
      if (String(getCell(data[i], header, ['id membre', 'id'])) === String(idMembre)) {
        return { id: idMembre, nom: getCell(data[i], header, ['nom & prénom', 'nom']), photo: getCell(data[i], header, ['photo']) };
      }
    }
    return null;
  } catch (err) {
    logErreur('getMembrePublic', err);
    return { __error: err.message };
  }
}

function saveMembre(data, pin) {
  try {
    verifierPinOuErreur(pin);
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CONFIG.SHEET_MEMBRE);
    const header = getHeaderMap(sheet);
    const headerRow = header.__headerRow || 1;
    if (data.id) {
      const lastRow = sheet.getLastRow();
      const idCol = colIndexOf(header, ['id membre', 'id']);
      if (lastRow > headerRow) {
        const ids = sheet.getRange(headerRow + 1, idCol, lastRow - headerRow, 1).getValues();
        for (let i = 0; i < ids.length; i++) {
          if (String(ids[i][0]) === String(data.id)) {
            const r = i + headerRow + 1;
            const valeurs = {
              'nom & prénom': data.nom,
              'téléphone': data.telephone,
              'adresse e-mail': data.email,
              'fonction': data.fonction || CONFIG.FONCTION_DEFAUT,
              'statut': data.statut || CONFIG.STATUT_NOUVEAU,
              'date de naissance': data.dateNaissance,
              'adresse postale': data.adressePostale
            };
            ecrireLigne(sheet, r, header, valeurs);
            return { success: true, id: data.id };
          }
        }
      }
    }
    const lastRow2 = sheet.getLastRow();
    let maxId = 0;
    if (lastRow2 > headerRow) {
      const idCol = colIndexOf(header, ['id membre', 'id']);
      const ids = sheet.getRange(headerRow + 1, idCol, lastRow2 - headerRow, 1).getValues();
      ids.forEach(function (r) {
        const n = parseInt(String(r[0]).replace(/[^0-9]/g, ''), 10);
        if (!isNaN(n) && n > maxId) maxId = n;
      });
    }
    const nextId = CONFIG.MEMBRE_ID_PREFIX + String(maxId + 1).padStart(CONFIG.MEMBRE_ID_PAD, '0');
    const ligne = new Array(sheet.getLastColumn()).fill('');
    setCell(ligne, header, ['id membre', 'id'], nextId);
    setCell(ligne, header, ['nom & prénom', 'nom'], data.nom);
    setCell(ligne, header, ['téléphone', 'telephone'], data.telephone);
    setCell(ligne, header, ['adresse e-mail', 'email', 'e-mail'], data.email);
    setCell(ligne, header, ['fonction'], data.fonction || CONFIG.FONCTION_DEFAUT);
    setCell(ligne, header, ['statut'], data.statut || CONFIG.STATUT_NOUVEAU);
    setCell(ligne, header, ["date d'adhésion"], new Date());
    setCell(ligne, header, ['date de naissance'], data.dateNaissance);
    setCell(ligne, header, ['adresse postale'], data.adressePostale);
    sheet.appendRow(ligne);
    return { success: true, id: nextId };
  } catch (err) {
    logErreur('saveMembre', err);
    return { success: false, message: err.message };
  }
}

function uploadPhotoMembre(idMembre, base64Data, mimeType, pin) {
  try {
    verifierPinOuErreur(pin);
    if (!idMembre) throw new Error('ID membre manquant.');
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CONFIG.SHEET_MEMBRE);
    const header = getHeaderMap(sheet);
    const headerRow = header.__headerRow || 1;
    const lastRow = sheet.getLastRow();
    let ligneCible = -1, nomMembre = '';
    if (lastRow > headerRow) {
      const data = sheet.getRange(headerRow + 1, 1, lastRow - headerRow, sheet.getLastColumn()).getValues();
      for (let i = 0; i < data.length; i++) {
        if (String(getCell(data[i], header, ['id membre', 'id'])) === String(idMembre)) {
          ligneCible = i + headerRow + 1;
          nomMembre = getCell(data[i], header, ['nom & prénom', 'nom']);
          break;
        }
      }
    }
    if (ligneCible === -1) throw new Error('Membre introuvable.');
    if (!base64Data) throw new Error('Aucune image reçue.');
    const contenuBase64 = String(base64Data).split(',').pop();
    const decoded = Utilities.base64Decode(contenuBase64);
    const blob = Utilities.newBlob(decoded, mimeType || 'image/jpeg', idMembre + ' - ' + nomMembre + '.jpg');
    const folder = DriveApp.getFolderById(CONFIG.DRIVE_FOLDER_PHOTOS);
    const file = folder.createFile(blob);
    file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
    const photoUrl = 'https://drive.google.com/thumbnail?id=' + file.getId() + '&sz=w500';
    setCellDansLigne(sheet, ligneCible, header, ['photo'], photoUrl);
    return { success: true, photoUrl: photoUrl };
  } catch (err) {
    logErreur('uploadPhotoMembre', err);
    return { success: false, message: err.message };
  }
}

function supprimerMembre(pin, idMembre) {
  try {
    verifierPinOuErreur(pin);
    if (!idMembre) throw new Error('ID membre manquant.');
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheetMembre = ss.getSheetByName(CONFIG.SHEET_MEMBRE);
    const headerMembre = getHeaderMap(sheetMembre);
    const headerRowMembre = headerMembre.__headerRow || 1;
    const lastRowMembre = sheetMembre.getLastRow();
    const colIdMembre = colIndexOf(headerMembre, ['id membre', 'id']);
    let ligneMembre = -1;
    if (lastRowMembre > headerRowMembre) {
      const ids = sheetMembre.getRange(headerRowMembre + 1, colIdMembre, lastRowMembre - headerRowMembre, 1).getValues();
      for (let i = 0; i < ids.length; i++) {
        if (String(ids[i][0]) === String(idMembre)) { ligneMembre = i + headerRowMembre + 1; break; }
      }
    }
    if (ligneMembre === -1) throw new Error('Membre introuvable.');
    sheetMembre.deleteRow(ligneMembre);
    const sheetApports = ss.getSheetByName(CONFIG.SHEET_APPORTS);
    const headerApports = getHeaderMap(sheetApports);
    const headerRowApports = headerApports.__headerRow || 1;
    const lastRowApports = sheetApports.getLastRow();
    const colIdApports = colIndexOf(headerApports, ['id membre', 'id']);
    if (lastRowApports > headerRowApports && colIdApports) {
      const idsApports = sheetApports.getRange(headerRowApports + 1, colIdApports, lastRowApports - headerRowApports, 1).getValues();
      for (let i = idsApports.length - 1; i >= 0; i--) {
        if (String(idsApports[i][0]) === String(idMembre)) sheetApports.deleteRow(i + headerRowApports + 1);
      }
    }
    const sheetFinance = ss.getSheetByName(CONFIG.SHEET_FINANCE);
    if (sheetFinance) {
      const headerFinance = getHeaderMap(sheetFinance);
      const headerRowFinance = headerFinance.__headerRow || 1;
      const lastRowFinance = sheetFinance.getLastRow();
      const colDesc = colIndexOf(headerFinance, ['description']);
      if (lastRowFinance > headerRowFinance && colDesc) {
        const descs = sheetFinance.getRange(headerRowFinance + 1, colDesc, lastRowFinance - headerRowFinance, 1).getValues();
        for (let i = descs.length - 1; i >= 0; i--) {
          if (String(descs[i][0]).indexOf(String(idMembre)) !== -1) sheetFinance.deleteRow(i + headerRowFinance + 1);
        }
      }
    }
    ngmadaSynchroniserTableauDeBord();
    return { success: true, message: 'Membre supprimé avec succès.' };
  } catch (err) {
    logErreur('supprimerMembre', err);
    return { success: false, message: err.message };
  }
}

function supprimerMembreAvecCascade(pin, idMembre) {
  try {
    verifierPinOuErreur(pin);
    if (!idMembre) throw new Error('ID membre manquant.');
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const resultat = supprimerMembre(pin, idMembre);
    if (!resultat.success) return resultat;
    const apportsSheet = ss.getSheetByName(CONFIG.SHEET_APPORTS);
    if (apportsSheet) {
      const header = getHeaderMap(apportsSheet);
      const headerRow = header.__headerRow || 1;
      const lastRow = apportsSheet.getLastRow();
      if (lastRow > headerRow) {
        const colId = colIndexOf(header, ['id membre', 'id']);
        const ids = apportsSheet.getRange(headerRow + 1, colId, lastRow - headerRow, 1).getValues();
        for (let i = ids.length - 1; i >= 0; i--) {
          if (String(ids[i][0]) === String(idMembre)) {
            apportsSheet.deleteRow(i + headerRow + 1);
          }
        }
      }
    }
    const financeSheet = ss.getSheetByName(CONFIG.SHEET_FINANCE);
    if (financeSheet) {
      const header = getHeaderMap(financeSheet);
      const headerRow = header.__headerRow || 1;
      const lastRow = financeSheet.getLastRow();
      if (lastRow > headerRow) {
        const colDesc = colIndexOf(header, ['description']);
        const descs = financeSheet.getRange(headerRow + 1, colDesc, lastRow - headerRow, 1).getValues();
        for (let i = descs.length - 1; i >= 0; i--) {
          if (String(descs[i][0]).indexOf(String(idMembre)) !== -1) {
            financeSheet.deleteRow(i + headerRow + 1);
          }
        }
      }
    }
    ngmadaSynchroniserTableauDeBord();
    return { success: true, message: 'Membre et ses relations supprimés avec succès.' };
  } catch (err) {
    logErreur('supprimerMembreAvecCascade', err);
    return { success: false, message: err.message };
  }
}

// ============================================================
//  FACTURES — suppression
// ============================================================
function supprimerFacture(pin, row) {
  try {
    verifierPinOuErreur(pin);
    if (!row) throw new Error('Ligne invalide.');
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CONFIG.SHEET_FACTURE);
    if (!sheet) throw new Error('Onglet FACTURE introuvable.');
    const lastRow = sheet.getLastRow();
    if (row < 2 || row > lastRow) throw new Error('Ligne invalide.');
    sheet.deleteRow(row);
    ngmadaSynchroniserTableauDeBord();
    return { success: true, message: 'Facture supprimée avec succès.' };
  } catch (err) {
    logErreur('supprimerFacture', err);
    return { success: false, message: err.message };
  }
}

function supprimerMouvementCaisse(pin, row) {
  try {
    verifierPinOuErreur(pin);
    if (!row) throw new Error('Ligne invalide.');
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CONFIG.SHEET_FINANCE);
    if (!sheet) throw new Error('Onglet Gestion Financière introuvable.');
    const lastRow = sheet.getLastRow();
    if (row < 2 || row > lastRow) throw new Error('Ligne invalide.');
    sheet.deleteRow(row);
    ngmadaSynchroniserTableauDeBord();
    return { success: true, message: 'Mouvement supprimé avec succès.' };
  } catch (err) {
    logErreur('supprimerMouvementCaisse', err);
    return { success: false, message: err.message };
  }
}

// ============================================================
//  SUPPRESSION EN CASCADE - FACTURE
// ============================================================
function supprimerFactureAvecCascade(pin, row) {
  try {
    verifierPinOuErreur(pin);
    if (!row) throw new Error('Ligne invalide.');
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const factureSheet = ss.getSheetByName(CONFIG.SHEET_FACTURE);
    if (!factureSheet) throw new Error('Onglet FACTURE introuvable.');
    const lastRow = factureSheet.getLastRow();
    if (row < 2 || row > lastRow) throw new Error('Ligne invalide.');
    const header = getHeaderMap(factureSheet);
    const colNum = colIndexOf(header, ['n° facture', 'numero facture']);
    const numFacture = factureSheet.getRange(row, colNum).getValue();
    supprimerLigneFinanceParNumero(ss, numFacture);
    supprimerFactureDesApports(ss, numFacture);
    factureSheet.deleteRow(row);
    ngmadaSynchroniserTableauDeBord();
    return { success: true, message: 'Facture et ses relations supprimées avec succès.' };
  } catch (err) {
    logErreur('supprimerFactureAvecCascade', err);
    return { success: false, message: err.message };
  }
}

function supprimerLigneFinanceParNumero(ss, numFacture) {
  if (!numFacture) return;
  const financeSheet = ss.getSheetByName(CONFIG.SHEET_FINANCE);
  if (!financeSheet) return;
  const header = getHeaderMap(financeSheet);
  const headerRow = header.__headerRow || 1;
  const lastRow = financeSheet.getLastRow();
  if (lastRow <= headerRow) return;
  const colNum = colIndexOf(header, ['n° facture', 'numero facture']);
  if (!colNum) return;
  const data = financeSheet.getRange(headerRow + 1, colNum, lastRow - headerRow, 1).getValues();
  for (let i = data.length - 1; i >= 0; i--) {
    if (String(data[i][0]).trim() === String(numFacture).trim()) {
      financeSheet.deleteRow(i + headerRow + 1);
    }
  }
}

function supprimerFactureDesApports(ss, numFacture) {
  if (!numFacture) return;
  const apportsSheet = ss.getSheetByName(CONFIG.SHEET_APPORTS);
  if (!apportsSheet) return;
  const header = getHeaderMap(apportsSheet);
  const headerRow = header.__headerRow || 1;
  const lastRow = apportsSheet.getLastRow();
  if (lastRow <= headerRow) return;
  const colFactures = colIndexOf(header, ['n° facture', 'numero facture']);
  if (!colFactures) return;
  const data = apportsSheet.getRange(headerRow + 1, colFactures, lastRow - headerRow, 1).getValues();
  for (let i = 0; i < data.length; i++) {
    const factures = String(data[i][0] || '').split('\n').filter(Boolean);
    const nouvelleListe = factures.filter(f => String(f).trim() !== String(numFacture).trim());
    if (nouvelleListe.length !== factures.length) {
      const row = i + headerRow + 1;
      apportsSheet.getRange(row, colFactures).setValue(nouvelleListe.join('\n'));
      recalculerMontantApport(apportsSheet, row);
    }
  }
}

function recalculerMontantApport(apportsSheet, row) {
  const header = getHeaderMap(apportsSheet);
  const colFactures = colIndexOf(header, ['n° facture', 'numero facture']);
  const colMontantVerse = colIndexOf(header, ['montant versé', 'montant verse']);
  const colReste = colIndexOf(header, ['reste à verser', 'reste a verser']);
  const colPourcentage = colIndexOf(header, ['% versé', '% verse']);
  const colStatut = colIndexOf(header, ['statut']);
  const colApportPrevu = colIndexOf(header, ['apport prévu', 'apport prevu']);
  if (!colMontantVerse) return;
  const facturesText = apportsSheet.getRange(row, colFactures).getValue();
  const factures = String(facturesText || '').split('\n').filter(Boolean);
  let nouveauTotal = 0;
  const ss = apportsSheet.getParent();
  const factureSheet = ss.getSheetByName(CONFIG.SHEET_FACTURE);
  if (factureSheet) {
    const headerFacture = getHeaderMap(factureSheet);
    const colNumFacture = colIndexOf(headerFacture, ['n° facture', 'numero facture']);
    const colMontantFacture = colIndexOf(headerFacture, ['montant']);
    const dataFactures = factureSheet.getRange(headerFacture.__headerRow + 1, 1, factureSheet.getLastRow() - headerFacture.__headerRow, factureSheet.getLastColumn()).getValues();
    factures.forEach(function(num) {
      for (let i = 0; i < dataFactures.length; i++) {
        if (String(dataFactures[i][colNumFacture - 1]).trim() === String(num).trim()) {
          nouveauTotal += toNombre(dataFactures[i][colMontantFacture - 1]);
          break;
        }
      }
    });
  }
  const seuil = toNombre(apportsSheet.getRange(row, colApportPrevu).getValue()) || CONFIG.SEUIL_VALIDATION_DEFAUT;
  const reste = Math.max(seuil - nouveauTotal, 0);
  const pourcentage = seuil > 0 ? Math.min(nouveauTotal / seuil, 1) : 0;
  const statut = nouveauTotal >= seuil ? CONFIG.STATUT_VALIDE : CONFIG.STATUT_PARTIEL;
  apportsSheet.getRange(row, colMontantVerse).setValue(nouveauTotal);
  if (colReste) apportsSheet.getRange(row, colReste).setValue(reste);
  if (colPourcentage) apportsSheet.getRange(row, colPourcentage).setValue(pourcentage);
  if (colStatut) apportsSheet.getRange(row, colStatut).setValue(statut);
}

// ============================================================
//  ACTUALITÉS
// ============================================================
const SHEET_ACTUALITES = 'Actualites';
const SHEET_COMMENTAIRES = 'Commentaires';

function creerActualite(pin, titre, description, photoURL) {
  try {
    verifierPinOuErreur(pin);
    if (!titre || !description) return { success: false, message: 'Titre et description requis.' };
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName(SHEET_ACTUALITES);
    if (!sheet) return { success: false, message: 'Onglet "Actualites" introuvable.' };
    var header = getHeaderMap(sheet);
    var headerRow = header.__headerRow || 1;
    var lastRow = sheet.getLastRow();
    var dernierNumero = 0;
    if (lastRow > headerRow) {
      var colNum = colIndexOf(header, ['n°', 'numero', 'n']);
      var numeros = sheet.getRange(headerRow + 1, colNum, lastRow - headerRow, 1).getValues();
      numeros.forEach(function (r) { var n = parseInt(r[0], 10); if (!isNaN(n) && n > dernierNumero) dernierNumero = n; });
    }
    var idActualite = 'ACT-' + String(dernierNumero + 1).padStart(5, '0');
    var userEmail = Session.getActiveUser().getEmail() || 'anonymous';
    var nomUtilisateur = getNomUtilisateur(userEmail);
    var ligne = new Array(sheet.getLastColumn()).fill('');
    setCell(ligne, header, ['n°', 'numero', 'n'], dernierNumero + 1);
    setCell(ligne, header, ['date'], new Date());
    setCell(ligne, header, ['titre'], titre);
    setCell(ligne, header, ['description'], description);
    setCell(ligne, header, ['photourl', 'photo url', 'photo'], photoURL || '');
    setCell(ligne, header, ['auteur'], nomUtilisateur);
    setCell(ligne, header, ['auteur_email'], userEmail);
    setCell(ligne, header, ['likes'], 0);
    setCell(ligne, header, ['idactualite', 'id actualite'], idActualite);
    sheet.appendRow(ligne);
    return { success: true, idActualite: idActualite };
  } catch (err) {
    logErreur('creerActualite', err);
    return { success: false, message: err.message };
  }
}

function getActualites() {
  try {
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_ACTUALITES);
    if (!sheet) return [];
    var header = getHeaderMap(sheet);
    var headerRow = header.__headerRow || 1;
    var lastRow = sheet.getLastRow();
    if (lastRow <= headerRow) return [];
    var data = sheet.getRange(headerRow + 1, 1, lastRow - headerRow, sheet.getLastColumn()).getValues();
    var email = Session.getActiveUser().getEmail() || 'anonymous';
    var actualites = [];
    var colLikers = colIndexOf(header, ['likers']);
    var commentairesSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_COMMENTAIRES);
    var commentairesData = commentairesSheet ? getCommentairesCount(commentairesSheet) : {};
    data.forEach(function (r) {
      if (r.join('').toString().trim() === '') return;
      var ligne = { __row: r, __header: header };
      var likers = colLikers ? String(r[colLikers - 1] || '') : '';
      var liked = likers.split(',').includes(email);
      var id = getVal(ligne, ['idactualite', 'id actualite']);
      var auteur = getVal(ligne, ['auteur']) || 'Anonyme';
      var auteurEmail = getVal(ligne, ['auteur_email']) || '';
      if ((auteur === 'Admin' || auteur === 'Anonyme') && auteurEmail) {
        var nomTrouve = getNomUtilisateur(auteurEmail);
        if (nomTrouve && nomTrouve !== 'Anonyme') {
          auteur = nomTrouve;
        } else {
          auteur = auteurEmail.split('@')[0] || 'Anonyme';
        }
      }
      actualites.push({
        id: id,
        numero: getVal(ligne, ['n°', 'numero', 'n']),
        date: formatDateAffichage(getVal(ligne, ['date'])),
        titre: getVal(ligne, ['titre']),
        description: getVal(ligne, ['description']),
        photoURL: getVal(ligne, ['photourl', 'photo url', 'photo']),
        auteur: auteur,
        auteurEmail: auteurEmail,
        likes: parseInt(getVal(ligne, ['likes'])) || 0,
        liked: liked,
        commentaires: commentairesData[id] || 0
      });
    });
    actualites.reverse();
    return actualites;
  } catch (err) {
    logErreur('getActualites', err);
    return [];
  }
}

function getCommentairesCount(sheet) {
  try {
    if (!sheet) return {};
    const header = getHeaderMap(sheet);
    const headerRow = header.__headerRow || 1;
    const lastRow = sheet.getLastRow();
    if (lastRow <= headerRow) return {};
    const data = sheet.getRange(headerRow + 1, 1, lastRow - headerRow, sheet.getLastColumn()).getValues();
    const colId = colIndexOf(header, ['idactualite', 'id actualite']);
    const counts = {};
    data.forEach(function (r) {
      if (r.join('').toString().trim() === '') return;
      const id = String(r[colId - 1] || '');
      if (id) counts[id] = (counts[id] || 0) + 1;
    });
    return counts;
  } catch (err) { return {}; }
}

function toggleLikePublic(idActualite) {
  try {
    if (!idActualite) return { success: false, message: 'ID actualité manquant.' };
    const email = Session.getActiveUser().getEmail() || 'anonymous';
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_ACTUALITES);
    if (!sheet) return { success: false, message: 'Onglet Actualites introuvable.' };
    const header = getHeaderMap(sheet);
    const headerRow = header.__headerRow || 1;
    const lastRow = sheet.getLastRow();
    if (lastRow <= headerRow) return { success: false, message: 'Aucune actualité.' };
    const colId = colIndexOf(header, ['idactualite', 'id actualite']);
    const colLikes = colIndexOf(header, ['likes']);
    const colLikers = colIndexOf(header, ['likers']);
    const data = sheet.getRange(headerRow + 1, 1, lastRow - headerRow, sheet.getLastColumn()).getValues();
    for (let i = 0; i < data.length; i++) {
      if (String(data[i][colId - 1]) === String(idActualite)) {
        const row = i + headerRow + 1;
        const likesActuels = parseInt(data[i][colLikes - 1]) || 0;
        const likers = String(data[i][colLikers - 1] || '');
        const likersList = likers ? likers.split(',') : [];
        const alreadyLiked = likersList.includes(email);
        let nouveauLikes, nouveauLikers;
        if (alreadyLiked) {
          nouveauLikes = likesActuels - 1;
          nouveauLikers = likersList.filter(e => e !== email).join(',');
        } else {
          nouveauLikes = likesActuels + 1;
          nouveauLikers = likersList.concat(email).join(',');
        }
        sheet.getRange(row, colLikes).setValue(nouveauLikes);
        if (colLikers) sheet.getRange(row, colLikers).setValue(nouveauLikers);
        return { success: true, likes: nouveauLikes, liked: !alreadyLiked };
      }
    }
    return { success: false, message: 'Actualité introuvable.' };
  } catch (err) {
    logErreur('toggleLikePublic', err);
    return { success: false, message: err.message };
  }
}

function ajouterCommentaire(idActualite, texte, email) {
  try {
    if (!idActualite || !texte) return { success: false, message: 'ID actualité et texte requis.' };
    let userEmail = email || Session.getActiveUser().getEmail() || 'anonymous';
    const nom = getNomUtilisateur(userEmail);
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(SHEET_COMMENTAIRES);
    if (!sheet) return { success: false, message: 'Onglet "Commentaires" introuvable.' };
    const header = getHeaderMap(sheet);
    const headerRow = header.__headerRow || 1;
    const lastRow = sheet.getLastRow();
    let dernierNumero = 0;
    if (lastRow > headerRow) {
      const colNum = colIndexOf(header, ['n°', 'numero', 'n']);
      const numeros = sheet.getRange(headerRow + 1, colNum, lastRow - headerRow, 1).getValues();
      numeros.forEach(function (r) { const n = parseInt(r[0], 10); if (!isNaN(n) && n > dernierNumero) dernierNumero = n; });
    }
    const ligne = new Array(sheet.getLastColumn()).fill('');
    setCell(ligne, header, ['n°', 'numero', 'n'], dernierNumero + 1);
    setCell(ligne, header, ['idactualite', 'id actualite'], idActualite);
    setCell(ligne, header, ['auteur'], nom);
    setCell(ligne, header, ['auteur_email'], userEmail);
    setCell(ligne, header, ['texte'], texte);
    setCell(ligne, header, ['date'], new Date());
    sheet.appendRow(ligne);
    return { success: true };
  } catch (err) {
    logErreur('ajouterCommentaire', err);
    return { success: false, message: err.message };
  }
}

function getCommentaires(idActualite) {
  try {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_COMMENTAIRES);
    if (!sheet) return [];
    const header = getHeaderMap(sheet);
    const headerRow = header.__headerRow || 1;
    const lastRow = sheet.getLastRow();
    if (lastRow <= headerRow) return [];
    const data = sheet.getRange(headerRow + 1, 1, lastRow - headerRow, sheet.getLastColumn()).getValues();
    const commentaires = [];
    data.forEach(function (r) {
      if (r.join('').toString().trim() === '') return;
      const ligne = { __row: r, __header: header };
      const idActu = String(getVal(ligne, ['idactualite', 'id actualite']) || '');
      if (idActu === String(idActualite)) {
        commentaires.push({ auteur: getVal(ligne, ['auteur']), texte: getVal(ligne, ['texte']), date: formatDateAffichage(getVal(ligne, ['date'])) });
      }
    });
    return commentaires;
  } catch (err) {
    logErreur('getCommentaires', err);
    return [];
  }
}

function getNomUtilisateur(email) {
  try {
    if (!email) return 'Anonyme';
    var emailClean = String(email).trim().toLowerCase();
    var sheet = getUtilisateursSheet_();
    if (!sheet) {
      return emailClean.split('@')[0] || 'Anonyme';
    }
    var header = getHeaderMap(sheet);
    var headerRow = header.__headerRow || 1;
    var lastRow = sheet.getLastRow();
    if (lastRow <= headerRow) {
      return emailClean.split('@')[0] || 'Anonyme';
    }
    var data = sheet.getRange(headerRow + 1, 1, lastRow - headerRow, sheet.getLastColumn()).getValues();
    var colEmail = colIndexOf(header, ['email']);
    var colNom = colIndexOf(header, ['nom']);
    for (var i = 0; i < data.length; i++) {
      var emailLigne = String(data[i][colEmail - 1] || '').trim().toLowerCase();
      if (emailLigne === emailClean) {
        var nom = String(data[i][colNom - 1] || '').trim();
        if (nom) return nom;
      }
    }
    return emailClean.split('@')[0] || 'Anonyme';
  } catch (err) {
    console.error('Erreur getNomUtilisateur:', err);
    return 'Anonyme';
  }
}

function uploadPhotoActualite(pin, base64Data, mimeType) {
  try {
    verifierPinOuErreur(pin);
    if (!base64Data) return { success: false, message: 'Aucune image reçue.' };
    const contenuBase64 = String(base64Data).split(',').pop();
    const decoded = Utilities.base64Decode(contenuBase64);
    const blob = Utilities.newBlob(decoded, mimeType || 'image/jpeg', 'actu_' + new Date().getTime() + '.jpg');
    const folder = DriveApp.getFolderById(CONFIG.DRIVE_FOLDER_AFFICHES_PHOTOS || '1-1aYDgllz0sCWShD4w1a0V1fW7i7y9t3');
    const file = folder.createFile(blob);
    file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
    const photoUrl = 'https://drive.google.com/thumbnail?id=' + file.getId() + '&sz=w800';
    return { success: true, photoUrl: photoUrl };
  } catch (err) {
    logErreur('uploadPhotoActualite', err);
    return { success: false, message: err.message };
  }
}

function modifierActualite(pin, idActualite, titre, description, photoURL) {
  try {
    verifierPinOuErreur(pin);
    if (!idActualite) throw new Error('ID actualité manquant.');
    if (!titre || !description) throw new Error('Titre et description requis.');
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(SHEET_ACTUALITES);
    if (!sheet) throw new Error('Onglet Actualites introuvable.');
    const header = getHeaderMap(sheet);
    const headerRow = header.__headerRow || 1;
    const lastRow = sheet.getLastRow();
    const colId = colIndexOf(header, ['idactualite', 'id actualite']);
    const colTitre = colIndexOf(header, ['titre']);
    const colDesc = colIndexOf(header, ['description']);
    const colPhoto = colIndexOf(header, ['photourl', 'photo url', 'photo']);
    if (lastRow > headerRow && colId) {
      const data = sheet.getRange(headerRow + 1, colId, lastRow - headerRow, 1).getValues();
      for (let i = 0; i < data.length; i++) {
        if (String(data[i][0]) === String(idActualite)) {
          const row = i + headerRow + 1;
          sheet.getRange(row, colTitre).setValue(titre);
          sheet.getRange(row, colDesc).setValue(description);
          if (colPhoto) sheet.getRange(row, colPhoto).setValue(photoURL || '');
          return { success: true };
        }
      }
    }
    return { success: false, message: 'Actualité introuvable.' };
  } catch (err) {
    logErreur('modifierActualite', err);
    return { success: false, message: err.message };
  }
}

function supprimerActualite(pin, idActualite) {
  try {
    verifierPinOuErreur(pin);
    if (!idActualite) throw new Error('ID actualité manquant.');
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheetActu = ss.getSheetByName(SHEET_ACTUALITES);
    if (!sheetActu) throw new Error('Onglet Actualites introuvable.');
    const header = getHeaderMap(sheetActu);
    const headerRow = header.__headerRow || 1;
    const lastRow = sheetActu.getLastRow();
    const colId = colIndexOf(header, ['idactualite', 'id actualite']);
    if (lastRow > headerRow && colId) {
      const data = sheetActu.getRange(headerRow + 1, colId, lastRow - headerRow, 1).getValues();
      for (let i = data.length - 1; i >= 0; i--) {
        if (String(data[i][0]) === String(idActualite)) { sheetActu.deleteRow(i + headerRow + 1); break; }
      }
    }
    const sheetComment = ss.getSheetByName(SHEET_COMMENTAIRES);
    if (sheetComment) {
      const headerC = getHeaderMap(sheetComment);
      const headerRowC = headerC.__headerRow || 1;
      const lastRowC = sheetComment.getLastRow();
      const colIdC = colIndexOf(headerC, ['idactualite', 'id actualite']);
      if (lastRowC > headerRowC && colIdC) {
        const dataC = sheetComment.getRange(headerRowC + 1, colIdC, lastRowC - headerRowC, 1).getValues();
        for (let i = dataC.length - 1; i >= 0; i--) {
          if (String(dataC[i][0]) === String(idActualite)) sheetComment.deleteRow(i + headerRowC + 1);
        }
      }
    }
    return { success: true };
  } catch (err) {
    logErreur('supprimerActualite', err);
    return { success: false, message: err.message };
  }
}

// ============================================================
//  PROFIL UTILISATEUR
// ============================================================
function getProfilUtilisateur(email) {
  try {
    if (!email) return { success: false, message: 'Email manquant.' };
    const sheet = getUtilisateursSheet_();
    if (!sheet) return { success: false, message: 'Onglet Utilisateures introuvable.' };
    const header = getHeaderMap(sheet);
    const headerRow = header.__headerRow || 1;
    const lastRow = sheet.getLastRow();
    if (lastRow <= headerRow) return { success: false, message: 'Aucun utilisateur.' };
    const data = sheet.getRange(headerRow + 1, 1, lastRow - headerRow, sheet.getLastColumn()).getValues();
    for (let i = 0; i < data.length; i++) {
      const emailLigne = normalizeEmail(getCell(data[i], header, ['email']));
      if (emailLigne === email) {
        return {
          success: true,
          nom: getCell(data[i], header, ['nom']),
          email: email,
          photo: getCell(data[i], header, ['photo']) || '',
          dateCreation: formatDateAffichage(getCell(data[i], header, ['date de création', 'date de creation']))
        };
      }
    }
    return { success: false, message: 'Utilisateur introuvable.' };
  } catch (err) {
    logErreur('getProfilUtilisateur', err);
    return { success: false, message: err.message };
  }
}

function mettreAJourNomUtilisateur(nouveauNom, email) {
  try {
    if (!email) return { success: false, message: 'Email manquant.' };
    if (!nouveauNom || nouveauNom.trim().length < 2) return { success: false, message: 'Nom trop court.' };
    const sheet = getUtilisateursSheet_();
    if (!sheet) return { success: false, message: 'Onglet Utilisateures introuvable.' };
    const header = getHeaderMap(sheet);
    const headerRow = header.__headerRow || 1;
    const lastRow = sheet.getLastRow();
    if (lastRow <= headerRow) return { success: false, message: 'Aucun utilisateur.' };
    const data = sheet.getRange(headerRow + 1, 1, lastRow - headerRow, sheet.getLastColumn()).getValues();
    const colNom = colIndexOf(header, ['nom']);
    const colEmail = colIndexOf(header, ['email']);
    for (let i = 0; i < data.length; i++) {
      if (normalizeEmail(data[i][colEmail - 1]) === normalizeEmail(email)) {
        const row = i + headerRow + 1;
        sheet.getRange(row, colNom).setValue(nouveauNom.trim());
        return { success: true, nom: nouveauNom.trim() };
      }
    }
    return { success: false, message: 'Utilisateur introuvable.' };
  } catch (err) {
    logErreur('mettreAJourNomUtilisateur', err);
    return { success: false, message: err.message };
  }
}

function uploadPhotoProfil(base64Data, mimeType, email) {
  try {
    if (!email) return { success: false, message: 'Email manquant.' };
    if (!base64Data) return { success: false, message: 'Aucune image reçue.' };
    const sheet = getUtilisateursSheet_();
    if (!sheet) return { success: false, message: 'Onglet Utilisateures introuvable.' };
    const header = getHeaderMap(sheet);
    const headerRow = header.__headerRow || 1;
    const lastRow = sheet.getLastRow();
    if (lastRow <= headerRow) return { success: false, message: 'Aucun utilisateur.' };
    const data = sheet.getRange(headerRow + 1, 1, lastRow - headerRow, sheet.getLastColumn()).getValues();
    const colEmail = colIndexOf(header, ['email']);
    const colPhoto = colIndexOf(header, ['photo']);
    let ligneTrouvee = -1;
    for (let i = 0; i < data.length; i++) {
      if (normalizeEmail(data[i][colEmail - 1]) === normalizeEmail(email)) { ligneTrouvee = i + headerRow + 1; break; }
    }
    if (ligneTrouvee === -1) return { success: false, message: 'Utilisateur introuvable.' };
    const contenuBase64 = String(base64Data).split(',').pop();
    const decoded = Utilities.base64Decode(contenuBase64);
    const blob = Utilities.newBlob(decoded, mimeType || 'image/jpeg', 'profil_' + email + '.jpg');
    const folderId = CONFIG.DRIVE_FOLDER_PHOTOS_PROFIL || CONFIG.DRIVE_FOLDER_PHOTOS || '11I5ajvNiu0qRgtCChHtIWxEU122bFmrU';
    const folder = DriveApp.getFolderById(folderId);
    const file = folder.createFile(blob);
    file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
    const photoUrl = 'https://drive.google.com/thumbnail?id=' + file.getId() + '&sz=w200';
    sheet.getRange(ligneTrouvee, colPhoto).setValue(photoUrl);
    return { success: true, photoUrl: photoUrl };
  } catch (err) {
    logErreur('uploadPhotoProfil', err);
    return { success: false, message: err.message };
  }
}

function getListeUtilisateurs() {
  try {
    const sheet = getUtilisateursSheet_();
    if (!sheet) return [];
    const header = getHeaderMap(sheet);
    const headerRow = header.__headerRow || 1;
    const lastRow = sheet.getLastRow();
    if (lastRow <= headerRow) return [];
    const data = sheet.getRange(headerRow + 1, 1, lastRow - headerRow, sheet.getLastColumn()).getValues();
    const utilisateurs = [];
    const now = new Date();
    const colNom = colIndexOf(header, ['nom']);
    const colEmail = colIndexOf(header, ['email']);
    const colConnexion = colIndexOf(header, ['date et heure connecté', 'date et heure connecte']);
    const colPhoto = colIndexOf(header, ['photo']);
    const colDateCreation = colIndexOf(header, ['date de création', 'date de creation']);
    for (let i = 0; i < data.length; i++) {
      const row = data[i];
      if (row.join('').trim() === '') continue;
      const nom = (colNom ? row[colNom - 1] : '') || 'Inconnu';
      const email = (colEmail ? row[colEmail - 1] : '') || '';
      if (!email) continue;
      const photo = (colPhoto ? row[colPhoto - 1] : '') || '';
      const dateCreation = (colDateCreation ? row[colDateCreation - 1] : null);
      let estEnLigne = false;
      if (colConnexion) {
        const dateConnexion = row[colConnexion - 1];
        if (dateConnexion) {
          let d = (dateConnexion instanceof Date) ? dateConnexion : new Date(dateConnexion);
          if (!isNaN(d.getTime())) { const diffMinutes = (now - d) / (1000 * 60); estEnLigne = diffMinutes < 5; }
        }
      }
      utilisateurs.push({ nom, email, photo, dateCreation: dateCreation ? Utilities.formatDate(dateCreation, Session.getScriptTimeZone(), 'dd/MM/yyyy') : '', estEnLigne });
    }
    return utilisateurs;
  } catch (err) {
    logErreur('getListeUtilisateurs', err);
    return [];
  }
}

function supprimerUtilisateur(pin, emailASupprimer) {
  try {
    verifierPinOuErreur(pin);
    const email = Session.getActiveUser().getEmail();
    if (!email) return { success: false, message: 'Non connecté.' };
    if (email === emailASupprimer) return { success: false, message: 'Vous ne pouvez pas supprimer votre propre compte.' };
    const sheet = getUtilisateursSheet_();
    const header = getHeaderMap(sheet);
    const headerRow = header.__headerRow || 1;
    const lastRow = sheet.getLastRow();
    if (lastRow <= headerRow) return { success: false, message: 'Aucun utilisateur.' };
    const data = sheet.getRange(headerRow + 1, 1, lastRow - headerRow, sheet.getLastColumn()).getValues();
    const colEmail = colIndexOf(header, ['email']);
    for (let i = data.length - 1; i >= 0; i--) {
      if (normalizeEmail(data[i][colEmail - 1]) === normalizeEmail(emailASupprimer)) { sheet.deleteRow(i + headerRow + 1); return { success: true, message: 'Utilisateur supprimé.' }; }
    }
    return { success: false, message: 'Utilisateur introuvable.' };
  } catch (err) {
    logErreur('supprimerUtilisateur', err);
    return { success: false, message: err.message };
  }
}

function mettreAJourDateConnexion(email) {
  try {
    if (!email) return { success: false };
    const sheet = getUtilisateursSheet_();
    if (!sheet) return { success: false };
    const header = getHeaderMap(sheet);
    const headerRow = header.__headerRow || 1;
    const lastRow = sheet.getLastRow();
    if (lastRow <= headerRow) return { success: false };
    const data = sheet.getRange(headerRow + 1, 1, lastRow - headerRow, sheet.getLastColumn()).getValues();
    const colEmail = colIndexOf(header, ['email']);
    for (let i = 0; i < data.length; i++) {
      if (normalizeEmail(data[i][colEmail - 1]) === normalizeEmail(email)) {
        const row = i + headerRow + 1;
        setCellDansLigne(sheet, row, header, ['date et heure connecté', 'date et heure connecte'], new Date());
        return { success: true };
      }
    }
    return { success: false };
  } catch (err) { return { success: false }; }
}

// ============================================================
//  VÉRIFICATION PUBLIQUE PAR QR
// ============================================================
function pageVerificationMembre(id) {
  let trouve = null;
  try {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CONFIG.SHEET_MEMBRE);
    const header = getHeaderMap(sheet);
    const headerRow = header.__headerRow || 1;
    const lastRow = sheet.getLastRow();
    if (lastRow > headerRow) {
      const data = sheet.getRange(headerRow + 1, 1, lastRow - headerRow, sheet.getLastColumn()).getValues();
      for (let i = 0; i < data.length; i++) {
        const idLigne = getCell(data[i], header, ['id membre', 'id']);
        if (String(idLigne).trim() === String(id).trim()) {
          trouve = {
            id: idLigne,
            nom: getCell(data[i], header, ['nom & prénom', 'nom']),
            fonction: getCell(data[i], header, ['fonction']),
            statut: getCell(data[i], header, ['statut']),
            photo: getCell(data[i], header, ['photo'])
          };
          break;
        }
      }
    }
  } catch (err) { logErreur('pageVerificationMembre', err); }
  const ok = trouve && String(trouve.statut || '').toLowerCase().indexOf('valid') !== -1;
  const html = trouve
    ? '<html><head><meta charset="UTF-8"><meta name="viewport" content="width=device-width, initial-scale=1">' +
      '<style>body{font-family:Arial,sans-serif;background:linear-gradient(135deg,#0f172a,#1e293b);color:#fff;text-align:center;padding:40px 20px;}' +
      'img{width:100px;height:100px;border-radius:50%;border:3px solid #facc15;object-fit:cover;margin-bottom:14px;}' +
      'h1{color:' + (ok ? '#4ade80' : '#facc15') + ';font-size:20px;}p{opacity:0.85;margin:4px 0;}.badge{display:inline-block;margin-top:10px;padding:6px 16px;border-radius:20px;font-weight:bold;' +
      'background:' + (ok ? 'rgba(74,222,128,0.15)' : 'rgba(250,204,21,0.15)') + ';color:' + (ok ? '#4ade80' : '#facc15') + ';}</style></head><body>' +
      '<img src="' + (trouve.photo || 'https://via.placeholder.com/100') + '">' +
      '<h1>' + (ok ? '✔ Associé vérifié' : '⏳ Statut : ' + (trouve.statut || 'à vérifier')) + '</h1>' +
      '<p><b>' + (trouve.nom || '') + '</b></p><p>' + (trouve.fonction || '') + '</p><p>ID : ' + trouve.id + '</p>' +
      '<div class="badge">' + (ok ? 'ASSOCIÉ VALIDÉ' : (trouve.statut || 'EN ATTENTE')) + '</div></body></html>'
    : '<html><head><meta charset="UTF-8"><meta name="viewport" content="width=device-width, initial-scale=1">' +
      '<style>body{font-family:Arial,sans-serif;background:#1e293b;color:#f87171;text-align:center;padding:60px 20px;}</style>' +
      '</head><body><h1>✖ ID introuvable</h1><p>Aucun membre ne correspond à cet identifiant.</p></body></html>';
  return HtmlService.createHtmlOutput(html).setTitle('Vérification — NG-MADA').addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

// ============================================================
//  DASHBOARD — synchronisation
// ============================================================
function trouverDateLimite(dashboardSheet) {
  if (!dashboardSheet) return null;
  const lastRow = dashboardSheet.getLastRow();
  const lastCol = dashboardSheet.getLastColumn();
  if (lastRow < 1 || lastCol < 1) return null;
  const valeurs = dashboardSheet.getRange(1, 1, lastRow, lastCol).getValues();
  for (let i = 0; i < valeurs.length; i++) {
    for (let j = 0; j < valeurs[i].length; j++) {
      const cellule = valeurs[i][j];
      if (typeof cellule === 'string' && cellule.toLowerCase().indexOf('date limite') !== -1) {
        for (let k = j + 1; k < valeurs[i].length; k++) {
          const v = valeurs[i][k];
          if (v === '' || v === null || v === undefined) continue;
          const d = (v instanceof Date) ? v : new Date(v);
          if (!isNaN(d.getTime())) return d;
        }
      }
    }
  }
  return null;
}

function ngmadaSynchroniserTableauDeBord() {
  SpreadsheetApp.flush();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(CONFIG.SHEET_DASHBOARD);
  if (!sheet) { Logger.log('❌ Onglet Tableau de bord introuvable.'); return; }
  const d = getDashboardData();
  if (d.__error) { Logger.log('❌ Erreur getDashboardData : ' + d.__error); return; }
  const capitalVise = toNombre(sheet.getRange('D7').getValue());
  sheet.getRange('B10').setValue(d.joursRestants != null ? d.joursRestants : '—');
  sheet.getRange('D10').setValue(d.associesValides || 0);
  sheet.getRange('F10').setValue((d.totalVerse || 0) + ' Ar');
  sheet.getRange('H10').setValue((d.totalReste || 0) + ' Ar');
  sheet.getRange('J9').setValue('En attente');
  sheet.getRange('J9').setFontSize(9);
  sheet.getRange('J9').setFontColor(ngmadaPalette.texteAtt);
  sheet.getRange('J9').setFontWeight('bold');
  sheet.getRange('J9').setWrap(true);
  sheet.getRange('J9').setBackground(ngmadaPalette.carte);
  sheet.getRange('J10').setValue(d.enAttente || 0);
  sheet.getRange('J10').setFontSize(18);
  sheet.getRange('J10').setFontWeight('bold');
  sheet.getRange('J10').setFontColor(ngmadaPalette.bleu);
  sheet.getRange('J10').setBackground(ngmadaPalette.carte);
  sheet.getRange('J9:J10').setBorder(true, true, true, true, false, false, ngmadaPalette.bordure, SpreadsheetApp.BorderStyle.SOLID);
  if (capitalVise > 0) {
    const pourcentage = Math.min((d.totalVerse || 0) / capitalVise, 1);
    const pourcentageAffichage = Math.round(pourcentage * 100);
    sheet.getRange('D14').setValue(pourcentageAffichage + '%');
    sheet.getRange('D14').setFontColor(ngmadaPalette.or);
    const nbPleins = Math.round(pourcentage * 30);
    const barre = '█'.repeat(Math.min(nbPleins, 30)) + '░'.repeat(Math.max(30 - nbPleins, 0));
    sheet.getRange('B14').setValue(barre);
    sheet.getRange('B14').setFontColor(ngmadaPalette.vert);
    sheet.getRange('B15').setValue((d.totalVerse || 0).toLocaleString('fr-FR') + ' Ar collectés sur ' + capitalVise.toLocaleString('fr-FR') + ' Ar objectif');
  } else {
    sheet.getRange('D14').setValue('— (objectif non défini)');
    sheet.getRange('D14').setFontColor(ngmadaPalette.texteAtt);
    sheet.getRange('B14').setValue('██████████████████████████████');
    sheet.getRange('B14').setFontColor(ngmadaPalette.texteAtt);
    sheet.getRange('B15').setValue((d.totalVerse || 0).toLocaleString('fr-FR') + ' Ar collectés (objectif non défini)');
  }
  const statutBandeau = sheet.getRange('B24:H25');
  if (statutBandeau) {
    const soldeCaisse = d.soldeCaisse || 0;
    const statutSolde = soldeCaisse >= 0 ? '🟢' : '🔴';
    const texteSolde = soldeCaisse >= 0 ? 'TOUT EST BON' : 'ATTENTION : SOLDE NÉGATIF';
    const couleurSolde = soldeCaisse >= 0 ? ngmadaPalette.vert : ngmadaPalette.rouge;
    const jours = d.joursRestants != null ? d.joursRestants : '—';
    const pct = capitalVise > 0 ? Math.round(Math.min((d.totalVerse || 0) / capitalVise, 1) * 100) : 0;
    const enAttente = d.enAttente || 0;
    const texteFinal = statutSolde + '  ' + texteSolde + '  |  ⚡ RAPPEL : ' + jours + ' jours restants  |  📊 ' + pct + '% atteint  |  ⏳ ' + enAttente + ' en attente';
    statutBandeau.setValue(texteFinal);
    statutBandeau.setFontColor(couleurSolde);
  }
  Logger.log('✅ Tableau de bord synchronisé !');
}

// ============================================================
//  MODULE D'AIDE
// ============================================================
function afficherAideUtilisateur() {
  const html = HtmlService.createHtmlOutput(
    '<div style="font-family:Segoe UI,sans-serif;padding:20px;color:#2b2b2b;">' +
    '<h2 style="color:#1b4332;margin-top:0;">📖 Guide Rapide - NG-MADA</h2>' +
    '<hr style="border-color:#d4a373;">' +
    '<p><strong>1. Navigation</strong><br>Utilisez les onglets pour basculer entre les sections.</p>' +
    '<p><strong>2. Sécurité</strong><br>Cliquez sur "🔒 Admin" pour déverrouiller les actions de modification.</p>' +
    '<p><strong>3. Enregistrement</strong><br>Remplissez tous les champs marqués d\'un * puis cliquez sur Enregistrer.</p>' +
    '<p><strong>4. Chargement</strong><br>Patientez pendant l\'indicateur "ZH" : ne cliquez pas deux fois.</p>' +
    '<p><strong>5. Problèmes ?</strong><br>Contactez votre administrateur si une erreur persiste.</p>' +
    '<hr style="border-color:#d4a373;">' +
    '<p style="font-size:12px;color:#6c757d;text-align:center;">NEXT GENERATION MADA — "NG-MADA"</p>' +
    '</div>'
  ).setWidth(440).setHeight(400).setTitle('Aide & Instructions NG-MADA');
  SpreadsheetApp.getUi().showModalDialog(html, '📖 Aide & Instructions');
}

// ============================================================
//  UTILITAIRES GÉNÉRIQUES
// ============================================================
function detecterLigneEntetes(sheet, maxRows) {
  maxRows = maxRows || 10;
  const numRows = Math.min(maxRows, sheet.getLastRow());
  if (numRows < 1) return 1;
  const values = sheet.getRange(1, 1, numRows, sheet.getLastColumn()).getValues();
  for (let i = 0; i < values.length; i++) {
    const nonEmpty = values[i].filter(function (v) { return String(v).trim() !== ''; }).length;
    if (nonEmpty >= 3) return i + 1;
  }
  return 1;
}

function getHeaderMap(sheet) {
  const headerRow = detecterLigneEntetes(sheet);
  const values = sheet.getRange(headerRow, 1, 1, sheet.getLastColumn()).getValues()[0];
  const map = {};
  values.forEach(function (h, i) {
    const key = String(h).replace(/\n/g, ' ').trim().toLowerCase();
    if (key) map[key] = i;
  });
  Object.defineProperty(map, '__headerRow', { value: headerRow, enumerable: false });
  return map;
}

function lireFeuilleObjets(sheet) {
  if (!sheet) return [];
  const header = getHeaderMap(sheet);
  const headerRow = header.__headerRow || 1;
  const lastRow = sheet.getLastRow();
  const lastCol = sheet.getLastColumn();
  if (lastRow <= headerRow) return [];
  const data = sheet.getRange(headerRow + 1, 1, lastRow - headerRow, lastCol).getValues();
  return data.filter(function (row) { return row.join('').toString().trim() !== ''; })
    .map(function (row) { return { __row: row, __header: header }; });
}

function colIndexOf(header, keywords) {
  for (const k of keywords) {
    for (const key in header) {
      if (key.indexOf(k) === 0 || key === k) return header[key] + 1;
    }
  }
  return 1;
}

function getCell(rowValues, header, keywords) {
  for (const k of keywords) {
    for (const key in header) {
      if (key.indexOf(k) === 0 || key === k) return rowValues[header[key]];
    }
  }
  return '';
}

function setCell(rowArray, header, keywords, value) {
  for (const k of keywords) {
    for (const key in header) {
      if (key.indexOf(k) === 0 || key === k) { rowArray[header[key]] = value; return; }
    }
  }
}

function normalizePhone(phone) { if (!phone) return ''; const digits = String(phone).replace(/[^0-9]/g, ''); return digits.length >= 9 ? digits.slice(-9) : digits; }
function normalizeEmail(email) { return email ? String(email).trim().toLowerCase() : ''; }
function normaliserNomComplet(str) { return String(str || '').trim().replace(/\s+/g, ' ').toUpperCase(); }
function getVal(objetLigne, keywords) { return getCell(objetLigne.__row, objetLigne.__header, keywords); }
function toNombre(v) { if (typeof v === 'number') return v; if (!v) return 0; const n = parseFloat(String(v).replace(/[^\d.,-]/g, '').replace(',', '.')); return isNaN(n) ? 0 : n; }
function formatDateAffichage(v) { if (!v) return ''; try { const d = (v instanceof Date) ? v : new Date(v); if (isNaN(d.getTime())) return String(v); return Utilities.formatDate(d, Session.getScriptTimeZone(), 'dd/MM/yyyy'); } catch (e) { return String(v); } }
function logErreur(fonction, err) {
  Logger.log('Erreur dans ' + fonction + ' : ' + err);
  if (ADMIN_EMAIL) MailApp.sendEmail(ADMIN_EMAIL, 'Erreur script NG-MADA', 'Erreur dans ' + fonction + ' :\n' + err);
}
function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

const ngmadaPalette = {
  texteAtt: '#94a3b8',
  carte: '#1e293b',
  bleu: '#60a5fa',
  bordure: '#334155',
  or: '#facc15',
  vert: '#4ade80',
  rouge: '#ef4444'
};

// ============================================================
//  MENU DU SHEET (onOpen)
// ============================================================
function onOpen() {
  const ui = SpreadsheetApp.getUi();
  const menu = ui.createMenu('📊 Dashboard NG-MADA');
  menu.addItem('📊 Ouvrir le tableau de bord', 'showDashboard');
  menu.addSeparator();
  menu.addItem('📖 Aide / Guide', 'afficherAideUtilisateur');
  menu.addSeparator();
  menu.addItem('🎨 Générer le design', 'ngmadaReconstruireTableauDeBord');
  menu.addItem('🔄 Synchroniser les données', 'ngmadaSynchroniserTableauDeBord');
  menu.addItem('🚀 Reconstruire + Synchroniser', 'ngmadaReinitialiserDashboardComplet');
  menu.addToUi();
}

function showDashboard() {
  const html = HtmlService.createHtmlOutputFromFile('Dashboard')
    .setWidth(900).setHeight(650).setTitle('Suivi Apports Associés - NG-MADA');
  SpreadsheetApp.getUi().showModalDialog(html, '📊 Suivi des Apports');
}

function ngmadaReconstruireTableauDeBord() {
  Logger.log('Design reconstruit (simulation)');
}

function ngmadaReinitialiserDashboardComplet() {
  ngmadaReconstruireTableauDeBord();
  ngmadaSynchroniserTableauDeBord();
}

function getEnvironnement() {
  const nomFichier = SpreadsheetApp.getActiveSpreadsheet().getName();
  if (nomFichier.toLowerCase().includes('[dev]')) return 'DEV';
  const url = ScriptApp.getService().getUrl();
  if (url && url.indexOf('dev') !== -1) return 'DEV';
  return 'PROD';
}

// ============================================================
//  PHOTO MEMBRE - VERSION PUBLIQUE (SANS PIN)
// ============================================================
function uploadPhotoMembrePublic(idMembre, base64Data, mimeType) {
  try {
    if (!idMembre) throw new Error('ID membre manquant.');
    if (!base64Data) throw new Error('Aucune image reçue.');
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CONFIG.SHEET_MEMBRE);
    const header = getHeaderMap(sheet);
    const headerRow = header.__headerRow || 1;
    const lastRow = sheet.getLastRow();
    let ligneCible = -1;
    let nomMembre = '';
    if (lastRow > headerRow) {
      const data = sheet.getRange(headerRow + 1, 1, lastRow - headerRow, sheet.getLastColumn()).getValues();
      for (let i = 0; i < data.length; i++) {
        if (String(getCell(data[i], header, ['id membre', 'id'])) === String(idMembre)) {
          ligneCible = i + headerRow + 1;
          nomMembre = getCell(data[i], header, ['nom & prénom', 'nom']);
          break;
        }
      }
    }
    if (ligneCible === -1) throw new Error('Membre introuvable.');
    const contenuBase64 = String(base64Data).split(',').pop();
    const decoded = Utilities.base64Decode(contenuBase64);
    const blob = Utilities.newBlob(decoded, mimeType || 'image/jpeg', idMembre + ' - ' + nomMembre + '.jpg');
    const folder = DriveApp.getFolderById(CONFIG.DRIVE_FOLDER_PHOTOS);
    const file = folder.createFile(blob);
    file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
    const photoUrl = 'https://drive.google.com/thumbnail?id=' + file.getId() + '&sz=w500';
    setCellDansLigne(sheet, ligneCible, header, ['photo'], photoUrl);
    return { success: true, photoUrl: photoUrl };
  } catch (err) {
    logErreur('uploadPhotoMembrePublic', err);
    return { success: false, message: err.message };
  }
}

function testerConnexion() {
  var pepper = PropertiesService.getScriptProperties().getProperty('PASSWORD_PEPPER');
  Logger.log('Pepper actuel : "' + pepper + '"');
  var motDePasse = 'VOTRE_MOT_DE_PASSE';
  var contenu = pepper + '::' + motDePasse;
  var octets = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, contenu, Utilities.Charset.UTF_8);
  var hash = octets.map(function (o) { return ('0' + (o & 0xFF).toString(16)).slice(-2); }).join('');
  Logger.log('Hash calculé : ' + hash);
  var hashStocke = '039d005a494b768a4ed4f89a5f8d662017ef42a04a73e24138fabddb2d6a89f9';
  Logger.log('Hash stocké  : ' + hashStocke);
  Logger.log('Correspond ? ' + (hash === hashStocke));
}

// ============================================================
//  CONFIGURATION DES SECRETS
// ============================================================
function configurerSecrets() {
  var userEmail = Session.getActiveUser().getEmail();
  var authorizedEmails = [
    'tovoherinirinaeliceronaldo@gmail.com',
    'chamssidinyfast@gmail.com'
  ];
  if (authorizedEmails.indexOf(userEmail) === -1) {
    Logger.log('⛔ Accès refusé pour ' + userEmail);
    throw new Error('Vous n\'êtes pas autorisé à configurer les secrets.');
  }
  var ui = SpreadsheetApp.getUi();
  var response = ui.alert(
    '⚠️ Configuration des secrets',
    'Voulez-vous vraiment réinitialiser les secrets ?\n\n' +
    'PIN_ADMIN : ' + (PropertiesService.getScriptProperties().getProperty('PIN_ADMIN') || 'Non défini') + '\n' +
    'PASSWORD_PEPPER : ' + (PropertiesService.getScriptProperties().getProperty('PASSWORD_PEPPER') ? 'Défini' : 'Non défini'),
    ui.ButtonSet.YES_NO
  );
  if (response !== ui.Button.YES) {
    Logger.log('❌ Configuration annulée par l\'utilisateur');
    return '❌ Configuration annulée.';
  }
  var nouveauPin = ui.prompt(
    '🔑 Nouveau PIN Admin',
    'Entrez le nouveau code PIN (4 chiffres recommandé) :',
    ui.ButtonSet.OK_CANCEL
  );
  if (nouveauPin.getSelectedButton() !== ui.Button.OK || !nouveauPin.getResponseText()) {
    Logger.log('❌ Configuration annulée - PIN non fourni');
    return '❌ PIN non fourni. Configuration annulée.';
  }
  var nouveauPinValue = nouveauPin.getResponseText().trim();
  if (nouveauPinValue.length < 4 || !/^\d+$/.test(nouveauPinValue)) {
    Logger.log('❌ PIN invalide');
    return '❌ PIN invalide. Utilisez au moins 4 chiffres.';
  }
  var nouveauPepper = ui.prompt(
    '🧂 Nouveau Pepper',
    'Entrez un nouveau pepper (chaîne aléatoire) :\n' +
    'Laissez vide pour garder le pepper existant.',
    ui.ButtonSet.OK_CANCEL
  );
  if (nouveauPepper.getSelectedButton() !== ui.Button.OK) {
    Logger.log('❌ Configuration annulée - pepper non fourni');
    return '❌ Pepper non fourni. Configuration annulée.';
  }
  var nouveauPepperValue = nouveauPepper.getResponseText().trim();
  var scriptProperties = PropertiesService.getScriptProperties();
  var updated = {};
  if (nouveauPinValue) {
    updated['PIN_ADMIN'] = nouveauPinValue;
  }
  if (nouveauPepperValue) {
    updated['PASSWORD_PEPPER'] = nouveauPepperValue;
  }
  if (Object.keys(updated).length === 0) {
    Logger.log('❌ Aucune modification');
    return '❌ Aucune modification effectuée.';
  }
  scriptProperties.setProperties(updated);
  Logger.log('✅ Secrets configurés par ' + userEmail);
  for (var key in updated) {
    Logger.log('   ✅ ' + key + ' : ' + (key === 'PIN_ADMIN' ? '***' : 'Défini'));
  }
  ui.alert(
    '✅ Configuration réussie',
    'Les secrets ont été configurés avec succès !\n\n' +
    'PIN_ADMIN : ' + '***' + '\n' +
    'PASSWORD_PEPPER : ' + (nouveauPepperValue ? 'Nouveau pepper défini' : 'Conservé'),
    ui.ButtonSet.OK
  );
  return '✅ Secrets configurés avec succès.';
}

// ============================================================
//  VÉRIFICATION DES SECRETS
// ============================================================
function verifierSecrets() {
  var scriptProperties = PropertiesService.getScriptProperties();
  var pin = scriptProperties.getProperty('PIN_ADMIN');
  var pepper = scriptProperties.getProperty('PASSWORD_PEPPER');
  Logger.log('🔐 Vérification des secrets :');
  Logger.log('   PIN_ADMIN : ' + (pin ? '✅ Défini (***)' : '❌ Non défini'));
  Logger.log('   PASSWORD_PEPPER : ' + (pepper ? '✅ Défini' : '❌ Non défini'));
  return {
    pinConfigured: !!pin,
    pepperConfigured: !!pepper
  };
}

// ============================================================
//  SUPPRIMER LES SECRETS (URGENCE)
// ============================================================
function supprimerSecrets() {
  var userEmail = Session.getActiveUser().getEmail();
  var authorizedEmails = [
    'tovoherinirinaeliceronaldo@gmail.com',
    'chamssidinyfast@gmail.com'
  ];
  if (authorizedEmails.indexOf(userEmail) === -1) {
    throw new Error('⛔ Accès refusé.');
  }
  var ui = SpreadsheetApp.getUi();
  var response = ui.alert(
    '⚠️ SUPPRESSION DES SECRETS',
    '⚠️ ATTENTION : Cette action est IRREVERSIBLE !\n\n' +
    'Voulez-vous vraiment supprimer tous les secrets ?\n' +
    'Cela rendra l\'application inaccessible.',
    ui.ButtonSet.YES_NO
  );
  if (response !== ui.Button.YES) {
    return '❌ Suppression annulée.';
  }
  var confirmation = ui.prompt(
    '⚠️ Confirmation finale',
    'Tapez "SUPPRIMER" pour confirmer :',
    ui.ButtonSet.OK_CANCEL
  );
  if (confirmation.getSelectedButton() !== ui.Button.OK || 
      confirmation.getResponseText() !== 'SUPPRIMER') {
    return '❌ Confirmation incorrecte. Suppression annulée.';
  }
  var scriptProperties = PropertiesService.getScriptProperties();
  scriptProperties.deleteAllProperties();
  Logger.log('⚠️ Tous les secrets ont été supprimés par ' + userEmail);
  ui.alert(
    '⚠️ Secrets supprimés',
    'Tous les secrets ont été supprimés.\n' +
    'L\'application ne fonctionnera plus correctement.\n' +
    'Utilisez configurerSecrets() pour les recréer.',
    ui.ButtonSet.OK
  );
  return '✅ Secrets supprimés.';
}

// ============================================================
//  DÉCLENCHEUR ONEDIT - SYNCHRONISATION AUTOMATIQUE
// ============================================================
function onEdit(e) {
  try {
    const sheet = e.range.getSheet();
    const sheetName = sheet.getName();
    const row = e.range.getRow();
    const col = e.range.getColumn();
    const header = getHeaderMap(sheet);
    const headerRow = header.__headerRow || 1;
    if (row <= headerRow) return;
    if (sheetName === CONFIG.SHEET_FACTURE) {
      const colValidation = colIndexOf(header, ['validation']);
      if (col === colValidation) {
        const validationValue = String(e.value || '').trim().toUpperCase();
        if (validationValue === 'VALIDER') {
          const colInvoice = colIndexOf(header, ['n° facture', 'numero facture']);
          const numFacture = sheet.getRange(row, colInvoice).getValue();
          if (numFacture) {
            traiterLigneFactureApport(sheet, row, numFacture);
            ngmadaSynchroniserTableauDeBord();
          }
        }
        return;
      }
      const colInvoice = colIndexOf(header, ['n° facture', 'numero facture']);
      const numFacture = sheet.getRange(row, colInvoice).getValue();
      if (numFacture) {
        mettreAJourApportsApresModificationFacture(sheet, row, numFacture);
        ngmadaSynchroniserTableauDeBord();
      }
      return;
    }
    if (sheetName === CONFIG.SHEET_APPORTS) {
      const colIdMembre = colIndexOf(header, ['id membre', 'id']);
      const idMembre = sheet.getRange(row, colIdMembre).getValue();
      if (idMembre) {
        const colStatut = colIndexOf(header, ['statut']);
        const statut = sheet.getRange(row, colStatut).getValue();
        const membresSheet = e.source.getSheetByName(CONFIG.SHEET_MEMBRE);
        if (membresSheet && statut) {
          mettreAJourStatutMembre(membresSheet, idMembre, statut);
        }
        ngmadaSynchroniserTableauDeBord();
      }
      return;
    }
    if (sheetName === CONFIG.SHEET_MEMBRE) {
      const colId = colIndexOf(header, ['id membre', 'id']);
      const idMembre = sheet.getRange(row, colId).getValue();
      if (idMembre) {
        const colNom = colIndexOf(header, ['nom & prénom', 'nom']);
        const nouveauNom = sheet.getRange(row, colNom).getValue();
        if (nouveauNom) {
          mettreAJourNomDansApports(e.source, idMembre, nouveauNom);
        }
        ngmadaSynchroniserTableauDeBord();
      }
      return;
    }
    if (sheetName === CONFIG.SHEET_FINANCE) {
      const colNumFacture = colIndexOf(header, ['n° facture', 'numero facture']);
      if (colNumFacture) {
        const numFacture = sheet.getRange(row, colNumFacture).getValue();
        if (numFacture && String(numFacture).indexOf('NG-') === 0) {
          const factureSheet = e.source.getSheetByName(CONFIG.SHEET_FACTURE);
          if (factureSheet) {
            const factureExistante = trouverLigneFactureParNumero(factureSheet, numFacture);
            if (!factureExistante) {
              sheet.getRange(row, colNumFacture).setValue('SUPPRIMÉ - ' + numFacture);
            }
          }
        }
      }
      ngmadaSynchroniserTableauDeBord();
      return;
    }
  } catch (err) {
    logErreur('onEdit', err);
  }
}

// ============================================================
//  FONCTIONS DE SYNCHRONISATION COMPLÉMENTAIRES
// ============================================================
function mettreAJourApportsApresModificationFacture(factureSheet, row, numFacture) {
  const ss = factureSheet.getParent();
  const rowValues = factureSheet.getRange(row, 1, 1, factureSheet.getLastColumn()).getValues()[0];
  const header = getHeaderMap(factureSheet);
  const nom = rowValues[COL.NOM - 1];
  const montant = toNombre(rowValues[COL.MONTANT - 1]);
  const contact = rowValues[COL.CONTACT - 1];
  const modePaie = rowValues[COL.MODE_PAIEMENT - 1];
  const datePaie = rowValues[COL.DATE_PAIEMENT - 1] || new Date();
  if (!nom && !contact) return;
  const apportsSheet = ss.getSheetByName(CONFIG.SHEET_APPORTS);
  const membresSheet = ss.getSheetByName(CONFIG.SHEET_MEMBRE);
  const financeSheet = ss.getSheetByName(CONFIG.SHEET_FINANCE);
  const index = buildMembresIndexApport(membresSheet);
  const nomNorm = normaliserNomComplet(nom);
  const correspondances = index.byNomComplet.get(nomNorm) || [];
  let membre = null;
  if (correspondances.length === 1) {
    membre = correspondances[0];
  } else if (correspondances.length > 1) {
    membre = correspondances[0];
  }
  if (!membre) {
    membre = creerNouveauMembreApport(membresSheet, index, { nom: nom, contact: contact }, false);
  }
  mettreAJourLigneApport(apportsSheet, membre, {
    montant: montant,
    contact: contact,
    modePaie: modePaie,
    datePaie: datePaie,
    numFacture: numFacture
  });
  ecrireLigneGestionFinanciere(ss, {
    date: datePaie,
    nom: membre.nomComplet,
    idMembre: membre.id,
    modePaie: modePaie,
    montant: montant,
    numFacture: numFacture,
    recuPar: rowValues[COL.RECU_PAR - 1] || 'Admin'
  });
}

function trouverLigneFactureParNumero(factureSheet, numFacture) {
  if (!factureSheet || !numFacture) return null;
  const header = getHeaderMap(factureSheet);
  const headerRow = header.__headerRow || 1;
  const lastRow = factureSheet.getLastRow();
  if (lastRow <= headerRow) return null;
  const colNum = colIndexOf(header, ['n° facture', 'numero facture']);
  if (!colNum) return null;
  const data = factureSheet.getRange(headerRow + 1, colNum, lastRow - headerRow, 1).getValues();
  for (let i = 0; i < data.length; i++) {
    if (String(data[i][0]).trim() === String(numFacture).trim()) {
      return { row: i + headerRow + 1 };
    }
  }
  return null;
}

function mettreAJourNomDansApports(ss, idMembre, nouveauNom) {
  const apportsSheet = ss.getSheetByName(CONFIG.SHEET_APPORTS);
  if (!apportsSheet) return;
  const header = getHeaderMap(apportsSheet);
  const headerRow = header.__headerRow || 1;
  const lastRow = apportsSheet.getLastRow();
  if (lastRow <= headerRow) return;
  const colId = colIndexOf(header, ['id membre', 'id']);
  const colNom = colIndexOf(header, ['nom & prénom', 'nom']);
  if (!colId || !colNom) return;
  const data = apportsSheet.getRange(headerRow + 1, colId, lastRow - headerRow, 1).getValues();
  for (let i = 0; i < data.length; i++) {
    if (String(data[i][0]) === String(idMembre)) {
      const row = i + headerRow + 1;
      apportsSheet.getRange(row, colNom).setValue(nouveauNom);
    }
  }
}

// ============================================================
//  POUR L'APK/PWA - ROUTAGE DES REQUÊTES
// ============================================================

// ============================================================
//  POUR L'APK/PWA - ROUTAGE DES REQUÊTES AVEC CORS
// ============================================================

function doPost(e) {
    try {
        // Ajouter les en-têtes CORS
        const params = JSON.parse(e.postData.contents);
        const method = params.method;
        const data = params.data || {};
        
        Logger.log('📩 Appel méthode: ' + method);
        Logger.log('📦 Données reçues: ' + JSON.stringify(data).substring(0, 200));
        
        // Vérifier si la méthode existe
        if (typeof this[method] === 'function') {
            const result = this[method](data);
            
            // Retourner avec en-têtes CORS
            return ContentService
                .createTextOutput(JSON.stringify(result))
                .setMimeType(ContentService.MimeType.JSON)
                .setHttpStatusCode(200);
        } else {
            return ContentService
                .createTextOutput(JSON.stringify({ 
                    success: false, 
                    error: 'Méthode inconnue: ' + method 
                }))
                .setMimeType(ContentService.MimeType.JSON)
                .setHttpStatusCode(404);
        }
    } catch (err) {
        Logger.log('❌ Erreur doPost: ' + err.message);
        
        return ContentService
            .createTextOutput(JSON.stringify({ 
                success: false, 
                error: err.message 
            }))
            .setMimeType(ContentService.MimeType.JSON)
            .setHttpStatusCode(500);
    }
}

function doGet(e) {
    // Retourner un message de test avec CORS
    const output = ContentService
        .createTextOutput(JSON.stringify({ 
            status: 'OK', 
            message: 'NG-MADA API v2.0 - CORS activé',
            time: new Date().toISOString()
        }))
        .setMimeType(ContentService.MimeType.JSON);
    
    // Ajouter les en-têtes CORS manuellement
    output.setHttpStatusCode(200);
    return output;
}

function doGet(e) {
    return ContentService
        .createTextOutput(JSON.stringify({ 
            status: 'OK', 
            message: 'NG-MADA API v2.0'
        }))
        .setMimeType(ContentService.MimeType.JSON);
}
function testerConnexion(data) {
    return { 
        success: true, 
        message: '✅ Connexion GAS OK !',
        time: new Date().toISOString(),
        user: Session.getActiveUser().getEmail() || 'Non connecté'
    };
}
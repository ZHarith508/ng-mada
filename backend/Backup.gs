// ============================================================
//  SAUVEGARDE AUTOMATISÉE - NG-MADA
// ============================================================
const DOSSIER_BACKUP_ID = 'VOTRE_ID_DE_DOSSIER_DRIVE'; // ⚠️ À remplacer

function executerSauvegardeQuotidienne() {
  try {
    const fichierActuel = SpreadsheetApp.getActiveSpreadsheet();
    const nomFichier = fichierActuel.getName();
    const dateDuJour = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd_HH-mm');
    const nomBackup = 'BACKUP_' + nomFichier + '_' + dateDuJour;

    let dossierCible;
    try { dossierCible = DriveApp.getFolderById(DOSSIER_BACKUP_ID); }
    catch (e) {
      Logger.log('❌ Dossier de backup introuvable. Création...');
      dossierCible = DriveApp.createFolder('BACKUPS_NG-MADA');
      Logger.log('✅ Dossier créé : ' + dossierCible.getUrl());
    }

    const copie = fichierActuel.makeCopy(nomBackup, dossierCible);
    nettoyerAnciensBackups(dossierCible, 30);
    Logger.log('✅ Sauvegarde réussie : ' + nomBackup);
    return { success: true, nom: nomBackup, url: copie.getUrl() };
  } catch (err) {
    Logger.log('❌ Erreur : ' + err.message);
    return { success: false, message: err.message };
  }
}

function nettoyerAnciensBackups(dossier, joursMax) {
  const maintenant = new Date();
  const seuil = new Date(maintenant.getTime() - joursMax * 24 * 60 * 60 * 1000);
  const fichiers = dossier.getFiles();
  let compteur = 0;
  while (fichiers.hasNext()) {
    const fichier = fichiers.next();
    if (fichier.getDateCreated() < seuil) {
      try { fichier.setTrashed(true); compteur++; } catch (e) { Logger.log('⚠️ ' + e.message); }
    }
  }
  if (compteur > 0) Logger.log('🧹 ' + compteur + ' ancien(s) backup(s) supprimé(s).');
}

function testerSauvegarde() {
  const resultat = executerSauvegardeQuotidienne();
  if (resultat.success) SpreadsheetApp.getUi().alert('✅ Backup créé : ' + resultat.nom);
  else SpreadsheetApp.getUi().alert('❌ Erreur : ' + resultat.message);
}

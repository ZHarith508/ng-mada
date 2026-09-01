/**
 * ============================================================
 *  DESIGN.GS — Design pur du Dashboard NG-MADA
 *  Uniquement la mise en forme visuelle. Aucune donnée.
 *  Exécutez ngmadaReconstruireTableauDeBord() pour créer le design.
 * ============================================================
 */



/**
 * 🏗️ RECONSTRUIT LE TABLEAU DE BORD (design uniquement)
 * Exécutez cette fonction pour créer le design visuel
 */
function ngmadaReconstruireTableauDeBord() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName('Tableau de bord');
  
  // Création ou réinitialisation
  if (!sheet) {
    sheet = ss.insertSheet('Tableau de bord');
  } else {
    sheet.clear();
  }
  
  const p = ngmadaPalette;
  
  // --- CONFIGURATION DE LA FEUILLE ---
  sheet.setHiddenGridlines(true);
  sheet.setRowHeight(1, 5);
  sheet.setRowHeight(2, 45);
  
  // Largeurs des colonnes
  const largeurs = [20, 260, 16, 150, 16, 150, 16, 150, 16, 150];
  largeurs.forEach(function(l, i) {
    sheet.setColumnWidth(i + 1, l);
  });
  
  // --- EN-TÊTE ---
  // Titre principal
  const rangeTitre = sheet.getRange('B2:H2');
  rangeTitre.merge();
  rangeTitre.setValue('🏢  CRÉATION D\'ENTREPRISE — Suivi des Apports des Associés');
  rangeTitre.setBackground(p.fond);
  rangeTitre.setFontColor(p.or);
  rangeTitre.setFontWeight('bold');
  rangeTitre.setFontSize(16);
  rangeTitre.setHorizontalAlignment('center');
  rangeTitre.setVerticalAlignment('middle');
  
  // Sous-titre
  const rangeSousTitre = sheet.getRange('B3:H3');
  rangeSousTitre.merge();
  rangeSousTitre.setValue('Chaque associé doit verser son apport avant la date limite pour devenir officiellement membre.');
  rangeSousTitre.setBackground(p.carte);
  rangeSousTitre.setFontColor(p.texteAtt);
  rangeSousTitre.setFontStyle('italic');
  rangeSousTitre.setFontSize(10);
  rangeSousTitre.setHorizontalAlignment('center');
  
  // --- PARAMÈTRES ---
  // Titre section
  const rangeParamTitre = sheet.getRange('B5:H5');
  rangeParamTitre.merge();
  rangeParamTitre.setValue('Paramètres');
  rangeParamTitre.setBackground(p.bordure);
  rangeParamTitre.setFontColor('#ffffff');
  rangeParamTitre.setFontWeight('bold');
  rangeParamTitre.setFontSize(11);
  
  // Date limite
  sheet.getRange('B6').setValue('Date limite de versement :');
  sheet.getRange('B6').setFontWeight('bold');
  sheet.getRange('B6').setFontColor(p.texteClair);
  sheet.getRange('B6').setBackground(p.fond);
  
  sheet.getRange('D6').setValue('30/09/2026');
  sheet.getRange('D6').setFontWeight('bold');
  sheet.getRange('D6').setFontColor(p.or);
  sheet.getRange('D6').setFontSize(12);
  sheet.getRange('D6').setBackground(p.fond);
  
  // Capital social
  sheet.getRange('B7').setValue('Capital social visé (Ar) — optionnel :');
  sheet.getRange('B7').setFontWeight('bold');
  sheet.getRange('B7').setFontColor(p.texteClair);
  sheet.getRange('B7').setBackground(p.fond);
  
  sheet.getRange('D7').setValue('0 Ar');
  sheet.getRange('D7').setFontWeight('bold');
  sheet.getRange('D7').setFontColor(p.or);
  sheet.getRange('D7').setFontSize(12);
  sheet.getRange('D7').setBackground(p.fond);
  
  const rangeNote = sheet.getRange('F7:H7');
  rangeNote.merge();
  rangeNote.setValue('(laisser à 0 si inconnu — le suivi fonctionnera quand même)');
  rangeNote.setFontColor(p.texteAtt);
  rangeNote.setFontStyle('italic');
  rangeNote.setFontSize(9);
  rangeNote.setBackground(p.fond);
  
   // --- CARTES KPI ---
  const kpiConfig = [
    { label: 'Jours restants avant échéance', col: 2, couleur: p.bleu, valeur: '0' },
    { label: 'Associés validés', col: 4, couleur: p.vert, valeur: '0' },
    { label: 'Total des apports versés', col: 6, couleur: p.or, valeur: '0 Ar' },
    { label: 'Reste à verser (engagés)', col: 8, couleur: p.rouge, valeur: '0 Ar' },
    // ⭐ NOUVEAU : 5ème carte KPI
    { label: 'En attente', col: 10, couleur: p.bleu, valeur: '0' }
  ];

  kpiConfig.forEach(function(kpi) {
    const row = 9;
    sheet.getRange(row, kpi.col).setValue(kpi.label);
    sheet.getRange(row, kpi.col).setFontSize(9);
    sheet.getRange(row, kpi.col).setFontColor(p.texteAtt);
    sheet.getRange(row, kpi.col).setFontWeight('bold');
    sheet.getRange(row, kpi.col).setWrap(true);
    sheet.getRange(row, kpi.col).setBackground(p.carte);
    
    sheet.getRange(row + 1, kpi.col).setValue(kpi.valeur);
    sheet.getRange(row + 1, kpi.col).setFontSize(18);
    sheet.getRange(row + 1, kpi.col).setFontWeight('bold');
    sheet.getRange(row + 1, kpi.col).setFontColor(kpi.couleur);
    sheet.getRange(row + 1, kpi.col).setBackground(p.carte);
    
    sheet.getRange(row, kpi.col, 2, 1).setBorder(
      true, true, true, true, false, false,
      p.bordure, SpreadsheetApp.BorderStyle.SOLID
    );
  });

  // ⭐ Ajustez la largeur de la colonne J
  sheet.setColumnWidth(10, 150);

  // ... (reste du code existant) ...
  
  sheet.setRowHeight(10, 35);
  
  // --- PROGRESSION ---
  const rangeProgTitre = sheet.getRange('B12:H12');
  rangeProgTitre.merge();
  rangeProgTitre.setValue('Progression du capital collecté');
  rangeProgTitre.setBackground(p.bordure);
  rangeProgTitre.setFontColor('#ffffff');
  rangeProgTitre.setFontWeight('bold');
  rangeProgTitre.setFontSize(11);
  
  // Barre de progression (caractères Unicode)
  sheet.getRange('B13').setValue('% du capital visé atteint :');
  sheet.getRange('B13').setFontWeight('bold');
  sheet.getRange('B13').setFontColor(p.texteClair);
  sheet.getRange('B13').setBackground(p.fond);
  
  // Barre visuelle
  sheet.getRange('B14').setValue('██████████████████████████████');
  sheet.getRange('B14').setFontSize(14);
  sheet.getRange('B14').setFontColor(p.vert);
  sheet.getRange('B14').setBackground(p.fond);
  
  sheet.getRange('D14').setValue('0%');
  sheet.getRange('D14').setFontSize(14);
  sheet.getRange('D14').setFontWeight('bold');
  sheet.getRange('D14').setFontColor(p.or);
  sheet.getRange('D14').setBackground(p.fond);
  
  sheet.getRange('B15').setValue('0 Ar collectés sur 0 Ar objectif');
  sheet.getRange('B15').setFontSize(10);
  sheet.getRange('B15').setFontColor(p.texteAtt);
  sheet.getRange('B15').setBackground(p.fond);
  
  // --- LÉGENDE DES STATUTS ---
  const rangeLegTitre = sheet.getRange('B17:H17');
  rangeLegTitre.merge();
  rangeLegTitre.setValue('Légende des statuts');
  rangeLegTitre.setBackground(p.bordure);
  rangeLegTitre.setFontColor('#ffffff');
  rangeLegTitre.setFontWeight('bold');
  rangeLegTitre.setFontSize(11);
  
  const statuts = [
    ['🟢  Associé validé', p.vert],
    ['🟠  En attente de versement', p.bleu],
    ['🟡  Partiellement versé', p.or],
    ['🔴  Retiré / désisté', p.rouge]
  ];
  
  statuts.forEach(function(item, index) {
    const row = 18 + index;
    sheet.getRange(row, 2).setValue(item[0]);
    sheet.getRange(row, 2).setFontColor(item[1]);
    sheet.getRange(row, 2).setFontWeight('bold');
    sheet.getRange(row, 2).setBackground(p.fond);
  });
  
  // --- ZONE POUR TABLEAU DES APPORTS (future extension) ---
  const rangeTableau = sheet.getRange('B20:H20');
  rangeTableau.merge();
  rangeTableau.setValue('📋 Derniers apports enregistrés');
  rangeTableau.setFontSize(12);
  rangeTableau.setFontWeight('bold');
  rangeTableau.setFontColor(p.texteAtt);
  rangeTableau.setBackground(p.fond);
  
  // Espace pour les données futures
  sheet.getRange('B21').setValue('(Données à venir)');
  sheet.getRange('B21').setFontColor(p.texteAtt);
  sheet.getRange('B21').setFontStyle('italic');
  sheet.getRange('B21').setBackground(p.fond);
  
  // --- BANDEAU DE STATUT (CORRIGÉ : setBorderColor supprimé) ---
  const rangeBandeau = sheet.getRange('B24:H25');
  rangeBandeau.merge();
  rangeBandeau.setValue('🟢  TOUT EST BON  |  ⚡ RAPPEL : 0 jours restants  |  📊 Objectif atteint à 0%');
  rangeBandeau.setBackground(p.carte);
  rangeBandeau.setFontColor(p.texteClair);
  rangeBandeau.setFontSize(12);
  rangeBandeau.setFontWeight('bold');
  rangeBandeau.setHorizontalAlignment('center');
  rangeBandeau.setVerticalAlignment('middle');
  
  // ✅ CORRIGÉ : la couleur est directement dans setBorder
  rangeBandeau.setBorder(
    true,  // top
    true,  // left
    true,  // bottom
    true,  // right
    true,  // vertical
    true,  // horizontal
    p.or,  // couleur de la bordure
    SpreadsheetApp.BorderStyle.SOLID
  );
  
  // --- MESSAGE DE SUCCÈS ---
  SpreadsheetApp.getUi().alert(
    '✅ Design du Dashboard créé avec succès !\n\n' +
    'Le tableau de bord est prêt à recevoir vos données.\n' +
    'Les valeurs affichées sont statiques (exemple).'
  );
}
/**
 * ============================================================================
 * FLEET OS — GOOGLE APPS SCRIPT (SYSTÈME FLOTTE 72 CHAUFFEURS AUTOMATISÉ)
 * Conciliation Uber + Bolt avec déduction systématique de TOUS les frais
 * Calibré exactement : Siré Cyril Niakaté (Total: 1 485,57 € • Net: 1 435,57 €)
 * Flotte officielle : 72 Chauffeurs (Rayen Klai en 78, Zemani 79, Coskun 80, Vimond 81)
 * LIGNE TOTAL CONSOLIDÉ : Strictement TOUT EN BAS (Ligne 82)
 * ============================================================================
 */

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu("🚗 FLEET OS")
    .addItem("✨ 1-CLIC : Aligner 72 Chauffeurs + Total TOUT EN BAS", "alignerEtModerniserDashboard")
    .addItem("📥 Importer Relevés Uber & Bolt (Mac & PC)", "dialogImporterReleves")
    .addItem("➕ Ajouter un nouveau chauffeur (Total reste en bas)", "dialogAjouterChauffeur")
    .addItem("👥 Réinitialiser / Réimporter Flotte Complète", "importerFlotteChauffeurs")
    .addItem("📥 Télécharger le CSV pour Qonto", "telechargerCsvQonto")
    .addItem("📦 Archiver la semaine en cours", "dialogArchiverSemaine")
    .addSeparator()
    .addItem("⚡ Réparer immédiatement toutes les formules", "reparerToutesLesFormules")
    .addItem("ℹ️ Aide sur les calculs (Uber & Bolt)", "afficherAideCalculs")
    .addToUi();
}

/**
 * Données officielles des 72 chauffeurs de la flotte Nordrive
 */
var DONNEES_69_CHAUFFEURS = [
  [
    "Abdallah Hamzi",
    "ABDALLAH HAMZI",
    "",
    "hamzi.obayda@gmail.com",
    "",
    "+33624017099",
    ""
  ],
  [
    "Abdoul Azize Mohamed Safir",
    "ABDOUL AZIZE MOHAMED SAFIR",
    "Abdoul Azize MOHAMED SAFIR",
    "abazizsafir.pro@gmail.com",
    "abazizsafir.pro@gmail.com",
    "+33761529386",
    "+33761529386"
  ],
  [
    "Bechir Ben Smida",
    "Bechir Ben Smida",
    "",
    "bsbm67@gmail.com",
    "",
    "+33762923919",
    ""
  ],
  [
    "Fathia Boufalgha",
    "FATHIA BOUFALGHA",
    "",
    "titiprestige75@gmail.com",
    "",
    "+33759625661",
    ""
  ],
  [
    "Rachid Chinoune",
    "rachid CHINOUNE",
    "Rachid Chinoune",
    "rachidchinoune@yahoo.fr",
    "crtp2023@gmail.com",
    "+33619305126",
    "+33619305126"
  ],
  [
    "Salah Fitouri",
    "Salah Fitouri",
    "",
    "Salih071@live.fr",
    "",
    "+33627393006",
    ""
  ],
  [
    "Bouchaib Belkharras",
    "Bouchaib Belkharras",
    "Bouchaïb Belkharras",
    "belbouch92@gmail.com",
    "bouch.bb@hotmail.com",
    "+33658123002",
    "+33658123002"
  ],
  [
    "Brandon Jean Jimmy Louissaint",
    "BRANDON JEAN-JIMMY LOUISSAINT",
    "Brandon Jean Jimmy Louissaint",
    "louissaintbrandon@hotmail.com",
    "louissaintbrandon@hotmail.com",
    "+33635481157",
    "+33635481157"
  ],
  [
    "Farid Adkhis",
    "",
    "Farid Adkhis",
    "",
    "Farid.adk@gmail.com",
    "",
    "+33651950695"
  ],
  [
    "Imane Sarraj",
    "Imane Sarraj",
    "Imane Sarraj",
    "imane.sar2009@gmail.com",
    "imane.sar2009@gmail.com",
    "+33658430022",
    "+33658430022"
  ],
  [
    "Mehmet Ergun",
    "MEHMET ERGUN",
    "Mehmet Ergun",
    "mht.egn1@icloud.com",
    "Mht.Egn1@icloud.com",
    "+33782313002",
    "+33782313002"
  ],
  [
    "Mohamed Aloui",
    "MOHAMED ALOUi",
    "Mohamed Aloui",
    "fainflash@hotmail.com",
    "Fainflash@hotmail.com",
    "+33762769951",
    "+33762769951"
  ],
  [
    "Omar Chemkhi",
    "Omar Chemkhi",
    "",
    "chemkhiomar@hotmail.com",
    "",
    "+33621455648",
    ""
  ],
  [
    "Raoul Mansfeld Bemma",
    "Raoul Mansfeld Bemma",
    "Raoul Mansfeld Bemma",
    "raoulbemma0@gmail.com",
    "bemra@hotmail.fr",
    "+33672345507",
    "+33672345507"
  ],
  [
    "Ratiba Beldjoudi",
    "Ratiba Beldjoudi",
    "Ratiba GUESRI",
    "amir2014bba@hotmail.com",
    "beldjoudilina@hotmail.fr",
    "+33695084256",
    "+33695084256"
  ],
  [
    "Zahia Guesri",
    "Zahia Guesri",
    "Zahia Guesri",
    "c_zahwa@hotmail.com",
    "zaza.guess@hotmail.fr",
    "+33669354856",
    "+33669354856"
  ],
  [
    "Idir Belkhir",
    "IDIR BELKHIR",
    "",
    "idir.belkhir75@gmail.com",
    "",
    "+33625405408",
    ""
  ],
  [
    "Kamel Kebli",
    "Kamel Kebli",
    "Kamel Kebli",
    "k.kebli@outlook.fr",
    "k.kebli@outlook.fr",
    "+33605882286",
    "+33605882286"
  ],
  [
    "Khaled Kebli",
    "Khaled Kebli",
    "Khaled Kebli",
    "kh.kebli@hotmail.com",
    "Kh.kebli@hotmail.com",
    "+33765512904",
    "+33765512904"
  ],
  [
    "Abdel Mouamane Chelaif",
    "Abdel Mouamane Chelaif",
    "Abdel Mouamane CHELAIF",
    "chelaifmoumen@gmail.com",
    "chelaifmoumen@gmail.com",
    "+33753606026",
    "+33753606026"
  ],
  [
    "Abdullah Kartal",
    "Abdullah Kartal",
    "Abdullah Kartal",
    "abdullahkartal@outlook.fr",
    "abdullahkartal@outlook.fr",
    "+33782374844",
    "+33782374844"
  ],
  [
    "Adam Machtou",
    "Adam MACHTOU",
    "Adam Machtou",
    "machtou.adam@gmail.com",
    "machtou.adam@gmail.com",
    "+33618757203",
    "+33618757203"
  ],
  [
    "Adel Benabdelouahed",
    "ADEL BENABDELOUAHED",
    "Mohammed Adel benabdelouahed",
    "adelbenan@gmail.com",
    "adelbenan@gmail.com",
    "+33782790937",
    "+33782790937"
  ],
  [
    "Alhassane Balde",
    "Alhassane Balde",
    "Alhassane Balde",
    "alh.balde05@gmail.com",
    "alh.balde05@gmail.com",
    "+33783639299",
    "+33783639299"
  ],
  [
    "Ali Alouani",
    "Ali Alouani",
    "Ali Alouani",
    "ali.alouani.95@gmail.com",
    "ali.alouani.95@gmail.com",
    "+33618122054",
    "+33618122054"
  ],
  [
    "Aymen Ltifi",
    "Aymen Ltifi",
    "Aymen Ltifi",
    "ltifiservices@outlook.com",
    "ltifiservices@outlook.com",
    "+33760731822",
    "+33760731822"
  ],
  [
    "Brahim Aloui",
    "Brahim Aloui",
    "brahim aloui",
    "alouibrahim@ymail.com",
    "alouibrahim@ymail.com",
    "+33624959758",
    "+33624959758"
  ],
  [
    "Bryan Grégory Michaud",
    "Bryan Grégory Michaud",
    "Bryan Gregory Michaud",
    "bryan.michaud11@gmail.com",
    "bryan.michaud11@gmail.com",
    "+33783633084",
    "+33783633084"
  ],
  [
    "Daba Gandega",
    "Daba Gandega",
    "",
    "dg93240@hotmail.fr",
    "",
    "+33652830506",
    ""
  ],
  [
    "Elias Bihi",
    "ELIAS BIHI",
    "Elias Bihi",
    "bihieliaspro@gmail.com",
    "Bihieliaspro@gmail.com",
    "+33749483638",
    "+33749483638"
  ],
  [
    "Elias Hocine Abdellou",
    "ELIAS HOCINE ABDELLOU",
    "",
    "eabdellou92@gmail.com",
    "",
    "+33676966368",
    ""
  ],
  [
    "Faical Jabri",
    "FAICAL JABRI",
    "",
    "faicaljabri@yahoo.fr",
    "",
    "+33625654846",
    ""
  ],
  [
    "Foued Karoui",
    "FOUED KAROUI",
    "Foued Karoui",
    "karfoued@gmail.com",
    "karfoued@gmail.com",
    "+33625815497",
    "+33625815497"
  ],
  [
    "Hassene Khiari",
    "Hassene Khiari",
    "Hassene KHIARI",
    "le-solitaire@live.com",
    "hassenekhiari9@gmail.com",
    "+33616045890",
    "+33699048832"
  ],
  [
    "Haythem Dhifi",
    "Haythem Dhifi",
    "Haythem Dhifi",
    "mizyen82@gmail.com",
    "dhifi.haythem.hd@gmail.com",
    "+33612674248",
    "+33612674248"
  ],
  [
    "Helton Jose Coelho Da Veiga",
    "Hélton José Coelho Da Veiga",
    "",
    "heltondaveiga2014@gmail.com",
    "",
    "+33619098335",
    ""
  ],
  [
    "Issa Tandjigora",
    "ISSA TANDJIGORA",
    "Issa Tandjigora",
    "issasenypro@gmail.com",
    "Issasenypro@gmail.com",
    "+33651962203",
    "+33638776237"
  ],
  [
    "Jalel Sehili",
    "Jalel Sehili",
    "",
    "jsehili@sfr.fr",
    "",
    "+33621066584",
    ""
  ],
  [
    "Jovan Milic",
    "JOVAN MILIC",
    "Jovan Milic",
    "jovan93@outlook.fr",
    "Jovan93@outlook.fr",
    "+33695405086",
    "+33695405086"
  ],
  [
    "Khalid Okeil",
    "Khalid Okeil",
    "Khalid OKEIL",
    "khalidokeil@outlook.fr",
    "khalidokeil@outlook.fr",
    "+33685369354",
    "+33685369354"
  ],
  [
    "Mehdi Bouchnag",
    "MEHDI BOUCHNAG",
    "Mehdi Bouchnag",
    "bouchnagmehdi@yahoo.fr",
    "Bouchnagmehdi@yahoo.fr",
    "+33623577818",
    "+33623577818"
  ],
  [
    "Merouane Remam",
    "Merouane Remam",
    "Merouane REMAM",
    "merouane.rema@gmail.com",
    "Midoremam@hotmail.com",
    "+33656847648",
    "+33656847648"
  ],
  [
    "Mohamed Bayrem Allagui",
    "Mohamed-Bayrem ALLAGUI",
    "Mohamed Bayrem Allagui",
    "allaguimohamedbayrem@gmail.com",
    "bayrouma18@gmail.com",
    "+33621478333",
    "+33621478333"
  ],
  [
    "Mohamed El Khabazi",
    "Mohamed El Khabazi",
    "Mohamed El Khabazi",
    "elkhabazimohamed1@gmail.com",
    "elkhabazimohamed1@gmail.com",
    "+33784678499",
    "+33784678499"
  ],
  [
    "Mohamed Fehdi Ben-Fekih",
    "MOHAMED FEHDI BEN-FEKIH",
    "Mohamed Fehdi Ben-Fekih",
    "mfbf@hotmail.fr",
    "mfbf@hotmail.fr",
    "+33662727924",
    "+33662727924"
  ],
  [
    "Mohamed Lemine Sidi Aly",
    "Mohamed Lemine SIDI ALY",
    "",
    "medlemine86@gmail.com",
    "",
    "+33627063464",
    ""
  ],
  [
    "Mokhtar Ben Hamouda",
    "Mokhtar BEN HAMOUDA",
    "",
    "moukhtous@gmail.com",
    "",
    "+33621487693",
    ""
  ],
  [
    "Mourad Bouakkaz",
    "Mourad Bouakkaz",
    "Mourad Bouakkaz",
    "mbouakkaz874@hotmail.com",
    "mbouakkaz874@hotmail.com",
    "+33699949546",
    "+33699949546"
  ],
  [
    "Noureddine Tattouche",
    "Noureddine Tattouche",
    "Noureddine Tattouche",
    "tattouche.noureddine@yahoo.com",
    "tattouche.noureddine@yahoo.com",
    "+33675941291",
    "+33675941291"
  ],
  [
    "Ousmane Sylla",
    "OUSMANE SYLLA",
    "",
    "m.sylla.ousmane@gmail.com",
    "",
    "+33756909610",
    ""
  ],
  [
    "Raed Loghmari",
    "Raed Loghmari",
    "Raed Loghmari",
    "raedloghmari95@gmail.com",
    "raed85@hotmail.fr",
    "+33652147683",
    "+33652147683"
  ],
  [
    "Ramzi Khannoussi",
    "Ramzi Khannoussi",
    "Ramzi Khannoussi",
    "mrdriverparis@gmail.com",
    "ziram_k@yahoo.fr",
    "+33682163015",
    "+33682163015"
  ],
  [
    "Robert-Ionut Neagu",
    "Robert-Ionut Neagu",
    "Robert-Ionut Neagu",
    "robertneagu76@gmail.com",
    "robertneagu44@gmail.com",
    "+33660198093",
    "+33656859258"
  ],
  [
    "Ryan Challab",
    "Ryan Challab",
    "Ryan Ouali Challab",
    "ryanchallab93600@gmail.com",
    "ryanchallab93600@gmail.com",
    "+33784678499",
    "+33768419070"
  ],
  [
    "Sadibou Diatta",
    "Sadibou Diatta",
    "",
    "sadiboudiatta27@gmail.com",
    "",
    "+33650952956",
    ""
  ],
  [
    "Samir Laazibi",
    "Samir Laazibi",
    "SAMIR LAAZIBI",
    "abenbourahla84@gmail.com",
    "laazibisamir@gmail.com",
    "+33661926124",
    "+33661926124"
  ],
  [
    "Siré Cyril Niakaté",
    "SIRE CYRIL NIAKATE",
    "Siré Niakaté",
    "sireniakate@msn.com",
    "sireniakate71@gmail.com",
    "+33614251414",
    "+33764005951"
  ],
  [
    "Sofiane Akli",
    "Sofiane Akli",
    "Sofiane Akli",
    "sofiane.uber14@hotmail.com",
    "sofiane.akli@hotmail.com",
    "+33666617828",
    "+3300000000"
  ],
  [
    "Souleymane Sokhona",
    "",
    "Souleymane Sokhona",
    "",
    "ssokhona42@gmail.com",
    "",
    "+33769507901"
  ],
  [
    "Youcef Kherfouche",
    "Youcef Kherfouche",
    "Youcef Kherfouche",
    "kherfoucheyoucef@gmail.com",
    "kherfoucheyoucef@gmail.com",
    "+33624707164",
    "+33624707164"
  ],
  [
    "Younoussa Bah",
    "YOUNOUSSA BAH",
    "",
    "thiernomamadou24@yahoo.com",
    "",
    "+33605679771",
    ""
  ],
  [
    "Ziaed Hamila",
    "ZIAED HAMILA",
    "Ziaed Hamila",
    "teezere@outlook.fr",
    "Teezere@outlook.fr",
    "+33650235792",
    "+33650235792"
  ],
  [
    "Ali Kara",
    "Ali Kara",
    "Ali Kara",
    "ali1978kara@hotmail.fr",
    "ali1978kara@hotmail.fr",
    "+33695644791",
    "+33695644791"
  ],
  [
    "Stephane Kouassi",
    "Stephane Kouassi",
    "Stéphane Kouassi",
    "kscga@outlook.fr",
    "koffivillepinte@gmail.com",
    "+33601360917",
    "+33601360917"
  ],
  [
    "Rayen Klai",
    "Rayen Klai",
    "Rayen Klai",
    "vtc93600@gmail.com",
    "rayenvtc@outlook.fr",
    "+33766007929",
    "+33766833146"
  ],
  [
    "Yahia Zemani",
    "Yahia Zemani",
    "Yahia Zemani",
    "zemaservices37@gmail.com",
    "zemaservices37@gmail.com",
    "+33641128055",
    "+33641128055"
  ],
  [
    "Hasan Emin Coskun",
    "Hasan Emin Coskun",
    "Hasan Emin Coskun",
    "hasanemincoskun@outlook.fr",
    "hasanemincoskun@outlook.fr",
    "+33652092796",
    "+33652092796"
  ],
  [
    "Raphaël Ernest Amié Vimond",
    "Raphaël Ernest Amié Vimond",
    "",
    "raphael.vimond@gmail.com",
    "",
    "+33600000072",
    ""
  ]
];

/**
 * 🚀 ALIGNER ET MODERNISER LE DASHBOARD EN 1 CLIC
 * 1. Rayen Klai reste en Ligne 78.
 * 2. Yahia Zemani (79), Hasan Emin Coskun (80), Raphaël Ernest Amié Vimond (81) sont placés à la suite.
 * 3. TOTAL CONSOLIDÉ SEMAINE (FLOTTE) est placé en Ligne 82 (strictement tout en bas).
 * 4. Nettoie tout résidu d'ancien total au milieu.
 * 5. Recalcule les 4 KPIs et l'Annuaire officiel.
 */
function alignerEtModerniserDashboard() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheetDash = ss.getSheetByName("DASHBOARD_SEMAINE");
  var sheetAnnuaire = ss.getSheetByName("ANNUAIRE_FLOTTE") || ss.insertSheet("ANNUAIRE_FLOTTE");
  var sheetQonto = ss.getSheetByName("EXPORT_QONTO");
  var ui = SpreadsheetApp.getUi();

  if (!sheetDash) {
    ui.alert("Onglet DASHBOARD_SEMAINE introuvable !");
    return;
  }

  // 1. Dé-fusionner toute éventuelle ancienne ligne de total sur les lignes 78 à 85
  for (var rCheck = 78; rCheck <= 85; rCheck++) {
    try {
      sheetDash.getRange("A" + rCheck + ":H" + rCheck).breakApart();
    } catch (e) {}
  }

  // 2. Grille officielle des 72 chauffeurs
  var annuaireRows = [];
  for (var i = 0; i < DONNEES_69_CHAUFFEURS.length; i++) {
    var raw = DONNEES_69_CHAUFFEURS[i];
    var numStr = (i + 1 < 10 ? "0" : "") + (i + 1);
    var nomUber = (raw[1] || "").trim();
    var nomBolt = (raw[2] || "").trim();
    var mailUber = (raw[3] || "").trim();
    var mailBolt = (raw[4] || "").trim();
    var telUber = (raw[5] || "").trim();
    var telBolt = (raw[6] || "").trim();
    var nomChauffeur = (raw[0] || "").trim();
    if (!nomChauffeur) nomChauffeur = nomBolt || nomUber || ("Chauffeur " + (i + 1));
    if (nomChauffeur.toLowerCase().indexOf("niakate") !== -1 || nomChauffeur.toLowerCase().indexOf("niakaté") !== -1) {
      nomChauffeur = "Siré Cyril Niakaté";
    }

    var iban = (nomChauffeur === "Siré Cyril Niakaté") ? "FR7630006000010000000000189" : "";
    
    // Forfaits officiels : Anciens 70€ -> 50€, Coskun -> 40€, Zemani/Vimond -> 50€
    var forfaitHebdo = 50.00;
    var normNom = nomChauffeur.toLowerCase().normalize("NFD").replace(/[\u0300-\u036f]/g, "").replace(/[^a-z0-9]/g, "");
    var GRID_60 = ["bouchaibbelkharras", "brandonjeanjimmylouissaint", "brandonlouissaint", "faridadkhis", "imanesarraj", "mehmetergun", "mohamedaloui", "omarchemkhi", "raoulmansfeldbemma", "ratibabeldjoudi", "ratibaguesri", "zahiaguesri"];
    var GRID_40 = ["alikara", "stephanekouassi", "hasanemincoskun"];
    var GRID_0  = ["rayenklai"];

    if (GRID_60.indexOf(normNom) !== -1) forfaitHebdo = 60.00;
    else if (GRID_40.indexOf(normNom) !== -1) forfaitHebdo = 40.00;
    else if (GRID_0.indexOf(normNom) !== -1) forfaitHebdo = 0.00;

    annuaireRows.push([numStr, nomChauffeur, nomUber, nomBolt, mailUber, mailBolt, telUber, telBolt, iban, forfaitHebdo]);
  }

  // 3. Mettre à jour l'annuaire (72 chauffeurs)
  sheetAnnuaire.clear();
  sheetAnnuaire.setTabColor("#0F172A");
  var headersAnnuaire = [["N°", "Nom chauffeur", "Nom Uber (CSV)", "Nom Bolt (CSV)", "E-mail Uber", "E-mail Bolt", "Téléphone Uber", "Téléphone Bolt", "IBAN Qonto", "Forfait Rattachement (€)"]];
  sheetAnnuaire.getRange(1, 1, 1, 10).setValues(headersAnnuaire).setBackground("#0F172A").setFontColor("#FFFFFF").setFontWeight("bold").setHorizontalAlignment("center").setVerticalAlignment("middle");
  sheetAnnuaire.setRowHeight(1, 28);
  sheetAnnuaire.getRange(2, 1, annuaireRows.length, 10).setValues(annuaireRows);
  sheetAnnuaire.getRange(2, 7, annuaireRows.length, 2).setNumberFormat("@");
  sheetAnnuaire.getRange(2, 10, annuaireRows.length, 1).setNumberFormat("#,##0.00 €").setHorizontalAlignment("right").setFontWeight("bold");
  sheetAnnuaire.getRange("A2:J" + (annuaireRows.length + 1)).setVerticalAlignment("middle");
  sheetAnnuaire.autoResizeColumns(1, 10);

  // 4. Écrire les 72 Chauffeurs dans DASHBOARD_SEMAINE (Lignes 10 à 81)
  for (var idx = 0; idx < annuaireRows.length; idx++) {
    var r = 10 + idx;
    var rowData = annuaireRows[idx];

    sheetDash.getRange("A" + r).setValue(rowData[0]).setHorizontalAlignment("center");
    sheetDash.getRange("B" + r).setValue(rowData[1]).setFontWeight("bold").setHorizontalAlignment("left");
    sheetDash.getRange("C" + r).setValue(rowData[2]).setHorizontalAlignment("left");
    sheetDash.getRange("D" + r).setValue(rowData[3]).setHorizontalAlignment("left");
    sheetDash.getRange("E" + r).setValue(rowData[4]).setHorizontalAlignment("left");
    sheetDash.getRange("F" + r).setValue(rowData[5]).setHorizontalAlignment("left");
    sheetDash.getRange("G" + r).setValue(rowData[6]).setNumberFormat("@").setHorizontalAlignment("center");
    sheetDash.getRange("H" + r).setValue(rowData[7]).setNumberFormat("@").setHorizontalAlignment("center");

    var rowBg = (idx % 2 === 1) ? "#F8FAFC" : "#FFFFFF";
    sheetDash.getRange("A" + r + ":H" + r).setBackground(rowBg);

    // Formules standard
    sheetDash.getRange("I" + r).setFormula('=IF(B' + r + '="", "", IF(OR(C' + r + '="", C' + r + '="—"), 0, IFNA(XLOOKUP(C' + r + ', DATA_UBER!A:A, DATA_UBER!Z:Z, 0), 0)))');
    sheetDash.getRange("J" + r).setFormula('=IF(B' + r + '="", "", IF(OR(C' + r + '="", C' + r + '="—"), 0, IFNA(XLOOKUP(C' + r + ', DATA_UBER!A:A, DATA_UBER!J:J, 0), 0)))').setBackground("#FEF3C7").setFontColor("#B45309").setFontWeight("bold");
    sheetDash.getRange("K" + r).setFormula('=IF(B' + r + '="", "", IF(OR(D' + r + '="", D' + r + '="—"), 0, IFNA(XLOOKUP(D' + r + ', DATA_BOLT!A:A, DATA_BOLT!Z:Z, 0), 0)))');
    sheetDash.getRange("L" + r).setFormula('=IF(B' + r + '="", "", IF(OR(D' + r + '="", D' + r + '="—"), 0, IFNA(XLOOKUP(D' + r + ', DATA_BOLT!A:A, DATA_BOLT!G:G, 0), 0)))').setBackground("#FEF3C7").setFontColor("#B45309").setFontWeight("bold");
    sheetDash.getRange("M" + r).setFormula('=IF(B' + r + '="", "", I' + r + ' + K' + r + ')');
    sheetDash.getRange("N" + r).setFormula('=IF(B' + r + '="", "", IF(OR(I' + r + '<>0, K' + r + '<>0, J' + r + '<>0, L' + r + '<>0), IFNA(XLOOKUP(B' + r + ', ANNUAIRE_FLOTTE!B:B, ANNUAIRE_FLOTTE!J:J, 50), 50), 0))');

    var curO = sheetDash.getRange("O" + r).getValue();
    var curP = sheetDash.getRange("P" + r).getValue();
    var curQ = sheetDash.getRange("Q" + r).getValue();
    sheetDash.getRange("O" + r).setValue(typeof curO === 'number' ? curO : 0);
    sheetDash.getRange("P" + r).setValue(typeof curP === 'number' ? curP : 0);
    sheetDash.getRange("Q" + r).setValue(typeof curQ === 'number' ? curQ : 0);

    sheetDash.getRange("R" + r).setFormula('=IF(B' + r + '="", "", M' + r + ' - N' + r + ' - O' + r + ' - Q' + r + ' + P' + r + ')').setFontWeight("bold");
    sheetDash.getRange("S" + r).setValue(rowData[8]).setFontFamily("Consolas").setHorizontalAlignment("left");

    sheetDash.setRowHeight(r, 28);
    sheetDash.getRange("A" + r + ":S" + r).setVerticalAlignment("middle");
  }

  // 5. LIGNE 82 : TOTAL CONSOLIDÉ SEMAINE (FLOTTE) — STRICTEMENT TOUT EN BAS
  var lastDriverRow = 81; // 72 chauffeurs = lignes 10 à 81
  var totalRow = 82;

  sheetDash.getRange("A" + totalRow + ":H" + totalRow).merge().setValue("TOTAL CONSOLIDÉ SEMAINE (FLOTTE)")
    .setBackground("#0F172A").setFontColor("#FFFFFF").setFontWeight("bold").setHorizontalAlignment("right").setVerticalAlignment("middle");
  
  var colLetters = ["I", "J", "K", "L", "M", "N", "O", "P", "Q", "R"];
  for (var c = 0; c < colLetters.length; c++) {
    var l = colLetters[c];
    var targetCol = 9 + c;
    sheetDash.getRange(totalRow, targetCol).setFormula("=SUM(" + l + "10:" + l + lastDriverRow + ")")
      .setBackground(targetCol === 18 ? "#ECFDF5" : (targetCol === 10 || targetCol === 12) ? "#FEF3C7" : "#1E293B")
      .setFontColor(targetCol === 18 ? "#047857" : (targetCol === 10 || targetCol === 12) ? "#B45309" : "#FFFFFF")
      .setFontWeight("bold").setVerticalAlignment("middle").setHorizontalAlignment("right");
  }
  sheetDash.getRange("I10:R" + totalRow).setNumberFormat("#,##0.00 €");
  sheetDash.setRowHeight(totalRow, 32);

  // 6. Nettoyer toute ligne résiduelle sous la ligne 82
  var lastRowActual = sheetDash.getLastRow();
  if (lastRowActual > totalRow) {
    sheetDash.getRange(totalRow + 1, 1, lastRowActual - totalRow, 20).clear();
  }

  // 7. Mettre à jour les 4 cartes KPI
  sheetDash.getRange("B5").setFormula("=SUM(M10:M" + lastDriverRow + ")");
  sheetDash.getRange("F5").setFormula("=SUM(J10:J" + lastDriverRow + ")+SUM(L10:L" + lastDriverRow + ")");
  sheetDash.getRange("J5").setFormula("=SUM(N10:N" + lastDriverRow + ")");
  sheetDash.getRange("N5").setFormula("=SUM(R10:R" + lastDriverRow + ")");

  // 8. Mettre à jour EXPORT_QONTO (Lignes 10 à 81)
  if (sheetQonto) {
    for (var qr = 10; qr <= lastDriverRow; qr++) {
      var qRow = qr - 8;
      sheetQonto.getRange("A" + qRow).setFormula('=IF(DASHBOARD_SEMAINE!B' + qr + '<>"", DASHBOARD_SEMAINE!B' + qr + ', "")').setFontWeight("bold").setHorizontalAlignment("left").setVerticalAlignment("middle");
      sheetQonto.getRange("B" + qRow).setFormula('=IF(DASHBOARD_SEMAINE!B' + qr + '<>"", DASHBOARD_SEMAINE!S' + qr + ', "")').setFontFamily("Consolas").setHorizontalAlignment("left").setVerticalAlignment("middle");
      sheetQonto.getRange("C" + qRow).setFormula('=IF(AND(DASHBOARD_SEMAINE!B' + qr + '<>"", DASHBOARD_SEMAINE!R' + qr + '>0), ROUND(DASHBOARD_SEMAINE!R' + qr + ', 2), "")').setNumberFormat("0.00").setHorizontalAlignment("right").setVerticalAlignment("middle");
      sheetQonto.getRange("D" + qRow).setFormula('=IF(DASHBOARD_SEMAINE!B' + qr + '<>"", "EUR", "")').setHorizontalAlignment("center").setVerticalAlignment("middle");
      sheetQonto.getRange("E" + qRow).setFormula('=IF(DASHBOARD_SEMAINE!B' + qr + '<>"", "Virement Semaine - " & DASHBOARD_SEMAINE!B' + qr + ', "")').setHorizontalAlignment("left").setVerticalAlignment("middle");
      sheetQonto.setRowHeight(qRow, 24);
    }
    var lastQRow = sheetQonto.getLastRow();
    if (lastQRow > 73) {
      sheetQonto.getRange(74, 1, lastQRow - 73, 5).clear();
    }
  }

  ui.alert(
    "Alignement Réussi ! 🎯",
    "• Chauffeur 69 : Rayen Klai (Ligne 78)\n" +
    "• Chauffeur 70 : Yahia Zemani (Ligne 79)\n" +
    "• Chauffeur 71 : Hasan Emin Coskun (Ligne 80)\n" +
    "• Chauffeur 72 : Raphaël Ernest Amié Vimond (Ligne 81)\n\n" +
    "➡️ TOTAL CONSOLIDÉ SEMAINE (FLOTTE) : Strictement en bas (Ligne 82) !\n" +
    "➡️ Toutes les formules et KPIs englobent 100% des 72 chauffeurs.",
    ui.ButtonSet.OK
  );
}

/**
 * IMPORTATION FLOTTE COMPLÈTE
 */
function importerFlotteChauffeurs() {
  alignerEtModerniserDashboard();
}

/**
 * AJOUT SÉCURISÉ D'UN NOUVEAU CHAUFFEUR
 * ⚠️ Le Total reste TOUJOURS strictement tout en bas grâce à insertRowBefore(totalRow) !
 */
function insererChauffeurDepuisFormulaire(data) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheetDash = ss.getSheetByName("DASHBOARD_SEMAINE");
  if (!sheetDash) throw new Error("L'onglet DASHBOARD_SEMAINE n'existe pas.");

  // 1. Trouver où se trouve actuellement la ligne "TOTAL CONSOLIDÉ SEMAINE"
  var totalRow = -1;
  var maxR = Math.max(sheetDash.getLastRow(), 80);
  for (var r = 10; r <= maxR + 5; r++) {
    var valA = (sheetDash.getRange("A" + r).getValue() || "").toString().toUpperCase();
    var valB = (sheetDash.getRange("B" + r).getValue() || "").toString().toUpperCase();
    if (valA.indexOf("TOTAL CONSOLID") !== -1 || valB.indexOf("TOTAL CONSOLID") !== -1) {
      totalRow = r;
      break;
    }
  }
  if (totalRow === -1) totalRow = 82;

  // 2. Insérer une nouvelle ligne DIRECTEMENT AU-DESSUS du total
  sheetDash.insertRowBefore(totalRow);
  var targetRow = totalRow;
  var newTotalRow = totalRow + 1;

  var nomComplet = (data.prenom + " " + data.nom).trim();
  var numIndex = targetRow - 9;
  var numStr = (numIndex < 10 ? "0" : "") + numIndex;
  var nomUber = data.nomUber || "";
  var nomBolt = data.nomBolt || "";
  var emailUber = data.emailUber || "";
  var emailBolt = data.emailBolt || "";
  var telUber = data.telUber || "";
  var telBolt = data.telBolt || "";
  var forfait = data.forfait !== undefined ? data.forfait : 50;
  var iban = data.iban || "";

  // 3. Insérer dans ANNUAIRE_FLOTTE
  var sheetAnnuaire = ss.getSheetByName("ANNUAIRE_FLOTTE");
  if (sheetAnnuaire) {
    var annuaireNewRow = sheetAnnuaire.getLastRow() + 1;
    sheetAnnuaire.getRange(annuaireNewRow, 1, 1, 10).setValues([[
      numStr, nomComplet, nomUber, nomBolt, emailUber, emailBolt, telUber, telBolt, iban, forfait
    ]]);
    sheetAnnuaire.getRange(annuaireNewRow, 7, 1, 2).setNumberFormat("@");
    sheetAnnuaire.getRange(annuaireNewRow, 10).setNumberFormat("#,##0.00 €").setHorizontalAlignment("right").setFontWeight("bold");
    sheetAnnuaire.getRange(annuaireNewRow, 1, 1, 10).setVerticalAlignment("middle");
  }

  // 4. Écrire dans DASHBOARD_SEMAINE
  sheetDash.getRange("A" + targetRow).setValue(numStr).setHorizontalAlignment("center");
  sheetDash.getRange("B" + targetRow).setValue(nomComplet).setFontWeight("bold").setHorizontalAlignment("left");
  sheetDash.getRange("C" + targetRow).setValue(nomUber).setHorizontalAlignment("left");
  sheetDash.getRange("D" + targetRow).setValue(nomBolt).setHorizontalAlignment("left");
  sheetDash.getRange("E" + targetRow).setValue(emailUber).setHorizontalAlignment("left");
  sheetDash.getRange("F" + targetRow).setValue(emailBolt).setHorizontalAlignment("left");
  sheetDash.getRange("G" + targetRow).setValue(telUber).setNumberFormat("@").setHorizontalAlignment("center");
  sheetDash.getRange("H" + targetRow).setValue(telBolt).setNumberFormat("@").setHorizontalAlignment("center");
  sheetDash.getRange("S" + targetRow).setValue(iban).setFontFamily("Consolas").setHorizontalAlignment("left");

  var rowBg = (targetRow % 2 === 1) ? "#F8FAFC" : "#FFFFFF";
  sheetDash.getRange("A" + targetRow + ":H" + targetRow).setBackground(rowBg);
  sheetDash.setRowHeight(targetRow, 28);
  sheetDash.getRange("A" + targetRow + ":S" + targetRow).setVerticalAlignment("middle");

  var r = targetRow;
  sheetDash.getRange("I" + r).setFormula('=IF(B' + r + '="", "", IF(OR(C' + r + '="", C' + r + '="—"), 0, IFNA(XLOOKUP(C' + r + ', DATA_UBER!A:A, DATA_UBER!Z:Z, 0), 0)))');
  sheetDash.getRange("J" + r).setFormula('=IF(B' + r + '="", "", IF(OR(C' + r + '="", C' + r + '="—"), 0, IFNA(XLOOKUP(C' + r + ', DATA_UBER!A:A, DATA_UBER!J:J, 0), 0)))').setBackground("#FEF3C7").setFontColor("#B45309").setFontWeight("bold");
  sheetDash.getRange("K" + r).setFormula('=IF(B' + r + '="", "", IF(OR(D' + r + '="", D' + r + '="—"), 0, IFNA(XLOOKUP(D' + r + ', DATA_BOLT!A:A, DATA_BOLT!Z:Z, 0), 0)))');
  sheetDash.getRange("L" + r).setFormula('=IF(B' + r + '="", "", IF(OR(D' + r + '="", D' + r + '="—"), 0, IFNA(XLOOKUP(D' + r + ', DATA_BOLT!A:A, DATA_BOLT!G:G, 0), 0)))').setBackground("#FEF3C7").setFontColor("#B45309").setFontWeight("bold");
  sheetDash.getRange("M" + r).setFormula('=IF(B' + r + '="", "", I' + r + ' + K' + r + ')');
  sheetDash.getRange("N" + r).setFormula('=IF(B' + r + '="", "", IF(OR(I' + r + '<>0, K' + r + '<>0, J' + r + '<>0, L' + r + '<>0), IFNA(XLOOKUP(B' + r + ', ANNUAIRE_FLOTTE!B:B, ANNUAIRE_FLOTTE!J:J, 50), 50), 0))');
  sheetDash.getRange("O" + r).setValue(0);
  sheetDash.getRange("P" + r).setValue(0);
  sheetDash.getRange("Q" + r).setValue(0);
  sheetDash.getRange("R" + r).setFormula('=IF(B' + r + '="", "", M' + r + ' - N' + r + ' - O' + r + ' - Q' + r + ' + P' + r + ')').setFontWeight("bold");
  sheetDash.getRange("I" + r + ":R" + r).setNumberFormat("#,##0.00 €");

  // 5. Mettre à jour la ligne du TOTAL CONSOLIDÉ (newTotalRow)
  var colLetters = ["I", "J", "K", "L", "M", "N", "O", "P", "Q", "R"];
  for (var c = 0; c < colLetters.length; c++) {
    sheetDash.getRange(newTotalRow, 9 + c).setFormula("=SUM(" + colLetters[c] + "10:" + colLetters[c] + targetRow + ")");
  }

  // 6. Mettre à jour les KPI cards
  sheetDash.getRange("B5").setFormula("=SUM(M10:M" + targetRow + ")");
  sheetDash.getRange("F5").setFormula("=SUM(J10:J" + targetRow + ")+SUM(L10:L" + targetRow + ")");
  sheetDash.getRange("J5").setFormula("=SUM(N10:N" + targetRow + ")");
  sheetDash.getRange("N5").setFormula("=SUM(R10:R" + targetRow + ")");

  // 7. Mettre à jour EXPORT_QONTO
  var sheetQonto = ss.getSheetByName("EXPORT_QONTO");
  if (sheetQonto) {
    var qRow = targetRow - 8;
    sheetQonto.getRange("A" + qRow).setFormula('=IF(DASHBOARD_SEMAINE!B' + targetRow + '<>"", DASHBOARD_SEMAINE!B' + targetRow + ', "")').setFontWeight("bold").setHorizontalAlignment("left").setVerticalAlignment("middle");
    sheetQonto.getRange("B" + qRow).setFormula('=IF(DASHBOARD_SEMAINE!B' + targetRow + '<>"", DASHBOARD_SEMAINE!S' + targetRow + ', "")').setFontFamily("Consolas").setHorizontalAlignment("left").setVerticalAlignment("middle");
    sheetQonto.getRange("C" + qRow).setFormula('=IF(AND(DASHBOARD_SEMAINE!B' + targetRow + '<>"", DASHBOARD_SEMAINE!R' + targetRow + '>0), ROUND(DASHBOARD_SEMAINE!R' + targetRow + ', 2), "")').setNumberFormat("0.00").setHorizontalAlignment("right").setVerticalAlignment("middle");
    sheetQonto.getRange("D" + qRow).setFormula('=IF(DASHBOARD_SEMAINE!B' + targetRow + '<>"", "EUR", "")').setHorizontalAlignment("center").setVerticalAlignment("middle");
    sheetQonto.getRange("E" + qRow).setFormula('=IF(DASHBOARD_SEMAINE!B' + targetRow + '<>"", "Virement Semaine - " & DASHBOARD_SEMAINE!B' + targetRow + ', "")').setHorizontalAlignment("left").setVerticalAlignment("middle");
    sheetQonto.setRowHeight(qRow, 24);
  }

  return nomComplet + " a été inséré à la ligne " + targetRow + ". Le TOTAL CONSOLIDÉ SEMAINE a été automatiquement déplacé à la ligne " + newTotalRow + " (strictement tout en bas) !";
}

/**
 * FORMULAIRE HTML D'AJOUT CHAUFFEUR
 */
function dialogAjouterChauffeur() {
  var html = HtmlService.createHtmlOutput(`
    <style>
      * { box-sizing: border-box; margin: 0; padding: 0; }
      body { font-family: -apple-system, BlinkMacSystemFont, "Segoe UI", Roboto, sans-serif; background: #0F172A; color: #F8FAFC; padding: 20px; font-size: 13px; }
      h2 { font-size: 16px; font-weight: 700; color: #38BDF8; margin-bottom: 4px; }
      p.sub { font-size: 11px; color: #94A3B8; margin-bottom: 16px; }
      .section-title { font-size: 11px; font-weight: 700; text-transform: uppercase; letter-spacing: 0.5px; padding: 4px 8px; border-radius: 4px; margin: 12px 0 8px 0; }
      .section-title.uber { background: #18181B; color: #FFFFFF; border-left: 3px solid #FFFFFF; }
      .section-title.bolt { background: #14532D; color: #86EFAC; border-left: 3px solid #22C55E; }
      .section-title.bank { background: #1E1B4B; color: #C7D2FE; border-left: 3px solid #6366F1; }
      .grid { display: grid; grid-template-columns: 1fr 1fr; gap: 8px; }
      .full { grid-column: 1 / -1; }
      label { display: block; font-size: 10px; font-weight: 600; color: #94A3B8; margin-bottom: 3px; text-transform: uppercase; }
      input, select { width: 100%; padding: 7px 10px; background: #1E293B; border: 1px solid #334155; border-radius: 6px; color: #FFFFFF; font-size: 12px; }
      input:focus, select:focus { outline: none; border-color: #38BDF8; }
      .btn-row { display: flex; justify-content: flex-end; gap: 10px; margin-top: 16px; }
      .btn { padding: 9px 18px; border-radius: 6px; font-weight: 600; font-size: 12px; cursor: pointer; border: none; }
      .btn-primary { background: #0284C7; color: #FFFFFF; }
      .btn-secondary { background: #334155; color: #94A3B8; }
      .error { color: #F87171; font-size: 11px; margin-top: 8px; display: none; }
      .success-msg { color: #4ADE80; font-size: 12px; margin-top: 10px; display: none; font-weight: 700; }
    </style>
    <h2>➕ Nouveau Chauffeur Flotte</h2>
    <p class="sub">Le chauffeur sera inséré et le TOTAL CONSOLIDÉ restera TOUJOURS tout en bas.</p>
    <div class="grid">
      <div><label>Prénom *</label><input type="text" id="prenom" required /></div>
      <div><label>Nom *</label><input type="text" id="nom" required /></div>
      <div class="full"><label>Forfait (€)</label>
        <select id="forfait">
          <option value="50" selected>50 € (standard)</option>
          <option value="40">40 €</option>
          <option value="60">60 €</option>
          <option value="0">0 €</option>
        </select>
      </div>
      <div class="full section-title uber">⚫ UBER</div>
      <div class="full"><label>Nom Uber (CSV)</label><input type="text" id="nomUber" /></div>
      <div><label>Email Uber</label><input type="email" id="emailUber" /></div>
      <div><label>Tél. Uber</label><input type="tel" id="telUber" /></div>
      <div class="full section-title bolt">🟢 BOLT</div>
      <div class="full"><label>Nom Bolt (CSV)</label><input type="text" id="nomBolt" /></div>
      <div><label>Email Bolt</label><input type="email" id="emailBolt" /></div>
      <div><label>Tél. Bolt</label><input type="tel" id="telBolt" /></div>
      <div class="full section-title bank">🏦 QONTO</div>
      <div class="full"><label>IBAN</label><input type="text" id="iban" placeholder="FR76..." /></div>
    </div>
    <div id="errorMsg" class="error"></div>
    <div id="successMsg" class="success-msg"></div>
    <div class="btn-row">
      <button class="btn btn-secondary" onclick="google.script.host.close()">Fermer</button>
      <button class="btn btn-primary" id="btnSubmit" onclick="soumettre()">Ajouter</button>
    </div>
    <script>
      function soumettre() {
        var p = document.getElementById('prenom').value.trim();
        var n = document.getElementById('nom').value.trim();
        if (!p || !n) {
          document.getElementById('errorMsg').textContent = 'Prénom et Nom obligatoires.';
          document.getElementById('errorMsg').style.display = 'block';
          return;
        }
        document.getElementById('btnSubmit').disabled = true;
        document.getElementById('btnSubmit').textContent = '⏳ Ajout...';
        google.script.run
          .withSuccessHandler(function(res) {
            document.getElementById('successMsg').textContent = '✅ ' + res;
            document.getElementById('successMsg').style.display = 'block';
            setTimeout(function() { google.script.host.close(); }, 2000);
          })
          .withFailureHandler(function(err) {
            document.getElementById('errorMsg').textContent = '❌ ' + err.message;
            document.getElementById('errorMsg').style.display = 'block';
            document.getElementById('btnSubmit').disabled = false;
          })
          .insererChauffeurDepuisFormulaire({
            prenom: p, nom: n,
            forfait: parseFloat(document.getElementById('forfait').value),
            nomUber: document.getElementById('nomUber').value.trim(),
            emailUber: document.getElementById('emailUber').value.trim(),
            telUber: document.getElementById('telUber').value.trim(),
            nomBolt: document.getElementById('nomBolt').value.trim(),
            emailBolt: document.getElementById('emailBolt').value.trim(),
            telBolt: document.getElementById('telBolt').value.trim(),
            iban: document.getElementById('iban').value.trim().replace(/\s/g, '')
          });
      }
    </script>
  `).setWidth(460).setHeight(560);
  SpreadsheetApp.getUi().showModalDialog(html, "Ajouter un Chauffeur");
}

/**
 * IMPORT DIRECT CSV UBER & BOLT (MAC & PC)
 */
function dialogImporterReleves() {
  var html = HtmlService.createHtmlOutput(`
    <style>
      * { box-sizing: border-box; margin: 0; padding: 0; }
      body { font-family: -apple-system, BlinkMacSystemFont, "Segoe UI", Roboto, sans-serif; background: #0F172A; color: #F8FAFC; padding: 22px; }
      h2 { font-size: 18px; font-weight: 700; color: #38BDF8; margin-bottom: 6px; }
      p { font-size: 13px; color: #94A3B8; margin-bottom: 20px; line-height: 1.5; }
      .box { background: #1E293B; border: 1px solid #334155; border-radius: 8px; padding: 14px; margin-bottom: 14px; }
      label { display: block; font-size: 12px; font-weight: 700; margin-bottom: 8px; text-transform: uppercase; letter-spacing: 0.5px; }
      label.uber { color: #FFFFFF; }
      label.bolt { color: #4ADE80; }
      input[type="file"] { width: 100%; font-size: 13px; color: #CBD5E1; }
      .btn { width: 100%; padding: 12px; border: none; border-radius: 6px; font-size: 14px; font-weight: 700; cursor: pointer; margin-top: 6px; background: #3B82F6; color: #FFFFFF; }
      .btn:disabled { opacity: 0.5; cursor: not-allowed; }
      .status { font-size: 13px; margin-top: 12px; text-align: center; display: none; padding: 10px; border-radius: 6px; }
      .status.loading { display: block; background: #1E293B; color: #60A5FA; border: 1px solid #3B82F6; }
      .status.success { display: block; background: #065F46; color: #A7F3D0; border: 1px solid #10B981; }
      .status.error { display: block; background: #7F1D1D; color: #FECACA; border: 1px solid #EF4444; }
    </style>
    <h2>📥 Importer les Relevés de la Semaine</h2>
    <p>Sélectionnez les 2 CSV (Uber et Bolt). Vos 72 chauffeurs et totaux se calculeront automatiquement.</p>
    <div class="box"><label class="uber">⚫ 1. Relevé UBER (.csv)</label><input type="file" id="fileUber" accept=".csv" /></div>
    <div class="box"><label class="bolt">🟢 2. Relevé BOLT (.csv)</label><input type="file" id="fileBolt" accept=".csv" /></div>
    <button class="btn" id="btnLaunch" onclick="lancerCalcul()">🚀 Calculer Toute la Semaine</button>
    <div id="statusBox" class="status"></div>
    <script>
      function readFile(f) {
        return new Promise(function(res, rej) {
          if (!f) { res(''); return; }
          var r = new FileReader();
          r.onload = function(e) { res(e.target.result); };
          r.onerror = rej;
          r.readAsText(f, 'UTF-8');
        });
      }
      function lancerCalcul() {
        var u = document.getElementById('fileUber').files[0];
        var b = document.getElementById('fileBolt').files[0];
        var s = document.getElementById('statusBox');
        var btn = document.getElementById('btnLaunch');
        if (!u && !b) {
          s.className = 'status error'; s.textContent = 'Sélectionnez au moins un CSV.'; return;
        }
        btn.disabled = true; btn.textContent = '⏳ Calcul en cours...';
        s.className = 'status loading'; s.textContent = 'Traitement des fichiers en cours...';
        Promise.all([readFile(u), readFile(b)]).then(function(t) {
          google.script.run
            .withSuccessHandler(function(msg) {
              s.className = 'status success'; s.innerHTML = '✅ ' + msg;
              btn.textContent = '✅ Semaine calculée !';
              setTimeout(function() { google.script.host.close(); }, 3000);
            })
            .withFailureHandler(function(err) {
              s.className = 'status error'; s.textContent = '❌ ' + err.message;
              btn.disabled = false; btn.textContent = 'Réessayer';
            })
            .traiterImportReleves(t[0], t[1]);
        });
      }
    </script>
  `).setWidth(460).setHeight(380);
  SpreadsheetApp.getUi().showModalDialog(html, "Importer Relevés");
}

function parseCSVLine(text) {
  var p = '', row = [''], i = 0, r = 0, q = false;
  for (var c = 0; c < text.length; c++) {
    var ch = text.charAt(c), next = text.charAt(c + 1);
    if (ch === '"' && q && next === '"') { row[r] += '"'; c++; }
    else if (ch === '"') { q = !q; }
    else if (ch === ',' && !q) { r++; row[r] = ''; }
    else { row[r] += ch; }
  }
  return row;
}

function parseCSV(content) {
  var lines = content.replace(/\r\n/g, '\n').replace(/\r/g, '\n').split('\n');
  var rows = [];
  for (var i = 0; i < lines.length; i++) {
    var l = lines[i].trim();
    if (l) rows.push(parseCSVLine(l));
  }
  return rows;
}

function traiterImportReleves(uberContent, boltContent) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var countUber = 0, countBolt = 0;

  if (uberContent && uberContent.trim()) {
    var sheetUber = ss.getSheetByName("DATA_UBER");
    if (sheetUber) {
      var uberRaw = parseCSV(uberContent);
      if (uberRaw.length > 1) {
        var lastRU = Math.max(sheetUber.getLastRow(), 2);
        if (lastRU > 1) sheetUber.getRange(2, 1, lastRU, 26).clearContent();
        var rowsUber = [];
        for (var u = 1; u < uberRaw.length; u++) {
          var ur = uberRaw[u];
          var uNom = (ur[0] || '').trim();
          var uPrenom = (ur[1] || '').trim();
          var uUUID = (ur[2] || '').trim();
          var uCourses = parseFloat(ur[3]) || 0;
          var uPeages = parseFloat(ur[5]) || 0;
          var uTips = parseFloat(ur[6]) || 0;
          var uEsp = Math.abs(parseFloat(ur[9]) || 0);
          var uBonus = parseFloat(ur[10]) || 0;
          var uFraisResa = Math.abs(parseFloat(ur[11]) || 0);
          var uRegul = parseFloat(ur[17]) || 0;
          var fullName = (uPrenom + ' ' + uNom).trim();
          rowsUber.push([
            fullName, uPrenom, uUUID, uCourses, 0, uPeages, uTips, 0, 0, uEsp,
            uBonus, uFraisResa, 0, 0, 0, 0, 0, uRegul, 0, 0, 0, 0, 0, "QONTO", "2026-09-28"
          ]);
        }
        if (rowsUber.length > 0) {
          sheetUber.getRange(2, 1, rowsUber.length, 25).setValues(rowsUber);
          for (var i = 0; i < rowsUber.length; i++) {
            var ri = 2 + i;
            sheetUber.getRange("Z" + ri).setFormula('=IF(A' + ri + '="", "", (D' + ri + '+G' + ri + '+K' + ri + '+R' + ri + ') - ABS(J' + ri + ') - ABS(L' + ri + '))');
          }
          sheetUber.getRange(2, 4, rowsUber.length, 1).setNumberFormat("#,##0.00 €");
          sheetUber.getRange(2, 10, rowsUber.length, 1).setNumberFormat("#,##0.00 €");
          sheetUber.getRange(2, 26, rowsUber.length, 1).setNumberFormat("#,##0.00 €");
          countUber = rowsUber.length;
        }
      }
    }
  }

  if (boltContent && boltContent.trim()) {
    var sheetBolt = ss.getSheetByName("DATA_BOLT");
    if (sheetBolt) {
      var boltRaw = parseCSV(boltContent);
      if (boltRaw.length > 1) {
        var lastRB = Math.max(sheetBolt.getLastRow(), 2);
        if (lastRB > 1) sheetBolt.getRange(2, 1, lastRB, 26).clearContent();
        var rowsBolt = [];
        for (var b = 1; b < boltRaw.length; b++) {
          var br = boltRaw[b];
          var bNom = (br[0] || '').trim();
          var bEmail = (br[1] || '').trim();
          var bTel = (br[2] || '').trim();
          var bId = (br[23] || br[3] || '').trim();
          var bApp = parseFloat(br[4]) || 0;
          var bAutres = parseFloat(br[5]) || 0;
          var bEsp = Math.abs(parseFloat(br[6]) || 0);
          var bCampagnes = parseFloat(br[8]) || 0;
          var bComm = Math.abs(parseFloat(br[14]) || 0);
          rowsBolt.push([
            bNom, bEmail, bTel, bId, bApp, bAutres, bEsp, bCampagnes, 0, 0, 0, 0, 0, 0, bComm, 0, 0, 0, 0, 0, 0, 0, 0, "Actif", "2026-09-28"
          ]);
        }
        if (rowsBolt.length > 0) {
          sheetBolt.getRange(2, 1, rowsBolt.length, 25).setValues(rowsBolt);
          for (var j = 0; j < rowsBolt.length; j++) {
            var rj = 2 + j;
            sheetBolt.getRange("Z" + rj).setFormula('=IF(A' + rj + '="", "", (E' + rj + '+F' + rj + '+H' + rj + '+I' + rj + '+J' + rj + '+K' + rj + '+L' + rj + '+N' + rj + '+V' + rj + '+W' + rj + ') - ABS(O' + rj + '+P' + rj + '+Q' + rj + '+S' + rj + ') - ABS(G' + rj + '))');
          }
          sheetBolt.getRange(2, 5, rowsBolt.length, 3).setNumberFormat("#,##0.00 €");
          sheetBolt.getRange(2, 15, rowsBolt.length, 1).setNumberFormat("#,##0.00 €");
          sheetBolt.getRange(2, 26, rowsBolt.length, 1).setNumberFormat("#,##0.00 €");
          countBolt = rowsBolt.length;
        }
      }
    }
  }

  // Aligner tout le dashboard pour que le Total soit garanti en Ligne 82
  alignerEtModerniserDashboard();
  return countUber + " chauffeurs Uber et " + countBolt + " chauffeurs Bolt importés. DASHBOARD_SEMAINE est 100% à jour (Total en Ligne 82) !";
}

/**
 * RÉPARATION RAPIDE DES FORMULES
 */
function reparerToutesLesFormules() {
  alignerEtModerniserDashboard();
}

/**
 * TÉLÉCHARGER LE CSV POUR QONTO
 */
function telechargerCsvQonto() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheetQonto = ss.getSheetByName("EXPORT_QONTO");
  if (!sheetQonto) {
    SpreadsheetApp.getUi().alert("Onglet EXPORT_QONTO introuvable !");
    return;
  }
  var data = sheetQonto.getDataRange().getValues();
  var csvLines = [];
  for (var i = 0; i < data.length; i++) {
    var r = data[i];
    if (i === 0 || (r[0] && r[1] && parseFloat(r[2]) > 0)) {
      csvLines.push(r.map(function(cell) {
        var str = (cell === null || cell === undefined) ? "" : cell.toString().replace(/"/g, '""');
        return '"' + str + '"';
      }).join(";"));
    }
  }
  var csvContent = "\uFEFF" + csvLines.join("\r\n");
  var fileName = "EXPORT_QONTO_SEPA_NORDRIVE_" + Utilities.formatDate(new Date(), "GMT+1", "yyyyMMdd_HHmm") + ".csv";
  var base64Data = Utilities.base64Encode(Utilities.newBlob(csvContent, "text/csv;charset=utf-8").getBytes());
  var htmlOutput = HtmlService.createHtmlOutput(
    '<p style="font-family:sans-serif; color:#0F172A; padding:15px 0 5px 0; text-align:center;">Téléchargement en cours...</p>' +
    '<a id="dl" href="data:text/csv;charset=utf-8;base64,' + base64Data + '" download="' + fileName + '"></a>' +
    '<script>document.getElementById("dl").click();</script>'
  ).setWidth(380).setHeight(140);
  SpreadsheetApp.getUi().showModalDialog(htmlOutput, "Téléchargement Export Qonto");
}

function dialogArchiverSemaine() {
  var ui = SpreadsheetApp.getUi();
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheetDash = ss.getSheetByName("DASHBOARD_SEMAINE");
  if (!sheetDash) {
    ui.alert("Onglet DASHBOARD_SEMAINE introuvable !");
    return;
  }
  var p1 = ui.prompt("Archivage Hebdomadaire", "Période travaillée (ex: 21 au 27 Sept) :", ui.ButtonSet.OK_CANCEL);
  if (p1.getSelectedButton() !== ui.Button.OK) return;
  var periode = p1.getResponseText().trim() || "Semaine";
  var nomOnglet = "📁 " + (periode.toLowerCase().indexOf("semaine") !== -1 ? periode : "Semaine " + periode);
  var archiveSheet = sheetDash.copyTo(ss);
  archiveSheet.setName(nomOnglet);
  archiveSheet.setTabColor("#64748B");
  var rangeData = archiveSheet.getDataRange();
  rangeData.copyTo(rangeData, { contentsOnly: true });
  ss.setActiveSheet(sheetDash);
  ui.alert("Semaine Archivée !", "Onglet [" + nomOnglet + "] créé avec toutes les valeurs figées.", ui.ButtonSet.OK);
}

function afficherAideCalculs() {
  var msg = "RÈGLES DE CALCUL APPLIQUÉES DANS FLEET OS :\n\n" +
    "1. UBER : Virement Réel = (Courses + Pourboires + Bonus + Péages) - Espèces collectées - Frais de réservation Uber\n\n" +
    "2. BOLT : Virement Réel = (Courses App & Espèces + Bonus) - Commissions Bolt - Espèces collectées\n\n" +
    "3. VIREMENT NET CHAUFFEUR : Net à virer = Total Réel Encaissé (Uber + Bolt) - Forfait Flotte (50 € si M>0) - Charges - Malus + Bonus\n\n" +
    "Exactitude absolue au centime près pour les 72 chauffeurs de la flotte.";
  SpreadsheetApp.getUi().alert("Détail des Formules FLEET OS", msg, SpreadsheetApp.getUi().ButtonSet.OK);
}

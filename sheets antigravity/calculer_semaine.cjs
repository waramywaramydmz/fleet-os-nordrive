const ExcelJS = require('exceljs');
const fs = require('fs');
const path = require('path');

console.log('========================================================================');
console.log('🚗 FLEET OS — MOTEUR DE CONCILIATION & CALCUL DE PAIE HEBDOMADAIRE');
console.log('========================================================================\n');

function parseCSV(text) {
  const rows = [];
  let row = [];
  let inQuotes = false;
  let cur = '';
  for (let i = 0; i < text.length; i++) {
    const ch = text[i];
    if (ch === '"') {
      if (inQuotes && text[i + 1] === '"') { cur += '"'; i++; }
      else { inQuotes = !inQuotes; }
    } else if (ch === ',' && !inQuotes) {
      row.push(cur.trim()); cur = '';
    } else if ((ch === '\r' || ch === '\n') && !inQuotes) {
      if (ch === '\r' && text[i + 1] === '\n') i++;
      row.push(cur.trim());
      if (row.some(c => c !== '')) rows.push(row);
      row = []; cur = '';
    } else cur += ch;
  }
  if (cur || row.length > 0) {
    row.push(cur.trim());
    if (row.some(c => c !== '')) rows.push(row);
  }
  return rows;
}

function norm(s) {
  return (s || '').toLowerCase()
    .normalize('NFD').replace(/[\u0300-\u036f]/g, '')
    .replace(/[^a-z0-9]/g, '');
}

// 1. Détection intelligente des fichiers CSV
function findFile(patterns, exclude = []) {
  const files = fs.readdirSync('.');
  const matches = [];
  for (const f of files) {
    if (!f.endsWith('.csv')) continue;
    if (exclude.some(ex => f.toLowerCase() === ex.toLowerCase())) continue;
    const lower = f.toLowerCase();
    for (const pat of patterns) {
      if (lower.includes(pat.toLowerCase())) {
        const stats = fs.statSync(f);
        matches.push({ name: f, mtime: stats.mtimeMs });
        break;
      }
    }
  }
  if (matches.length === 0) return null;
  // Prioriser 'uber.csv' ou 'bolt.csv' exactement s'il existe
  const exact = matches.find(m => patterns.some(p => m.name.toLowerCase() === p.toLowerCase() + '.csv'));
  if (exact) return exact.name;
  // Sinon trier par date de modification la plus récente
  matches.sort((a, b) => b.mtime - a.mtime);
  return matches[0].name;
}

const customUber = process.argv[2];
const customBolt = process.argv[3];

const uberFile = customUber && fs.existsSync(customUber) 
  ? customUber 
  : findFile(['uber', 'rapport-uber', 'rapport_uber'], ['DATA_UBER.csv', 'EXPORT_QONTO_SEPA.csv', 'DASHBOARD_SEMAINE.csv']);

const boltFile = customBolt && fs.existsSync(customBolt)
  ? customBolt 
  : findFile(['bolt', 'driver_activity', 'rapport-bolt', 'rapport_bolt'], ['DATA_BOLT.csv', 'EXPORT_QONTO_SEPA.csv', 'DASHBOARD_SEMAINE.csv']);

if (!uberFile || !boltFile) {
  console.error('❌ Erreur : Impossible de trouver les fichiers CSV Uber et Bolt dans le dossier.');
  console.error(`- Fichier Uber trouvé : ${uberFile || 'AUCUN (Déposez votre CSV et nommez-le uber.csv)'}`);
  console.error(`- Fichier Bolt trouvé : ${boltFile || 'AUCUN (Déposez votre CSV et nommez-le bolt.csv)'}`);
  process.exit(1);
}

console.log(`📁 Fichier Uber utilisé : ${uberFile}`);
console.log(`📁 Fichier Bolt utilisé : ${boltFile}\n`);

// 2. Parsing Bolt
const boltRaw = parseCSV(fs.readFileSync(boltFile, 'utf8'));
const boltDrivers = boltRaw.slice(1).map(r => ({
  nom: r[0],
  email: r[1],
  tel: r[2],
  brutTotal: parseFloat(r[3]) || 0,
  app: parseFloat(r[4]) || 0,
  especesBrut: parseFloat(r[5]) || 0,
  especesPercu: parseFloat(r[6]) || 0,
  pourboires: parseFloat(r[7]) || 0,
  campagnes: parseFloat(r[8]) || 0,
  remboursements: parseFloat(r[9]) || 0,
  annulations: parseFloat(r[10]) || 0,
  peages: parseFloat(r[11]) || 0,
  reservation: parseFloat(r[12]) || 0,
  totalFrais: parseFloat(r[13]) || 0,
  fraisService: parseFloat(r[14]) || 0,
  rembClients: parseFloat(r[15]) || 0,
  autresFrais: parseFloat(r[16]) || 0,
  revenusNets: parseFloat(r[17]) || 0,
  paiementPrevu: parseFloat(r[18]) || 0,
  idChauffeur: r[23] || '',
  idIndividuel: r[24] || ''
}));

// 3. Parsing Uber
const uberRaw = parseCSV(fs.readFileSync(uberFile, 'utf8'));
const uberHeaders = (uberRaw[0] || []).map(h => (h || '').trim().toLowerCase());

const findUberCol = (predicate, defaultIdx) => {
  const idx = uberHeaders.findIndex(predicate);
  return idx !== -1 ? idx : defaultIdx;
};

const colUuid = findUberCol(h => h.includes('uuid'), 0);
const colPrenom = findUberCol(h => h.includes('prénom') || h.includes('prenom'), 1);
const colNom = findUberCol(h => (h.includes('nom') && !h.includes('prénom') && !h.includes('prenom')), 2);
const colRevTotaux = findUberCol(h => h === 'revenus totaux' || (h.includes('revenus') && !h.includes(':')), 3);
const colPrixNet = findUberCol(h => h.includes('prix net'), 4);
const colBonus = findUberCol(h => h.includes('bonus'), 5);
const colRemb = findUberCol(h => h.startsWith('remboursements') && !h.includes('péage') && !h.includes('peage'), 6);
const colTransfBanque = findUberCol(h => h.includes('transferts'), 8);
const colEspeces = findUberCol(h => h.includes('espèces') || h.includes('especes'), 9);
const colFraisResa = findUberCol(h => h.includes('réservation') || h.includes('reservation'), 12);
const colFraisUber = findUberCol(h => (h.includes('montant versé à uber') || h.includes('montant verse')) && !h.includes('réservation') && !h.includes('reservation'), 10);
const colPeage = findUberCol(h => h.includes('péage') || h.includes('peage'), 13);
const colPourboire = findUberCol(h => h.includes('pourboire'), 11);
const colObjet = findUberCol(h => h.includes('objet'), 14);

const uberDrivers = uberRaw.slice(1).map(r => {
  const uuid = r[colUuid] || '';
  const prenom = (r[colPrenom] || '').replace(/[\r\n]+/g, ' ').trim();
  const nom = (r[colNom] || '').replace(/[\r\n]+/g, ' ').trim();
  const fullName = (prenom + ' ' + nom).trim();
  const revTotaux = parseFloat(r[colRevTotaux]) || 0;
  const prixNetCourse = parseFloat(r[colPrixNet]) || 0;
  const bonus = parseFloat(r[colBonus]) || 0;
  const remb = parseFloat(r[colRemb]) || 0;
  const versements = parseFloat(r[7]) || 0;
  const transfBanque = parseFloat(r[colTransfBanque]) || 0;
  const especes = Math.abs(parseFloat(r[colEspeces]) || 0);
  const frais = parseFloat(r[colFraisUber]) || 0;
  const fraisResa = Math.abs(parseFloat(r[colFraisResa]) || 0);
  const peage = parseFloat(r[colPeage]) || 0;
  const pourboire = parseFloat(r[colPourboire]) || 0;
  const objet = parseFloat(r[colObjet]) || 0;

  // Calcul exact du virement Uber
  let virementReel = 0;
  if (revTotaux > 0 || remb > 0 || especes > 0) {
    virementReel = Math.round(((revTotaux + remb + frais) - especes) * 100) / 100;
  }

  return {
    uuid,
    prenom,
    nom,
    fullName,
    revTotaux,
    prixNetCourse,
    bonus,
    remb,
    especes,
    frais,
    fraisResa,
    peage,
    pourboire,
    objet,
    virementReel
  };
});

// 4. Liste officielle des 69 Chauffeurs
const CHAUFFEURS_69_RAW = [
  ["Abdallah Hamzi", "ABDALLAH HAMZI", "", "hamzi.obayda@gmail.com", "", "+33624017099", ""],
  ["Abdoul Azize Mohamed Safir", "ABDOUL AZIZE MOHAMED SAFIR", "Abdoul Azize MOHAMED SAFIR", "abazizsafir.pro@gmail.com", "abazizsafir.pro@gmail.com", "+33761529386", "+33761529386"],
  ["Bechir Ben Smida", "Bechir Ben Smida", "", "bsbm67@gmail.com", "", "+33762923919", ""],
  ["Fathia Boufalgha", "FATHIA BOUFALGHA", "", "titiprestige75@gmail.com", "", "+33759625661", ""],
  ["Rachid Chinoune", "rachid CHINOUNE", "Rachid Chinoune", "rachidchinoune@yahoo.fr", "crtp2023@gmail.com", "+33619305126", "+33619305126"],
  ["Salah Fitouri", "Salah Fitouri", "", "Salih071@live.fr", "", "+33627393006", ""],
  ["Bouchaib Belkharras", "Bouchaib Belkharras", "Bouchaïb Belkharras", "belbouch92@gmail.com", "bouch.bb@hotmail.com", "+33658123002", "+33658123002"],
  ["Brandon Jean Jimmy Louissaint", "BRANDON JEAN-JIMMY LOUISSAINT", "Brandon Jean Jimmy Louissaint", "louissaintbrandon@hotmail.com", "louissaintbrandon@hotmail.com", "+33635481157", "+33635481157"],
  ["Farid Adkhis", "Farid Adkhis", "Farid Adkhis", "Farid.adk@gmail.com", "Farid.adk@gmail.com", "+33651950695", "+33651950695"],
  ["Imane Sarraj", "Imane Sarraj", "Imane Sarraj", "imane.sar2009@gmail.com", "imane.sar2009@gmail.com", "+33658430022", "+33658430022"],
  ["Mehmet Ergun", "MEHMET ERGUN", "Mehmet Ergun", "mht.egn1@icloud.com", "Mht.Egn1@icloud.com", "+33782313002", "+33782313002"],
  ["Mohamed Aloui", "MOHAMED ALOUi", "Mohamed Aloui", "fainflash@hotmail.com", "Fainflash@hotmail.com", "+33762769951", "+33762769951"],
  ["Omar Chemkhi", "Omar Chemkhi", "", "omarchemkhioc@gmail.com", "", "+33783433834", ""],
  ["Raoul Mansfeld Bemma", "Raoul Mansfeld Bemma", "Raoul Mansfeld Bemma", "raoulbemma0@gmail.com", "bemra@hotmail.fr", "+33672345507", "+33672345507"],
  ["Ratiba Beldjoudi", "Ratiba Beldjoudi", "Ratiba GUESRI", "amir2014bba@hotmail.com", "beldjoudilina@hotmail.fr", "+33695084256", "+33695084256"],
  ["Zahia Guesri", "Zahia Guesri", "Zahia Guesri", "c_zahwa@hotmail.com", "zaza.guess@hotmail.fr", "+33669354856", "+33669354856"],
  ["Abdel Mouamane Chelaif", "Abdel Mouamane Chelaif", "Abdel Mouamane CHELAIF", "chelaifmoumen@gmail.com", "chelaifmoumen@gmail.com", "+33753606026", "+33753606026"],
  ["Abdelaziz Guerfi", "Abdelaziz Guerfi", "", "guerfiaziz75@gmail.com", "", "+33671648042", ""],
  ["Abdullah Kartal", "Abdullah Kartal", "Abdullah Kartal", "abdullahkartal@outlook.fr", "abdullahkartal@outlook.fr", "+33782374844", "+33782374844"],
  ["Adam Machtou", "Adam MACHTOU", "Adam Machtou", "machtou.adam@gmail.com", "machtou.adam@gmail.com", "+33618757203", "+33618757203"],
  ["Adel Benabdelouahed", "ADEL BENABDELOUAHED", "Mohammed Adel benabdelouahed", "adelbenan@gmail.com", "adelbenan@gmail.com", "+33782790937", "+33782790937"],
  ["Ahmed Kherfouche", "Ahmed Kherfouche", "", "cena2309@hotmail.fr", "", "+33612943281", ""],
  ["Alhassane Balde", "Alhassane Balde", "Alhassane Balde", "alh.balde05@gmail.com", "alh.balde05@gmail.com", "+33783639299", "+33783639299"],
  ["Ali Alouani", "Ali Alouani", "Ali Alouani", "ali.alouani.95@gmail.com", "ali.alouani.95@gmail.com", "+33618122054", "+33618122054"],
  ["Asma Kissi", "ASMA KISSI", "", "kissiasma333@gmail.com", "", "+33626907046", ""],
  ["Aymen Ltifi", "Aymen Ltifi", "Aymen Ltifi", "ltifiservices@outlook.com", "ltifiservices@outlook.com", "+33760731822", "+33760731822"],
  ["Bilel Bouhali", "Bilel Bouhali", "", "thecompany2275@gmail.com", "", "+33743767400", ""],
  ["Brahim Aloui", "Brahim Aloui", "brahim aloui", "alouibrahim@ymail.com", "alouibrahim@ymail.com", "+33624959758", "+33624959758"],
  ["Bryan Grégory Michaud", "Bryan Grégory Michaud", "Bryan Gregory Michaud", "bryan.michaud11@gmail.com", "bryan.michaud11@gmail.com", "+33783633084", "+33783633084"],
  ["Christopher Andrew Mongis", "Christopher Andrew Mongis", "", "mongis.christopher77@gmail.com", "", "+33605838100", ""],
  ["Daba Gandega", "Daba Gandega", "", "dg93240@hotmail.fr", "", "+33652830506", ""],
  ["Elias Bihi", "ELIAS BIHI", "Elias Bihi", "bihieliaspro@gmail.com", "Bihieliaspro@gmail.com", "+33749483638", "+33749483638"],
  ["Elias Hocine Abdellou", "ELIAS HOCINE ABDELLOU", "", "eabdellou92@gmail.com", "", "+33676966368", ""],
  ["Faical Jabri", "FAICAL JABRI", "", "faicaljabri@yahoo.fr", "", "+33625654846", ""],
  ["Foued Karoui", "FOUED KAROUI", "Foued Karoui", "karfoued@gmail.com", "karfoued@gmail.com", "+33625815497", "+33625815497"],
  ["Hassene Khiari", "Hassene Khiari", "Hassene KHIARI", "le-solitaire@live.com", "hassenekhiari9@gmail.com", "+33616045890", "+33699048832"],
  ["Haythem Dhifi", "Haythem Dhifi", "Haythem Dhifi", "mizyen82@gmail.com", "dhifi.haythem.hd@gmail.com", "+33612674248", "+33612674248"],
  ["Helton Jose Coelho Da Veiga", "Hélton José Coelho Da Veiga", "", "heltondaveiga2014@gmail.com", "", "+33619098335", ""],
  ["Issa Tandjigora", "ISSA TANDJIGORA", "Issa Tandjigora", "issasenypro@gmail.com", "Issasenypro@gmail.com", "+33651962203", "+33638776237"],
  ["Jalel Sehili", "Jalel Sehili", "", "jsehili@sfr.fr", "", "+33621066584", ""],
  ["Jovan Milic", "JOVAN MILIC", "Jovan Milic", "jovan93@outlook.fr", "Jovan93@outlook.fr", "+33695405086", "+33695405086"],
  ["Kamel Kebli", "Kamel Kebli", "Kamel Kebli", "k.kebli@outlook.fr", "k.kebli@outlook.fr", "+33605882286", "+33605882286"],
  ["Karim Sidibe", "Karim Sidibe", "", "karimsidibe18@outlook.fr", "", "+33625270201", ""],
  ["Khaled Kebli", "Khaled Kebli", "Khaled Kebli", "kh.kebli@hotmail.com", "Kh.kebli@hotmail.com", "+33765512904", "+33765512904"],
  ["Khalid Okeil", "Khalid Okeil", "Khalid OKEIL", "khalidokeil@outlook.fr", "khalidokeil@outlook.fr", "+33685369354", "+33685369354"],
  ["Mehdi Bouchnag", "MEHDI BOUCHNAG", "Mehdi Bouchnag", "bouchnagmehdi@yahoo.fr", "Bouchnagmehdi@yahoo.fr", "+33623577818", "+33623577818"],
  ["Merouane Remam", "Merouane Remam", "Merouane REMAM", "merouane.rema@gmail.com", "Midoremam@hotmail.com", "+33656847648", "+33656847648"],
  ["Mohamed Bayrem Allagui", "Mohamed-Bayrem ALLAGUI", "Mohamed Bayrem Allagui", "allaguimohamedbayrem@gmail.com", "bayrouma18@gmail.com", "+33621478333", "+33621478333"],
  ["Mohamed Bouali", "MOHAMED BOUALI", "", "bouali75med@gmail.com", "", "+33751079527", ""],
  ["Mohamed El Khabazi", "Mohamed El Khabazi", "Mohamed El Khabazi", "elkhabazimohamed1@gmail.com", "elkhabazimohamed1@gmail.com", "+33784678499", "+33784678499"],
  ["Mohamed Fehdi Ben-Fekih", "MOHAMED FEHDI BEN-FEKIH", "Mohamed Fehdi Ben-Fekih", "mfbf@hotmail.fr", "mfbf@hotmail.fr", "+33662727924", "+33662727924"],
  ["Mourad Bouakkaz", "Mourad Bouakkaz", "Mourad Bouakkaz", "mbouakkaz874@hotmail.com", "mbouakkaz874@hotmail.com", "+33699949546", "+33699949546"],
  ["Noureddine Tattouche", "Noureddine Tattouche", "Noureddine Tattouche", "tattouche.noureddine@yahoo.com", "tattouche.noureddine@yahoo.com", "+33675941291", "+33675941291"],
  ["Ousmane Sylla", "OUSMANE SYLLA", "", "m.sylla.ousmane@gmail.com", "", "+33756909610", ""],
  ["Raed Loghmari", "Raed Loghmari", "Raed Loghmari", "raedloghmari95@gmail.com", "raed85@hotmail.fr", "+33652147683", "+33652147683"],
  ["Ramzi Khannoussi", "Ramzi Khannoussi", "Ramzi Khannoussi", "mrdriverparis@gmail.com", "ziram_k@yahoo.fr", "+33682163015", "+33682163015"],
  ["Robert-Ionut Neagu", "Robert-Ionut Neagu", "Robert-Ionut Neagu", "robertneagu76@gmail.com", "robertneagu44@gmail.com", "+33660198093", "+33656859258"],
  ["Ryan Challab", "Ryan Challab", "Ryan Ouali Challab", "ryanchallab93600@gmail.com", "ryanchallab93600@gmail.com", "+33784678499", "+33768419070"],
  ["Sadibou Diatta", "Sadibou Diatta", "", "sadiboudiatta27@gmail.com", "", "+33650952956", ""],
  ["Samir Laazibi", "Samir Laazibi", "SAMIR LAAZIBI", "abenbourahla84@gmail.com", "laazibisamir@gmail.com", "+33661926124", "+33661926124"],
  ["Siré Cyril Niakaté", "SIRE CYRIL NIAKATE", "Siré Niakaté", "sireniakate@msn.com", "sireniakate71@gmail.com", "+33614251414", "+33764005951"],
  ["Sofiane Akli", "Sofiane Akli", "Sofiane Akli", "sofiane.uber14@hotmail.com", "sofiane.akli@hotmail.com", "+33666617828", "+3300000000"],
  ["Souleymane Sokhona", "", "Souleymane Sokhona", "", "ssokhona42@gmail.com", "", "+33769507901"],
  ["Youcef Kherfouche", "Youcef Kherfouche", "Youcef Kherfouche", "kherfoucheyoucef@gmail.com", "kherfoucheyoucef@gmail.com", "+33624707164", "+33624707164"],
  ["Younoussa Bah", "YOUNOUSSA BAH", "", "thiernomamadou24@yahoo.com", "", "+33605679771", ""],
  ["Ziaed Hamila", "ZIAED HAMILA", "Ziaed Hamila", "teezere@outlook.fr", "Teezere@outlook.fr", "+33650235792", "+33650235792"],
  ["Ali Kara", "Ali Kara", "Ali Kara", "ali1978kara@hotmail.fr", "ali1978kara@hotmail.fr", "+33695644791", "+33695644791"],
  ["Stephane Kouassi", "Stephane Kouassi", "Stéphane Kouassi", "kscga@outlook.fr", "koffivillepinte@gmail.com", "+33601360917", "+33601360917"],
  ["Rayen Klai", "Rayen Klai", "Rayen Klai", "vtc93600@gmail.com", "rayenvtc@outlook.fr", "+33766007929", "+33766833146"],["Yahia Zemani", "Yahia Zemani", "Yahia Zemani", "zemaservices37@gmail.com", "zemaservices37@gmail.com", "+33641128055", "+33641128055"],
  ["Hasan Emin Coskun", "Hasan Emin Coskun", "Hasan Emin Coskun", "hasanemincoskun@outlook.fr", "hasanemincoskun@outlook.fr", "+33652092796", "+33652092796"],
  ["Raphaël Ernest Amié Vimond", "Raphaël Ernest Amié Vimond", "", "raphael.vimond@gmail.com", "", "+33600000072", ""],
];

const CHAUFFEURS_69 = CHAUFFEURS_69_RAW.map((item, idx) => ({
  num: String(idx + 1).padStart(2, '0'),
  nomChauffeur: item[0],
  nomUber: item[1] || '',
  nomBolt: item[2] || '',
  mailUber: item[3] || '',
  mailBolt: item[4] || '',
  telUber: item[5] || '',
  telBolt: item[6] || '',
  iban: (item[0] === 'Siré Cyril Niakaté') ? 'FR7630006000010000000000189' : ''
}));

// 4bis. IBAN / RIB persistants : lus depuis IBANS_CHAUFFEURS.csv (format : Nom Prénom;IBAN)
// Ce fichier n'est JAMAIS écrasé s'il existe → les RIB saisis restent à chaque nouveau calcul / import.
const IBAN_FILE = 'IBANS_CHAUFFEURS.csv';
if (fs.existsSync(IBAN_FILE)) {
  const ibanLines = fs.readFileSync(IBAN_FILE, 'utf8').replace(/^\uFEFF/, '').split(/\r?\n/).slice(1);
  const ibanMap = new Map();
  for (const line of ibanLines) {
    if (!line.trim()) continue;
    const sep = line.includes(';') ? ';' : ',';
    const parts = line.split(sep);
    const name = (parts[0] || '').replace(/^"|"$/g, '').trim();
    const iban = (parts[1] || '').replace(/^"|"$/g, '').replace(/\s+/g, '').toUpperCase();
    if (name && iban) ibanMap.set(norm(name), iban);
  }
  let nbIban = 0;
  CHAUFFEURS_69.forEach(c => {
    const found = ibanMap.get(norm(c.nomChauffeur));
    if (found) { c.iban = found; nbIban++; }
  });
  console.log(`🏦 IBAN chargés depuis ${IBAN_FILE} : ${nbIban}/${CHAUFFEURS_69.length}\n`);
} else {
  const tpl = ['Nom Prénom;IBAN'].concat(CHAUFFEURS_69.map(c => `${c.nomChauffeur};${c.iban}`));
  fs.writeFileSync(IBAN_FILE, '\ufeff' + tpl.join('\r\n'), 'utf8');
  console.log(`🏦 Modèle ${IBAN_FILE} créé : remplis la colonne IBAN, elle sera conservée à chaque calcul.\n`);
}


// Grille officielle des forfaits contractuels personnalisés validée
const DRIVER_FEE_GRID = {
  // Ancienne Catégorie 70 € passée à 50 €
  'abdallah hamzi': 50,
  'abdoul azize mohamed safir': 50,
  'bechir ben smida': 50,
  'fathia boufalgha': 50,
  'rachid chinoune': 50,
  'salah fitouri': 50,
  'idir belkhir': 50,

  // Catégorie 60 €
  'bouchaib belkharras': 60,
  'brandon jean jimmy louissaint': 60,
  'brandon jean-jimmy louissaint': 60,
  'brandon louissaint': 60,
  'farid adkhis': 60,
  'imane sarraj': 60,
  'mehmet ergun': 60,
  'mohamed aloui': 60,
  'omar chemkhi': 60,
  'raoul mansfeld bemma': 60,
  'ratiba beldjoudi': 60,
  'ratiba guesri': 60,
  'zahia guesri': 60,

  // Catégorie 40 €
  'ali kara': 40,
  'stephane kouassi': 40,
  'hasan emin coskun': 40,

  // Catégorie 0 €
  'rayen klai': 0,

  // Catégorie 50 € confirmée
  'brahim aloui': 50,
  'haythem dhifi': 50,
  'jovan milic': 50,
  'raed loghmari': 50,
  'issa tandjigora': 50
};

// Cotisations sociales / charges déductibles de la semaine
const DRIVER_CHARGES_WEEK = {
  'fathia boufalgha': 137.50,
  'karim sidibe': 20.00,
  'mohamed el khabazi': 137.50
};

function getDriverContractFee(driverName) {
  const n = norm(driverName);
  for (const [key, fee] of Object.entries(DRIVER_FEE_GRID)) {
    if (norm(key) === n) return fee;
  }
  return 50.00; // Par défaut 50 €
}

function getDriverCharges(driverName) {
  const n = norm(driverName);
  for (const [key, charge] of Object.entries(DRIVER_CHARGES_WEEK)) {
    if (norm(key) === n) return charge;
  }
  return 0.00;
}

// 5. Calculs par chauffeur
let totalUberVirement = 0;
let totalUberEspeces = 0;
let totalBoltVirement = 0;
let totalBoltEspeces = 0;
let totalEncaisse = 0;
let totalRattachement = 0;
let totalNet = 0;

const calculations = CHAUFFEURS_69.map((c, idx) => {
  // Match Uber
  const uMatch = uberDrivers.find(u => {
    if (u.fullName === 'Bilel Bouhali') return false; // Ligne société
    const uNorm1 = norm(u.fullName);
    const uNorm2 = norm(u.nom + ' ' + u.prenom);
    if (c.nomUber && (norm(c.nomUber) === uNorm1 || norm(c.nomUber) === uNorm2)) return true;
    if (norm(c.nomChauffeur) === uNorm1 || norm(c.nomChauffeur) === uNorm2) return true;
    return false;
  });

  // Match Bolt
  const bMatch = boltDrivers.find(b => {
    if (c.nomBolt && norm(b.nom) === norm(c.nomBolt)) return true;
    if (norm(b.nom) === norm(c.nomChauffeur)) return true;
    return false;
  });

  const virementUber = uMatch ? uMatch.virementReel : 0;
  const especesUber = uMatch ? uMatch.especes : 0;
  const virementBolt = bMatch ? bMatch.paiementPrevu : 0;
  const especesBolt = bMatch ? bMatch.especesPercu : 0;

  const encaisse = Math.round((virementUber + virementBolt) * 100) / 100;
  const hasWorked = (virementUber !== 0 || virementBolt !== 0 || especesUber > 0 || especesBolt > 0 || (uMatch && uMatch.revTotaux > 0) || (bMatch && bMatch.brutTotal > 0));

  // Règle d'or : rattachement personnalisé UNIQUEMENT si a travaillé
  const tarifContractuel = getDriverContractFee(c.nomChauffeur);
  const rattachement = hasWorked ? tarifContractuel : 0.00;
  const charges = getDriverCharges(c.nomChauffeur);
  const bonus = 0.00;
  const malus = 0.00;
  const net = hasWorked ? (Math.round((encaisse - rattachement - charges - malus + bonus) * 100) / 100) : 0.00;

  totalUberVirement += virementUber;
  totalUberEspeces += especesUber;
  totalBoltVirement += virementBolt;
  totalBoltEspeces += especesBolt;
  totalEncaisse += encaisse;
  totalRattachement += rattachement;
  totalNet += net;

  return {
    ...c,
    virementUber,
    especesUber,
    virementBolt,
    especesBolt,
    encaisse,
    rattachement,
    charges,
    bonus,
    malus,
    net,
    hasWorked
  };
});

totalUberVirement = Math.round(totalUberVirement * 100) / 100;
totalUberEspeces = Math.round(totalUberEspeces * 100) / 100;
totalBoltVirement = Math.round(totalBoltVirement * 100) / 100;
totalBoltEspeces = Math.round(totalBoltEspeces * 100) / 100;
totalEncaisse = Math.round(totalEncaisse * 100) / 100;
totalRattachement = Math.round(totalRattachement * 100) / 100;
totalNet = Math.round(totalNet * 100) / 100;

const activeCount = calculations.filter(c => c.hasWorked).length;
const inactiveCount = calculations.length - activeCount;

// Affichage Console
console.log('------------------------------------------------------------------------');
console.log('📊 RÉSULTAT DE LA SEMAINE CONSOLIDÉE (69 CHAUFFEURS)');
console.log('------------------------------------------------------------------------');
console.log(`• Chauffeurs actifs ayant roulé : ${activeCount}`);
console.log(`• Chauffeurs inactifs (0,00 €)   : ${inactiveCount}`);
console.log(`• Total Encaissé Plateformes     : ${totalEncaisse.toLocaleString('fr-FR', { minimumFractionDigits: 2 })} € (Uber: ${totalUberVirement.toLocaleString('fr-FR', { minimumFractionDigits: 2 })} € | Bolt: ${totalBoltVirement.toLocaleString('fr-FR', { minimumFractionDigits: 2 })} €)`);
console.log(`• Suivi Espèces Flotte en poche  : ${(totalUberEspeces + totalBoltEspeces).toLocaleString('fr-FR', { minimumFractionDigits: 2 })} € (Uber: ${totalUberEspeces.toLocaleString('fr-FR', { minimumFractionDigits: 2 })} € | Bolt: ${totalBoltEspeces.toLocaleString('fr-FR', { minimumFractionDigits: 2 })} €)`);
console.log(`• Total Redevance Rattachement   : ${totalRattachement.toLocaleString('fr-FR', { minimumFractionDigits: 2 })} € (${activeCount} chauffeurs x 50,00 €)`);
console.log(`• TOTAL NET À VIRER (QONTO)      : ${totalNet.toLocaleString('fr-FR', { minimumFractionDigits: 2 })} €`);
console.log('------------------------------------------------------------------------\n');

// 6. Génération Fichier EXPORT_QONTO_SEPA.csv
const qontoLines = ['beneficiary_name,iban,amount,currency,reference'];
calculations.forEach(c => {
  if (c.hasWorked && c.net > 0) {
    qontoLines.push(`"${c.nomChauffeur}","${c.iban}",${c.net.toFixed(2)},"EUR","Virement Semaine - ${c.nomChauffeur}"`);
  }
});
fs.writeFileSync('EXPORT_QONTO_SEPA.csv', '\ufeff' + qontoLines.join('\r\n'), 'utf8');
console.log(`✅ Fichier Qonto généré : EXPORT_QONTO_SEPA.csv (${qontoLines.length - 1} virements prêts)`);

// 7. Génération DASHBOARD_SEMAINE.csv
const dashHeaders = [
  'N°', 'Nom Prénom (Chauffeur)', 'Nom Uber', 'Nom Bolt', 
  'Mail Uber', 'Mail Bolt', 'Tél. Uber', 'Tél. Bolt', 
  'Virement Réel Uber (€)', 'Espèces Uber (€)',
  'Virement Réel Bolt (€)', 'Espèces Bolt (€)',
  'Total Réel Encaissé (Uber + Bolt) (€)', 
  'Rattachement (€)', 'Charges (€)', 'Bonus (€)', 'Malus (€)', 
  'CASE VIREMENT (NET À ENVOYER) (€)', 'RIB / IBAN (QONTO)'
];

const dashCsvRows = calculations.map(c => [
  c.num,
  `"${c.nomChauffeur}"`,
  `"${c.nomUber}"`,
  `"${c.nomBolt}"`,
  `"${c.mailUber}"`,
  `"${c.mailBolt}"`,
  `"${c.telUber}"`,
  `"${c.telBolt}"`,
  c.virementUber.toFixed(2),
  c.especesUber.toFixed(2),
  c.virementBolt.toFixed(2),
  c.especesBolt.toFixed(2),
  c.encaisse.toFixed(2),
  c.rattachement.toFixed(2),
  c.charges.toFixed(2),
  c.bonus.toFixed(2),
  c.malus.toFixed(2),
  c.net.toFixed(2),
  `"${c.iban}"`
]);

dashCsvRows.push([
  'TOTAL', '"TOTAL CONSOLIDÉ SEMAINE"', '""', '""', '""', '""', '""', '""',
  totalUberVirement.toFixed(2),
  totalUberEspeces.toFixed(2),
  totalBoltVirement.toFixed(2),
  totalBoltEspeces.toFixed(2),
  totalEncaisse.toFixed(2),
  totalRattachement.toFixed(2),
  '0.00', '0.00', '0.00',
  totalNet.toFixed(2),
  '""'
]);

const dashCsvContent = [dashHeaders.join(';')].concat(dashCsvRows.map(r => r.join(';'))).join('\r\n');
fs.writeFileSync('DASHBOARD_SEMAINE.csv', '\ufeff' + dashCsvContent, 'utf8');
console.log(`✅ Fichier CSV Dashboard généré : DASHBOARD_SEMAINE.csv (69 chauffeurs + Totaux)`);

// 8. Sauvegarde du fichier Excel Complet FLEET_OS_Nordrive_Reconciliation_Qonto.xlsx
async function buildFullExcelWorkbook() {
  const wb = new ExcelJS.Workbook();
  wb.creator = 'FLEET OS — Ingénierie Financière & Conciliation VTC';
  wb.lastModifiedBy = 'FLEET OS Engine';
  wb.created = new Date();
  wb.modified = new Date();

  const NAVY = '0F172A';
  const SLATE = '1E293B';
  const UBER_BLACK = '18181B';
  const BOLT_GREEN = '16A34A';
  const QONTO_PURPLE = '4F46E5';
  const KPI_BG_1 = 'EFF6FF';
  const KPI_BG_2 = 'F5F3FF';
  const KPI_BG_4 = 'ECFDF5';

  const fontHeader = { name: 'Segoe UI', size: 10, bold: true, color: { argb: 'FFFFFFFF' } };
  const fontMonospaceBold = { name: 'Consolas', size: 10, bold: true, color: { argb: 'FF047857' } };
  const fontMonospace = { name: 'Consolas', size: 10, color: { argb: 'FF0F172A' } };

  const borderThin = {
    top: { style: 'thin', color: { argb: 'FFCBD5E1' } },
    left: { style: 'thin', color: { argb: 'FFCBD5E1' } },
    bottom: { style: 'thin', color: { argb: 'FFCBD5E1' } },
    right: { style: 'thin', color: { argb: 'FFCBD5E1' } }
  };
  const currencyFormat = '#,##0.00 €';

  // --- ONGLET 1 : DATA_UBER ---
  const wsUber = wb.addWorksheet('DATA_UBER', { views: [{ showGridLines: true }] });
  const uberCols = [
    'Nom chauffeur Uber', 'Prénom', 'UUID Chauffeur', 'Courses (D)', 'Frais service Uber', 
    'Péages (F)', 'Pourboires (G)', 'Remboursements', 'Frais annulation', 'Espèces collectées (J)', 
    'Promotions / Bonus (K)', 'Frais réservation Uber (L)', 'Ajustements courses', 'Déductions diverses', 
    'Recharges', 'Équipements', 'Régularisations diverses', 'Régularisations / Péages (R)', 
    'Retenues fiscales', 'Acomptes', 'Autres déductions', 'Montant brut', 'Montant net', 
    'Banque', 'Date versement', 'Virement Réel Uber (Z)'
  ];
  wsUber.addRow(uberCols);
  const uberHeaderRow = wsUber.getRow(1);
  uberHeaderRow.height = 28;
  uberHeaderRow.eachCell((cell, colNum) => {
    cell.font = fontHeader;
    cell.alignment = { horizontal: 'center', vertical: 'middle' };
    cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: colNum === 26 ? 'FF059669' : 'FF' + UBER_BLACK } };
    cell.border = borderThin;
  });

  uberDrivers.forEach((u, i) => {
    const rIdx = 2 + i;
    const isCompany = (u.fullName === 'Bilel Bouhali');
    const uRow = wsUber.addRow([
      u.fullName, u.prenom, u.uuid,
      u.prixNetCourse, 0, u.peage, u.pourboire, 0, 0, u.especes,
      u.bonus, u.fraisResa, 0, 0, 0, 0, 0, u.remb + u.objet, 0, 0, 0,
      u.revTotaux, u.virementReel, 'QONTO', '2026-09-28'
    ]);
    uRow.height = 22;
    [4, 6, 7, 10, 11, 12, 18, 22, 23].forEach(c => { uRow.getCell(c).numFmt = currencyFormat; });

    const zCell = uRow.getCell(26);
    zCell.value = {
      formula: `IF(A${rIdx}="", "", (D${rIdx}+G${rIdx}+K${rIdx}+R${rIdx}) - ABS(J${rIdx}) - ABS(L${rIdx}))`,
      result: u.virementReel
    };
    zCell.numFmt = currencyFormat;
    zCell.font = { name: 'Segoe UI', size: 10, bold: true, color: { argb: 'FF065F46' } };
    zCell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFECFDF5' } };
    uRow.eachCell({ includeEmpty: true }, c => { c.border = borderThin; });
  });

  wsUber.columns.forEach((col, idx) => {
    col.width = (idx === 0) ? 28 : (idx === 25) ? 24 : 16;
  });

  // --- ONGLET 2 : DATA_BOLT ---
  const wsBolt = wb.addWorksheet('DATA_BOLT', { views: [{ showGridLines: true }] });
  const boltCols = [
    'Nom chauffeur Bolt', 'Email', 'Téléphone', 'ID Bolt', 
    'Brut Courses App (E)', 'Brut Courses Autres (F)', 'Espèces collectées chauffeur (G)',
    'Bonus campagnes (H)', 'Compensations (I)', 'Primes objectifs (J)', 'Pourboires Bolt (K)', 'Dédommagements (L)',
    'Péages (M)', 'Frais annulation (N)', 'Commission Bolt principale (O)', 'Commission Bolt secondaire (P)', 'Autres frais Bolt (Q)',
    'Régularisation commission (R)', 'Remboursements clients (S)', 'Frais attente (T)', 'Ajustements manuels (U)',
    'Autres crédits 1 (V)', 'Autres crédits 2 (W)', 'Statut (X)', 'Date relevé (Y)', 'Virement Réel Bolt (Z)'
  ];
  wsBolt.addRow(boltCols);
  const boltHeaderRow = wsBolt.getRow(1);
  boltHeaderRow.height = 28;
  boltHeaderRow.eachCell((cell, colNum) => {
    cell.font = fontHeader;
    cell.alignment = { horizontal: 'center', vertical: 'middle' };
    cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: colNum === 26 ? 'FF059669' : 'FF' + BOLT_GREEN } };
    cell.border = borderThin;
  });

  boltDrivers.forEach((b, i) => {
    const rIdx = 2 + i;
    const bRow = wsBolt.addRow([
      b.nom, b.email, b.tel, b.idChauffeur,
      b.app, b.especesBrut, b.especesPercu, b.campagnes, 0, 0, b.pourboires, b.remboursements,
      b.peages, b.annulations, b.fraisService, 0, b.autresFrais, 0, b.rembClients, 0, 0,
      0, 0, 'Actif', '2026-09-28'
    ]);
    bRow.height = 22;
    [5, 6, 7, 8, 11, 15, 17].forEach(c => { bRow.getCell(c).numFmt = currencyFormat; });

    const zCell = bRow.getCell(26);
    zCell.value = {
      formula: `IF(A${rIdx}="", "", (E${rIdx}+F${rIdx}+H${rIdx}+I${rIdx}+J${rIdx}+K${rIdx}+L${rIdx}+N${rIdx}+V${rIdx}+W${rIdx}) - ABS(O${rIdx}+P${rIdx}+Q${rIdx}+S${rIdx}) - ABS(G${rIdx}))`,
      result: b.paiementPrevu
    };
    zCell.numFmt = currencyFormat;
    zCell.font = { name: 'Segoe UI', size: 10, bold: true, color: { argb: 'FF065F46' } };
    zCell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFECFDF5' } };
    bRow.eachCell({ includeEmpty: true }, c => { c.border = borderThin; });
  });

  wsBolt.columns.forEach((col, idx) => {
    col.width = (idx === 0) ? 28 : (idx === 25) ? 24 : 16;
  });

  // --- ONGLET 3 : ANNUAIRE_FLOTTE ---
  const wsAnnuaire = wb.addWorksheet('ANNUAIRE_FLOTTE', { views: [{ showGridLines: true }] });
  const annuaireHeaders = [
    'N°', 'Nom chauffeur', 'Nom Uber (CSV)', 'Nom Bolt (CSV)', 
    'E-mail Uber', 'E-mail Bolt', 'Téléphone Uber', 'Téléphone Bolt', 'IBAN Qonto', 'Forfait Rattachement (€)'
  ];
  wsAnnuaire.addRow(annuaireHeaders);
  const annuaireHeaderRow = wsAnnuaire.getRow(1);
  annuaireHeaderRow.height = 28;
  annuaireHeaderRow.eachCell((cell) => {
    cell.font = fontHeader;
    cell.alignment = { horizontal: 'center', vertical: 'middle' };
    cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF' + NAVY } };
    cell.border = borderThin;
  });

  CHAUFFEURS_69.forEach((c, idx) => {
    const feeContract = getDriverContractFee(c.nomChauffeur);
    const row = wsAnnuaire.addRow([
      c.num, c.nomChauffeur, c.nomUber, c.nomBolt,
      c.mailUber, c.mailBolt, c.telUber, c.telBolt, c.iban, feeContract
    ]);
    row.height = 26;
    const isOdd = (idx % 2 === 1);
    const rowBg = isOdd ? 'FFF8FAFC' : 'FFFFFFFF';
    row.getCell(1).alignment = { horizontal: 'center', vertical: 'middle' };
    row.getCell(2).alignment = { horizontal: 'left', vertical: 'middle' };
    row.getCell(2).font = { name: 'Segoe UI', size: 10, bold: true, color: { argb: 'FF0F172A' } };
    row.getCell(3).alignment = { horizontal: 'left', vertical: 'middle' };
    row.getCell(4).alignment = { horizontal: 'left', vertical: 'middle' };
    row.getCell(5).alignment = { horizontal: 'left', vertical: 'middle' };
    row.getCell(6).alignment = { horizontal: 'left', vertical: 'middle' };
    row.getCell(7).alignment = { horizontal: 'center', vertical: 'middle' };
    row.getCell(7).numFmt = '@';
    row.getCell(8).alignment = { horizontal: 'center', vertical: 'middle' };
    row.getCell(8).numFmt = '@';
    row.getCell(9).alignment = { horizontal: 'left', vertical: 'middle' };
    row.getCell(9).font = fontMonospace;
    row.getCell(10).alignment = { horizontal: 'right', vertical: 'middle' };
    row.getCell(10).numFmt = currencyFormat;
    row.getCell(10).font = { name: 'Segoe UI', size: 10, bold: true, color: { argb: 'FF0F172A' } };
    row.eachCell({ includeEmpty: true }, cell => {
      cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: rowBg } };
      cell.border = borderThin;
    });
  });
  wsAnnuaire.columns.forEach((col, idx) => {
    col.width = (idx === 0) ? 6 : (idx === 1 || idx === 4 || idx === 5) ? 28 : (idx === 8) ? 34 : (idx === 9) ? 24 : 20;
  });


  // --- ONGLET 4 : DASHBOARD_SEMAINE ---
  const wsDash = wb.addWorksheet('DASHBOARD_SEMAINE', { views: [{ showGridLines: true }] });

  // Titres Banner
  wsDash.mergeCells('B1:S1');
  const titleCell = wsDash.getCell('B1');
  titleCell.value = 'FLEET OS — TABLEAU DE BORD HEBDOMADAIRE (SEMAINE DU 21 AU 27 SEPTEMBRE 2026)';
  titleCell.font = { name: 'Segoe UI', size: 16, bold: true, color: { argb: 'FFFFFFFF' } };
  titleCell.alignment = { horizontal: 'left', vertical: 'middle', indent: 1 };
  titleCell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF' + NAVY } };
  wsDash.getRow(1).height = 40;

  wsDash.mergeCells('B2:S2');
  const subCell = wsDash.getCell('B2');
  subCell.value = 'Conciliation Automatique Multi-Plateformes (Uber & Bolt) • 69 Chauffeurs • Rétrocession SEPA Qonto';
  subCell.font = { name: 'Segoe UI', size: 10, italic: true, color: { argb: 'FFE2E8F0' } };
  subCell.alignment = { horizontal: 'left', vertical: 'middle', indent: 1 };
  subCell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF' + SLATE } };
  wsDash.getRow(2).height = 22;
  wsDash.getRow(3).height = 12;

  // 4 CARTES KPI
  const lastDriverRow = 9 + calculations.length;
  const totRowIdx = lastDriverRow + 1;

  // KPI 1: TOTAL ENCAISSÉ
  wsDash.mergeCells('B4:D4');
  wsDash.getCell('B4').value = 'TOTAL RÉEL ENCAISSÉ (UBER+BOLT)';
  wsDash.getCell('B4').font = { name: 'Segoe UI', size: 9, bold: true, color: { argb: 'FF1E40AF' } };
  wsDash.getCell('B4').alignment = { horizontal: 'center', vertical: 'middle' };
  wsDash.getCell('B4').fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF' + KPI_BG_1 } };

  wsDash.mergeCells('B5:D5');
  wsDash.getCell('B5').value = { formula: 'SUM(M10:M' + lastDriverRow + ')', result: totalEncaisse };
  wsDash.getCell('B5').numFmt = currencyFormat;
  wsDash.getCell('B5').font = { name: 'Segoe UI', size: 16, bold: true, color: { argb: 'FF1E3A8A' } };
  wsDash.getCell('B5').alignment = { horizontal: 'center', vertical: 'middle' };
  wsDash.getCell('B5').fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF' + KPI_BG_1 } };

  // KPI 2: TOTAL ESPÈCES
  wsDash.mergeCells('F4:H4');
  wsDash.getCell('F4').value = 'SUIVI ESPÈCES TOTALES (UBER+BOLT)';
  wsDash.getCell('F4').font = { name: 'Segoe UI', size: 9, bold: true, color: { argb: 'FF92400E' } };
  wsDash.getCell('F4').alignment = { horizontal: 'center', vertical: 'middle' };
  wsDash.getCell('F4').fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFFEF3C7' } };

  wsDash.mergeCells('F5:H5');
  wsDash.getCell('F5').value = { formula: 'SUM(J10:J' + lastDriverRow + ')+SUM(L10:L' + lastDriverRow + ')', result: totalUberEspeces + totalBoltEspeces };
  wsDash.getCell('F5').numFmt = currencyFormat;
  wsDash.getCell('F5').font = { name: 'Segoe UI', size: 16, bold: true, color: { argb: 'FFB45309' } };
  wsDash.getCell('F5').alignment = { horizontal: 'center', vertical: 'middle' };
  wsDash.getCell('F5').fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFFEF3C7' } };

  // KPI 3: RATTACHEMENT
  wsDash.mergeCells('J4:L4');
  wsDash.getCell('J4').value = 'REVENUS RATTACHEMENT FLOTTE';
  wsDash.getCell('J4').font = { name: 'Segoe UI', size: 9, bold: true, color: { argb: 'FF5B21B6' } };
  wsDash.getCell('J4').alignment = { horizontal: 'center', vertical: 'middle' };
  wsDash.getCell('J4').fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF' + KPI_BG_2 } };

  wsDash.mergeCells('J5:L5');
  wsDash.getCell('J5').value = { formula: 'SUM(N10:N' + lastDriverRow + ')', result: totalRattachement };
  wsDash.getCell('J5').numFmt = currencyFormat;
  wsDash.getCell('J5').font = { name: 'Segoe UI', size: 16, bold: true, color: { argb: 'FF4C1D95' } };
  wsDash.getCell('J5').alignment = { horizontal: 'center', vertical: 'middle' };
  wsDash.getCell('J5').fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF' + KPI_BG_2 } };

  // KPI 4: NET À VIRER
  wsDash.mergeCells('N4:P4');
  wsDash.getCell('N4').value = 'TOTAL NET À VIRER (QONTO)';
  wsDash.getCell('N4').font = { name: 'Segoe UI', size: 9, bold: true, color: { argb: 'FF065F46' } };
  wsDash.getCell('N4').alignment = { horizontal: 'center', vertical: 'middle' };
  wsDash.getCell('N4').fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF' + KPI_BG_4 } };

  wsDash.mergeCells('N5:P5');
  wsDash.getCell('N5').value = { formula: 'SUM(R10:R' + lastDriverRow + ')', result: totalNet };
  wsDash.getCell('N5').numFmt = currencyFormat;
  wsDash.getCell('N5').font = { name: 'Segoe UI', size: 16, bold: true, color: { argb: 'FF047857' } };
  wsDash.getCell('N5').alignment = { horizontal: 'center', vertical: 'middle' };
  wsDash.getCell('N5').fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF' + KPI_BG_4 } };

  wsDash.getRow(4).height = 20;
  wsDash.getRow(5).height = 32;

  const applyCardBorders = (startCol, endCol) => {
    for (let r = 4; r <= 5; r++) {
      for (let c = startCol; c <= endCol; c++) {
        wsDash.getRow(r).getCell(c).border = borderThin;
      }
    }
  };
  applyCardBorders(2, 4);
  applyCardBorders(6, 8);
  applyCardBorders(10, 12);
  applyCardBorders(14, 16);

  // Groupes d'en-têtes Ligne 8
  const groupRanges = [
    { range: 'A8:B8', text: '👤 IDENTITÉ DU CHAUFFEUR', color: NAVY },
    { range: 'C8:D8', text: '🔗 IDENTIFIANTS PLATEFORMES', color: '1E293B' },
    { range: 'E8:H8', text: '📞 CONTACT & COORDONNÉES', color: '334155' },
    { range: 'I8:J8', text: '⚫ PRODUCTION UBER (€)', color: UBER_BLACK },
    { range: 'K8:L8', text: '🟢 PRODUCTION BOLT (€)', color: '15803D' },
    { range: 'M8:M8', text: '💰 TOTAL ENCAISSÉ (€)', color: '0E7490' },
    { range: 'N8:Q8', text: '⚙️ ARBITRAGES FLOTTE (€)', color: 'B45309' },
    { range: 'R8:R8', text: '🏦 CASE VIREMENT NET (€)', color: '047857' },
    { range: 'S8:S8', text: '💳 IBAN QONTO', color: QONTO_PURPLE }
  ];

  groupRanges.forEach(g => {
    if (g.range.split(':')[0] !== g.range.split(':')[1]) {
      wsDash.mergeCells(g.range);
    }
    const c = wsDash.getCell(g.range.split(':')[0]);
    c.value = g.text;
    c.font = { name: 'Segoe UI', size: 9, bold: true, color: { argb: 'FFFFFFFF' } };
    c.alignment = { horizontal: 'center', vertical: 'middle' };
    c.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF' + g.color } };
  });
  wsDash.getRow(8).height = 24;

  const row9 = wsDash.addRow(dashHeaders);
  row9.height = 32;
  row9.eachCell((cell, colNum) => {
    cell.font = fontHeader;
    cell.alignment = { horizontal: 'center', vertical: 'middle', wrapText: true };
    let bgCol = NAVY;
    if (colNum === 9) bgCol = UBER_BLACK;
    else if (colNum === 10) bgCol = '27272A';
    else if (colNum === 11) bgCol = '16A34A';
    else if (colNum === 12) bgCol = '14532D';
    else if (colNum === 13) bgCol = '0F766E';
    else if (colNum === 18) bgCol = '059669';
    else if (colNum === 19) bgCol = QONTO_PURPLE;

    cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF' + bgCol } };
    if (colNum === 10 || colNum === 12) {
      cell.font = { name: 'Segoe UI', size: 9, bold: true, color: { argb: 'FFFEF08A' } };
    }
    cell.border = borderThin;
  });

  // LIGNES 10 À 78 : 69 Chauffeurs
  calculations.forEach((c, idx) => {
    const rowIdx = 10 + idx;

    const rowData = [
      c.num,
      c.nomChauffeur,
      c.nomUber,
      c.nomBolt,
      c.mailUber,
      c.mailBolt,
      c.telUber,
      c.telBolt,
      // Col 9 (I): Virement Uber
      {
        formula: `IF(B${rowIdx}="", "", IF(OR(C${rowIdx}="", C${rowIdx}="—"), 0, IFNA(XLOOKUP(C${rowIdx}, DATA_UBER!A:A, DATA_UBER!Z:Z, 0), 0)))`,
        result: c.virementUber
      },
      // Col 10 (J): Espèces Uber
      {
        formula: `IF(B${rowIdx}="", "", IF(OR(C${rowIdx}="", C${rowIdx}="—"), 0, IFNA(XLOOKUP(C${rowIdx}, DATA_UBER!A:A, DATA_UBER!J:J, 0), 0)))`,
        result: c.especesUber
      },
      // Col 11 (K): Virement Bolt
      {
        formula: `IF(B${rowIdx}="", "", IF(OR(D${rowIdx}="", D${rowIdx}="—"), 0, IFNA(XLOOKUP(D${rowIdx}, DATA_BOLT!A:A, DATA_BOLT!Z:Z, 0), 0)))`,
        result: c.virementBolt
      },
      // Col 12 (L): Espèces Bolt
      {
        formula: `IF(B${rowIdx}="", "", IF(OR(D${rowIdx}="", D${rowIdx}="—"), 0, IFNA(XLOOKUP(D${rowIdx}, DATA_BOLT!A:A, DATA_BOLT!G:G, 0), 0)))`,
        result: c.especesBolt
      },
      // Col 13 (M): Total Encaissé
      {
        formula: `IF(B${rowIdx}="", "", I${rowIdx} + K${rowIdx})`,
        result: c.encaisse
      },
      // Col 14 (N): Rattachement dynamique selon le tarif contractuel de chaque chauffeur dans ANNUAIRE_FLOTTE (40, 50, 60, 70...)
      {
        formula: `IF(B${rowIdx}="", "", IF(OR(I${rowIdx}<>0, K${rowIdx}<>0, J${rowIdx}<>0, L${rowIdx}<>0), IFNA(XLOOKUP(B${rowIdx}, ANNUAIRE_FLOTTE!B:B, ANNUAIRE_FLOTTE!J:J, 50), 50), 0))`,
        result: c.rattachement
      },
      c.charges,
      c.bonus,
      c.malus,
      // Col 18 (R): Net à virer
      {
        formula: `IF(B${rowIdx}="", "", M${rowIdx} - N${rowIdx} - O${rowIdx} - Q${rowIdx} + P${rowIdx})`,
        result: c.net
      },
      c.iban
    ];

    const r = wsDash.addRow(rowData);
    r.height = 28;

    const isOdd = (idx % 2 === 1);
    const rowBg = isOdd ? 'FFF8FAFC' : 'FFFFFFFF';

    r.getCell(1).alignment = { horizontal: 'center', vertical: 'middle' };
    r.getCell(2).alignment = { horizontal: 'left', vertical: 'middle' };
    r.getCell(2).font = { name: 'Segoe UI', size: 10, bold: true, color: { argb: (c.net > 1000) ? 'FF047857' : 'FF0F172A' } };
    r.getCell(3).alignment = { horizontal: 'left', vertical: 'middle' };
    r.getCell(4).alignment = { horizontal: 'left', vertical: 'middle' };
    r.getCell(5).alignment = { horizontal: 'left', vertical: 'middle' };
    r.getCell(5).font = { name: 'Segoe UI', size: 9.5, color: { argb: 'FF475569' } };
    r.getCell(6).alignment = { horizontal: 'left', vertical: 'middle' };
    r.getCell(6).font = { name: 'Segoe UI', size: 9.5, color: { argb: 'FF475569' } };
    r.getCell(7).alignment = { horizontal: 'center', vertical: 'middle' };
    r.getCell(7).font = { name: 'Segoe UI', size: 9.5, color: { argb: 'FF475569' } };
    r.getCell(7).numFmt = '@';
    r.getCell(8).alignment = { horizontal: 'center', vertical: 'middle' };
    r.getCell(8).font = { name: 'Segoe UI', size: 9.5, color: { argb: 'FF475569' } };
    r.getCell(8).numFmt = '@';

    for (let col = 1; col <= 8; col++) {
      r.getCell(col).fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: rowBg } };
    }

    for (let col = 9; col <= 18; col++) {
      const cell = r.getCell(col);
      cell.numFmt = currencyFormat;
      cell.alignment = { horizontal: 'right', vertical: 'middle' };
      cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: rowBg } };
    }

    // Visuel espèces (Cols 10 et 12)
    [10, 12].forEach(colIdx => {
      const cell = r.getCell(colIdx);
      cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFFEF3C7' } };
      cell.font = { name: 'Segoe UI', size: 10, bold: true, color: { argb: 'FFB45309' } };
    });

    // Net à virer (Col 18)
    const nC = r.getCell(18);
    nC.font = { name: 'Segoe UI', size: 11, bold: true, color: { argb: c.net < 0 ? 'FFDC2626' : (c.net > 0 ? 'FF047857' : 'FF64748B') } };
    nC.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: c.net < 0 ? 'FFFEE2E2' : (c.net > 0 ? 'FFECFDF5' : rowBg) } };

    // IBAN (Col 19)
    const iC = r.getCell(19);
    iC.font = (c.nomChauffeur === 'Siré Cyril Niakaté') ? fontMonospaceBold : fontMonospace;
    iC.alignment = { horizontal: 'left', vertical: 'middle' };
    iC.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: rowBg } };

    r.eachCell({ includeEmpty: true }, cell => { cell.border = borderThin; });
  });

  // LIGNE DE TOTAL CONSOLIDÉ TOUT EN BAS (APRÈS TOUS LES CHAUFFEURS)
  wsDash.mergeCells(`A${totRowIdx}:H${totRowIdx}`);
  const totLabel = wsDash.getCell(`A${totRowIdx}`);
  totLabel.value = 'TOTAL CONSOLIDÉ SEMAINE (FLOTTE)';
  totLabel.font = { name: 'Segoe UI', size: 10, bold: true, color: { argb: 'FFFFFFFF' } };
  totLabel.alignment = { horizontal: 'right', vertical: 'middle', indent: 1 };
  totLabel.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF' + NAVY } };

  const totRow = wsDash.getRow(totRowIdx);
  totRow.height = 32;

  const colLetters = ['I', 'J', 'K', 'L', 'M', 'N', 'O', 'P', 'Q', 'R'];
  const expectedTotals = [
    totalUberVirement, totalUberEspeces, totalBoltVirement, totalBoltEspeces,
    totalEncaisse, totalRattachement, 0.00, 0.00, 0.00, totalNet
  ];

  colLetters.forEach((letter, i) => {
    const colNum = 9 + i;
    const cell = totRow.getCell(colNum);
    cell.value = {
      formula: `SUM(${letter}10:${letter}${lastDriverRow})`,
      result: expectedTotals[i]
    };
    cell.numFmt = currencyFormat;
    cell.alignment = { horizontal: 'right', vertical: 'middle' };
    cell.font = { name: 'Segoe UI', size: 11, bold: true, color: { argb: colNum === 18 ? 'FF047857' : (colNum === 10 || colNum === 12) ? 'FFB45309' : 'FFFFFFFF' } };
    cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: colNum === 18 ? 'FFECFDF5' : (colNum === 10 || colNum === 12) ? 'FFFEF3C7' : 'FF' + SLATE } };
  });

  totRow.getCell(19).fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF' + NAVY } };

  for (let col = 1; col <= 19; col++) {
    const cell = totRow.getCell(col);
    cell.border = {
      top: { style: 'medium', color: { argb: 'FF0F172A' } },
      bottom: { style: 'double', color: { argb: 'FF0F172A' } },
      left: { style: 'thin', color: { argb: 'FFCBD5E1' } },
      right: { style: 'thin', color: { argb: 'FFCBD5E1' } }
    };
  }

  // Largeurs de colonnes Dashboard
  wsDash.getColumn(1).width = 7;
  wsDash.getColumn(2).width = 28;
  wsDash.getColumn(3).width = 24;
  wsDash.getColumn(4).width = 24;
  wsDash.getColumn(5).width = 28;
  wsDash.getColumn(6).width = 28;
  wsDash.getColumn(7).width = 18;
  wsDash.getColumn(8).width = 18;
  wsDash.getColumn(9).width = 22;
  wsDash.getColumn(10).width = 18;
  wsDash.getColumn(11).width = 22;
  wsDash.getColumn(12).width = 18;
  wsDash.getColumn(13).width = 25;
  wsDash.getColumn(14).width = 16;
  wsDash.getColumn(15).width = 14;
  wsDash.getColumn(16).width = 14;
  wsDash.getColumn(17).width = 14;
  wsDash.getColumn(18).width = 28;
  wsDash.getColumn(19).width = 34;

  // --- ONGLET 5 : EXPORT_QONTO ---
  const wsQonto = wb.addWorksheet('EXPORT_QONTO', { views: [{ showGridLines: true }] });
  const qontoHeaders = ['beneficiary_name', 'iban', 'amount', 'currency', 'reference'];
  wsQonto.addRow(qontoHeaders);
  const qontoHeaderRow = wsQonto.getRow(1);
  qontoHeaderRow.height = 28;
  qontoHeaderRow.eachCell(cell => {
    cell.font = fontHeader;
    cell.alignment = { horizontal: 'center', vertical: 'middle' };
    cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF' + QONTO_PURPLE } };
    cell.border = borderThin;
  });

  calculations.forEach((c, idx) => {
    const r = 10 + idx;
    const qR = wsQonto.addRow([
      { formula: `IF(DASHBOARD_SEMAINE!B${r}="", "", DASHBOARD_SEMAINE!B${r})`, result: c.nomChauffeur },
      { formula: `IF(DASHBOARD_SEMAINE!B${r}="", "", DASHBOARD_SEMAINE!S${r})`, result: c.iban },
      { formula: `IF(OR(DASHBOARD_SEMAINE!B${r}="", DASHBOARD_SEMAINE!R${r}<=0), "", ROUND(DASHBOARD_SEMAINE!R${r}, 2))`, result: c.net > 0 ? c.net : '' },
      { formula: `IF(DASHBOARD_SEMAINE!B${r}="", "", "EUR")`, result: 'EUR' },
      { formula: `IF(DASHBOARD_SEMAINE!B${r}="", "", "Virement Semaine - " & DASHBOARD_SEMAINE!B${r})`, result: `Virement Semaine - ${c.nomChauffeur}` }
    ]);
    qR.height = 24;
    qR.getCell(1).font = { name: 'Segoe UI', size: 10, bold: true };
    qR.getCell(2).font = (c.nomChauffeur === 'Siré Cyril Niakaté') ? fontMonospaceBold : fontMonospace;
    qR.getCell(3).numFmt = '0.00';
    qR.getCell(3).alignment = { horizontal: 'right', vertical: 'middle' };
    qR.getCell(4).alignment = { horizontal: 'center', vertical: 'middle' };
    qR.eachCell({ includeEmpty: true }, cell => { cell.border = borderThin; });
  });

  wsQonto.getColumn(1).width = 28;
  wsQonto.getColumn(2).width = 34;
  wsQonto.getColumn(3).width = 16;
  wsQonto.getColumn(4).width = 12;
  wsQonto.getColumn(5).width = 36;

  // --- ONGLET 6 : ARCHIVE FIGÉE S39 ---
  const wsArchive = wb.addWorksheet('📁 S39 (21 au 27 Sept)', { views: [{ showGridLines: true }] });
  wsArchive.mergeCells('B1:S1');
  const arcTitle = wsArchive.getCell('B1');
  arcTitle.value = 'FLEET OS — ARCHIVE OFFICIELLE SEMAINE 39 (69 CHAUFFEURS)';
  arcTitle.font = { name: 'Segoe UI', size: 16, bold: true, color: { argb: 'FFFFFFFF' } };
  arcTitle.alignment = { horizontal: 'left', vertical: 'middle', indent: 1 };
  arcTitle.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF475569' } };
  wsArchive.getRow(1).height = 40;

  wsArchive.mergeCells('B2:S2');
  const arcSub = wsArchive.getCell('B2');
  arcSub.value = 'Archive Figée — Période travaillée : 21 au 27 Septembre 2026 • Virement bancaire Qonto exécuté le : 28/09/2026';
  arcSub.font = { name: 'Segoe UI', size: 10, italic: true, color: { argb: 'FFE2E8F0' } };
  arcSub.alignment = { horizontal: 'left', vertical: 'middle', indent: 1 };
  arcSub.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF334155' } };
  wsArchive.getRow(2).height = 22;

  // Copie conforme figée des valeurs du Dashboard dans l'archive
  for (let r = 4; r <= 79; r++) {
    const srcRow = wsDash.getRow(r);
    const destRow = wsArchive.getRow(r);
    destRow.height = srcRow.height;
    for (let col = 1; col <= 19; col++) {
      const srcCell = srcRow.getCell(col);
      const destCell = destRow.getCell(col);
      if (srcCell.value && typeof srcCell.value === 'object' && 'result' in srcCell.value) {
        destCell.value = srcCell.value.result;
      } else {
        destCell.value = srcCell.value;
      }
      destCell.numFmt = srcCell.numFmt;
      destCell.font = srcCell.font;
      destCell.fill = srcCell.fill;
      destCell.alignment = srcCell.alignment;
      destCell.border = srcCell.border;
    }
  }

  // Copier les fusions de cellules
  ['B4:D4', 'B5:D5', 'F4:H4', 'F5:H5', 'J4:L4', 'J5:L5', 'N4:P4', 'N5:P5', 'A8:B8', 'C8:D8', 'E8:H8', 'I8:J8', 'K8:L8', 'N8:Q8', 'A79:H79'].forEach(range => {
    try { wsArchive.mergeCells(range); } catch (e) {}
  });

  for (let colIdx = 1; colIdx <= 19; colIdx++) {
    wsArchive.getColumn(colIdx).width = wsDash.getColumn(colIdx).width;
  }

  // Sauvegarde
  const xlsxPath = 'FLEET_OS_Nordrive_Reconciliation_Qonto.xlsx';
  await wb.xlsx.writeFile(xlsxPath);
  console.log(`✅ Fichier XLSX complet généré : ${xlsxPath}`);
}

buildFullExcelWorkbook().then(() => {
  console.log('\n========================================================================');
  console.log('🎉 TOUS LES CALCULS ET FICHIERS SONT PRÊTS AVEC SUCCÈS !');
  console.log('========================================================================\n');
}).catch(err => {
  console.error('❌ Erreur lors de la génération Excel :', err);
});

const { app, BrowserWindow, ipcMain, dialog } = require('electron');
const path = require('path');
const fs = require('fs');
const XLSX = require('xlsx');
const docx = require('docx');

function getDBPath() {
  const portable = path.join(__dirname, 'aidat_database.json');
  if (fs.existsSync(portable)) return portable;
  const userData = app.getPath('userData');
  return path.join(userData, 'aidat_database.json');
}

function getSeedPath() {
  if (process.resourcesPath) {
    const seeded = path.join(process.resourcesPath, 'aidat_database.json');
    if (fs.existsSync(seeded)) return seeded;
  }
  return null;
}

let DB_PATH;
let db;
let mainWindow;

function loadDB() {
  DB_PATH = getDBPath();
  if (fs.existsSync(DB_PATH)) {
    db = JSON.parse(fs.readFileSync(DB_PATH, 'utf-8'));
  } else {
    const seed = getSeedPath();
    if (seed) {
      db = JSON.parse(fs.readFileSync(seed, 'utf-8'));
    } else {
      db = buildDefaultDB();
    }
    DB_PATH = path.join(app.getPath('userData'), 'aidat_database.json');
    saveDB();
  }
  if (!db.giderItems) db.giderItems = {};
  if (!db.siteGiderleri) db.siteGiderleri = {};
}

function saveDB() {
  fs.writeFileSync(DB_PATH, JSON.stringify(db, null, 2), 'utf-8');
}

function buildDefaultDB() {
  const residents = [
    {sira:1,name:"AYHAN ÖZYILMAZ",flat:"C/1",block:"C",m2:83},
    {sira:2,name:"OSMAN ÖZTÜRK",flat:"C/2",block:"C",m2:109},
    {sira:3,name:"FEYZULLAH YALÇIN",flat:"C/3",block:"C",m2:109},
    {sira:4,name:"İSMAİL ÇAĞLAR",flat:"C/4",block:"C",m2:109},
    {sira:5,name:"ARİF KARA",flat:"C/5",block:"C",m2:109},
    {sira:6,name:"EMRE KAYA",flat:"C/6",block:"C",m2:109},
    {sira:7,name:"HARUN HASANOĞLU",flat:"C/7",block:"C",m2:113},
    {sira:8,name:"TUNCAY DELİGÖZ",flat:"C/8",block:"C",m2:113},
    {sira:9,name:"ŞEYDA NUR ÖZDEMİR",flat:"C/9",block:"C",m2:109},
    {sira:10,name:"MUSTAFA KAÇAR",flat:"C/10",block:"C",m2:113},
    {sira:11,name:"MEHMET KARATEPE",flat:"C/11",block:"C",m2:109},
    {sira:12,name:"YUSUF KAYA",flat:"C/12",block:"C",m2:109},
    {sira:13,name:"İSMAİL KELEŞ",flat:"C/13",block:"C",m2:113},
    {sira:14,name:"MEHMET YILDIRIM",flat:"C/14",block:"C",m2:109},
    {sira:15,name:"FERHAT OKÇU",flat:"C/15",block:"C",m2:109},
    {sira:16,name:"İSA GÜRAY",flat:"C/16",block:"C",m2:113},
    {sira:17,name:"SATILMIŞ GÖKBAYIR",flat:"C/17",block:"C",m2:109},
    {sira:18,name:"HASAN AYAN",flat:"C/18",block:"C",m2:83},
    {sira:19,name:"KANAN KİBAR",flat:"D/1",block:"D",m2:83},
    {sira:20,name:"İRFAN AKBAŞ",flat:"D/2",block:"D",m2:83},
    {sira:21,name:"MURAT YEŞİLKAYA",flat:"D/3",block:"D",m2:109},
    {sira:22,name:"HASAN ÇAPAR",flat:"D/4",block:"D",m2:109},
    {sira:23,name:"SATI DURMUŞ",flat:"D/5",block:"D",m2:109},
    {sira:24,name:"CÜNEYT GÜL",flat:"D/6",block:"D",m2:109},
    {sira:25,name:"BİLAL YALÇIN",flat:"D/7",block:"D",m2:109},
    {sira:26,name:"SATILMIŞ YETİŞ",flat:"D/8",block:"D",m2:109},
    {sira:27,name:"SİNAN ARAL",flat:"D/9",block:"D",m2:109},
    {sira:28,name:"MUSTAFA ARMAĞAN",flat:"D/10",block:"D",m2:109},
    {sira:29,name:"GÜLENDER KELEŞ",flat:"D/11",block:"D",m2:109},
    {sira:30,name:"İBRAHİM ÇAPAR",flat:"D/12",block:"D",m2:109},
    {sira:31,name:"SERVET DERVİŞ",flat:"D/13",block:"D",m2:109},
    {sira:32,name:"KERİM COŞKUN",flat:"D/14",block:"D",m2:109},
    {sira:33,name:"ASLAN DİRİCAN",flat:"D/15",block:"D",m2:109},
    {sira:34,name:"ŞABAN AKGÜNDÜZ",flat:"D/16",block:"D",m2:109},
    {sira:35,name:"HİKMET SANCAR",flat:"D/17",block:"D",m2:109},
    {sira:36,name:"TECELLİ KURTOĞLU",flat:"D/18",block:"D",m2:109},
    {sira:37,name:"SEVCAN ASAR",flat:"D/11-A",block:"D",m2:29},
    {sira:38,name:"FAZLI AKYILDIZ",flat:"E/1",block:"E",m2:83},
    {sira:39,name:"NEFİSE VAZVAZ",flat:"E/2",block:"E",m2:83},
    {sira:40,name:"FAHRETTİN GÜVENDİ",flat:"E/3",block:"E",m2:113},
    {sira:41,name:"MURAT AÇIKGÖZ",flat:"E/4",block:"E",m2:109},
    {sira:42,name:"AHMET YALÇIN",flat:"E/5",block:"E",m2:109},
    {sira:43,name:"İBRAHİM ÖDEMİŞ",flat:"E/6",block:"E",m2:109},
    {sira:44,name:"ALİ ÖZYILMAZ",flat:"E/7",block:"E",m2:109},
    {sira:45,name:"HÜSEYİN KANDEMİR",flat:"E/8",block:"E",m2:109},
    {sira:46,name:"HÜSEYİN AKYILDIZ",flat:"E/9",block:"E",m2:113},
    {sira:47,name:"MUSTAFA KARALI",flat:"E/10",block:"E",m2:113},
    {sira:48,name:"İSMAİL KULA",flat:"E/11",block:"E",m2:109},
    {sira:49,name:"YAKUP KARTAL",flat:"E/12",block:"E",m2:109},
    {sira:50,name:"YÜKSEL ÖKSÜZ",flat:"E/13",block:"E",m2:109},
    {sira:51,name:"MUHİTTİN ÇAPAR",flat:"E/14",block:"E",m2:109},
    {sira:52,name:"EMİNE DEMİRÖZ",flat:"E/15",block:"E",m2:113},
    {sira:53,name:"İBRAHİM UÇAR",flat:"E/16",block:"E",m2:109},
    {sira:54,name:"MESUT MEMİŞ",flat:"E/17",block:"E",m2:113},
    {sira:55,name:"SATILMIŞ HOTOĞLU",flat:"E/18",block:"E",m2:109}
  ];
  const monthNames = [
    "TEMMUZ 2025","AĞUSTOS 2025","EYLÜL 2025","EKİM 2025","kasım 2025",
    "ARALIK 2025","OCAK 2026","ŞUBAT 2026","Mart 2026","nisan 2026",
    "mayıs 2026","HAZİRAN 2026","TEMMUZ 2026"
  ];
  const months = {};
  monthNames.forEach((mName, mIdx) => {
    months[mName] = residents.map((r, i) => ({
      flat:r.flat, name:r.name, yakit:0, genel:0, gider:0, toplam:0,
      tarih:"", makbuz: mIdx * residents.length + (i + 1)
    }));
  });
  return { site_info:{name:"SABIR SİTESİ C-D-E BLOKLARI YÖNETİCİLİĞİ",address:"Buğday Pazarı Mah. Menekşe Sk. ÇANKIRI",seri:"SERİ-A",banka:"QNB FİNANSBANK",iban_no:"TR61 0011 1000 0000 0081 9966 14",note:"SON ÖDEME TARİHİ HER AYIN 10'udur.",hesap_sahibi:"HESAP SAHİBİ; SABIR SİTESİ C-D-E BLOKLARI YÖNETİCİLİĞİ"}, residents, months, giderItems:{} };
}

function createWindow() {
  mainWindow = new BrowserWindow({
    width:1250, height:880, minWidth:1000, minHeight:700,
    title:'Aidat Yönetim — kuisoft.com',
    icon: path.join(__dirname, 'icon.ico'),
    autoHideMenuBar: true,
    webPreferences:{ preload:path.join(__dirname,'preload.js'), contextIsolation:true, nodeIntegration:false }
  });
  mainWindow.removeMenu();
  mainWindow.loadFile('index.html');
}

ipcMain.handle('get-data', () => db);
ipcMain.handle('save-data', (_, newDb) => { db = newDb; saveDB(); return true; });

ipcMain.handle('print-pdf', async (_, monthName) => {
  const safeName = (monthName || 'makbuz').replace(/[^a-zA-ZğüşıöçĞÜŞİÖÇ0-9 _-]/g, '');
  const result = await dialog.showSaveDialog(mainWindow, {
    title:'PDF Olarak Kaydet', defaultPath:path.join(__dirname, `Makbuzlar_${safeName}.pdf`),
    filters:[{name:'PDF',extensions:['pdf']}]
  });
  if (result.canceled) return {success:false};
  const pdfData = await mainWindow.webContents.printToPDF({
    printBackground:true, landscape:false, pageSize:'A4',
    margins:{top:0.3,bottom:0.3,left:0.3,right:0.3}
  });
  fs.writeFileSync(result.filePath, pdfData);
  return {success:true, path:result.filePath};
});

ipcMain.handle('backup', async () => {
  const now = new Date();
  const ds = `${now.getFullYear()}${String(now.getMonth()+1).padStart(2,'0')}${String(now.getDate()).padStart(2,'0')}_${String(now.getHours()).padStart(2,'0')}${String(now.getMinutes()).padStart(2,'0')}`;
  const result = await dialog.showSaveDialog(mainWindow, {
    title:'Veritabanını Yedekle', defaultPath:path.join(__dirname, `yedek_${ds}.json`),
    filters:[{name:'JSON',extensions:['json']}]
  });
  if (result.canceled) return {success:false};
  fs.writeFileSync(result.filePath, JSON.stringify(db, null, 2), 'utf-8');
  return {success:true, path:result.filePath};
});

ipcMain.handle('export-excel', async (_, monthName) => {
  const safeName = (monthName || 'aidat').replace(/[^a-zA-ZğüşıöçĞÜŞİÖÇ0-9 _-]/g, '');
  const result = await dialog.showSaveDialog(mainWindow, {
    title:'Excel Olarak Kaydet', defaultPath:path.join(__dirname, `Aidat_${safeName}.xlsx`),
    filters:[{name:'Excel',extensions:['xlsx']}]
  });
  if (result.canceled) return {success:false};

  const wbook = XLSX.utils.book_new();
  const records = db.months[monthName] || [];
  const rows = [['SIRA','ADI SOYADI','BLOK/DAİRE','M2','YAKIT BEDELİ','GENEL','GİDER','TOPLAM','TARİH','MAKBUZ NU.']];
  records.forEach((r,i) => {
    const res = db.residents.find(x => x.flat === r.flat);
    rows.push([i+1, r.name, r.flat, res?res.m2:109, r.yakit, r.genel, r.gider, r.toplam, r.tarih, r.makbuz]);
  });
  const totY = records.reduce((s,r)=>s+r.yakit,0);
  const totG = records.reduce((s,r)=>s+r.genel,0);
  const totGi = records.reduce((s,r)=>s+r.gider,0);
  const totT = records.reduce((s,r)=>s+r.toplam,0);
  rows.push(['TOPLAM','','','',totY,totG,totGi,totT,'','']);
  const ws1 = XLSX.utils.aoa_to_sheet(rows);
  ws1['!cols'] = [{wch:6},{wch:22},{wch:12},{wch:6},{wch:14},{wch:10},{wch:10},{wch:14},{wch:12},{wch:10}];
  XLSX.utils.book_append_sheet(wbook, ws1, monthName.substring(0,31));

  const resRows = [['SIRA','ADI SOYADI','BLOK/DAİRE','BLOK','M2']];
  db.residents.forEach(r => resRows.push([r.sira, r.name, r.flat, r.block, r.m2]));
  const ws2 = XLSX.utils.aoa_to_sheet(resRows);
  ws2['!cols'] = [{wch:6},{wch:22},{wch:12},{wch:8},{wch:6}];
  XLSX.utils.book_append_sheet(wbook, ws2, 'Sakinler');

  XLSX.writeFile(wbook, result.filePath);
  return {success:true, path:result.filePath};
});

ipcMain.handle('export-all-excel', async () => {
  const result = await dialog.showSaveDialog(mainWindow, {
    title:'Tüm Verileri Excel Olarak Kaydet', defaultPath:path.join(__dirname, 'Aidat_Tum_Veriler.xlsx'),
    filters:[{name:'Excel',extensions:['xlsx']}]
  });
  if (result.canceled) return {success:false};

  const wbook = XLSX.utils.book_new();
  Object.keys(db.months).forEach(mName => {
    const records = db.months[mName];
    const rows = [['SIRA','ADI SOYADI','BLOK/DAİRE','M2','YAKIT','GENEL','GİDER','TOPLAM','TARİH','MAKBUZ']];
    records.forEach((r,i) => {
      const res = db.residents.find(x => x.flat === r.flat);
      rows.push([i+1, r.name, r.flat, res?res.m2:109, r.yakit, r.genel, r.gider, r.toplam, r.tarih, r.makbuz]);
    });
    const ws = XLSX.utils.aoa_to_sheet(rows);
    ws['!cols'] = [{wch:5},{wch:22},{wch:11},{wch:5},{wch:10},{wch:8},{wch:8},{wch:10},{wch:11},{wch:8}];
    const sheetName = mName.substring(0,31);
    XLSX.utils.book_append_sheet(wbook, ws, sheetName);
  });

  const resRows = [['SIRA','ADI SOYADI','BLOK/DAİRE','BLOK','M2']];
  db.residents.forEach(r => resRows.push([r.sira, r.name, r.flat, r.block, r.m2]));
  const ws2 = XLSX.utils.aoa_to_sheet(resRows);
  XLSX.utils.book_append_sheet(wbook, ws2, 'Sakinler');

  XLSX.writeFile(wbook, result.filePath);
  return {success:true, path:result.filePath};
});

// Word export — makbuzları .docx olarak kaydet
ipcMain.handle('export-word', async (_, monthName, blockFilter) => {
  const safeName = (monthName || 'makbuz').replace(/[^a-zA-ZğüşıöçĞÜŞİÖÇ0-9 _-]/g, '');
  const result = await dialog.showSaveDialog(mainWindow, {
    title:'Word Olarak Kaydet', defaultPath:path.join(__dirname, `Makbuzlar_${safeName}.docx`),
    filters:[{name:'Word',extensions:['docx']}]
  });
  if (result.canceled) return {success:false};

  const { Document, Paragraph, Table, TableRow, TableCell, TextRun, WidthType, AlignmentType,
          BorderStyle, PageBreak } = docx;

  const records = db.months[monthName] || [];
  const giderItems = (db.giderItems && db.giderItems[monthName]) || [{name:'Gider',amount:0}];
  const filtered = records.filter(r => {
    if (blockFilter === 'ALL') return true;
    const res = db.residents.find(x => x.flat === r.flat);
    return res && res.block === blockFilter;
  });

  const thinBorder = { style: BorderStyle.SINGLE, size: 1 };
  const thickBorder = { style: BorderStyle.SINGLE, size: 3 };

  function fmMoney(v) { return '₺' + (v||0).toLocaleString('tr-TR', {minimumFractionDigits:2, maximumFractionDigits:2}); }

  function buildReceiptElements(r) {
    const isSp = r.flat === 'C/18';
    const gidTotal = isSp ? 0 : giderItems.reduce((s,g) => s + (g.amount||0), 0);
    const toplam = (r.yakit||0) + (r.genel||0) + gidTotal;

    const feeRows = [];
    feeRows.push({label:'Yakıt Bedeli:', val: fmMoney(r.yakit)});
    feeRows.push({label:'Genel:', val: fmMoney(r.genel)});
    if (isSp) {
      feeRows.push({label:'Gider:', val: fmMoney(0)});
    } else {
      giderItems.forEach(g => feeRows.push({label:(g.name||'Gider')+':', val: fmMoney(g.amount)}));
    }

    const tableRows = feeRows.map(f =>
      new TableRow({ children: [
        new TableCell({ children:[new Paragraph({children:[new TextRun({text:f.label, font:'Calibri', size:22})]})], width:{size:50, type:WidthType.PERCENTAGE}, borders:{top:thinBorder,bottom:thinBorder,left:thickBorder,right:thinBorder} }),
        new TableCell({ children:[new Paragraph({children:[new TextRun({text:f.val, font:'Calibri', size:22})], alignment:AlignmentType.CENTER})], width:{size:50, type:WidthType.PERCENTAGE}, borders:{top:thinBorder,bottom:thinBorder,left:thinBorder,right:thickBorder} })
      ]})
    );
    tableRows.push(new TableRow({ children: [
      new TableCell({ children:[new Paragraph({children:[new TextRun({text:'TOPLAM:', font:'Calibri', size:24, bold:true})], alignment:AlignmentType.CENTER})], borders:{top:thinBorder,bottom:thickBorder,left:thickBorder,right:thinBorder} }),
      new TableCell({ children:[new Paragraph({children:[new TextRun({text:fmMoney(toplam), font:'Calibri', size:24, bold:true})], alignment:AlignmentType.CENTER})], borders:{top:thinBorder,bottom:thickBorder,left:thinBorder,right:thickBorder} })
    ]}));

    return [
      new Paragraph({ alignment:AlignmentType.CENTER, spacing:{after:40}, children:[
        new TextRun({text:'SABIR SİTESİ', font:'Calibri', size:22, bold:false}), new TextRun({text:'\nC-D-E BLOKLARI YÖNETİCİLİĞİ', font:'Calibri', size:22, break:1}),
        new TextRun({text:'\nBuğday Pazarı Mah. Menekşe Sk. ÇANKIRI', font:'Calibri', size:20, break:1})
      ]}),
      new Paragraph({ alignment:AlignmentType.RIGHT, spacing:{after:20}, children:[
        new TextRun({text:'SERİ-A', font:'Calibri', size:22, bold:true}),
        new TextRun({text:'\nSeri No: ', font:'Calibri', size:22, bold:true, break:1}), new TextRun({text:String(r.makbuz), font:'Calibri', size:22}),
        new TextRun({text:'\nTarih: ', font:'Calibri', size:22, bold:true, break:1}), new TextRun({text:r.tarih||'—', font:'Calibri', size:22})
      ]}),
      new Paragraph({ alignment:AlignmentType.CENTER, spacing:{before:80, after:80}, children:[
        new TextRun({text:'ÖDEME MAKBUZU', font:'Calibri', size:24, bold:true, underline:{}})
      ]}),
      new Table({ rows: tableRows, width:{size:80, type:WidthType.PERCENTAGE}, alignment:AlignmentType.CENTER }),
      new Paragraph({ spacing:{before:100}, children:[
        new TextRun({text:'Daire Sahibinin:', font:'Calibri', size:22, bold:true, underline:{}})
      ]}),
      new Paragraph({ indent:{left:400}, children:[
        new TextRun({text:'Adı Soyadı:  ', font:'Calibri', size:22, bold:true}),
        new TextRun({text:r.name, font:'Calibri', size:22, bold:true})
      ]}),
      new Paragraph({ indent:{left:400}, children:[
        new TextRun({text:'Blok/Daire No:  ', font:'Calibri', size:22, bold:true}),
        new TextRun({text:r.flat, font:'Calibri', size:22, bold:true})
      ]}),
      new Paragraph({ alignment:AlignmentType.CENTER, spacing:{before:80}, children:[
        new TextRun({text:"SON ÖDEME TARİHİ HER AYIN 10'udur.", font:'Calibri', size:20, bold:true, underline:{}, color:'CC0000'})
      ]}),
      new Paragraph({ alignment:AlignmentType.CENTER, spacing:{after:0}, children:[
        new TextRun({text: (db.site_info && db.site_info.hesap_sahibi) || 'HESAP SAHİBİ; SABIR SİTESİ C-D-E BLOKLARI YÖNETİCİLİĞİ', font:'Calibri', size:18})
      ]}),
      new Paragraph({ alignment:AlignmentType.CENTER, spacing:{after:0}, children:[
        new TextRun({text: ((db.site_info && db.site_info.banka) || 'QNB FİNANSBANK') + ' IBAN; ' + ((db.site_info && db.site_info.iban_no) || 'TR61 0011 1000 0000 0081 9966 14'), font:'Calibri', size:18, bold:true})
      ]}),
      new Paragraph({ alignment:AlignmentType.CENTER, spacing:{after:40}, children:[
        new TextRun({text:'kuisoft.com', font:'Calibri', size:14, color:'AAAAAA'})
      ]})
    ];
  }

  const sections = [];
  filtered.forEach((r, idx) => {
    const elems = buildReceiptElements(r);
    const elems2 = buildReceiptElements(r);
    const children = [
      ...elems,
      new Paragraph({ spacing:{before:200, after:200}, children:[new TextRun({text:'─'.repeat(60), font:'Calibri', size:16, color:'999999'})], alignment:AlignmentType.CENTER }),
      ...elems2
    ];
    if (idx < filtered.length - 1) {
      children.push(new Paragraph({ children:[new PageBreak()] }));
    }
    sections.push(...children);
  });

  const doc = new Document({
    sections: [{ children: sections }]
  });

  const buffer = await docx.Packer.toBuffer(doc);
  fs.writeFileSync(result.filePath, buffer);
  return {success:true, path:result.filePath};
});

// Yedeği Excel olarak kaydet
ipcMain.handle('backup-excel', async () => {
  const now = new Date();
  const ds = `${now.getFullYear()}${String(now.getMonth()+1).padStart(2,'0')}${String(now.getDate()).padStart(2,'0')}`;
  const result = await dialog.showSaveDialog(mainWindow, {
    title:'Yedeği Excel Olarak Kaydet', defaultPath:path.join(__dirname, `Yedek_${ds}.xlsx`),
    filters:[{name:'Excel',extensions:['xlsx']}]
  });
  if (result.canceled) return {success:false};

  const wbook = XLSX.utils.book_new();

  // Her ay için sayfa
  Object.keys(db.months).forEach(mName => {
    const records = db.months[mName];
    const rows = [['SIRA','ADI SOYADI','BLOK/DAİRE','M2','YAKIT BEDELİ','GENEL','GİDER','TOPLAM','TARİH','MAKBUZ NU.']];
    records.forEach((r,i) => {
      const res = db.residents.find(x => x.flat === r.flat);
      rows.push([i+1, r.name, r.flat, res?res.m2:109, r.yakit, r.genel, r.gider, r.toplam, r.tarih, r.makbuz]);
    });
    const totY=records.reduce((s,r)=>s+(r.yakit||0),0), totG=records.reduce((s,r)=>s+(r.genel||0),0);
    const totGi=records.reduce((s,r)=>s+(r.gider||0),0), totT=records.reduce((s,r)=>s+(r.toplam||0),0);
    rows.push(['TOPLAM','','','',totY,totG,totGi,totT,'','']);
    const ws = XLSX.utils.aoa_to_sheet(rows);
    ws['!cols'] = [{wch:6},{wch:24},{wch:12},{wch:5},{wch:14},{wch:10},{wch:10},{wch:14},{wch:12},{wch:10}];
    XLSX.utils.book_append_sheet(wbook, ws, mName.substring(0,31));
  });

  // Sakinler sayfası
  const resRows = [['SIRA','ADI SOYADI','BLOK/DAİRE','BLOK','M2']];
  db.residents.forEach(r => resRows.push([r.sira, r.name, r.flat, r.block, r.m2]));
  const wsR = XLSX.utils.aoa_to_sheet(resRows);
  wsR['!cols'] = [{wch:6},{wch:24},{wch:12},{wch:8},{wch:6}];
  XLSX.utils.book_append_sheet(wbook, wsR, 'Sakinler');

  // Gider kalemleri sayfası
  if (db.giderItems && Object.keys(db.giderItems).length > 0) {
    const gRows = [['AY','GİDER ADI','TUTAR (TL)']];
    Object.keys(db.giderItems).forEach(mName => {
      (db.giderItems[mName]||[]).forEach(g => {
        gRows.push([mName, g.name, g.amount]);
      });
    });
    const wsG = XLSX.utils.aoa_to_sheet(gRows);
    wsG['!cols'] = [{wch:20},{wch:20},{wch:12}];
    XLSX.utils.book_append_sheet(wbook, wsG, 'Gider Kalemleri');
  }

  // Site giderleri sayfası
  if (db.siteGiderleri && Object.keys(db.siteGiderleri).length > 0) {
    const sgRows = [['AY','KATEGORİ','TUTAR (TL)','AÇIKLAMA']];
    Object.keys(db.siteGiderleri).forEach(mName => {
      (db.siteGiderleri[mName]||[]).forEach(g => {
        sgRows.push([mName, g.kategori, g.tutar, g.aciklama||'']);
      });
    });
    const wsSG = XLSX.utils.aoa_to_sheet(sgRows);
    wsSG['!cols'] = [{wch:20},{wch:18},{wch:14},{wch:30}];
    XLSX.utils.book_append_sheet(wbook, wsSG, 'Site Giderleri');
  }

  XLSX.writeFile(wbook, result.filePath);
  return {success:true, path:result.filePath};
});

// ===== AUTO-UPDATE =====
const https = require('https');
const REPO_OWNER = 'CagriKibar';
const REPO_NAME = 'aidat-programi';

function checkGitHubUpdate() {
  return new Promise((resolve, reject) => {
    const opts = {
      hostname: 'api.github.com',
      path: `/repos/${REPO_OWNER}/${REPO_NAME}/releases/latest`,
      headers: { 'User-Agent': 'AidatYonetim/' + app.getVersion() }
    };
    https.get(opts, res => {
      let data = '';
      res.on('data', c => data += c);
      res.on('end', () => {
        try {
          const json = JSON.parse(data);
          resolve({ version: (json.tag_name || '').replace(/^v/, ''), url: json.html_url || '', assets: json.assets || [] });
        } catch (e) { reject(e); }
      });
    }).on('error', reject);
  });
}

function isNewer(remote, local) {
  const r = remote.split('.').map(Number);
  const l = local.split('.').map(Number);
  for (let i = 0; i < 3; i++) {
    if ((r[i]||0) > (l[i]||0)) return true;
    if ((r[i]||0) < (l[i]||0)) return false;
  }
  return false;
}

async function setupAutoUpdater() {
  try {
    const current = app.getVersion();
    const release = await checkGitHubUpdate();
    if (!release.version || !isNewer(release.version, current)) return;

    const setupAsset = release.assets.find(a => a.name && a.name.endsWith('.exe') && a.name.includes('setup'));
    const downloadUrl = setupAsset ? setupAsset.browser_download_url : release.url;

    const result = await dialog.showMessageBox(mainWindow, {
      type: 'info',
      title: 'Güncelleme Mevcut',
      message: `Yeni sürüm mevcut: v${release.version}\nMevcut sürüm: v${current}\n\nGüncellemek ister misiniz?`,
      buttons: ['İndir', 'Daha Sonra'],
      defaultId: 0
    });

    if (result.response === 0) {
      require('electron').shell.openExternal(downloadUrl);
    }
  } catch (e) {
    // ağ hatası — sessiz
  }
}

ipcMain.handle('check-update', async () => {
  try {
    const current = app.getVersion();
    const release = await checkGitHubUpdate();
    if (release.version && isNewer(release.version, current)) {
      return { available: true, version: release.version };
    }
    return { available: false, current };
  } catch (e) {
    return { available: false, error: e.message };
  }
});

ipcMain.handle('get-version', () => app.getVersion());

app.whenReady().then(() => {
  loadDB();
  createWindow();
  setTimeout(setupAutoUpdater, 2000);
});
app.on('window-all-closed', () => { app.quit(); });

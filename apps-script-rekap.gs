// Cara pasang: buka sheets.new > Extensions > Apps Script > tempel kode ini > Deploy > New deployment
// Setiap kali kode diubah: Deploy > Manage deployments > ikon pensil > Version: New version > Deploy.
// Type: Web app | Execute as: Me | Who has access: Anyone > Deploy > salin URL "/exec" ke konstanta ENDPOINT di file HTML.
// Jika skrip dibuat dari script.google.com (bukan dari menu Extensions di dalam Sheets), isi ID spreadsheet di sini.
// ID ada di URL sheet: docs.google.com/spreadsheets/d/ID_ADA_DI_SINI/edit
const SHEET_ID="";
const HEAD=["Waktu","Nama","Kelas","NIS","Status","PG Benar","Skor PG (maks 60)","Pelanggaran","Esai 1","Esai 2","Esai 3","Esai 4","Esai 5","Skor Esai (isi guru, maks 40)","Nilai Akhir","Log Pelanggaran"];
function doGet(){
  return ContentService.createTextOutput("Server rekap CBT aktif. Siswa mengirim nilai lewat aplikasi ujian.");
}
function doPost(e){
  const lock=LockService.getScriptLock();lock.waitLock(20000);
  try{
    const d=JSON.parse(e.postData.contents);
    const ss=SHEET_ID?SpreadsheetApp.openById(SHEET_ID):SpreadsheetApp.getActiveSpreadsheet();
    let sh=ss.getSheetByName("Rekap")||ss.insertSheet("Rekap");
    if(sh.getLastRow()===0){sh.appendRow(HEAD);sh.setFrozenRows(1);}
    const nis=String(d.nis),kelas=String(d.kelas);
    const v=sh.getRange(1,1,sh.getLastRow(),5).getValues();
    let r=0;
    for(let i=1;i<v.length;i++){if(String(v[i][3])===nis&&String(v[i][2])===kelas){r=i+1;break;}}
    sh.getRange(1,8).setValue("Pelanggaran");sh.getRange(1,16).setValue("Log Pelanggaran");
    if(r&&String(v[r-1][4])==="selesai"){const c=sh.getRange(r,16);c.setValue((c.getValue()?c.getValue()+" | ":"")+"PERCOBAAN ULANG ditolak "+new Date().toLocaleString("id-ID"));return ContentService.createTextOutput("dup");}
    const row=[new Date(),d.nama,kelas,"'"+nis,d.status,d.benar,d.pg,d.sw].concat(d.es||[]);
    if(!r){r=sh.getLastRow()+1;}
    sh.getRange(r,1,1,row.length).setValues([row]);
    sh.getRange(r,15).setFormula("=G"+r+"+N"+r);
    sh.getRange(r,16).setValue(d.log||"");
    return ContentService.createTextOutput("ok");
  }finally{lock.releaseLock();}
}

// Uji dari editor: pilih fungsi "tes" di menu atas lalu klik Run. Satu baris "Siswa Uji" harus muncul di sheet Rekap.
function tes(){
  doPost({postData:{contents:JSON.stringify({nama:"Siswa Uji",kelas:"XI-0",nis:"000",status:"selesai",benar:10,pg:30,sw:0,es:["a","b","c","d","e"]})}});
}

// Cara pasang: buka sheets.new > Extensions > Apps Script > tempel kode ini > Deploy > New deployment
// Setiap kali kode diubah: Deploy > Manage deployments > ikon pensil > Version: New version > Deploy.
// Type: Web app | Execute as: Me | Who has access: Anyone > Deploy > salin URL "/exec" ke konstanta ENDPOINT di file HTML.
const HEAD=["Waktu","Nama","Kelas","NIS","Status","PG Benar","Skor PG (maks 60)","Pindah Tab","Esai 1","Esai 2","Esai 3","Esai 4","Esai 5","Skor Esai (isi guru, maks 40)","Nilai Akhir"];
function doGet(){
  return ContentService.createTextOutput("Server rekap CBT aktif. Siswa mengirim nilai lewat aplikasi ujian.");
}
function doPost(e){
  const lock=LockService.getScriptLock();lock.waitLock(20000);
  try{
    const d=JSON.parse(e.postData.contents);
    const ss=SpreadsheetApp.getActiveSpreadsheet();
    let sh=ss.getSheetByName("Rekap")||ss.insertSheet("Rekap");
    if(sh.getLastRow()===0){sh.appendRow(HEAD);sh.setFrozenRows(1);}
    const nis=String(d.nis),kelas=String(d.kelas);
    const v=sh.getRange(1,1,sh.getLastRow(),4).getValues();
    let r=0;
    for(let i=1;i<v.length;i++){if(String(v[i][3])===nis&&String(v[i][2])===kelas){r=i+1;break;}}
    const row=[new Date(),d.nama,kelas,"'"+nis,d.status,d.benar,d.pg,d.sw].concat(d.es||[]);
    if(!r){r=sh.getLastRow()+1;}
    sh.getRange(r,1,1,row.length).setValues([row]);
    sh.getRange(r,15).setFormula("=G"+r+"+N"+r);
    return ContentService.createTextOutput("ok");
  }finally{lock.releaseLock();}
}

const APP = {
  name: 'GARJEK Driver Registration',
  timezone: 'Asia/Jakarta',
  maxFileBytes: 5 * 1024 * 1024,
  allowedMime: ['image/jpeg','image/png','application/pdf'],
  sessionHours: 6
};

function doGet(e) {
  return json_({ok:true, service:APP.name, time:new Date().toISOString()});
}

function doPost(e) {
  try {
    const p = e && e.postData ? JSON.parse(e.postData.contents || '{}') : {};
    return json_(route_(p));
  } catch (err) {
    return json_({ok:false,error:safeError_(err)});
  }
}

/**
 * Run once from Apps Script editor.
 * Creates Spreadsheet + Drive folder + required sheets and an initial admin account.
 * The generated password is returned/logged ONLY during this setup step; it is never
 * sent to the public website and is not stored in plaintext.
 */
function setupSystem() {
  const props = PropertiesService.getScriptProperties();
  let ssId = props.getProperty('SPREADSHEET_ID');
  let folderId = props.getProperty('DRIVE_FOLDER_ID');

  if (!ssId) {
    const ss = SpreadsheetApp.create('GARJEK - Database Driver');
    ssId = ss.getId();
    props.setProperty('SPREADSHEET_ID', ssId);
  }
  if (!folderId) {
    const folder = DriveApp.createFolder('GARJEK - Dokumen Driver');
    folderId = folder.getId();
    props.setProperty('DRIVE_FOLDER_ID', folderId);
  }

  const ss = SpreadsheetApp.openById(ssId);
  ensureSheet_(ss,'Drivers',[
    'Timestamp','Nomor Pendaftaran','NIK','Nama Lengkap','Tempat Lahir','Tanggal Lahir',
    'Jenis Kelamin','WhatsApp','Email','Status Pernikahan','Negara','Provinsi',
    'Kabupaten/Kota','Kecamatan','Desa/Kelurahan','Dusun','RT','RW','Kode Pos',
    'Alamat Lengkap','Layanan','Nomor Polisi','Merek','Tipe','Tahun','Warna',
    'Nomor STNK','Nomor BPKB','Nama Pemilik','Status Kepemilikan','Status Dokumen',
    'Status Pembayaran','Status Pendaftaran','Catatan Admin','Folder Dokumen'
  ]);
  ensureSheet_(ss,'Documents',[
    'Timestamp','Nomor Pendaftaran','NIK','KTP','SIM','STNK','Foto Driver','Foto Kendaraan',
    'Bukti Pembayaran'
  ]);
  ensureSheet_(ss,'Settings',['Key','Value']);
  ensureSheet_(ss,'Admin',['Username','Password Hash','Created At','Active']);

  const settings = {
    FEE_GARBIKE:'150000',
    FEE_GARCAR:'0',
    DANA_NAME:'AHMAD SYAIFFUDDIN',
    DANA_NUMBER:'081932592253',
    SYSTEM_NAME:'GARJEK'
  };
  const shSet=ss.getSheetByName('Settings');
  const existing = shSet.getDataRange().getValues();
  const keys = existing.slice(1).map(r=>String(r[0]));
  Object.keys(settings).forEach(k=>{
    if (!keys.includes(k)) shSet.appendRow([k,settings[k]]);
  });

  let adminUser = props.getProperty('ADMIN_USERNAME');
  let adminHash = props.getProperty('ADMIN_PASSWORD_HASH');
  let generatedPassword = null;
  if (!adminUser || !adminHash) {
    adminUser = 'admin';
    generatedPassword = randomPassword_(14);
    adminHash = sha256_(generatedPassword);
    props.setProperties({
      ADMIN_USERNAME: adminUser,
      ADMIN_PASSWORD_HASH: adminHash
    });
    const ah=ss.getSheetByName('Admin');
    if (ah.getLastRow()===1) ah.appendRow([adminUser,adminHash,new Date(),true]);
  }

  const result = {
    ok:true,
    message:'Setup selesai.',
    spreadsheetId:ssId,
    spreadsheetUrl:ss.getUrl(),
    driveFolderId:folderId,
    driveFolderUrl:DriveApp.getFolderById(folderId).getUrl(),
    adminUsername:adminUser
  };
  if (generatedPassword) result.generatedAdminPassword = generatedPassword;
  Logger.log(JSON.stringify(result));
  return result;
}

function route_(p) {
  switch (String(p.action||'')) {
    case 'publicConfig': return publicConfig_();
    case 'register': return register_(p);
    case 'checkStatus': return checkStatus_(p);
    case 'login': return login_(p);
    case 'adminStats': return adminStats_(p);
    case 'adminList': return adminList_(p);
    case 'adminDetail': return adminDetail_(p);
    case 'updateStatus': return updateStatus_(p);
    default: return {ok:false,error:'Aksi tidak dikenali.'};
  }
}

function publicConfig_() {
  const s=getSettings_();
  return {ok:true,config:{
    systemName:s.SYSTEM_NAME||APP.name,
    danaName:s.DANA_NAME||'AHMAD SYAIFFUDDIN',
    danaNumber:s.DANA_NUMBER||'081932592253',
    feeGarbike:Number(s.FEE_GARBIKE||150000),
    feeGarcar:Number(s.FEE_GARCAR||0)
  }};
}

function register_(p) {
  const d=p.data||{};
  validateRegistration_(d,p.files||[]);
  const lock=LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    const ss=getSS_(), drivers=ss.getSheetByName('Drivers');
    const docs=ss.getSheetByName('Documents');
    const existing=drivers.getDataRange().getValues().slice(1);
    if (existing.some(r=>String(r[2])===String(d.nik))) {
      throw new Error('NIK sudah terdaftar. Gunakan menu Cek Status Pendaftaran.');
    }
    const no=nextRegistrationNumber_(drivers);
    const folder=DriveApp.getFolderById(PropertiesService.getScriptProperties().getProperty('DRIVE_FOLDER_ID'))
      .createFolder(no+' - '+safeName_(d.nama));
    const uploaded={};
    (p.files||[]).forEach(f=>{
      uploaded[f.field]=saveFile_(f,folder);
    });

    const row=[
      new Date(),no,d.nik,d.nama,d.tempatLahir,d.tanggalLahir,d.jenisKelamin,d.whatsapp,d.email,
      d.statusPernikahan,d.negara,d.provinsi,d.kabupatenKota,d.kecamatan,d.desaKelurahan,d.dusun,
      d.rt,d.rw,d.kodePos,d.alamatLengkap,d.layanan,d.nomorPolisi,d.merek,d.tipe,d.tahun,d.warna,
      d.nomorSTNK,d.nomorBPKB,d.namaPemilik,d.statusKepemilikan,'MENUNGGU VERIFIKASI',
      'MENUNGGU VERIFIKASI','MENUNGGU VERIFIKASI','',folder.getUrl()
    ];
    drivers.appendRow(row);
    docs.appendRow([
      new Date(),no,d.nik,uploaded.ktp||'',uploaded.sim||'',uploaded.stnk||'',
      uploaded.fotoDriver||'',uploaded.fotoKendaraan||'',uploaded.buktiPembayaran||''
    ]);
    return {ok:true,registrationNumber:no,status:'MENUNGGU VERIFIKASI'};
  } finally {
    lock.releaseLock();
  }
}

function checkStatus_(p) {
  const nik=String(p.nik||'').replace(/\D/g,'');
  if (!/^\d{16}$/.test(nik)) throw new Error('NIK harus tepat 16 digit.');
  const ss=getSS_(), rows=ss.getSheetByName('Drivers').getDataRange().getValues();
  const r=rows.slice(1).find(x=>String(x[2])===nik);
  if(!r) return {ok:true,found:false};
  return {ok:true,found:true,data:{
    registrationNumber:r[1], name:maskName_(r[3]), nik:maskNik_(r[2]),
    service:r[20], registeredAt:formatDate_(r[0]), documentStatus:r[30],
    paymentStatus:r[31], status:r[32], note:r[33]||''
  }};
}

function login_(p) {
  const props=PropertiesService.getScriptProperties();
  const user=String(p.username||'');
  const pass=String(p.password||'');
  if(user!==props.getProperty('ADMIN_USERNAME') || sha256_(pass)!==props.getProperty('ADMIN_PASSWORD_HASH')) {
    throw new Error('Username atau password salah.');
  }
  const token=Utilities.getUuid();
  CacheService.getScriptCache().put('ADMIN_'+token,user,APP.sessionHours*3600);
  return {ok:true,token};
}

function adminStats_(p) {
  requireAdmin_(p.token);
  const rows=getSS_().getSheetByName('Drivers').getDataRange().getValues().slice(1);
  const today=Utilities.formatDate(new Date(),APP.timezone,'yyyy-MM-dd');
  const stats={total:rows.length,today:0,garbike:0,garcar:0,pending:0,approved:0,rejected:0,needFix:0,paymentPending:0};
  rows.forEach(r=>{
    if(formatDate_(r[0]).slice(0,10)===today) stats.today++;
    if(r[20]==='GARBIKE') stats.garbike++;
    if(r[20]==='GARCAR') stats.garcar++;
    if(r[32]==='MENUNGGU VERIFIKASI') stats.pending++;
    if(r[32]==='DISETUJUI') stats.approved++;
    if(r[32]==='DITOLAK') stats.rejected++;
    if(r[30]==='PERLU PERBAIKAN' || r[32]==='PERLU PERBAIKAN') stats.needFix++;
    if(r[31]==='MENUNGGU VERIFIKASI') stats.paymentPending++;
  });
  const chart=monthlyCounts_(rows);
  return {ok:true,stats,chart};
}

function adminList_(p) {
  requireAdmin_(p.token);
  const rows=getSS_().getSheetByName('Drivers').getDataRange().getValues();
  const q=String(p.q||'').toLowerCase();
  const service=String(p.service||'');
  const status=String(p.status||'');
  const out=[];
  for(let i=1;i<rows.length;i++){
    const r=rows[i];
    if(q && ![r[1],r[2],r[3],r[7],r[21]].join(' ').toLowerCase().includes(q)) continue;
    if(service && r[20]!==service) continue;
    if(status && r[32]!==status) continue;
    out.push({row:i+1,registrationNumber:r[1],date:formatDate_(r[0]),name:r[3],nik:r[2],
      whatsapp:r[7],service:r[20],vehicle:[r[22],r[23]].filter(Boolean).join(' '),
      documentStatus:r[30],paymentStatus:r[31],status:r[32]});
  }
  return {ok:true,data:out};
}

function adminDetail_(p) {
  requireAdmin_(p.token);
  const no=String(p.registrationNumber||'');
  const ss=getSS_(), d=ss.getSheetByName('Drivers').getDataRange().getValues().slice(1);
  const r=d.find(x=>String(x[1])===no);
  if(!r) throw new Error('Data tidak ditemukan.');
  const doc=ss.getSheetByName('Documents').getDataRange().getValues().slice(1).find(x=>String(x[1])===no)||[];
  return {ok:true,data:{
    registrationNumber:r[1],timestamp:formatDate_(r[0]),nik:r[2],name:r[3],tempatLahir:r[4],
    tanggalLahir:r[5],jenisKelamin:r[6],whatsapp:r[7],email:r[8],statusPernikahan:r[9],
    negara:r[10],provinsi:r[11],kabupatenKota:r[12],kecamatan:r[13],desaKelurahan:r[14],
    dusun:r[15],rt:r[16],rw:r[17],kodePos:r[18],alamatLengkap:r[19],layanan:r[20],
    nomorPolisi:r[21],merek:r[22],tipe:r[23],tahun:r[24],warna:r[25],nomorSTNK:r[26],
    nomorBPKB:r[27],namaPemilik:r[28],statusKepemilikan:r[29],statusDokumen:r[30],
    statusPembayaran:r[31],statusPendaftaran:r[32],catatan:r[33],folder:r[34],
    documents:{ktp:doc[3]||'',sim:doc[4]||'',stnk:doc[5]||'',fotoDriver:doc[6]||'',
      fotoKendaraan:doc[7]||'',buktiPembayaran:doc[8]||''}
  }};
}

function updateStatus_(p) {
  requireAdmin_(p.token);
  const no=String(p.registrationNumber||'');
  const ss=getSS_(), sh=ss.getSheetByName('Drivers'), rows=sh.getDataRange().getValues();
  const idx=rows.findIndex((r,i)=>i>0 && String(r[1])===no);
  if(idx<1) throw new Error('Data tidak ditemukan.');
  if(p.documentStatus!==undefined) sh.getRange(idx+1,31).setValue(String(p.documentStatus));
  if(p.paymentStatus!==undefined) sh.getRange(idx+1,32).setValue(String(p.paymentStatus));
  if(p.status!==undefined) sh.getRange(idx+1,33).setValue(String(p.status));
  if(p.note!==undefined) sh.getRange(idx+1,34).setValue(String(p.note));
  return {ok:true,message:'Status berhasil diperbarui.'};
}

function validateRegistration_(d,files) {
  const required=['nama','nik','tempatLahir','tanggalLahir','jenisKelamin','whatsapp','negara','provinsi',
    'kabupatenKota','kecamatan','desaKelurahan','kodePos','alamatLengkap','layanan','nomorPolisi',
    'merek','tipe','tahun','warna','nomorSTNK','namaPemilik','statusKepemilikan'];
  required.forEach(k=>{if(!String(d[k]||'').trim()) throw new Error('Kolom '+k+' wajib diisi.');});
  if(!/^\d{16}$/.test(String(d.nik))) throw new Error('NIK harus tepat 16 digit.');
  if(!['GARBIKE','GARCAR'].includes(String(d.layanan))) throw new Error('Pilihan layanan tidak valid.');
  if(!/^\d{4}$/.test(String(d.tahun))) throw new Error('Tahun kendaraan tidak valid.');
  const needed=['ktp','sim','stnk','fotoDriver','fotoKendaraan','buktiPembayaran'];
  needed.forEach(k=>{if(!files.some(f=>f.field===k)) throw new Error('File '+k+' wajib diunggah.');});
  files.forEach(f=>{
    const bytes=Utilities.base64Decode(String(f.base64||''));
    if(bytes.length>APP.maxFileBytes) throw new Error('File '+f.field+' melebihi 5 MB.');
    if(!APP.allowedMime.includes(String(f.mimeType))) throw new Error('Format file '+f.field+' tidak didukung.');
  });
}

function saveFile_(f,folder) {
  const bytes=Utilities.base64Decode(String(f.base64||''));
  const blob=Utilities.newBlob(bytes,String(f.mimeType),safeFile_(f.name||f.field));
  const file=folder.createFile(blob);
  // Private by default; admin opens via Drive account permission.
  return file.getUrl();
}

function requireAdmin_(token) {
  if(!token || !CacheService.getScriptCache().get('ADMIN_'+token)) throw new Error('Sesi admin tidak valid atau sudah kedaluwarsa.');
}

function getSS_(){return SpreadsheetApp.openById(PropertiesService.getScriptProperties().getProperty('SPREADSHEET_ID'));}
function getSettings_(){
  const sh=getSS_().getSheetByName('Settings'), v=sh.getDataRange().getValues(), o={};
  v.slice(1).forEach(r=>o[String(r[0])]=String(r[1]));
  return o;
}
function ensureSheet_(ss,name,headers){
  let sh=ss.getSheetByName(name);
  if(!sh) sh=ss.insertSheet(name);
  if(sh.getLastRow()===0) sh.appendRow(headers);
  else {
    const current=sh.getRange(1,1,1,Math.max(sh.getLastColumn(),headers.length)).getValues()[0];
    headers.forEach((h,i)=>{if(current[i]!==h) sh.getRange(1,i+1).setValue(h);});
  }
  sh.setFrozenRows(1);
}
function nextRegistrationNumber_(sh){
  const year=Utilities.formatDate(new Date(),APP.timezone,'yyyyMMdd');
  const rows=sh.getDataRange().getValues().slice(1);
  let max=0;
  rows.forEach(r=>{const m=String(r[1]||'').match(/^GJ-\d{8}-(\d{4})$/);if(m) max=Math.max(max,Number(m[1]));});
  return 'GJ-'+year+'-'+String(max+1).padStart(4,'0');
}
function monthlyCounts_(rows){
  const out={};
  rows.forEach(r=>{const k=Utilities.formatDate(new Date(r[0]),APP.timezone,'yyyy-MM');out[k]=(out[k]||0)+1;});
  return Object.keys(out).sort().slice(-12).map(k=>({month:k,count:out[k]}));
}
function sha256_(s){
  const bytes=Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256,String(s),Utilities.Charset.UTF_8);
  return bytes.map(b=>(b<0?b+256:b).toString(16).padStart(2,'0')).join('');
}
function randomPassword_(n){
  const chars='ABCDEFGHJKLMNPQRSTUVWXYZabcdefghijkmnopqrstuvwxyz23456789!@#$%';
  let s=''; for(let i=0;i<n;i++) s+=chars[Math.floor(Math.random()*chars.length)]; return s;
}
function safeName_(s){return String(s||'DRIVER').replace(/[\\/:*?"<>|#%{}]/g,' ').trim().slice(0,80)||'DRIVER';}
function safeFile_(s){return String(s||'file').replace(/[\\/:*?"<>|#%{}]/g,'_').slice(0,120);}
function maskNik_(s){s=String(s);return s.slice(0,4)+'********'+s.slice(-4);}
function maskName_(s){const a=String(s).trim().split(/\s+/);return a[0].slice(0,2)+'***';}
function formatDate_(v){return v instanceof Date?Utilities.formatDate(v,APP.timezone,'yyyy-MM-dd HH:mm:ss'):String(v||'');}
function json_(o){return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON);}
function safeError_(e){return e&&e.message?e.message:'Terjadi kesalahan pada server.';}

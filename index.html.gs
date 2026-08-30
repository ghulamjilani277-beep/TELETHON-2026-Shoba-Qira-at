
const RATE = 7000;
const ADMIN_PASSWORD = 'GJhabibi89@';
const PARTICIPANTS_SHEET = 'Participants';
const COORDINATORS_SHEET = 'Coordinators';
const DAILY_SHEET = 'DailyCollections';

function doGet() {
  return HtmlService.createHtmlOutputFromFile('index').setTitle('TELETHON 2026');
}

function getHeaderMap(sheet) {
  const map = {};
  sheet.getRange(1,1,1,sheet.getLastColumn()).getValues()[0].forEach((h,i)=>{
    map[String(h).trim().toLowerCase().replace(/[\s_\-\/]/g,'')] = i;
  });
  return map;
}

function findColumn(map, names) {
  for (const name of names) {
    const key = String(name).toLowerCase().replace(/[\s_\-\/]/g,'');
    if (map[key] !== undefined) return map[key];
  }
  return -1;
}
/* removed duplicate function */




function getOrCreateDailySheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName(DAILY_SHEET);
  if (!sh) {
    sh = ss.insertSheet(DAILY_SHEET);
    sh.appendRow(['Date','ParticipantID','Amount','UpdatedAt']);
    sh.setFrozenRows(1);
  }
  return sh;
}

function getAllowedDates() {
  const out = [];
  const d = new Date(2026,7,1);
  const end = new Date(2026,8,15);
  while (d <= end) {
    out.push(Utilities.formatDate(d, Session.getScriptTimeZone(), 'yyyy-MM-dd'));
    d.setDate(d.getDate()+1);
  }
  return out;
}

function normaliseDate_(dateStr) {
  // Calendar-based date selection: accept any valid YYYY-MM-DD date.
  if (!dateStr || dateStr === 'all') return 'all';
  const date = String(dateStr).trim();
  if (!/^\d{4}-\d{2}-\d{2}$/.test(date)) throw new Error('Invalid collection date format');
  const d = new Date(date + 'T00:00:00');
  if (isNaN(d.getTime()) || Utilities.formatDate(d, Session.getScriptTimeZone(), 'yyyy-MM-dd') !== date) throw new Error('Invalid collection date');
  return date;
}

function getDailyMap_(dateStr) {
  const result = {};
  if (dateStr === 'all') return result;
  const sh = getOrCreateDailySheet_();
  if (sh.getLastRow() < 2) return result;
  const rows = sh.getDataRange().getValues();
  for (let i=1;i<rows.length;i++) {
    const raw = rows[i][0];
    const key = raw instanceof Date
      ? Utilities.formatDate(raw, Session.getScriptTimeZone(), 'yyyy-MM-dd')
      : String(raw);
    if (key === dateStr) result[String(rows[i][1])] = Number(rows[i][2]) || 0;
  }
  return result;
}

function canonicalRegion_(value) {
  const s = String(value || '').toLowerCase().trim();
  const compact = s.replace(/[^a-z0-9]/g,'');

  // Explicit known TELETHON regions. This also handles sheet keys like:
  // delhi1, delhi2, ajmer1, ajmer2, mumbai, kolkata, hyderabad, bangalore.
  const known = ['mumbai','kolkata','hyderabad','bangalore','ajmer','delhi'];
  for (const region of known) {
    if (compact === region || compact.indexOf(region) === 0 || s.indexOf(region) >= 0) {
      return region;
    }
  }

  // Generic cleanup fallback
  return compact.replace(/region$/,'').replace(/part\d+$/,'');
}

function canonicalPart_(value, fallbackSource) {
  const combined = String(value || '') + ' ' + String(fallbackSource || '');
  const s = combined.toLowerCase().trim();

  if (!s || /\bmain\b/.test(s)) return 'main';

  const m = s.match(/\bpart\s*0*([0-9]+)\b/i) ||
            s.match(/(?:^|[^a-z0-9])0*([12])(?:$|[^a-z0-9])/i) ||
            s.match(/(?:delhi|ajmer)\s*0*([12])\b/i);

  return m ? 'part' + String(Number(m[1])) : 'main';
}

function accessKey_(region, part, fallbackSource) {
  return canonicalRegion_(region) + '|' + canonicalPart_(part, fallbackSource);
}

function getCoordinatorMap_() {
  const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(COORDINATORS_SHEET);
  const result = {};
  if (!sh || sh.getLastRow() < 2) return result;

  const rows = sh.getDataRange().getValues();
  const hm = getHeaderMap(sh);
  const regionCol = findColumn(hm,['Region','RegionKey']);
  const partCol = findColumn(hm,['Part']);
  const nameCol = findColumn(hm,['Coordinator','Name']);
  const mobileCol = findColumn(hm,['Mobile','Phone','Number']);

  for (let i=1;i<rows.length;i++) {
    const r = rows[i];
    const regionRaw = regionCol>=0 ? String(r[regionCol]||'') : '';
    const partRaw = partCol>=0 ? String(r[partCol]||'') : '';
    const key = accessKey_(regionRaw, partRaw, regionRaw + ' ' + partRaw);

    result[key] = {
      name: nameCol>=0 ? String(r[nameCol]||'') : '',
      mobile: mobileCol>=0 ? String(r[mobileCol]||'') : ''
    };
  }
  return result;
}

function getCoordinatorFor_(regionKey, regionLabel, part, map) {
  const key = accessKey_(regionLabel || regionKey, part, regionKey + ' ' + regionLabel);
  return map[key] || {};
}
/* ================= QARI SAHIB PHOTO SYSTEM ================= */
const QARI_PHOTOS_SHEET = 'QariPhotos';
const QARI_PHOTOS_FOLDER = 'TELETHON 2026 Qari Sahiban Photos';

function getOrCreateQariPhotosSheet_(){
  const ss=SpreadsheetApp.getActiveSpreadsheet();
  let sh=ss.getSheetByName(QARI_PHOTOS_SHEET);
  if(!sh){sh=ss.insertSheet(QARI_PHOTOS_SHEET);sh.appendRow(['ParticipantID','PhotoUrl','PhotoFileId','UpdatedBy','UpdatedAt']);sh.setFrozenRows(1);}
  return sh;
}
function getQariPhotosFolder_(){
  const folders=DriveApp.getFoldersByName(QARI_PHOTOS_FOLDER);
  return folders.hasNext()?folders.next():DriveApp.createFolder(QARI_PHOTOS_FOLDER);
}
function getQariPhotoMap_(){
  const sh=getOrCreateQariPhotosSheet_(), out={};
  if(sh.getLastRow()<2)return out;
  sh.getDataRange().getValues().slice(1).forEach(r=>{if(r[0])out[String(r[0])]={photoUrl:String(r[1]||''),fileId:String(r[2]||'')};});
  return out;
}
function uploadQariPhoto(token,participantId,dataUrl,fileName,mimeType){
  const user=assertParticipantAccess_(token,participantId);
  const raw=String(dataUrl||''); if(!raw||raw.indexOf(',')<0)throw new Error('Invalid image data.');
  const type=String(mimeType||'image/jpeg'); if(type.indexOf('image/')!==0)throw new Error('Only image files are allowed.');
  const bytes=Utilities.base64Decode(raw.split(',')[1]);
  if(bytes.length>5*1024*1024)throw new Error('Image must be under 5 MB.');
  const safe=String(fileName||'qari.jpg').replace(/[^A-Za-z0-9._-]/g,'_');
  const file=getQariPhotosFolder_().createFile(Utilities.newBlob(bytes,type,safe));
  file.setName('QARI_'+String(participantId).replace(/[^A-Za-z0-9_-]/g,'_')+'_'+Date.now()+'_'+safe);
  file.setSharing(DriveApp.Access.ANYONE_WITH_LINK,DriveApp.Permission.VIEW);
  const url='https://drive.google.com/thumbnail?id='+file.getId()+'&sz=w800';
  const sh=getOrCreateQariPhotosSheet_(), vals=sh.getDataRange().getValues(); let rowNo=-1,oldFile='';
  for(let i=1;i<vals.length;i++)if(String(vals[i][0])===String(participantId)){rowNo=i+1;oldFile=String(vals[i][2]||'');break;}
  if(oldFile){try{DriveApp.getFileById(oldFile).setTrashed(true);}catch(e){}}
  const row=[String(participantId),url,file.getId(),user.email,new Date()];
  if(rowNo>0)sh.getRange(rowNo,1,1,row.length).setValues([row]); else sh.appendRow(row);
  try{logActivity_('QARI PHOTO UPDATED',String(participantId),'Qari Photo','PROFILE',0,0,'Updated by '+user.email);}catch(e){}
  return {success:true,photoUrl:url};
}
function removeQariPhoto(token,participantId){
  const user=assertParticipantAccess_(token,participantId), sh=getOrCreateQariPhotosSheet_();
  if(sh.getLastRow()<2)return {success:true};
  const vals=sh.getDataRange().getValues();
  for(let i=1;i<vals.length;i++)if(String(vals[i][0])===String(participantId)){
    const fileId=String(vals[i][2]||''); if(fileId){try{DriveApp.getFileById(fileId).setTrashed(true);}catch(e){}}
    sh.deleteRow(i+1); try{logActivity_('QARI PHOTO REMOVED',String(participantId),'Qari Photo','PROFILE',0,0,'Removed by '+user.email);}catch(e){}
    return {success:true};
  }
  return {success:true};
}


// Returns a same-origin data URL for poster rendering. This avoids Google Drive CORS issues in html2canvas downloads.
function getQariPhotoDataUrl(token, participantId){
  assertParticipantAccess_(token, participantId);
  const sh=getOrCreateQariPhotosSheet_();
  if(sh.getLastRow()<2) return {success:true,dataUrl:''};
  const vals=sh.getDataRange().getValues();
  for(let i=1;i<vals.length;i++){
    if(String(vals[i][0])===String(participantId)){
      const fileId=String(vals[i][2]||'');
      if(!fileId) return {success:true,dataUrl:''};
      try{
        const blob=DriveApp.getFileById(fileId).getBlob();
        const mime=blob.getContentType()||'image/jpeg';
        const dataUrl='data:'+mime+';base64,'+Utilities.base64Encode(blob.getBytes());
        return {success:true,dataUrl:dataUrl};
      }catch(e){ return {success:true,dataUrl:''}; }
    }
  }
  return {success:true,dataUrl:''};
}

function getDataInternal_(selectedDate) {
  selectedDate = normaliseDate_(selectedDate || 'all');
  const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(PARTICIPANTS_SHEET);
  if (!sh || sh.getLastRow() < 2) return [];

  const rows = sh.getDataRange().getValues();
  const hm = getHeaderMap(sh);
  const idCol = findColumn(hm,['ID']);
  const regionCol = findColumn(hm,['Region','RegionKey']);
  const regionLabelCol = findColumn(hm,['RegionLabel','RegionName']);
  const partCol = findColumn(hm,['Part']);
  const nameCol = findColumn(hm,['QariName','Name','Qari Sahab']);
  const centreCol = findColumn(hm,['CityCentre','City','Centre','Jamia']);
  const amountCol = findColumn(hm,['Amount']);
  const statusCol = findColumn(hm,['Status']);
  const dailyMap = getDailyMap_(selectedDate);
  const coordinators = getCoordinatorMap_();
  const qariPhotos = getQariPhotoMap_();
  const result = [];

  for (let i=1;i<rows.length;i++) {
    const r = rows[i];
    const id = idCol>=0 ? String(r[idCol]) : String(i);
    const rawRegion = regionCol>=0 ? String(r[regionCol]||'') : '';
    const region = regionLabelCol>=0 && r[regionLabelCol] ? String(r[regionLabelCol]) : rawRegion;
    const part = partCol>=0 ? String(r[partCol]||'') : '';
    const c = getCoordinatorFor_(rawRegion, region, part, coordinators);
    const overall = amountCol>=0 ? Number(r[amountCol])||0 : 0;
    result.push({
      id:id, region:region, regionKey:rawRegion, part:part,
      name:nameCol>=0 ? String(r[nameCol]||'') : '',
      centre:centreCol>=0 ? String(r[centreCol]||'') : '',
      overallAmount:overall,
      amount:selectedDate==='all' ? overall : (Number(dailyMap[id])||0),
      status:statusCol>=0 ? String(r[statusCol]||'active') : 'active',
      coordinator:c.name||'', phone:c.mobile||'',
      photoUrl:(qariPhotos[id]&&qariPhotos[id].photoUrl)||''
    });
  }
  return result;
}
/* removed duplicate function */















































/* removed duplicate function */




















/* removed duplicate function */











/* removed duplicate function */





/* =====================================================
   TELETHON 2026 PROFESSIONAL UPGRADE ADDON
   Paste this at the BOTTOM of your existing Code.gs
   ===================================================== */

const AUDIT_SHEET = 'ActivityLog';

function getOrCreateAuditSheet_(){
  const ss=SpreadsheetApp.getActiveSpreadsheet();
  let sh=ss.getSheetByName(AUDIT_SHEET);
  if(!sh){
    sh=ss.insertSheet(AUDIT_SHEET);
    sh.appendRow(['Timestamp','Action','ParticipantID','Qari Name','Date','Old Amount','New Amount','Details']);
    sh.setFrozenRows(1);
  }
  return sh;
}

function logActivity_(action,id,name,date,oldAmount,newAmount,details){
  getOrCreateAuditSheet_().appendRow([
    new Date(), action, String(id||''), String(name||''), String(date||''),
    Number(oldAmount)||0, Number(newAmount)||0, String(details||'')
  ]);
}

function getParticipantInfo_(id){
  const sh=SpreadsheetApp.getActiveSpreadsheet().getSheetByName(PARTICIPANTS_SHEET);
  if(!sh) throw new Error('Participants sheet not found');
  const hm=getHeaderMap(sh);
  const idCol=findColumn(hm,['ID']);
  const nameCol=findColumn(hm,['QariName','Name','Qari Sahab']);
  const amountCol=findColumn(hm,['Amount']);
  if(idCol<0||amountCol<0) throw new Error('Required participant columns not found');
  const ids=sh.getRange(2,idCol+1,sh.getLastRow()-1,1).getValues();
  for(let i=0;i<ids.length;i++){
    if(String(ids[i][0])===String(id)){
      return {sh:sh,hm:hm,row:i+2,name:nameCol>=0?String(sh.getRange(i+2,nameCol+1).getValue()||''):'' ,amountCol:amountCol};
    }
  }
  throw new Error('Participant not found');
}
/* removed duplicate function */
































/* removed duplicate function */















/* removed duplicate function */

















/* IMPORTANT:
   Replace your existing updateAmount function with this version
   so every correction is recorded in ActivityLog.
*/
/* removed duplicate function */























































/* =====================================================
   TELETHON 2026 - SECURE ADMIN PASSWORD UPDATE
   Paste this at the VERY BOTTOM of your existing Code.gs.

   IMPORTANT:
   Your old ADMIN_PASSWORD constant may remain in the file.
   These functions override the old password checking functions.
   The password is stored in Script Properties after initialization.
   ===================================================== */

const ADMIN_PASSWORD_PROPERTY = 'TELETHON_ADMIN_PASSWORD';

/* First-time setup:
   Uses your existing ADMIN_PASSWORD value only once, then saves it
   into Script Properties. After that, password changes are managed
   from the website and no Code.gs editing is needed.
*/
function getStoredAdminPassword_() {
  const props = PropertiesService.getScriptProperties();
  let password = props.getProperty(ADMIN_PASSWORD_PROPERTY);

  if (!password) {
    // Initial migration from your existing password constant
    password = ADMIN_PASSWORD;
    props.setProperty(ADMIN_PASSWORD_PROPERTY, password);
  }
  return password;
}

/* Override old verification function */
/* removed duplicate function */





/* Override old login checker */
/* removed duplicate function */



/* Change password directly from Admin Panel */
/* removed duplicate function */










































/* Optional emergency reset function.
   Run manually from Apps Script only if you completely forget the password.
   After running it, the password becomes: GJhabibi89@
*/
function emergencyResetAdminPassword() {
  PropertiesService.getScriptProperties()
    .setProperty(ADMIN_PASSWORD_PROPERTY, 'GJhabibi89@');
  return 'Admin password reset successfully.';
}


/* =====================================================
   FINAL STABILITY FIXES — TELETHON 2026
   These helpers make password migration and setup reliable.
   ===================================================== */

function initializeAdminPasswordStorage() {
  const props = PropertiesService.getScriptProperties();
  if (!props.getProperty(ADMIN_PASSWORD_PROPERTY)) {
    props.setProperty(ADMIN_PASSWORD_PROPERTY, ADMIN_PASSWORD);
  }
  return {success:true};
}

function getSystemHealth(password) {
  verifyAdmin(password);
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  return {
    success:true,
    participants: !!ss.getSheetByName(PARTICIPANTS_SHEET),
    coordinators: !!ss.getSheetByName(COORDINATORS_SHEET),
    dailyCollections: !!ss.getSheetByName(DAILY_SHEET),
    activityLog: !!ss.getSheetByName(AUDIT_SHEET)
  };
}




/* ============================================================
   TELETHON 2026 — MULTI-ADMIN ROLE & REGION ACCESS SYSTEM
   Super Admin + Region Admins
   ============================================================ */

const AUTH_SESSION_PREFIX = 'TELETHON_SESSION_';
const AUTH_SESSION_TTL = 21600; // 6 hours

const USER_ACCOUNTS = [
  {
    name: 'GJ Habibi',
    email: 'ghulamjilani277@gmail.com',
    role: 'SUPER_ADMIN',
    regions: ['ALL'],
    passwordType: 'SUPER'
  },
  {
    name: 'Mumbai Region Admin',
    email: 'jtmqiraatmumbairegion@gmail.com',
    role: 'REGION_ADMIN',
    regions: ['Mumbai'],
    password: 'Madina2526'
  },
  {
    name: 'Hyderabad & Bangalore Region Admin',
    email: 'jtmqirat.hydbadregion@gmail.com',
    role: 'REGION_ADMIN',
    regions: ['Hyderabad','Bangalore'],
    password: 'Madina2526'
  },
  {
    name: 'Delhi Region Part 1 Admin',
    email: 'jtmqiraatbareillyregion@gmail.com',
    role: 'REGION_ADMIN',
    regions: ['Delhi'],
    parts: ['Part 1'],
    password: 'Madina2526'
  },
  {
    name: 'Delhi Region Part 2 Admin',
    email: 'jtmqiraatdelhiregion@gmail.com',
    role: 'REGION_ADMIN',
    regions: ['Delhi'],
    parts: ['Part 2'],
    password: 'Madina2526'
  },
  {
    name: 'Kolkata Region Admin',
    email: 'jtmqiraatkolkataregion@gmail.com',
    role: 'REGION_ADMIN',
    regions: ['Kolkata'],
    password: 'Madina2526'
  },
  {
    name: 'Ajmer Region Part 1 Admin',
    email: 'jtmqiraatrajasthan@gmail.com',
    role: 'REGION_ADMIN',
    regions: ['Ajmer'],
    parts: ['Part 1'],
    password: 'Madina2526'
  },
  {
    name: 'Ajmer Region Part 2 Admin',
    email: 'jtmqiraatgujarat@gmail.com',
    role: 'REGION_ADMIN',
    regions: ['Ajmer'],
    parts: ['Part 2'],
    password: 'Madina2526'
  }
];

function normalizeEmail_(email) {
  return String(email || '').trim().toLowerCase();
}

function findUserAccount_(email) {
  const key = normalizeEmail_(email);
  return USER_ACCOUNTS.find(u => normalizeEmail_(u.email) === key) || null;
}

/* removed duplicate function */

































/* removed duplicate function */







function assertSuperAdmin_(token) {
  const user = getAuthUser_(token);
  if (user.role !== 'SUPER_ADMIN') throw new Error('Super Admin access required.');
  return user;
}

function normalizeRegionName_(value) {
  return canonicalRegion_(value);
}

function normalizePartName_(value) {
  return canonicalPart_(value);
}

// ===== STRICT REGION / PART ACCESS =====
// This is intentionally strict, but understands the actual sheet formats:
// Region keys such as "delhi1" / "ajmer2" and labels such as
// "Delhi Region – Part 1" / "Ajmer Region – Part 2".
function participantMatchesUser_(user, participant) {
  if (user.role === 'SUPER_ADMIN') return true;

  const participantRegion = canonicalRegion_(
    participant.region || participant.regionLabel || participant.regionKey || ''
  );
  const participantPart = canonicalPart_(
    participant.part || '',
    (participant.regionKey || '') + ' ' + (participant.region || participant.regionLabel || '')
  );

  const allowedRegions = (user.regions || []).map(canonicalRegion_);
  const allowedParts = (user.parts || []).map(p => canonicalPart_(p));

  if (!allowedRegions.includes(participantRegion)) return false;

  // For Part-wise admins BOTH Region AND Part must match exactly.
  if (allowedParts.length > 0 && !allowedParts.includes(participantPart)) return false;

  return true;
}

function regionMatches_(user, region, part) {
  return participantMatchesUser_(user, {
    region: region,
    regionKey: region,
    part: part
  });
}
function assertParticipantAccess_(token, participantId) {
  const user = getAuthUser_(token);
  if (user.role === 'SUPER_ADMIN') return user;

  const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(PARTICIPANTS_SHEET);
  if (!sh || sh.getLastRow() < 2) throw new Error('Participants sheet not found.');

  const hm = getHeaderMap(sh);
  const idCol = findColumn(hm,['ID']);
  const regionCol = findColumn(hm,['Region','RegionKey']);
  const regionLabelCol = findColumn(hm,['RegionLabel','RegionName']);
  const partCol = findColumn(hm,['Part']);

  const rows = sh.getDataRange().getValues();
  for (let i=1;i<rows.length;i++) {
    if (String(rows[i][idCol]) === String(participantId)) {
      const regionKey = regionCol>=0 ? rows[i][regionCol] : '';
      const regionLabel = regionLabelCol>=0 ? rows[i][regionLabelCol] : '';
      const part = partCol>=0 ? rows[i][partCol] : '';
      if (!participantMatchesUser_(user, {region:regionLabel, regionKey:regionKey, part:part})) {
        throw new Error('You do not have permission to modify this Region.');
      }
      return user;
    }
  }
  throw new Error('Participant not found.');
}

/* Override verification so legacy admin actions now require a valid session token */
function verifyAdmin(token) {
  return getAuthUser_(token);
}

/* Keep old compatibility endpoint disabled for password-only logins */
function checkAdminPassword(password) {
  return String(password) === String(getStoredAdminPassword_());
}

/* Region-filtered admin data */
function getAdminData(selectedDate, token) {
  const user = getAuthUser_(token);
  const rows = getDataInternal_(selectedDate || 'all');

  // Authoritative server-side filtering.
  const participants = rows.filter(x => participantMatchesUser_(user, x));

  // Coordinator information for the logged-in Region / Part.
  let coordinator = {name:'', mobile:''};
  if (user.role !== 'SUPER_ADMIN') {
    const cm = getCoordinatorMap_();
    const region = (user.regions || [])[0] || '';
    const part = (user.parts || [])[0] || 'Main';
    coordinator = getCoordinatorFor_(region, region, part, cm);
  }

  return {
    success: true,
    user: user,
    participants: participants,
    count: participants.length,
    coordinator: coordinator
  };
}

/* Override amount update with region permission check */
function updateAmount(id, dateStr, amount, token) {
  assertParticipantAccess_(token, id);

  dateStr = normaliseDate_(dateStr);
  const newAmount = Math.max(0, Number(amount)||0);

  const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(PARTICIPANTS_SHEET);
  if (!sh) throw new Error('Participants sheet not found');

  const hm = getHeaderMap(sh);
  const idCol = findColumn(hm,['ID']);
  const amountCol = findColumn(hm,['Amount']);
  const updatedCol = findColumn(hm,['UpdatedAt','Updated']);
  if (idCol<0 || amountCol<0) throw new Error('ID or Amount column not found');

  const ids = sh.getRange(2,idCol+1,sh.getLastRow()-1,1).getValues();
  let rowNo = -1;
  for (let i=0;i<ids.length;i++) {
    if (String(ids[i][0])===String(id)) { rowNo=i+2; break; }
  }
  if (rowNo<0) throw new Error('Participant not found');

  let oldAmount = 0;

  if (dateStr==='all') {
    oldAmount = Number(sh.getRange(rowNo,amountCol+1).getValue())||0;
    sh.getRange(rowNo,amountCol+1).setValue(newAmount);
    if (updatedCol>=0) sh.getRange(rowNo,updatedCol+1).setValue(new Date());

    try { logActivity_('AMOUNT CORRECTED', String(id), '', 'ALL', oldAmount, newAmount, 'Overall amount updated'); } catch(e) {}
    return {success:true,overall:newAmount};
  }

  const dsh = getOrCreateDailySheet_();
  let foundRow=-1;

  if (dsh.getLastRow()>=2) {
    const drows=dsh.getRange(2,1,dsh.getLastRow()-1,4).getValues();
    for (let i=0;i<drows.length;i++) {
      const raw=drows[i][0];
      const key=raw instanceof Date
        ? Utilities.formatDate(raw,Session.getScriptTimeZone(),'yyyy-MM-dd')
        : String(raw);
      if (key===dateStr && String(drows[i][1])===String(id)) {
        foundRow=i+2;
        oldAmount=Number(drows[i][2])||0;
        break;
      }
    }
  }

  if (foundRow>0) {
    dsh.getRange(foundRow,3).setValue(newAmount);
    dsh.getRange(foundRow,4).setValue(new Date());
  } else {
    dsh.appendRow([dateStr,String(id),newAmount,new Date()]);
  }

  const currentOverall=Number(sh.getRange(rowNo,amountCol+1).getValue())||0;
  const overall=Math.max(0,currentOverall+(newAmount-oldAmount));
  sh.getRange(rowNo,amountCol+1).setValue(overall);
  if (updatedCol>=0) sh.getRange(rowNo,updatedCol+1).setValue(new Date());

  try { logActivity_('DAILY AMOUNT UPDATED', String(id), '', dateStr, oldAmount, newAmount, 'Overall auto-adjusted'); } catch(e) {}

  return {success:true,daily:newAmount,overall:overall};
}

/* Override add participant: only Super Admin can add */
function addParticipant(region,name,centre,token) {
  assertSuperAdmin_(token);

  const sh=SpreadsheetApp.getActiveSpreadsheet().getSheetByName(PARTICIPANTS_SHEET);
  if (!sh) throw new Error('Participants sheet not found');

  const hm=getHeaderMap(sh), row=Array(sh.getLastColumn()).fill(''), lastRow=sh.getLastRow();
  const set=(names,val)=>{const c=findColumn(hm,names); if(c>=0) row[c]=val;};

  set(['ID'],'P'+String(lastRow).padStart(3,'0'));
  set(['Region','RegionKey'],region);
  set(['RegionLabel'],region);
  set(['Part'],'Main');
  set(['QariName','Name'],name);
  set(['CityCentre','City','Centre'],centre);
  set(['Amount'],0);
  set(['UnitRate'],RATE);
  set(['Status'],'active');
  set(['CreatedAt','Created'],new Date());
  set(['UpdatedAt','Updated'],new Date());

  sh.appendRow(row);
  try { logActivity_('PARTICIPANT ADDED', '', name, 'SYSTEM', 0, 0, region+' - '+centre); } catch(e) {}
  return {success:true};
}

/* Only Super Admin can reset all */
function resetAllAmounts(token) {
  assertSuperAdmin_(token);

  const sh=SpreadsheetApp.getActiveSpreadsheet().getSheetByName(PARTICIPANTS_SHEET);
  if (!sh) throw new Error('Participants sheet not found');

  const hm=getHeaderMap(sh), amountCol=findColumn(hm,['Amount']), updatedCol=findColumn(hm,['UpdatedAt','Updated']), rows=sh.getLastRow()-1;
  if (rows>0 && amountCol>=0) sh.getRange(2,amountCol+1,rows,1).setValues(Array.from({length:rows},()=>[0]));
  if (rows>0 && updatedCol>=0) sh.getRange(2,updatedCol+1,rows,1).setValues(Array.from({length:rows},()=>[new Date()]));

  const dsh=getOrCreateDailySheet_();
  if (dsh.getLastRow()>1) dsh.getRange(2,1,dsh.getLastRow()-1,4).clearContent();

  try { logActivity_('SYSTEM RESET', '', '', 'SYSTEM', 0, 0, 'All collection data reset by Super Admin'); } catch(e) {}
  return {success:true};
}

/* Override delete daily entry with region restriction */
function deleteDailyEntry(id, dateStr, token) {
  assertParticipantAccess_(token, id);
  dateStr = normaliseDate_(dateStr);
  if (dateStr === 'all') throw new Error('Select a specific date to delete.');

  const dsh = getOrCreateDailySheet_();
  if (dsh.getLastRow()<2) return {success:true,deleted:0};

  const rows = dsh.getRange(2,1,dsh.getLastRow()-1,4).getValues();
  let target=-1, old=0;

  for(let i=0;i<rows.length;i++){
    const raw=rows[i][0];
    const key=raw instanceof Date ? Utilities.formatDate(raw,Session.getScriptTimeZone(),'yyyy-MM-dd') : String(raw);
    if(key===dateStr && String(rows[i][1])===String(id)){
      target=i+2;
      old=Number(rows[i][2])||0;
      break;
    }
  }

  if(target<0) return {success:true,deleted:0};

  const psh=SpreadsheetApp.getActiveSpreadsheet().getSheetByName(PARTICIPANTS_SHEET);
  const hm=getHeaderMap(psh), idCol=findColumn(hm,['ID']), amountCol=findColumn(hm,['Amount']);
  const ids=psh.getRange(2,idCol+1,psh.getLastRow()-1,1).getValues();
  for(let i=0;i<ids.length;i++){
    if(String(ids[i][0])===String(id)){
      const row=i+2;
      const overall=Math.max(0,(Number(psh.getRange(row,amountCol+1).getValue())||0)-old);
      psh.getRange(row,amountCol+1).setValue(overall);
      break;
    }
  }

  dsh.deleteRow(target);
  try { logActivity_('DAILY ENTRY DELETED', String(id), '', dateStr, old, 0, 'Entry deleted and overall corrected'); } catch(e) {}
  return {success:true,deleted:old};
}

/* Region-filtered history */
function getParticipantHistory(id, token) {
  assertParticipantAccess_(token, id);
  return getParticipantHistoryLegacy_(id);
}

/* Region Admin can see only their own audit-relevant participants */
function getAuditLog(token, limit) {
  const user = getAuthUser_(token);
  const rows = getAuditLogLegacy_(limit);
  if (user.role === 'SUPER_ADMIN') return rows;

  const allowed = {};
  const participants = getDataInternal_('all');
  participants.forEach(p=>{
    if(regionMatches_(user,p.regionKey||p.region,p.part)) allowed[String(p.id)] = true;
  });

  return rows.filter(r => !r.participantId || allowed[String(r.participantId)]);
}

/* Super Admin password can be changed only by Super Admin session */
function changeAdminPassword(token, currentPassword, newPassword) {
  const actor = assertSuperAdmin_(token);

  if (String(currentPassword) !== String(getStoredAdminPassword_())) {
    throw new Error('Current password is incorrect.');
  }

  newPassword = String(newPassword || '');
  if (newPassword.length < 8) throw new Error('New password must contain at least 8 characters.');
  if (!/[A-Za-z]/.test(newPassword) || !/[0-9]/.test(newPassword)) {
    throw new Error('Password must contain at least one letter and one number.');
  }

  // Keep Script Properties and the V2 Users sheet in sync.
  // Without this, a Super Admin whose account exists in Users could be locked out
  // because V2 login checks PasswordHash first.
  PropertiesService.getScriptProperties().setProperty(ADMIN_PASSWORD_PROPERTY, newPassword);

  try {
    const sh = getOrCreateUsersSheet_();
    if (sh.getLastRow() >= 2) {
      const rows = sh.getDataRange().getValues();
      const superEmail = normalizeEmail_(actor.email || 'ghulamjilani277@gmail.com');
      for (let i = 1; i < rows.length; i++) {
        const email = normalizeEmail_(rows[i][2]);
        const role = String(rows[i][4] || '').toUpperCase();
        if (email === superEmail || role === 'SUPER_ADMIN') {
          sh.getRange(i + 1, 4).setValue(sha256_(newPassword));
          sh.getRange(i + 1, 12).setValue(new Date());
        }
      }
    }
  } catch (e) {
    // Script Properties remains authoritative for the legacy Super Admin fallback.
    // Do not fail the password change only because the optional Users sheet is unavailable.
  }

  try { logActivity_('SUPER ADMIN PASSWORD CHANGED', '', actor.name || 'Super Admin', 'SYSTEM', 0, 0, 'Password changed from Admin Panel'); } catch(e) {}

  return {success:true,message:'Password changed successfully.'};
}

/* Legacy function aliases captured safely */
function getParticipantHistoryLegacy_(id) {
  const dsh = getOrCreateDailySheet_();
  if (dsh.getLastRow()<2) return [];
  const rows=dsh.getRange(2,1,dsh.getLastRow()-1,4).getValues();
  return rows.filter(r=>String(r[1])===String(id)).map(r=>({
    date:r[0] instanceof Date ? Utilities.formatDate(r[0],Session.getScriptTimeZone(),'yyyy-MM-dd') : String(r[0]),
    amount:Number(r[2])||0,
    updatedAt:r[3] instanceof Date ? Utilities.formatDate(r[3],Session.getScriptTimeZone(),'yyyy-MM-dd HH:mm:ss') : String(r[3]||'')
  })).sort((a,b)=>String(b.date).localeCompare(String(a.date)));
}

function getAuditLogLegacy_(limit) {
  const ss=SpreadsheetApp.getActiveSpreadsheet();
  const sh=ss.getSheetByName(AUDIT_SHEET);
  if(!sh || sh.getLastRow()<2) return [];
  const rows=sh.getDataRange().getValues().slice(1).reverse();
  return rows.slice(0,Number(limit)||100).map(r=>({
    timestamp:r[0] instanceof Date ? Utilities.formatDate(r[0],Session.getScriptTimeZone(),'yyyy-MM-dd HH:mm:ss') : String(r[0]||''),
    action:String(r[1]||''),
    participantId:String(r[2]||''),
    name:String(r[3]||''),
    date:String(r[4]||''),
    oldAmount:Number(r[5])||0,
    newAmount:Number(r[6])||0,
    notes:String(r[7]||'')
  }));
}


/* ============================================================
   TELETHON 2026 — ROLE-BASED PRIVATE PORTALS + COORDINATOR PROFILE
   FINAL PORTAL LAYER
   ============================================================ */

const COORDINATOR_PROFILES_SHEET = 'CoordinatorProfiles';
const PROFILE_PHOTOS_FOLDER = 'TELETHON 2026 Coordinator Profile Photos';

function getOrCreateCoordinatorProfilesSheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName(COORDINATOR_PROFILES_SHEET);
  if (!sh) {
    sh = ss.insertSheet(COORDINATOR_PROFILES_SHEET);
    sh.appendRow(['Email','Name','Mobile','Role','Regions','Parts','PhotoUrl','PhotoFileId','UpdatedAt']);
    sh.setFrozenRows(1);
  }
  return sh;
}

function getProfileFolder_() {
  const folders = DriveApp.getFoldersByName(PROFILE_PHOTOS_FOLDER);
  return folders.hasNext() ? folders.next() : DriveApp.createFolder(PROFILE_PHOTOS_FOLDER);
}

function defaultProfileForUser_(user) {
  let name = user.name || '';
  let mobile = '';

  if (user.role !== 'SUPER_ADMIN') {
    const cm = getCoordinatorMap_();
    const region = (user.regions || [])[0] || '';
    const part = (user.parts || [])[0] || 'Main';
    const c = getCoordinatorFor_(region, region, part, cm);
    if (c.name) name = c.name;
    mobile = c.mobile || '';
  }

  return {
    email: user.email || '',
    name: name,
    mobile: mobile,
    role: user.role || '',
    regions: user.regions || [],
    parts: user.parts || [],
    photoUrl: ''
  };
}

function getProfileForUser_(user) {
  const base = defaultProfileForUser_(user);
  const sh = getOrCreateCoordinatorProfilesSheet_();
  if (sh.getLastRow() < 2) return base;

  const rows = sh.getDataRange().getValues();
  const email = normalizeEmail_(user.email);
  for (let i = 1; i < rows.length; i++) {
    if (normalizeEmail_(rows[i][0]) === email) {
      return {
        email: String(rows[i][0] || base.email),
        name: String(rows[i][1] || base.name),
        mobile: String(rows[i][2] || base.mobile),
        role: String(rows[i][3] || base.role),
        regions: String(rows[i][4] || '').split(',').map(x=>x.trim()).filter(Boolean),
        parts: String(rows[i][5] || '').split(',').map(x=>x.trim()).filter(Boolean),
        photoUrl: String(rows[i][6] || '')
      };
    }
  }
  return base;
}

function getMyProfile(token) {
  const user = getAuthUser_(token);
  return { success:true, profile:getProfileForUser_(user) };
}

function saveMyProfile(token, profile) {
  const user = getAuthUser_(token);
  profile = profile || {};

  const name = String(profile.name || '').trim();
  const mobile = String(profile.mobile || '').trim();
  if (!name) throw new Error('Profile name is required.');

  const sh = getOrCreateCoordinatorProfilesSheet_();
  const email = normalizeEmail_(user.email);
  const values = [
    user.email,
    name,
    mobile,
    user.role,
    (user.regions || []).join(', '),
    (user.parts || []).join(', '),
    getProfileForUser_(user).photoUrl || '',
    '',
    new Date()
  ];

  let rowNo = -1;
  if (sh.getLastRow() >= 2) {
    const emails = sh.getRange(2,1,sh.getLastRow()-1,1).getValues();
    for (let i=0;i<emails.length;i++) {
      if (normalizeEmail_(emails[i][0]) === email) { rowNo=i+2; break; }
    }
  }

  if (rowNo > 0) {
    // Preserve existing photo file ID.
    const oldFileId = String(sh.getRange(rowNo,8).getValue() || '');
    values[7] = oldFileId;
    sh.getRange(rowNo,1,1,values.length).setValues([values]);
  } else {
    sh.appendRow(values);
  }

  try { logActivity_('PROFILE UPDATED','',name,'PROFILE',0,0,'Coordinator profile updated'); } catch(e) {}
  return {success:true, profile:getProfileForUser_(user)};
}

function uploadMyProfilePhoto(token, dataUrl, fileName, mimeType) {
  const user = getAuthUser_(token);
  const raw = String(dataUrl || '');
  if (!raw || raw.indexOf(',') < 0) throw new Error('Invalid image data.');

  const base64 = raw.split(',')[1];
  const bytes = Utilities.base64Decode(base64);
  if (bytes.length > 5 * 1024 * 1024) throw new Error('Image is too large. Please upload an image under 5 MB.');

  const safeName = String(fileName || 'profile.jpg').replace(/[^A-Za-z0-9._-]/g,'_');
  const type = String(mimeType || 'image/jpeg');
  if (type.indexOf('image/') !== 0) throw new Error('Only image files are allowed.');

  const blob = Utilities.newBlob(bytes, type, safeName);
  const folder = getProfileFolder_();
  const file = folder.createFile(blob);
  file.setName(normalizeEmail_(user.email).replace(/[^a-z0-9]/g,'_') + '_' + new Date().getTime() + '_' + safeName);

  // Public view URL is used only for displaying the coordinator profile photo in this portal.
  file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
  const photoUrl = 'https://drive.google.com/thumbnail?id=' + file.getId() + '&sz=w500';

  const sh = getOrCreateCoordinatorProfilesSheet_();
  const email = normalizeEmail_(user.email);
  let rowNo = -1;
  if (sh.getLastRow() >= 2) {
    const emails = sh.getRange(2,1,sh.getLastRow()-1,1).getValues();
    for (let i=0;i<emails.length;i++) {
      if (normalizeEmail_(emails[i][0]) === email) { rowNo=i+2; break; }
    }
  }

  if (rowNo < 0) {
    const base = defaultProfileForUser_(user);
    sh.appendRow([
      user.email, base.name, base.mobile, user.role,
      (user.regions||[]).join(', '), (user.parts||[]).join(', '),
      photoUrl, file.getId(), new Date()
    ]);
  } else {
    const oldFileId = String(sh.getRange(rowNo,8).getValue() || '');
    if (oldFileId) {
      try { DriveApp.getFileById(oldFileId).setTrashed(true); } catch(e) {}
    }
    sh.getRange(rowNo,7,1,3).setValues([[photoUrl,file.getId(),new Date()]]);
  }

  try { logActivity_('PROFILE PHOTO UPDATED','',user.name,'PROFILE',0,0,'Coordinator profile photo uploaded'); } catch(e) {}
  return {success:true, photoUrl:photoUrl, profile:getProfileForUser_(user)};
}
/* removed duplicate function */


























// Final login response enriched with the coordinator profile.
/* removed duplicate function */



























// Protected compatibility endpoint. Public dashboard data can never be fetched
// without an authenticated session token.
function getData(selectedDate, token) {
  const user = getAuthUser_(token);
  const rows = getDataInternal_(selectedDate || 'all');
  return user.role === 'SUPER_ADMIN' ? rows : rows.filter(x => participantMatchesUser_(user, x));
}

// ================= TELETHON 2026 PORTAL V2 PROFESSIONAL LAYER =================
const USERS_SHEET = 'Users';
const TARGETS_SHEET = 'Targets';
const TRASH_SHEET = 'Trash';

function sha256_(text){
  const bytes=Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256,String(text),Utilities.Charset.UTF_8);
  return bytes.map(b=>{const v=(b+256)%256;return ('0'+v.toString(16)).slice(-2)}).join('');
}
function ensureSheet_(name,headers){
  const ss=SpreadsheetApp.getActiveSpreadsheet(); let sh=ss.getSheetByName(name);
  if(!sh){sh=ss.insertSheet(name);sh.appendRow(headers);sh.setFrozenRows(1);}
  return sh;
}
function getOrCreateUsersSheet_(){return ensureSheet_(USERS_SHEET,['UserID','Name','Email','PasswordHash','Role','Regions','Parts','Mobile','PhotoURL','Status','CreatedAt','UpdatedAt','LastLogin']);}
function getOrCreateTargetsSheet_(){return ensureSheet_(TARGETS_SHEET,['TargetID','Region','Part','TargetAmount','StartDate','EndDate','Active','UpdatedBy','UpdatedAt']);}
function getOrCreateTrashSheet_(){return ensureSheet_(TRASH_SHEET,['TrashID','OriginalID','Region','Part','Name','Centre','Amount','DeletedBy','DeletedAt','OriginalRowJSON']);}

function setupPortalV2(token){
  // May be run once by Super Admin; if no token is available, use this only from script editor.
  if(token) assertSuperAdmin_(token);
  const sh=getOrCreateUsersSheet_();
  if(sh.getLastRow()<2 && typeof USER_ACCOUNTS!=='undefined'){
    const rows=USER_ACCOUNTS.map((u,i)=>[
      'U'+String(i+1).padStart(3,'0'),u.name,u.email,
      sha256_(u.passwordType==='SUPER'?getStoredAdminPassword_():u.password||'Madina2526'),
      u.role||'REGION_ADMIN',(u.regions||[]).join(','),(u.parts||[]).join(','),u.mobile||'',u.photo||'',
      'ACTIVE',new Date(),new Date(),''
    ]);
    if(rows.length) sh.getRange(2,1,rows.length,rows[0].length).setValues(rows);
  }
  getOrCreateTargetsSheet_(); getOrCreateTrashSheet_(); getOrCreateAuditSheet_();
  return {success:true,message:'Portal V2 sheets ready'};
}
function parseUserRow_(r){return {id:String(r[0]),name:String(r[1]||''),email:normalizeEmail_(r[2]),passwordHash:String(r[3]||''),role:String(r[4]||'REGION_ADMIN'),regions:String(r[5]||'').split(',').map(x=>x.trim()).filter(Boolean),parts:String(r[6]||'').split(',').map(x=>x.trim()).filter(Boolean),mobile:String(r[7]||''),photoUrl:String(r[8]||''),status:String(r[9]||'ACTIVE').toUpperCase()};}
function getDynamicUser_(email){const sh=getOrCreateUsersSheet_();if(sh.getLastRow()<2)return null;const rows=sh.getDataRange().getValues();for(let i=1;i<rows.length;i++){const u=parseUserRow_(rows[i]);if(u.email===normalizeEmail_(email))return {...u,rowNo:i+1};}return null;}
function authenticateUser(email,password){
  // FINAL LOGIN FIX:
  // The portal must work even before the V2 Users sheet has been initialized.
  const loginEmail = normalizeEmail_(email);
  const loginPassword = String(password || '');
  let user = getDynamicUser_(loginEmail);

  if(!user){
    // Legacy/default account fallback.
    const legacy = findUserAccount_(loginEmail);
    if(!legacy) throw new Error('This email is not registered.');

    const valid = legacy.passwordType === 'SUPER'
      ? loginPassword === String(getStoredAdminPassword_())
      : loginPassword === String(legacy.password);

    if(!valid) throw new Error('Invalid email or password.');

    const result = createAuthSession_(legacy);
    const profile = getProfileForUser_(result.user);

    // IMPORTANT: old createAuthSession_ does not return success:true.
    // The login page checks res.success, so we must normalize the response.
    return {
      success:true,
      token:result.token,
      user:result.user,
      profile:profile
    };
  }

  if(user.status !== 'ACTIVE') {
    throw new Error('This account is disabled. Contact Super Admin.');
  }

  if(sha256_(loginPassword) !== String(user.passwordHash || '')) {
    throw new Error('Invalid email or password.');
  }

  const sh = getOrCreateUsersSheet_();
  sh.getRange(user.rowNo,13).setValue(new Date());

  const result = createAuthSessionV2_(user);
  try { logActivity_('LOGIN',user.id,user.name,'SYSTEM',0,0,user.email); } catch(e) {}

  return {
    success:true,
    token:result.token,
    user:result.user,
    profile:{name:user.name,mobile:user.mobile,photoUrl:user.photoUrl}
  };
}
function listUsers(token){assertSuperAdmin_(token);const sh=getOrCreateUsersSheet_();if(sh.getLastRow()<2)return [];return sh.getDataRange().getValues().slice(1).map(parseUserRow_).map(u=>({id:u.id,name:u.name,email:u.email,role:u.role,regions:u.regions,parts:u.parts,mobile:u.mobile,photoUrl:u.photoUrl,status:u.status}));}
function createUser(token,data){const actor=assertSuperAdmin_(token);const d=data||{};const email=normalizeEmail_(d.email);if(!email||!String(d.name||'').trim()||!String(d.password||''))throw new Error('Name, Email and Password are required.');if(String(d.password).length<8)throw new Error('Password must be at least 8 characters.');if(getDynamicUser_(email))throw new Error('This Email ID already exists.');const role=['SUPER_ADMIN','REGION_ADMIN','PART_ADMIN'].includes(d.role)?d.role:'REGION_ADMIN';const regions=Array.isArray(d.regions)?d.regions:[];const parts=Array.isArray(d.parts)?d.parts:[];const sh=getOrCreateUsersSheet_();const id='U'+Utilities.getUuid().slice(0,8).toUpperCase();sh.appendRow([id,String(d.name).trim(),email,sha256_(d.password),role,regions.join(','),parts.join(','),String(d.mobile||''),'','ACTIVE',new Date(),new Date(),'']);logActivity_('USER CREATED',id,d.name,'SYSTEM',0,0,'Created by '+actor.email);return {success:true,id};}
function updateUser(token,id,data){const actor=assertSuperAdmin_(token),sh=getOrCreateUsersSheet_();const rows=sh.getDataRange().getValues();let rowNo=-1;for(let i=1;i<rows.length;i++)if(String(rows[i][0])===String(id)){rowNo=i+1;break;}if(rowNo<0)throw new Error('User not found.');const old=parseUserRow_(rows[rowNo-1]);const d=data||{};const role=d.role||old.role,regions=(Array.isArray(d.regions)?d.regions:old.regions),parts=(Array.isArray(d.parts)?d.parts:old.parts);sh.getRange(rowNo,2,1,11).setValues([[String(d.name||old.name),normalizeEmail_(d.email||old.email),old.passwordHash,role,regions.join(','),parts.join(','),String(d.mobile??old.mobile),String(d.photoUrl??old.photoUrl),String(d.status||old.status).toUpperCase(),rows[rowNo-1][10],new Date()]]);if(d.password)sh.getRange(rowNo,4).setValue(sha256_(d.password));logActivity_('USER UPDATED',id,String(d.name||old.name),'SYSTEM',0,0,'Updated by '+actor.email);return {success:true};}
function toggleUserStatus(token,id,status){return updateUser(token,id,{status:status});}
function deleteUser(token,id){const actor=assertSuperAdmin_(token),sh=getOrCreateUsersSheet_();const rows=sh.getDataRange().getValues();for(let i=1;i<rows.length;i++){if(String(rows[i][0])===String(id)){const email=String(rows[i][2]);if(normalizeEmail_(email)===normalizeEmail_(actor.email))throw new Error('You cannot delete your own active Super Admin account.');sh.deleteRow(i+1);logActivity_('USER DELETED',id,email,'SYSTEM',0,0,'Deleted by '+actor.email);return {success:true};}}throw new Error('User not found.');}
function resetUserPassword(token,id,newPassword){assertSuperAdmin_(token);if(String(newPassword||'').length<8)throw new Error('Minimum 8 characters required.');const sh=getOrCreateUsersSheet_(),rows=sh.getDataRange().getValues();for(let i=1;i<rows.length;i++)if(String(rows[i][0])===String(id)){sh.getRange(i+1,4).setValue(sha256_(newPassword));sh.getRange(i+1,12).setValue(new Date());return {success:true};}throw new Error('User not found.');}
function getTargets(token){const user=getAuthUser_(token),sh=getOrCreateTargetsSheet_();const all=sh.getLastRow()<2?[]:sh.getDataRange().getValues().slice(1).map(r=>({id:r[0],region:r[1],part:r[2],target:Number(r[3])||0,start:r[4],end:r[5],active:String(r[6]).toUpperCase()==='TRUE'||String(r[6]).toUpperCase()==='ACTIVE'}));return user.role==='SUPER_ADMIN'?all:all.filter(t=>participantMatchesUser_(user,{region:t.region,part:t.part}));}
function saveTarget(token,data){
  const actor=getAuthUser_(token),d=data||{};
  if(!d.region||!Number(d.target)||Number(d.target)<0)throw new Error('Region and valid Target Amount are required.');
  const role=String(actor.role||'').toUpperCase();
  const actorRegions=(actor.regions||[]).map(String);
  const actorParts=(actor.parts||[]).map(String);
  let region=String(d.region).trim(), part=String(d.part||'Main').trim()||'Main';
  if(role!=='SUPER_ADMIN'){
    if(!['REGION_ADMIN','PART_ADMIN'].includes(role))throw new Error('You are not allowed to manage targets.');
    if(actorRegions.length && actorRegions.indexOf(region)===-1)throw new Error('You can set targets only for your assigned region(s).');
    if(role==='PART_ADMIN' && actorParts.length){
      if(actorParts.indexOf(part)===-1)throw new Error('You can set targets only for your assigned part(s).');
    }
  }
  const sh=getOrCreateTargetsSheet_();
  const id=d.id||('T'+Utilities.getUuid().slice(0,8).toUpperCase());
  if(d.id){
    const vals=sh.getDataRange().getValues();
    for(let i=1;i<vals.length;i++)if(String(vals[i][0])===String(d.id)){
      const oldRegion=String(vals[i][1]||''),oldPart=String(vals[i][2]||'Main');
      if(role!=='SUPER_ADMIN'){
        if(actorRegions.length && actorRegions.indexOf(oldRegion)===-1)throw new Error('You cannot edit another region target.');
        if(role==='PART_ADMIN' && actorParts.length && actorParts.indexOf(oldPart)===-1)throw new Error('You cannot edit another part target.');
      }
      sh.getRange(i+1,1,1,9).setValues([[id,region,part,Number(d.target),d.start||'',d.end||'',d.active!==false?'ACTIVE':'INACTIVE',actor.email,new Date()]]);
      try{logActivity_('TARGET UPDATED',id,region+' '+part,'SYSTEM',0,Number(d.target),actor.email);}catch(e){}
      return {success:true,id};
    }
  }
  sh.appendRow([id,region,part,Number(d.target),d.start||'',d.end||'',d.active!==false?'ACTIVE':'INACTIVE',actor.email,new Date()]);
  try{logActivity_('TARGET CREATED',id,region+' '+part,'SYSTEM',0,Number(d.target),actor.email);}catch(e){}
  return {success:true,id};
}
function softDeleteParticipant(token,id){const actor=getAuthUser_(token);assertParticipantAccess_(token,id);const sh=SpreadsheetApp.getActiveSpreadsheet().getSheetByName(PARTICIPANTS_SHEET);const hm=getHeaderMap(sh),idCol=findColumn(hm,['ID']);const rows=sh.getDataRange().getValues();let idx=-1;for(let i=1;i<rows.length;i++)if(String(rows[i][idCol])===String(id)){idx=i;break;}if(idx<0)throw new Error('Participant not found.');const r=rows[idx],trash=getOrCreateTrashSheet_();trash.appendRow(['TR'+Utilities.getUuid().slice(0,8),id,r[findColumn(hm,['Region','RegionKey'])],r[findColumn(hm,['Part'])],r[findColumn(hm,['QariName','Name'])],r[findColumn(hm,['CityCentre','City','Centre'])],r[findColumn(hm,['Amount'])],actor.email,new Date(),JSON.stringify(r)]);sh.deleteRow(idx+1);logActivity_('QARI MOVED TO TRASH',id,String(r[findColumn(hm,['QariName','Name'])]),'SYSTEM',0,0,'By '+actor.email);return {success:true};}
function listTrash(token){assertSuperAdmin_(token);const sh=getOrCreateTrashSheet_();if(sh.getLastRow()<2)return [];return sh.getDataRange().getValues().slice(1).map(r=>({trashId:r[0],id:r[1],region:r[2],part:r[3],name:r[4],centre:r[5],amount:r[6],deletedBy:r[7],deletedAt:r[8]}));}
function restoreParticipant(token,trashId){assertSuperAdmin_(token);const trash=getOrCreateTrashSheet_(),rows=trash.getDataRange().getValues();for(let i=1;i<rows.length;i++)if(String(rows[i][0])===String(trashId)){const data=JSON.parse(rows[i][9]);SpreadsheetApp.getActiveSpreadsheet().getSheetByName(PARTICIPANTS_SHEET).appendRow(data);trash.deleteRow(i+1);return {success:true};}throw new Error('Trash record not found.');}
function getProfessionalAnalytics(token,dateStr){const user=getAuthUser_(token),rows=getDataInternal_(dateStr||'all').filter(x=>participantMatchesUser_(user,x));const total=rows.reduce((s,x)=>s+(Number(x.amount)||0),0),active=rows.filter(x=>Number(x.amount)>0).length,top=rows.slice().sort((a,b)=>b.amount-a.amount).slice(0,10);const groups={};rows.forEach(x=>{const k=(x.region||'')+'|'+(x.part||'Main');if(!groups[k])groups[k]={region:x.region,part:x.part,total:0,count:0};groups[k].total+=Number(x.amount)||0;groups[k].count++;});const targets=getTargets(token);const targetTotal=targets.filter(t=>t.active).reduce((s,t)=>s+t.target,0);return {total,count:rows.length,active,average:rows.length?total/rows.length:0,top,regions:Object.values(groups).sort((a,b)=>b.total-a.total),targetTotal,remaining:Math.max(0,targetTotal-total),progress:targetTotal?Math.min(100,total*100/targetTotal):0};}
function getNotifications(token,dateStr){const a=getProfessionalAnalytics(token,dateStr||'all'),n=[];if(a.progress>=100)n.push({type:'success',text:'🎉 Target achieved!'});else if(a.progress>=75)n.push({type:'info',text:'🔥 Target is above 75%.'});if(a.active<a.count)n.push({type:'warning',text:'⚠️ '+(a.count-a.active)+' Qari Sahiban have no collection in this view.'});if(a.top[0])n.push({type:'top',text:'🏆 Current leader: '+a.top[0].name+' — ₹'+a.top[0].amount.toLocaleString('en-IN')});return n;}


/* ============================================================
   LOGIN EMERGENCY REPAIR / BOOTSTRAP
   Run repairPortalLoginAccounts() once from Apps Script editor
   only if the Users sheet was created with bad or incomplete data.
   This rebuilds the default portal accounts from USER_ACCOUNTS.
   ============================================================ */

function repairPortalLoginAccounts() {
  const sh = getOrCreateUsersSheet_();
  const existing = {};

  if (sh.getLastRow() >= 2) {
    const rows = sh.getDataRange().getValues();
    for (let i = 1; i < rows.length; i++) {
      existing[normalizeEmail_(rows[i][2])] = i + 1;
    }
  }

  // IMPORTANT: do not clear the Users sheet. Custom accounts created by Super Admin
  // must survive a login repair. We only insert/update the built-in default accounts.
  USER_ACCOUNTS.forEach((u,i) => {
    const email = normalizeEmail_(u.email);
    const row = [
      'U'+String(i+1).padStart(3,'0'),
      u.name,
      email,
      sha256_(u.passwordType==='SUPER' ? getStoredAdminPassword_() : (u.password || 'Madina2526')),
      u.role || 'REGION_ADMIN',
      (u.regions || []).join(','),
      (u.parts || []).join(','),
      u.mobile || '',
      u.photo || '',
      'ACTIVE',
      new Date(),
      new Date(),
      ''
    ];

    const rowNo = existing[email];
    if (rowNo) {
      // Preserve any profile/mobile/photo fields already customized in Users.
      const old = sh.getRange(rowNo,1,1,13).getValues()[0];
      row[0] = old[0] || row[0];
      row[7] = old[7] || row[7];
      row[8] = old[8] || row[8];
      row[10] = old[10] || row[10];
      row[12] = old[12] || row[12];
      sh.getRange(rowNo,1,1,13).setValues([row]);
    } else {
      sh.appendRow(row);
    }
  });

  return {
    success:true,
    message:'Default portal login accounts repaired successfully without deleting custom users.',
    accounts:USER_ACCOUNTS.length
  };
}

function portalLoginHealthCheck() {
  const sh = getOrCreateUsersSheet_();
  return {
    usersSheetExists:true,
    users:Math.max(0,sh.getLastRow()-1),
    defaultAccounts:USER_ACCOUNTS.length,
    superAdmin:'ghulamjilani277@gmail.com',
    defaultRegionPassword:'Madina2526'
  };
}

/* ============================================================
   FINAL STABILITY PATCH — MANAGEMENT BOOTSTRAP + EMAIL OTP
   ============================================================ */
const PASSWORD_OTP_PREFIX = 'TELETHON_PASSWORD_OTP_';
const PASSWORD_OTP_TTL = 600; // 10 minutes
/* removed duplicate function */









function requestPasswordResetOtp(email){
  const normalized=normalizeEmail_(email);
  if(!normalized) throw new Error('Registered email is required.');

  let user=getDynamicUser_(normalized);
  if(!user){
    const legacy=findUserAccount_(normalized);
    if(!legacy) throw new Error('This email is not registered.');
    user={id:'LEGACY_'+normalized,name:legacy.name,email:legacy.email,role:legacy.role,status:'ACTIVE',legacy:true};
  }
  if(String(user.status||'ACTIVE').toUpperCase()!=='ACTIVE') throw new Error('This account is disabled.');

  const cache=CacheService.getScriptCache();
  const rateKey=PASSWORD_OTP_PREFIX+'RATE_'+normalized;
  if(cache.get(rateKey)) throw new Error('Please wait 60 seconds before requesting another OTP.');

  const otp=String(Math.floor(100000+Math.random()*900000));
  const payload={hash:sha256_(otp),attempts:0};
  cache.put(PASSWORD_OTP_PREFIX+normalized,JSON.stringify(payload),PASSWORD_OTP_TTL);
  cache.put(rateKey,'1',60);

  MailApp.sendEmail({
    to:normalized,
    subject:'TELETHON 2026 — Password Reset OTP',
    htmlBody:'<div style="font-family:Arial,sans-serif;max-width:560px;margin:auto;padding:24px;border:1px solid #ddd;border-radius:14px"><h2>TELETHON 2026</h2><p>Hello '+String(user.name||'User').replace(/</g,'&lt;')+',</p><p>Your password reset OTP is:</p><div style="font-size:30px;font-weight:bold;letter-spacing:8px;padding:14px;background:#f3f4f6;text-align:center">'+otp+'</div><p>This OTP expires in <b>10 minutes</b>. Do not share it with anyone.</p></div>',
    body:'TELETHON 2026 Password Reset OTP: '+otp+'\nValid for 10 minutes. Do not share this code.'
  });
  return {success:true,message:'OTP sent to your registered email. It is valid for 10 minutes.'};
}

function resetPasswordWithOtp(email,otp,newPassword){
  const normalized=normalizeEmail_(email);
  const password=String(newPassword||'');
  if(!/^\d{6}$/.test(String(otp||''))) throw new Error('Enter a valid 6-digit OTP.');
  if(password.length<8) throw new Error('New password must be at least 8 characters.');

  const cache=CacheService.getScriptCache();
  const key=PASSWORD_OTP_PREFIX+normalized;
  const raw=cache.get(key);
  if(!raw) throw new Error('OTP expired or not found. Please request a new OTP.');
  const payload=JSON.parse(raw);
  if((Number(payload.attempts)||0)>=5){cache.remove(key);throw new Error('Too many incorrect attempts. Request a new OTP.');}
  if(sha256_(String(otp))!==String(payload.hash||'')){
    payload.attempts=(Number(payload.attempts)||0)+1;
    cache.put(key,JSON.stringify(payload),PASSWORD_OTP_TTL);
    throw new Error('Incorrect OTP.');
  }

  let user=getDynamicUser_(normalized);
  if(user){
    const sh=getOrCreateUsersSheet_();
    sh.getRange(user.rowNo,4).setValue(sha256_(password));
    sh.getRange(user.rowNo,12).setValue(new Date());
  }else{
    const legacy=findUserAccount_(normalized);
    if(!legacy) throw new Error('This email is not registered.');
    const sh=getOrCreateUsersSheet_();
    const id='U'+Utilities.getUuid().slice(0,8).toUpperCase();
    const regions=(legacy.regions||[]).join(',');
    const parts=(legacy.parts||[]).join(',');
    sh.appendRow([id,legacy.name,legacy.email,sha256_(password),legacy.role||'REGION_ADMIN',regions,parts,'','', 'ACTIVE',new Date(),new Date(),'']);
  }

  if(normalized===normalizeEmail_('ghulamjilani277@gmail.com')){
    PropertiesService.getScriptProperties().setProperty(ADMIN_PASSWORD_PROPERTY,password);
  }
  cache.remove(key);
  try{logActivity_('PASSWORD RESET','',normalized,'SYSTEM',0,0,'Password reset using Email OTP');}catch(e){}
  return {success:true,message:'Password changed successfully.'};
}

/* ================= PROFESSIONAL PERFORMANCE SUITE ================= */
const QARI_NOTES_SHEET='QariNotes';
function getOrCreateQariNotesSheet_(){return ensureSheet_(QARI_NOTES_SHEET,['NoteID','ParticipantID','Note','Status','CreatedBy','CreatedAt','UpdatedAt']);}
function getQariNotes(token,participantId){const u=getAuthUser_(token);assertParticipantAccess_(token,participantId);const sh=getOrCreateQariNotesSheet_();if(sh.getLastRow()<2)return [];return sh.getDataRange().getValues().slice(1).filter(r=>String(r[1])===String(participantId)).map(r=>({id:r[0],participantId:r[1],note:r[2],status:r[3],createdBy:r[4],createdAt:r[5],updatedAt:r[6]}));}
function saveQariNote(token,participantId,note,status){const u=getAuthUser_(token);assertParticipantAccess_(token,participantId);note=String(note||'').trim();if(!note)throw new Error('Note cannot be empty.');const sh=getOrCreateQariNotesSheet_();sh.appendRow(['N'+Utilities.getUuid().slice(0,8).toUpperCase(),String(participantId),note,String(status||'Follow-up Required'),u.email,new Date(),new Date()]);logActivity_('QARI NOTE',String(participantId),'Qari Note','SYSTEM',0,0,'Status: '+String(status||''));return {success:true};}
/* removed duplicate function */
function getPerformanceSuite(token,dateStr){const user=getAuthUser_(token),rows=getDataInternal_(dateStr||'all').filter(x=>participantMatchesUser_(user,x));const sorted=rows.slice().sort((a,b)=>Number(b.amount)-Number(a.amount));const avg=rows.length?rows.reduce((s,x)=>s+(Number(x.amount)||0),0)/rows.length:0;const low=rows.filter(x=>Number(x.amount)<avg*0.35).sort((a,b)=>Number(a.amount)-Number(b.amount)).slice(0,20);const badges={};sorted.forEach((x,i)=>{const b=[];if(i===0)b.push('🥇 Winner');if(i<3)b.push('🏆 Top Performer');if(Number(x.amount)>=RATE)b.push('💯 Target Achiever');if(Number(x.amount)>=avg*1.5&&Number(x.amount)>0)b.push('⚡ Fast Performer');badges[x.id]=b;});return {average:avg,lowPerformers:low,badges:badges,leaderboard:sorted.slice(0,10)};}
function getReportData(token,dateStr){const user=getAuthUser_(token);const rows=getDataInternal_(dateStr||'all').filter(x=>participantMatchesUser_(user,x));return {generatedAt:new Date(),date:dateStr||'all',rows:rows.map(x=>({name:x.name,centre:x.centre,region:x.region,part:x.part,amount:Number(x.amount)||0,units:(Number(x.amount)||0)/RATE,status:x.status||'Active'}))};}


// ===== QARI-WISE TARGETS (OVERALL + DAILY) =====
const QARI_TARGETS_SHEET = 'QariTargets';
function getOrCreateQariTargetsSheet_(){
  return ensureSheet_(QARI_TARGETS_SHEET,['TargetID','ParticipantID','TargetType','TargetAmount','TargetDate','Region','Part','Active','UpdatedBy','UpdatedAt']);
}
function getParticipantForTarget_(id){
  const rows=getDataInternal_('all');
  const p=rows.find(x=>String(x.id)===String(id));
  if(!p) throw new Error('Qari Sahab not found');
  return p;
}
function getQariTargets(token, participantId){
  const user=getAuthUser_(token), sh=getOrCreateQariTargetsSheet_();
  let rows=sh.getLastRow()<2?[]:sh.getDataRange().getValues().slice(1).map(r=>({id:r[0],participantId:String(r[1]),type:String(r[2]||'OVERALL'),target:Number(r[3])||0,date:r[4]||'',region:r[5]||'',part:r[6]||'',active:String(r[7]).toUpperCase()!=='INACTIVE'}));
  rows=rows.filter(x=>participantMatchesUser_(user,{region:x.region,part:x.part}));
  if(participantId) rows=rows.filter(x=>x.participantId===String(participantId));
  return rows;
}
function saveQariTarget(token,data){
  const actor=getAuthUser_(token), d=data||{};
  const participantId=String(d.participantId||'').trim();
  const type=String(d.type||'OVERALL').toUpperCase();
  const target=Number(d.target)||0;
  if(!participantId||!target||target<0) throw new Error('Select Qari Sahab and enter a valid target amount.');
  if(['OVERALL','DAILY'].indexOf(type)===-1) throw new Error('Invalid target type.');
  const p=getParticipantForTarget_(participantId);
  if(!participantMatchesUser_(actor,p)) throw new Error('You can set targets only for Qari Sahiban in your assigned region/part.');
  const targetDate=type==='DAILY'?normaliseDate_(d.date):'';
  if(type==='DAILY'&&targetDate==='all') throw new Error('Select a date for Daily Target.');
  const sh=getOrCreateQariTargetsSheet_();
  const vals=sh.getLastRow()<2?[]:sh.getDataRange().getValues();
  let rowNo=-1;
  for(let x=1;x<vals.length;x++){
    if(String(vals[x][1])===participantId && String(vals[x][2]).toUpperCase()===type && String(vals[x][4]||'')===String(targetDate||'')){rowNo=x+1;break;}
  }
  const id=rowNo>0?String(vals[rowNo-1][0]):('QT'+Utilities.getUuid().slice(0,8).toUpperCase());
  const row=[id,participantId,type,target,targetDate,p.region||'',p.part||'Main','ACTIVE',actor.email,new Date()];
  if(rowNo>0) sh.getRange(rowNo,1,1,row.length).setValues([row]); else sh.appendRow(row);
  try{logActivity_('QARI TARGET '+(rowNo>0?'UPDATED':'CREATED'),participantId,p.name,targetDate||'OVERALL',0,target,actor.email);}catch(e){}
  return {success:true,id:id};
}
function getQariTargetSummary(token, participantId, dateStr){
  const p=getParticipantForTarget_(participantId), user=getAuthUser_(token);
  if(!participantMatchesUser_(user,p)) throw new Error('Access denied');
  const date=normaliseDate_(dateStr||'all'), targets=getQariTargets(token,participantId);
  const overall=targets.filter(x=>x.type==='OVERALL'&&x.active).sort((a,b)=>String(b.id).localeCompare(String(a.id)))[0];
  const daily=date==='all'?null:targets.filter(x=>x.type==='DAILY'&&String(x.date)===date&&x.active)[0];
  const data=getDataInternal_(date).find(x=>String(x.id)===String(participantId))||{};
  return {overallTarget:overall?overall.target:0,dailyTarget:daily?daily.target:0,amount:Number(data.amount)||0};
}

/* ===== FINAL CLEAN ACCESS / CALENDAR REPAIR ===== */
function getManagementBootstrap(token){
  const u=getAuthUser_(token);
  const role=String(u.role||'').toUpperCase();
  if(!['SUPER_ADMIN','REGION_ADMIN','PART_ADMIN'].includes(role)) throw new Error('Management access denied.');
  return {success:true,role:role,user:{name:u.name,email:u.email,regions:u.regions||[],parts:u.parts||[]}};
}

function getQariPerformanceGraph(token,participantId){
  assertParticipantAccess_(token,participantId);
  const sh=getOrCreateDailySheet_(), map={};
  if(sh.getLastRow()>=2){
    sh.getDataRange().getValues().slice(1).forEach(r=>{
      const d=r[0] instanceof Date?Utilities.formatDate(r[0],Session.getScriptTimeZone(),'yyyy-MM-dd'):String(r[0]||'');
      if(d && String(r[1])===String(participantId)) map[d]=Number(r[2])||0;
    });
  }
  return Object.keys(map).sort().map(d=>({date:d,amount:map[d]}));
}


/* ================= QUICK CARD COLLECTION MODE ================= */
function quickAdjustAmount(participantId, dateStr, delta, token) {
  const user = getAuthUser_(token);
  assertParticipantAccess_(token, participantId);
  dateStr = normaliseDate_(dateStr || 'all');
  delta = Number(delta) || 0;
  if (!delta) throw new Error('Amount must be greater than zero.');

  let current = 0;
  if (dateStr === 'all') {
    const rows = getDataInternal_('all');
    const x = rows.find(r=>String(r.id)===String(participantId));
    if (!x) throw new Error('Participant not found');
    current = Number(x.amount)||0;
  } else {
    const daily = getDataInternal_(dateStr);
    const x = daily.find(r=>String(r.id)===String(participantId));
    if (!x) throw new Error('Participant not found');
    current = Number(x.amount)||0;
  }
  const next = Math.max(0, current + delta);
  const result = updateAmount(participantId, dateStr, next, token);
  try { logActivity_(delta>0?'QUICK ADD':'QUICK MINUS',String(participantId),user.name||'',dateStr,current,next,'Quick card update'); } catch(e) {}
  return Object.assign({success:true, previous:current, current:next}, result||{});
}

function getPortalData(selectedDate, token) {
  const user = getAuthUser_(token);
  selectedDate = normaliseDate_(selectedDate || 'all');
  const allRows = getDataInternal_(selectedDate);
  const participants = user.role === 'SUPER_ADMIN' ? allRows : allRows.filter(x => participantMatchesUser_(user, x));

  // Qari's own phone number comes from Participants sheet when available.
  const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(PARTICIPANTS_SHEET);
  if (sh && sh.getLastRow() > 1) {
    const rows = sh.getDataRange().getValues();
    const hm = getHeaderMap(sh);
    const idCol = findColumn(hm,['ID']);
    const mobileCol = findColumn(hm,['Mobile','Phone','MobileNumber','QariMobile','Contact','Number']);
    if (idCol >= 0 && mobileCol >= 0) {
      const phoneMap = {};
      for (let i=1;i<rows.length;i++) phoneMap[String(rows[i][idCol])] = String(rows[i][mobileCol]||'').trim();
      participants.forEach(x=>x.qariPhone=phoneMap[String(x.id)]||'');
    }
  }
  participants.forEach(x=>{ if(x.qariPhone===undefined) x.qariPhone=''; });

  const profile = getProfileForUser_(user);
  const scopeTitle = user.role === 'SUPER_ADMIN' ? 'All India — Super Admin Portal' : ((user.regions || []).join(' & ') + ((user.parts && user.parts.length) ? ' — ' + user.parts.join(' & ') : '') + ' Portal');
  return {success:true,user:user,profile:profile,scopeTitle:scopeTitle,selectedDate:selectedDate,participants:participants};
}


/* ================= FINAL DIRECT CARD ACTION REPAIR ================= */
function ensureParticipantColumn_(sheet, headerMap, acceptedNames, newHeader) {
  let col = findColumn(headerMap, acceptedNames);
  if (col >= 0) return col;
  const next = sheet.getLastColumn() + 1;
  sheet.getRange(1, next).setValue(newHeader);
  headerMap[String(newHeader).trim().toLowerCase().replace(/[\s_\-\/]/g,'')] = next - 1;
  return next - 1;
}

function updateParticipantContact(token, participantId, details) {
  const user = assertParticipantAccess_(token, participantId);
  const d = details || {};
  const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(PARTICIPANTS_SHEET);
  if (!sh || sh.getLastRow() < 2) throw new Error('Participants sheet not found.');
  const hm = getHeaderMap(sh);
  const idCol = findColumn(hm,['ID']);
  if (idCol < 0) throw new Error('ID column not found.');

  let rowNo = -1;
  const ids = sh.getRange(2,idCol+1,sh.getLastRow()-1,1).getValues();
  for (let i=0;i<ids.length;i++) {
    if (String(ids[i][0]) === String(participantId)) { rowNo=i+2; break; }
  }
  if (rowNo < 0) throw new Error('Participant not found.');

  const phone = String(d.phone || '').trim();
  if (phone && !/^[+0-9()\-\s]{7,20}$/.test(phone)) {
    throw new Error('Please enter a valid phone number.');
  }

  if (Object.prototype.hasOwnProperty.call(d,'phone')) {
    const mobileCol = ensureParticipantColumn_(sh, hm, ['Mobile','Phone','MobileNumber','QariMobile','Contact','Number'], 'Mobile');
    sh.getRange(rowNo,mobileCol+1).setValue(phone);
  }
  if (Object.prototype.hasOwnProperty.call(d,'centre')) {
    const centreCol = ensureParticipantColumn_(sh, hm, ['CityCentre','City','Centre','Jamia'], 'Centre');
    sh.getRange(rowNo,centreCol+1).setValue(String(d.centre||'').trim());
  }
  if (Object.prototype.hasOwnProperty.call(d,'name') && String(d.name||'').trim()) {
    const nameCol = ensureParticipantColumn_(sh, hm, ['QariName','Name','Qari Sahab'], 'QariName');
    sh.getRange(rowNo,nameCol+1).setValue(String(d.name).trim());
  }
  try { logActivity_('QARI DETAILS UPDATED',String(participantId),'', 'PROFILE',0,0,'Updated by '+(user.email||'')); } catch(e) {}
  return {success:true};
}
// Dedicated direct-card endpoint. Keeps daily and overall logic in one secure path.
function quickCardUpdate(token, participantId, dateStr, delta) {
  const user = assertParticipantAccess_(token, participantId);
  dateStr = normaliseDate_(dateStr || 'all');
  delta = Number(delta);
  if (!isFinite(delta) || delta === 0) throw new Error('Amount must be greater than zero.');
  const rows = getDataInternal_(dateStr);
  const row = rows.find(function(x){ return String(x.id)===String(participantId); });
  if (!row) throw new Error('Participant not found.');
  const current = Number(row.amount)||0;
  const next = Math.max(0,current+delta);
  const result = updateAmount(participantId,dateStr,next,token);
  try { logActivity_(delta>0?'CARD ADD':'CARD MINUS',String(participantId),String(row.name||''),dateStr,current,next,'Direct Qari card'); } catch(e) {}
  return Object.assign({success:true,previous:current,current:next},result||{});
}


/* ================= AUTH SESSION V4 — SHEET-BACKED STABLE SESSIONS =================
   Session records are stored in the bound spreadsheet instead of CacheService.
   This avoids random cache eviction and makes the same token available to every
   Apps Script execution while keeping server-side authorization intact. */
const AUTH_SESSION_SHEET_V4 = '_TELETHON_AUTH_SESSIONS_';
const AUTH_SESSION_TTL_V4_MS = 24 * 60 * 60 * 1000;

function getAuthSessionSheetV4_(){
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName(AUTH_SESSION_SHEET_V4);
  if(!sh){
    sh = ss.insertSheet(AUTH_SESSION_SHEET_V4);
    sh.appendRow(['Token','User JSON','Expires At','Updated At']);
    try { sh.hideSheet(); } catch(e) {}
  }
  return sh;
}

function cleanupAuthSessionsV4_(sh){
  sh = sh || getAuthSessionSheetV4_();
  const last = sh.getLastRow();
  if(last < 2) return;
  const values = sh.getRange(2,1,last-1,4).getValues();
  const now = Date.now();
  const deleteRows = [];
  values.forEach(function(r,idx){
    const exp = r[2] instanceof Date ? r[2].getTime() : new Date(r[2]).getTime();
    if(!r[0] || !isFinite(exp) || now > exp) deleteRows.push(idx+2);
  });
  for(let j=deleteRows.length-1;j>=0;j--) sh.deleteRow(deleteRows[j]);
}

function findAuthSessionRowV4_(sh, token){
  const last = sh.getLastRow();
  if(last < 2) return -1;
  const tokens = sh.getRange(2,1,last-1,1).getValues();
  for(let i=0;i<tokens.length;i++) if(String(tokens[i][0])===String(token)) return i+2;
  return -1;
}

function createAuthSessionV4_(user){
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try{
    const sh = getAuthSessionSheetV4_();
    cleanupAuthSessionsV4_(sh);
    const now = Date.now();
    const token = Utilities.getUuid() + Utilities.getUuid().replace(/-/g,'');
    const safe = {
      id:String(user.id||''), name:String(user.name||''), email:String(user.email||''),
      role:String(user.role||''), regions:Array.isArray(user.regions)?user.regions:[],
      parts:Array.isArray(user.parts)?user.parts:[], mobile:String(user.mobile||''),
      photoUrl:String(user.photoUrl||''), status:String(user.status||'ACTIVE'),
      loginAt:new Date(now).toISOString(), expiresAt:now + AUTH_SESSION_TTL_V4_MS
    };
    sh.appendRow([token, JSON.stringify(safe), new Date(safe.expiresAt), new Date(now)]);
    return {token:token,user:safe};
  } finally { lock.releaseLock(); }
}

function getAuthUser_(token){
  token = String(token||'').trim();
  if(!token) throw new Error('Session expired. Please login again.');
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try{
    const sh = getAuthSessionSheetV4_();
    const rowNo = findAuthSessionRowV4_(sh,token);
    if(rowNo < 0) throw new Error('Session expired. Please login again.');
    const row = sh.getRange(rowNo,1,1,4).getValues()[0];
    let user;
    try { user = JSON.parse(String(row[1]||'')); } catch(e) { sh.deleteRow(rowNo); throw new Error('Session expired. Please login again.'); }
    const expires = row[2] instanceof Date ? row[2].getTime() : new Date(row[2]).getTime();
    if(!isFinite(expires) || Date.now() > expires){ sh.deleteRow(rowNo); throw new Error('Session expired. Please login again.'); }
    if(String(user.status||'ACTIVE').toUpperCase() !== 'ACTIVE') throw new Error('This account is disabled. Contact Super Admin.');
    // Sliding expiry while the user is active.
    user.expiresAt = Date.now() + AUTH_SESSION_TTL_V4_MS;
    sh.getRange(rowNo,2,1,3).setValues([[JSON.stringify(user),new Date(user.expiresAt),new Date()]]);
    return user;
  } finally { lock.releaseLock(); }
}

function createAuthSession_(user){ return createAuthSessionV4_(user); }
function createAuthSessionV2_(user){ return createAuthSessionV4_(user); }

function logoutUser(token){
  token = String(token||'').trim();
  if(!token) return {success:true};
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try{
    const sh = getAuthSessionSheetV4_();
    const rowNo = findAuthSessionRowV4_(sh,token);
    if(rowNo >= 0) sh.deleteRow(rowNo);
  } finally { lock.releaseLock(); }
  return {success:true};
}



/* ================= FOLLOW-UP NOTIFICATION CENTER =================
   Super Admin creates a Qari follow-up request. Responsible Region/Part Admins
   see it in their own portal and update live status. Super Admin can see both
   the status and the responsible Admin's contact details for direct follow-up. */
const CONTACT_ALERTS_SHEET_V1 = 'Contact Followups';

function getContactAlertsSheetV1_(){
  const ss=SpreadsheetApp.getActiveSpreadsheet(); let sh=ss.getSheetByName(CONTACT_ALERTS_SHEET_V1);
  const headers=['AlertID','ParticipantID','QariName','Region','Part','Centre','Phone','AmountSnapshot','Note','Status','CreatedByEmail','CreatedByName','CreatedAt','ResolvedByEmail','ResolvedByName','ResolvedAt','ResolutionNote'];
  if(!sh){sh=ss.insertSheet(CONTACT_ALERTS_SHEET_V1);sh.appendRow(headers);sh.setFrozenRows(1);}
  else {
    const existing=sh.getRange(1,1,1,Math.max(1,sh.getLastColumn())).getValues()[0].map(String);
    headers.forEach(function(h){if(existing.indexOf(h)<0){sh.getRange(1,sh.getLastColumn()+1).setValue(h);existing.push(h);}});
  }
  return sh;
}
function getParticipantPhoneForAlert_(participantId){
  const sh=SpreadsheetApp.getActiveSpreadsheet().getSheetByName(PARTICIPANTS_SHEET);if(!sh||sh.getLastRow()<2)return '';
  const hm=getHeaderMap(sh),idCol=findColumn(hm,['ID']),mobileCol=findColumn(hm,['Mobile','Phone','MobileNumber','QariMobile','Contact','Number']);
  if(idCol<0||mobileCol<0)return '';const vals=sh.getDataRange().getValues();
  for(let i=1;i<vals.length;i++)if(String(vals[i][idCol])===String(participantId))return String(vals[i][mobileCol]||'').trim();return '';
}
function getAllPortalUsersForFollowup_(){
  const out=[];try{const sh=getOrCreateUsersSheet_();if(sh.getLastRow()>=2)sh.getDataRange().getValues().slice(1).forEach(function(r){const u=parseUserRow_(r);if(String(u.status||'ACTIVE').toUpperCase()==='ACTIVE')out.push(u);});}catch(e){}
  if(!out.length&&typeof USER_ACCOUNTS!=='undefined')USER_ACCOUNTS.forEach(function(u){out.push({name:u.name||'',email:normalizeEmail_(u.email),role:u.role||'REGION_ADMIN',regions:u.regions||[],parts:u.parts||[],mobile:u.mobile||'',status:'ACTIVE'});});
  return out;
}
function findResponsibleAdminsForParticipant_(p){
  return getAllPortalUsersForFollowup_().filter(function(u){return String(u.role||'').toUpperCase()!=='SUPER_ADMIN'&&participantMatchesUser_(u,p);}).map(function(u){return {name:String(u.name||''),email:String(u.email||''),mobile:String(u.mobile||''),role:String(u.role||'REGION_ADMIN')};});
}
function createContactAlert(token,participantId,note){
  const user=getAuthUser_(token);if(String(user.role||'').toUpperCase()!=='SUPER_ADMIN')throw new Error('Only Super Admin can create contact follow-up requests.');
  const rows=getDataInternal_('all'),p=rows.find(function(x){return String(x.id)===String(participantId);});if(!p)throw new Error('Qari Sahab not found.');
  const sh=getContactAlertsSheetV1_(),lock=LockService.getScriptLock();lock.waitLock(10000);try{
    const id='CF-'+Utilities.getUuid(),hm=getHeaderMap(sh),row=[];const data={AlertID:id,ParticipantID:String(p.id||''),QariName:String(p.name||''),Region:String(p.region||p.regionKey||''),Part:String(p.part||''),Centre:String(p.centre||''),Phone:getParticipantPhoneForAlert_(p.id),AmountSnapshot:Number(p.overallAmount!==undefined?p.overallAmount:p.amount)||0,Note:String(note||'').trim(),Status:'OPEN',CreatedByEmail:String(user.email||''),CreatedByName:String(user.name||'Super Admin'),CreatedAt:new Date(),ResolvedByEmail:'',ResolvedByName:'',ResolvedAt:'',ResolutionNote:''};
    const headers=sh.getRange(1,1,1,sh.getLastColumn()).getValues()[0];headers.forEach(function(h){row.push(data[String(h)]!==undefined?data[String(h)]:'');});sh.appendRow(row);
    try{logActivity_('CONTACT FOLLOW-UP CREATED',String(p.id||''),String(p.name||''),'FOLLOWUP',0,Number(p.overallAmount||p.amount||0),'Sent to responsible Region Admin');}catch(e){}
    return {success:true,id:id,assignedAdmins:findResponsibleAdminsForParticipant_(p)};
  }finally{lock.releaseLock();}
}
function getMyContactAlerts(token){
  const user=getAuthUser_(token),sh=getContactAlertsSheetV1_();if(sh.getLastRow()<2)return {success:true,alerts:[]};
  const values=sh.getDataRange().getValues(),headers=values[0].map(String),idx={};headers.forEach(function(h,i){idx[h]=i;});const alerts=[];
  const allRows=getDataInternal_('all');
  for(let r=1;r<values.length;r++){
    const row=values[r],participantId=String(row[idx.ParticipantID]||''),p=allRows.find(function(x){return String(x.id)===participantId;});
    if(String(user.role||'').toUpperCase()!=='SUPER_ADMIN'&&(!p||!participantMatchesUser_(user,p)))continue;
    const a={id:String(row[idx.AlertID]||''),participantId:participantId,qariName:String(row[idx.QariName]||''),region:String(row[idx.Region]||''),part:String(row[idx.Part]||''),centre:String(row[idx.Centre]||''),phone:String(row[idx.Phone]||''),amountSnapshot:Number(row[idx.AmountSnapshot])||0,note:String(row[idx.Note]||''),status:String(row[idx.Status]||'OPEN').toUpperCase(),createdByEmail:String(row[idx.CreatedByEmail]||''),createdByName:String(row[idx.CreatedByName]||''),createdAt:row[idx.CreatedAt]||'',resolvedByEmail:String(row[idx.ResolvedByEmail]||''),resolvedByName:String(row[idx.ResolvedByName]||''),resolvedAt:row[idx.ResolvedAt]||'',resolutionNote:String(row[idx.ResolutionNote]||'')};
    a.assignedAdmins=p?findResponsibleAdminsForParticipant_(p):[];alerts.push(a);
  }
  alerts.sort(function(a,b){return new Date(b.createdAt||0).getTime()-new Date(a.createdAt||0).getTime();});return {success:true,alerts:alerts};
}
function resolveContactAlert(token,alertId,status,resolutionNote){
  const user=getAuthUser_(token),sh=getContactAlertsSheetV1_();if(sh.getLastRow()<2)throw new Error('Follow-up request not found.');
  const values=sh.getDataRange().getValues(),headers=values[0].map(String),idx={};headers.forEach(function(h,i){idx[h]=i;});let rowNo=-1;for(let i=1;i<values.length;i++)if(String(values[i][idx.AlertID])===String(alertId)){rowNo=i+1;break;}if(rowNo<0)throw new Error('Follow-up request not found.');
  const participantId=String(values[rowNo-1][idx.ParticipantID]||''),p=getDataInternal_('all').find(function(x){return String(x.id)===participantId;});if(String(user.role||'').toUpperCase()!=='SUPER_ADMIN'&&(!p||!participantMatchesUser_(user,p)))throw new Error('You do not have permission to update this follow-up.');
  const next=String(status||'OPEN').toUpperCase(),allowed=['OPEN','CONTACTED','DONE','NO_ANSWER','NOT_RECEIVING','CALLBACK'];if(allowed.indexOf(next)<0)throw new Error('Invalid follow-up status.');
  const set=function(name,val){if(idx[name]!==undefined)sh.getRange(rowNo,idx[name]+1).setValue(val);};set('Status',next);set('ResolvedByEmail',String(user.email||''));set('ResolvedByName',String(user.name||''));set('ResolvedAt',new Date());set('ResolutionNote',String(resolutionNote||''));
  try{logActivity_('CONTACT FOLLOW-UP '+next,participantId,'','FOLLOWUP',0,0,'Updated by '+(user.email||''));}catch(e){}return {success:true,status:next};
}


/**
 * e-Rise Library API v3
 * Sheets: Students, Books, Transactions, Reviews, Settings
 *
 * SETUP: paste into Extensions > Apps Script, Deploy > New deployment > Web app
 * (Execute as: Me, Who has access: Anyone), copy the /exec URL into API_URL in the HTML.
 * If you already have a deployment, use Deploy > Manage deployments > edit > New version.
 */

function sheet_(name, headers){
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName(name);
  if(!sh){ sh = ss.insertSheet(name); sh.appendRow(headers); }
  return sh;
}
function ensureColumn_(sh, name){
  const lastCol = Math.max(sh.getLastColumn(),1);
  const headers = sh.getRange(1,1,1,lastCol).getValues()[0];
  if(headers.indexOf(name)===-1){
    sh.getRange(1, lastCol+1).setValue(name);
  }
}
function studentsSheet_(){
  const sh = sheet_('Students', ['id','code','name','grade','parentEmail','photo','remindersEnabled','createdAt','archived']);
  ensureColumn_(sh, 'photo');
  ensureColumn_(sh, 'remindersEnabled');
  ensureColumn_(sh, 'archived');
  return sh;
}
function booksSheet_(){
  const sh = sheet_('Books', ['id','code','title','author','category','level','shelf','photo','createdAt','archived']);
  ensureColumn_(sh, 'photo');
  ensureColumn_(sh, 'archived');
  return sh;
}
function txSheet_(){
  const sh = sheet_('Transactions', ['id','studentId','studentName','bookId','bookTitle','borrowedAt','dueAt','returnedAt','status','lastReminderAt','preDeleteStatus']);
  ensureColumn_(sh, 'preDeleteStatus');
  return sh;
}
function reviewsSheet_(){ return sheet_('Reviews', ['id','bookId','bookTitle','studentId','studentName','rating','comment','createdAt']); }
function settingsSheet_(){ return sheet_('Settings', ['key','value']); }

function rowsToObjects_(sh){
  const data = sh.getDataRange().getValues();
  const headers = data[0];
  const out = [];
  for(let i=1;i<data.length;i++){
    if(data[i].every(c=>c==='')) continue;
    const obj = {};
    headers.forEach((h,idx)=>{ let v=data[i][idx]; if(v instanceof Date) v=v.toISOString(); obj[h]=v; });
    out.push(obj);
  }
  return out;
}
function findRowIndex_(sh, colIdx, id){
  const data = sh.getDataRange().getValues();
  for(let i=1;i<data.length;i++){ if(String(data[i][colIdx])===String(id)) return i+1; }
  return -1;
}
function genId_(){ return Utilities.getUuid().slice(0,8); }
function genCode_(prefix){ return prefix+'-'+Utilities.getUuid().slice(0,5).toUpperCase(); }
function respond_(obj){ return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON); }

/* ---------- settings ---------- */
function getSettings_(){
  const rows = rowsToObjects_(settingsSheet_());
  const s = {};
  rows.forEach(r=> s[r.key]=r.value);
  s.libraryName = s.libraryName || 'e-Rise Library';
  s.loanDays = Number(s.loanDays)||14;
  s.maxBooks = Number(s.maxBooks)||2;
  s.pin = String(s.pin || '1234');
  s.dailyPin = String(s.dailyPin || '0000');
  s.language = s.language || 'en';
  s.reminderDays = Number(s.reminderDays)||2;
  s.remindersEnabled = (s.remindersEnabled===true || s.remindersEnabled==='true');
  s.fineEnabled = (s.fineEnabled===true || s.fineEnabled==='true');
  s.fineAmount = Number(s.fineAmount)||0;
  if(!s.apiKey){
    s.apiKey = Utilities.getUuid();
    saveSettings_(s); // persist the newly generated key immediately
  }
  return s;
}
function saveSettings_(newSettings){
  const sh = settingsSheet_();
  const last = sh.getLastRow();
  if(last>1) sh.getRange(2,1,last-1,2).clearContent();
  let r=2;
  Object.keys(newSettings).forEach(k=>{
    const cell = sh.getRange(r,2);
    if(k==='pin' || k==='dailyPin') cell.setNumberFormat('@'); // force plain text, stops auto-number conversion
    sh.getRange(r,1,1,2).setValues([[k, String(newSettings[k])]]);
    r++;
  });
}

/* ---------- router ---------- */
function doGet(e){ return respond_({error:'Use POST'}); }
function doPost(e){
  let body;
  try{ body = JSON.parse(e.postData.contents); } catch(err){ return respond_({success:false,error:'bad body'}); }

  // --- secret key check (blocks casual/automated requests that don't have it) ---
  const settings = getSettings_();
  if(body.key !== settings.apiKey){
    return respond_({success:false, error:'unauthorized'});
  }

  try{
    switch(body.action){
      case 'getAll': return respond_(getAll_());
      case 'addStudent': return respond_(addStudent_(body));
      case 'updateStudent': return respond_(updateStudent_(body));
      case 'deleteStudent': return respond_(deleteStudent_(body.id));
      case 'addBook': return respond_(addBook_(body));
      case 'updateBook': return respond_(updateBook_(body));
      case 'deleteBook': return respond_(deleteBook_(body.id));
      case 'borrow': return respond_(borrow_(body));
      case 'returnBook': return respond_(returnBook_(body));
      case 'deleteTransaction': return respond_(deleteTransaction_(body.id));
      case 'getArchived': return respond_(getArchived_());
      case 'restoreStudent': return respond_(restoreStudent_(body.id));
      case 'restoreBook': return respond_(restoreBook_(body.id));
      case 'restoreTransaction': return respond_(restoreTransaction_(body.id));
      case 'addReview': return respond_(addReview_(body));
      case 'saveSettings': saveSettings_(body.settings); return respond_({success:true});
      case 'installReminders': installReminders_(); return respond_({success:true});
      case 'uninstallReminders': uninstallReminders_(); return respond_({success:true});
      case 'installDailyBackup': installDailyBackup_(); return respond_({success:true});
      case 'uninstallDailyBackup': uninstallDailyBackup_(); return respond_({success:true});
      default: return respond_({success:false, error:'unknown action'});
    }
  }catch(err){ return respond_({success:false, error:String(err)}); }
}

function getAll_(){
  return {
    students: rowsToObjects_(studentsSheet_()).filter(s=> s.archived!==true && s.archived!=='true'),
    books: rowsToObjects_(booksSheet_()).filter(b=> b.archived!==true && b.archived!=='true'),
    transactions: rowsToObjects_(txSheet_()).filter(tx=> tx.status!=='deleted'),
    reviews: rowsToObjects_(reviewsSheet_()),
    settings: getSettings_()
  };
}

/* ---------- students ---------- */
function addStudent_(body){
  const sh = studentsSheet_();
  const id = genId_(), code = genCode_('ST');
  const headers = sh.getRange(1,1,1,sh.getLastColumn()).getValues()[0];
  const values = { id, code, name:body.name||'', grade:body.grade||'', parentEmail:body.parentEmail||'', photo:body.photo||'', remindersEnabled:true, createdAt:new Date().toISOString() };
  const row = headers.map(h => values[h]!==undefined ? values[h] : '');
  sh.appendRow(row);
  return {success:true, student:{id,code,name:body.name,grade:body.grade,parentEmail:body.parentEmail}};
}
function updateStudent_(body){
  const sh = studentsSheet_();
  const idx = findRowIndex_(sh,0,body.id);
  if(idx===-1) return {success:false, error:'not_found'};
  const headers = sh.getRange(1,1,1,sh.getLastColumn()).getValues()[0];
  const row = sh.getRange(idx,1,1,headers.length).getValues()[0];
  const map = {name:body.name, grade:body.grade, parentEmail:body.parentEmail, photo:body.photo, remindersEnabled:body.remindersEnabled};
  headers.forEach((h,i)=>{ if(map[h]!==undefined && map[h]!==null) row[i]=map[h]; });
  sh.getRange(idx,1,1,headers.length).setValues([row]);
  return {success:true};
}
function deleteStudent_(id){
  const sh = studentsSheet_();
  const idx = findRowIndex_(sh,0,id);
  if(idx===-1) return {success:false, error:'not_found'};
  const txs = rowsToObjects_(txSheet_());
  if(txs.some(t=>t.studentId===id && t.status==='active')) return {success:false, error:'has_active_loan'};
  const headers = sh.getRange(1,1,1,sh.getLastColumn()).getValues()[0];
  sh.getRange(idx, headers.indexOf('archived')+1).setValue(true);
  return {success:true};
}

/* ---------- books ---------- */
function addBook_(body){
  const sh = booksSheet_();
  const id = genId_(), code = genCode_('BK');
  const headers = sh.getRange(1,1,1,sh.getLastColumn()).getValues()[0];
  const values = { id, code, title:body.title||'', author:body.author||'', category:body.category||'', level:body.level||'', shelf:body.shelf||'', photo:body.photo||'', createdAt:new Date().toISOString() };
  const row = headers.map(h => values[h]!==undefined ? values[h] : '');
  sh.appendRow(row);
  return {success:true, book:{id,code,title:body.title,author:body.author,category:body.category,level:body.level,shelf:body.shelf,photo:body.photo}};
}
function updateBook_(body){
  const sh = booksSheet_();
  const idx = findRowIndex_(sh,0,body.id);
  if(idx===-1) return {success:false, error:'not_found'};
  const headers = sh.getRange(1,1,1,sh.getLastColumn()).getValues()[0];
  const row = sh.getRange(idx,1,1,headers.length).getValues()[0];
  const map = {title:body.title, author:body.author, category:body.category, level:body.level, shelf:body.shelf, photo:body.photo};
  headers.forEach((h,i)=>{ if(map[h]!==undefined && map[h]!==null) row[i]=map[h]; });
  sh.getRange(idx,1,1,headers.length).setValues([row]);
  return {success:true};
}
function deleteBook_(id){
  const sh = booksSheet_();
  const idx = findRowIndex_(sh,0,id);
  if(idx===-1) return {success:false, error:'not_found'};
  const txs = rowsToObjects_(txSheet_());
  if(txs.some(t=>t.bookId===id && t.status==='active')) return {success:false, error:'has_active_loan'};
  const headers = sh.getRange(1,1,1,sh.getLastColumn()).getValues()[0];
  sh.getRange(idx, headers.indexOf('archived')+1).setValue(true);
  return {success:true};
}

/* ---------- borrow / return / delete tx ---------- */
function borrow_(body){
  const settings = getSettings_();
  const students = rowsToObjects_(studentsSheet_());
  const books = rowsToObjects_(booksSheet_());
  const txSh = txSheet_();
  const txs = rowsToObjects_(txSh);
  const student = students.find(s=>s.id===body.studentId);
  if(!student) return {success:false, reason:'unknown_student'};
  const book = books.find(b=>b.id===body.bookId);
  if(!book) return {success:false, reason:'unknown_book'};
  const bookActive = txs.find(t=>t.bookId===book.id && t.status==='active');
  if(bookActive) return {success:false, reason:'already_borrowed', borrower: bookActive.studentName};
  const activeCount = txs.filter(t=>t.studentId===student.id && t.status==='active').length;
  if(activeCount >= settings.maxBooks) return {success:false, reason:'max_reached', max:settings.maxBooks};
  const now = new Date();
  const due = new Date(now.getTime() + settings.loanDays*24*60*60*1000);
  const id = genId_();
  txSh.appendRow([id, student.id, student.name, book.id, book.title, now.toISOString(), due.toISOString(), '', 'active', '']);
  return {success:true, transaction:{id, studentId:student.id, studentName:student.name, bookId:book.id, bookTitle:book.title, borrowedAt:now.toISOString(), dueAt:due.toISOString(), status:'active'}};
}
function returnBook_(body){
  const txSh = txSheet_();
  const data = txSh.getDataRange().getValues();
  const headers = data[0];
  const bookIdCol = headers.indexOf('bookId'), statusCol = headers.indexOf('status'), returnedAtCol = headers.indexOf('returnedAt');
  for(let i=1;i<data.length;i++){
    if(data[i][bookIdCol]===body.bookId && data[i][statusCol]==='active'){
      txSh.getRange(i+1, returnedAtCol+1).setValue(new Date().toISOString());
      txSh.getRange(i+1, statusCol+1).setValue('returned');
      return {success:true};
    }
  }
  return {success:false, reason:'no_active_loan'};
}
function deleteTransaction_(id){
  const sh = txSheet_();
  const idx = findRowIndex_(sh,0,id);
  if(idx===-1) return {success:false, error:'not_found'};
  const headers = sh.getRange(1,1,1,sh.getLastColumn()).getValues()[0];
  const row = sh.getRange(idx,1,1,headers.length).getValues()[0];
  const currentStatus = row[headers.indexOf('status')];
  sh.getRange(idx, headers.indexOf('preDeleteStatus')+1).setValue(currentStatus);
  sh.getRange(idx, headers.indexOf('status')+1).setValue('deleted');
  return {success:true};
}
function restoreTransaction_(id){
  const sh = txSheet_();
  const idx = findRowIndex_(sh,0,id);
  if(idx===-1) return {success:false, error:'not_found'};
  const headers = sh.getRange(1,1,1,sh.getLastColumn()).getValues()[0];
  const row = sh.getRange(idx,1,1,headers.length).getValues()[0];
  const prevStatus = row[headers.indexOf('preDeleteStatus')] || 'returned';
  sh.getRange(idx, headers.indexOf('status')+1).setValue(prevStatus);
  sh.getRange(idx, headers.indexOf('preDeleteStatus')+1).setValue('');
  return {success:true};
}
function restoreStudent_(id){
  const sh = studentsSheet_();
  const idx = findRowIndex_(sh,0,id);
  if(idx===-1) return {success:false, error:'not_found'};
  const headers = sh.getRange(1,1,1,sh.getLastColumn()).getValues()[0];
  sh.getRange(idx, headers.indexOf('archived')+1).setValue(false);
  return {success:true};
}
function restoreBook_(id){
  const sh = booksSheet_();
  const idx = findRowIndex_(sh,0,id);
  if(idx===-1) return {success:false, error:'not_found'};
  const headers = sh.getRange(1,1,1,sh.getLastColumn()).getValues()[0];
  sh.getRange(idx, headers.indexOf('archived')+1).setValue(false);
  return {success:true};
}
function getArchived_(){
  return {
    students: rowsToObjects_(studentsSheet_()).filter(s=> s.archived===true || s.archived==='true'),
    books: rowsToObjects_(booksSheet_()).filter(b=> b.archived===true || b.archived==='true'),
    transactions: rowsToObjects_(txSheet_()).filter(tx=> tx.status==='deleted')
  };
}

/* ---------- reviews ---------- */
function addReview_(body){
  const sh = reviewsSheet_();
  const id = genId_();
  sh.appendRow([id, body.bookId, body.bookTitle||'', body.studentId||'', body.studentName||'', Number(body.rating)||0, body.comment||'', new Date().toISOString()]);
  return {success:true};
}

/* ---------- reminders ---------- */
function installReminders_(){
  uninstallReminders_();
  ScriptApp.newTrigger('sendDueReminders').timeBased().everyDays(1).atHour(8).create();
  saveSettings_(Object.assign(getSettings_(), {remindersEnabled:true}));
}
function uninstallReminders_(){
  ScriptApp.getProjectTriggers().forEach(t=>{ if(t.getHandlerFunction()==='sendDueReminders') ScriptApp.deleteTrigger(t); });
  saveSettings_(Object.assign(getSettings_(), {remindersEnabled:false}));
}
function sendDueReminders(){
  const settings = getSettings_();
  const students = rowsToObjects_(studentsSheet_());
  const txSh = txSheet_();
  const data = txSh.getDataRange().getValues();
  const headers = data[0];
  const col = {}; headers.forEach((h,i)=>col[h]=i);
  const today = new Date();
  const todayStr = today.toDateString();
  for(let i=1;i<data.length;i++){
    if(data[i][col.status] !== 'active') continue;
    const due = new Date(data[i][col.dueAt]);
    const daysUntil = Math.ceil((due-today)/(24*60*60*1000));
    const lastReminder = data[i][col.lastReminderAt];
    if(lastReminder && new Date(lastReminder).toDateString()===todayStr) continue;
    const isUpcoming = daysUntil === settings.reminderDays;
    const isOverdue = daysUntil < 0;
    if(!isUpcoming && !isOverdue) continue;
    const student = students.find(s=>s.id===data[i][col.studentId]);
    if(!student || !student.parentEmail) continue;
    // per-student opt-out
    if(student.remindersEnabled===false || student.remindersEnabled==='false') continue;
    const subject = isOverdue ? `[${settings.libraryName}] Overdue book reminder` : `[${settings.libraryName}] Book due soon`;
    const message = isOverdue
      ? `Hi,\n\n${data[i][col.studentName]} still has "${data[i][col.bookTitle]}" checked out. It was due on ${due.toDateString()}. Please help them return it soon.\n\nThanks,\n${settings.libraryName}`
      : `Hi,\n\n${data[i][col.studentName]} has "${data[i][col.bookTitle]}" due on ${due.toDateString()}. Just a friendly reminder!\n\nThanks,\n${settings.libraryName}`;
    MailApp.sendEmail(student.parentEmail, subject, message);
    txSh.getRange(i+1, col.lastReminderAt+1).setValue(new Date().toISOString());
  }
}

/* ---------- daily backup ---------- */
function backupFolder_(){
  const folders = DriveApp.getFoldersByName('e-Rise Library Backups');
  return folders.hasNext() ? folders.next() : DriveApp.createFolder('e-Rise Library Backups');
}
function installDailyBackup_(){
  uninstallDailyBackup_();
  ScriptApp.newTrigger('runDailyBackup').timeBased().everyDays(1).atHour(2).create();
  runDailyBackup(); // also take one immediately
}
function uninstallDailyBackup_(){
  ScriptApp.getProjectTriggers().forEach(t=>{ if(t.getHandlerFunction()==='runDailyBackup') ScriptApp.deleteTrigger(t); });
}
function runDailyBackup(){
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const file = DriveApp.getFileById(ss.getId());
  const dateStr = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd');
  file.makeCopy(`e-Rise Library Backup ${dateStr}`, backupFolder_());

  // Clean up backups older than 90 days to avoid clutter
  const cutoff = new Date(Date.now() - 90*24*60*60*1000);
  const files = backupFolder_().getFiles();
  while(files.hasNext()){
    const f = files.next();
    if(f.getDateCreated() < cutoff) f.setTrashed(true);
  }
}

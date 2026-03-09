// ============================================================
// HVAC O&M Manager v3 — Google Apps Script (Code.gs)
// VERSION: 4.0 — Morning/Evening shifts + Schedule page
// ============================================================
// HOW TO DEPLOY:
// 1. Paste this file into Apps Script → Ctrl+S (Save)
// 2. Deploy → Manage Deployments → Edit (pencil) → New Version → Deploy
// ============================================================

var SS = SpreadsheetApp.getActiveSpreadsheet();
var TZ = Session.getScriptTimeZone();

function hdrReadings()    { return ['id','date','shift','equipment','equipId','parameter','value','unit','tech','createdAt']; }
function hdrBreakdowns()  { return ['id','date','time','machine','location','fault','tech','priority','action','status','downtime','restored','remarks','facility','createdAt']; }
function hdrHandovers()   { return ['id','date','shift','from','to','equipStatus','pending','completed','safety','remarks','facility','createdAt']; }
function hdrPMTasks()     { return ['id','name','equipment','frequency','estMins','procedure','assigned','createdAt']; }
function hdrPMLogs()      { return ['id','taskId','tech','date','remarks','createdAt']; }
function hdrAttendance()  { return ['id','emp','date','shift','status','createdAt']; }
function hdrROLogs()      { return ['id','date','shift','session','tech','feedTDS','productTDS','rejectTDS','recovery','saltRej','feedPress','hpPress','rejectPress','productFlow','rejectFlow','acfCl','turbidity','status','remarks','createdAt']; }
function hdrEquipment()   { return ['id','name','type','createdAt']; }
function hdrTechnicians() { return ['name']; }
function hdrSettings()    { return ['key','value']; }
function hdrSchedule()    { return ['id','emp','date','shift','createdAt']; }

function jsonOut(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}
function ok(extra) {
  var base = { success: true };
  if (extra) { var keys = Object.keys(extra); for (var k=0;k<keys.length;k++) base[keys[k]]=extra[keys[k]]; }
  return jsonOut(base);
}
function err(msg) { return jsonOut({ success: false, error: String(msg) }); }

function doGet(e)  { return handleRequest(e); }
function doPost(e) { return handleRequest(e); }

function handleRequest(e) {
  try {
    var p = {};
    if (e.postData && e.postData.contents) { try { p = JSON.parse(e.postData.contents); } catch(x) {} }
    if (e.parameter) { var pk=Object.keys(e.parameter); for (var k=0;k<pk.length;k++) p[pk[k]]=e.parameter[pk[k]]; }
    var a = String(p.action || '');
    if (a==='getData')             return getData();
    if (a==='saveReadings')        return saveReadings(p);
    if (a==='saveBreakdown')       return saveBreakdown(p);
    if (a==='updateBreakdown')     return updateBreakdown(p);
    if (a==='deleteBreakdown')     return deleteRowById('Breakdowns', p.id);
    if (a==='savePMTask')          return savePMTask(p);
    if (a==='updatePMTask')        return updatePMTask(p);
    if (a==='deletePMTask')        return deletePMTask(p);
    if (a==='completePM')          return completePM(p);
    if (a==='saveHandover')        return saveHandover(p);
    if (a==='updateHandover')      return updateHandover(p);
    if (a==='saveROLog')           return saveROLog(p);
    if (a==='updateROLog')         return updateROLog(p);
    if (a==='deleteROLog')         return deleteRowById('ROLogs', p.id);
    if (a==='saveAttendance')      return saveAttendanceSingle(p);
    if (a==='saveAttendanceBatch') return saveAttendanceBatch(p);
    if (a==='saveScheduleEntry')   return saveScheduleEntry(p);
    if (a==='saveScheduleBatch')   return saveScheduleBatch(p);
    if (a==='saveTechnician')      return saveTechnician(p);
    if (a==='deleteTechnician')    return deleteTechnician(p);
    if (a==='saveEquipment')       return saveEquipment(p);
    if (a==='deleteEquipment')     return deleteRowById('Equipment', p.id);
    if (a==='saveSettings')        return saveSettings(p);
    return err('Unknown action: ' + a);
  } catch (ex) { return err('Server error: ' + ex.message); }
}

function getData() {
  var techSh = getSheet('Technicians', hdrTechnicians());
  var techRows = techSh.getDataRange().getValues().slice(1);
  var technicians = [];
  for (var t=0;t<techRows.length;t++) { var n=String(techRows[t][0]||'').trim(); if(n) technicians.push(n); }
  var setsSh = getSheet('Settings', hdrSettings());
  var setsRows = setsSh.getDataRange().getValues().slice(1);
  var settings = {};
  for (var s=0;s<setsRows.length;s++) { if(setsRows[s][0]) settings[String(setsRows[s][0])]=String(setsRows[s][1]||''); }
  return ok({
    readings:    sheetData(getSheet('Readings',   hdrReadings())),
    breakdowns:  sheetData(getSheet('Breakdowns', hdrBreakdowns())),
    handovers:   sheetData(getSheet('Handovers',  hdrHandovers())),
    pmTasks:     sheetData(getSheet('PMTasks',    hdrPMTasks())),
    pmLogs:      sheetData(getSheet('PMLogs',     hdrPMLogs())),
    attendance:  sheetData(getSheet('Attendance', hdrAttendance())),
    roLogs:      sheetData(getSheet('ROLogs',     hdrROLogs())),
    equipment:   sheetData(getSheet('Equipment',  hdrEquipment())),
    schedule:    sheetData(getSheet('Schedule',   hdrSchedule())),
    technicians: technicians,
    settings:    settings
  });
}

function saveReadings(p) {
  var sh=getSheet('Readings',hdrReadings()); var rows=Array.isArray(p.rows)?p.rows:[];
  for(var i=0;i<rows.length;i++) sh.appendRow(objToRow(rows[i],hdrReadings()));
  return ok({saved:rows.length});
}
function saveBreakdown(p)  { getSheet('Breakdowns',hdrBreakdowns()).appendRow(objToRow(p,hdrBreakdowns())); return ok({id:p.id}); }
function updateBreakdown(p){ var h=hdrBreakdowns();var sh=getSheet('Breakdowns',h);var row=findRowById(sh,p.id);if(row<0){sh.appendRow(objToRow(p,h));return ok({id:p.id,note:'inserted'});}sh.getRange(row,1,1,h.length).setValues([objToRow(p,h)]);return ok({id:p.id}); }
function saveHandover(p)   { getSheet('Handovers',hdrHandovers()).appendRow(objToRow(p,hdrHandovers())); return ok({id:p.id}); }
function updateHandover(p) { var h=hdrHandovers();var sh=getSheet('Handovers',h);var row=findRowById(sh,p.id);if(row<0){sh.appendRow(objToRow(p,h));return ok({id:p.id,note:'inserted'});}sh.getRange(row,1,1,h.length).setValues([objToRow(p,h)]);return ok({id:p.id}); }
function savePMTask(p)     { getSheet('PMTasks',hdrPMTasks()).appendRow(objToRow(p,hdrPMTasks())); return ok({id:p.id}); }
function updatePMTask(p)   { var h=hdrPMTasks();var sh=getSheet('PMTasks',h);var row=findRowById(sh,p.id);if(row<0)sh.appendRow(objToRow(p,h));else sh.getRange(row,1,1,h.length).setValues([objToRow(p,h)]);return ok({id:p.id}); }
function deletePMTask(p) {
  deleteRowById('PMTasks',p.id);
  var logSh=getSheet('PMLogs',hdrPMLogs()); var data=logSh.getDataRange().getValues();
  for(var i=data.length-1;i>=1;i--) { if(String(data[i][1])===String(p.id)) logSh.deleteRow(i+1); }
  return ok({id:p.id});
}
function completePM(p)  { getSheet('PMLogs',hdrPMLogs()).appendRow(objToRow(p,hdrPMLogs())); return ok({id:p.id}); }
function saveROLog(p)   { getSheet('ROLogs',hdrROLogs()).appendRow(objToRow(p,hdrROLogs())); return ok({id:p.id}); }
function updateROLog(p) { var h=hdrROLogs();var sh=getSheet('ROLogs',h);var row=findRowById(sh,p.id);if(row<0){sh.appendRow(objToRow(p,h));return ok({id:p.id,note:'inserted'});}sh.getRange(row,1,1,h.length).setValues([objToRow(p,h)]);return ok({id:p.id}); }

function saveAttendanceSingle(p) {
  var h=hdrAttendance(); var sh=getSheet('Attendance',h); var data=sh.getDataRange().getValues();
  for(var i=1;i<data.length;i++) {
    if(String(data[i][1])===String(p.emp)&&cellToDateStr(data[i][2])===String(p.date)&&String(data[i][3])===String(p.shift)) {
      sh.getRange(i+1,1,1,h.length).setValues([objToRow(p,h)]); return ok({id:p.id,upsert:'updated'});
    }
  }
  sh.appendRow(objToRow(p,h)); return ok({id:p.id,upsert:'inserted'});
}
function saveAttendanceBatch(p) {
  var h=hdrAttendance(); var entries=Array.isArray(p.entries)?p.entries:[];
  if(!entries.length) return ok({saved:0});
  var sh=getSheet('Attendance',h); var data=sh.getDataRange().getValues();
  var lookup={};
  for(var i=1;i<data.length;i++) { var key=String(data[i][1])+'||'+cellToDateStr(data[i][2])+'||'+String(data[i][3]); lookup[key]=i+1; }
  var updated=0,inserted=0;
  for(var j=0;j<entries.length;j++) {
    var entry=entries[j]; var ekey=String(entry.emp)+'||'+String(entry.date)+'||'+String(entry.shift);
    var row=objToRow(entry,h);
    if(lookup[ekey]){sh.getRange(lookup[ekey],1,1,h.length).setValues([row]);updated++;}
    else{sh.appendRow(row);inserted++;}
  }
  return ok({saved:entries.length,updated:updated,inserted:inserted});
}

// ── Schedule ─────────────────────────────────────────────────────────────────

function saveScheduleEntry(p) {
  var h=hdrSchedule(); var sh=getSheet('Schedule',h);
  // If clearEntry flag, delete the row
  if(String(p.clearEntry)==='1'||p.shift==='') {
    var delRow=findRowById(sh,p.id);
    if(delRow>=0) sh.deleteRow(delRow);
    return ok({id:p.id,note:'cleared'});
  }
  // Upsert by id
  var row=findRowById(sh,p.id);
  if(row<0) sh.appendRow(objToRow(p,h));
  else sh.getRange(row,1,1,h.length).setValues([objToRow(p,h)]);
  return ok({id:p.id});
}

function saveScheduleBatch(p) {
  var h=hdrSchedule(); var entries=Array.isArray(p.entries)?p.entries:[];
  if(!entries.length) return ok({saved:0});
  var sh=getSheet('Schedule',h); var data=sh.getDataRange().getValues();
  // Build lookup by id
  var lookup={};
  for(var i=1;i<data.length;i++) { if(data[i][0]) lookup[String(data[i][0])]=i+1; }
  var updated=0,inserted=0;
  for(var j=0;j<entries.length;j++) {
    var entry=entries[j]; var row=objToRow(entry,h);
    if(lookup[String(entry.id)]){sh.getRange(lookup[String(entry.id)],1,1,h.length).setValues([row]);updated++;}
    else{sh.appendRow(row);inserted++;}
  }
  return ok({saved:entries.length,updated:updated,inserted:inserted});
}

function saveTechnician(p) {
  var sh=getSheet('Technicians',hdrTechnicians()); var rows=sh.getDataRange().getValues().slice(1);
  var name=String(p.name||'').trim();
  for(var i=0;i<rows.length;i++) { if(String(rows[i][0]).trim()===name) return ok({name:name,note:'exists'}); }
  sh.appendRow([name]); return ok({name:name});
}
function deleteTechnician(p) {
  var sh=getSheet('Technicians',hdrTechnicians()); var data=sh.getDataRange().getValues();
  for(var i=data.length-1;i>=1;i--) {
    if(String(data[i][0]).trim()===String(p.name||'').trim()){sh.deleteRow(i+1);return ok({name:p.name});}
  }
  return ok({name:p.name,note:'not found'});
}
function saveEquipment(p) {
  var h=hdrEquipment();var sh=getSheet('Equipment',h);var row=findRowById(sh,p.id);
  if(row<0)sh.appendRow(objToRow(p,h));else sh.getRange(row,1,1,h.length).setValues([objToRow(p,h)]);
  return ok({id:p.id});
}
function saveSettings(p) {
  var sh=getSheet('Settings',hdrSettings()); var data=sh.getDataRange().getValues();
  var pairs=[['facility',p.facility],['dept',p.dept],['user',p.user]];
  for(var f=0;f<pairs.length;f++) {
    var key=pairs[f][0],val=pairs[f][1]; if(val===undefined||val===null) continue;
    var found=false;
    for(var i=1;i<data.length;i++){if(String(data[i][0])===key){sh.getRange(i+1,2).setValue(val);found=true;break;}}
    if(!found) sh.appendRow([key,val]);
  }
  return ok({saved:true});
}

// ── Helpers ───────────────────────────────────────────────────────────────────

function getSheet(name, headers) {
  var sh=SS.getSheetByName(name);
  if(!sh) {
    sh=SS.insertSheet(name);
    if(headers&&headers.length>0) {
      sh.appendRow(headers); sh.setFrozenRows(1);
      sh.getRange(1,1,1,headers.length).setBackground('#1e293b').setFontColor('#ffffff').setFontWeight('bold');
    }
  }
  return sh;
}
function sheetData(sh) {
  var rows=sh.getDataRange().getValues(); if(rows.length<2) return [];
  var headers=[]; for(var h=0;h<rows[0].length;h++) headers.push(String(rows[0][h]).trim());
  var result=[];
  for(var i=1;i<rows.length;i++) {
    var r=rows[i]; var empty=true;
    for(var c=0;c<r.length;c++){if(r[c]!==''&&r[c]!==null&&r[c]!==undefined){empty=false;break;}}
    if(empty) continue;
    var obj={}; for(var j=0;j<headers.length;j++) obj[headers[j]]=cellToStr(r[j]);
    result.push(obj);
  }
  return result;
}
function cellToStr(val) {
  if(val===null||val===undefined||val==='') return '';
  if(val instanceof Date) return cellToDateStr(val);
  return String(val);
}
function cellToDateStr(val) {
  if(!val&&val!==0) return '';
  if(val instanceof Date) return Utilities.formatDate(val,TZ,'yyyy-MM-dd');
  var s=String(val).trim();
  if(/^\d{4}-\d{2}-\d{2}T/.test(s)) return s.slice(0,10);
  return s;
}
function objToRow(obj,headers) {
  var row=[]; for(var i=0;i<headers.length;i++){var v=obj[headers[i]];row.push((v!==undefined&&v!==null)?v:'');} return row;
}
function findRowById(sh,id) {
  var data=sh.getDataRange().getValues();
  for(var i=1;i<data.length;i++){if(String(data[i][0]).trim()===String(id).trim()) return i+1;}
  return -1;
}
function deleteRowById(sheetName,id) {
  var sh=SS.getSheetByName(sheetName); if(!sh) return ok({id:id,note:'sheet not found'});
  var row=findRowById(sh,id); if(row<0) return ok({id:id,note:'not found'});
  sh.deleteRow(row); return ok({id:id});
}

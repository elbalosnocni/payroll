function jsonOutput(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}
function successResponse(data){ return jsonOutput({ok:true,data:data}); }
function errorResponse(message,code){ return jsonOutput({ok:false,error:message,code:code||'ERROR'}); }

function formatCurrencyVND(v) {
  var n = Number(v); if (!isFinite(n)) n = 0;
  n = Math.round(n);
  var sign = n < 0 ? '-' : '';
  var s = String(Math.abs(n));
  return sign + s.replace(/\B(?=(\d{3})+(?!\d))/g,'.') + ' đ';
}
function formatDaysHours(v) {
  var n = Number(v); if (!isFinite(n)) return '';
  return Number.isInteger(n) ? String(n) : n.toFixed(2);
}
function formatDateVN(d) {
  return Utilities.formatDate(d, getSS().getSpreadsheetTimeZone() || Session.getScriptTimeZone() || 'Asia/Ho_Chi_Minh', 'dd/MM/yyyy HH:mm:ss');
}
function normalizeCCCD(v){ return String(v == null ? '' : v).replace(/[^0-9]/g,''); }
function normalizeText(v) {
  return String(v == null ? '' : v).normalize('NFD').replace(/[\u0300-\u036f]/g,'')
    .replace(/đ/gi,'d').toLowerCase().trim().replace(/\s+/g,' ');
}
function safeCompare(a,b) {
  a=String(a||''); b=String(b||'');
  if(a.length!==b.length) return false;
  var x=0; for(var i=0;i<a.length;i++) x |= a.charCodeAt(i)^b.charCodeAt(i);
  return x===0;
}
function bytesToHex(bytes){
  return bytes.map(function(b){ var n=b<0?b+256:b; return ('0'+n.toString(16)).slice(-2); }).join('');
}
function generateSalt(){ return Utilities.getUuid().replace(/-/g,'') + Utilities.getUuid().replace(/-/g,''); }
function hashPassword(password,salt) {
  var data = String(salt) + '|' + String(password) + '|' + CONFIG.PASSWORD_PEPPER;
  var digest;
  for(var i=0;i<CONFIG.HASH_ITERATIONS;i++){
    digest = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, data, Utilities.Charset.UTF_8);
    data = bytesToHex(digest) + '|' + salt + '|' + CONFIG.PASSWORD_PEPPER;
  }
  return data.split('|')[0];
}
function newToken() {
  return Utilities.getUuid().replace(/-/g,'') + Utilities.getUuid().replace(/-/g,'');
}
function tokenCacheKey(token){ return 'SESSION_' + Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256,String(token),Utilities.Charset.UTF_8).map(function(b){return ('0'+(b<0?b+256:b).toString(16)).slice(-2);}).join(''); }
function sheetToObjects(sheet) {
  var v=sheet.getDataRange().getValues(); if(v.length<2) return [];
  var h=v[0], out=[];
  for(var i=1;i<v.length;i++){
    var empty=true; for(var j=0;j<h.length;j++){ if(v[i][j]!=='' && v[i][j]!==null){empty=false;break;} }
    if(empty) continue;
    var o={}; for(j=0;j<h.length;j++) o[h[j]]=v[i][j];
    o.__row=i+1; out.push(o);
  }
  return out;
}
function setField(sh,row,field,value){
  var headers=sh.getRange(1,1,1,sh.getLastColumn()).getValues()[0], c=headers.indexOf(field)+1;
  if(c>0) sh.getRange(row,c).setValue(value);
}
function writeAudit(actor,action,target,detail){
  getOrCreateSheet(CONFIG.SHEET_AUDIT,['Timestamp','Actor','Action','Target','Detail'])
    .appendRow([formatDateVN(new Date()),actor,action,target,detail||'']);
}

/* Chokha Ghar – 21 Villa Township tracking → Google Sheet
   Do sheet banti hain apne aap:
   1) Events   – har click / event ki ek line
   2) Visitors – har visitor ki ek line (kitna interested, kaun sa ad, kitni baar aaya)
   Visitor ID wahi hai jo WhatsApp message ke (Ref: ...) mein aata hai. */
var EV_HEAD = ['Time (IST)','Visitor ID','Event','Standard event','Intent score','Tier','Ad (utm_content)','Campaign','Source','Visit #','Device','Details'];
var VI_HEAD = ['Visitor ID','First seen','Last seen','Visits','Ad','Max score','Tier','WhatsApp clicks','Call clicks','Site visit requests','Last event','Property'];

function sheet_(ss, name, head){
  var sh = ss.getSheetByName(name);
  if(!sh){ sh = ss.insertSheet(name); sh.appendRow(head); sh.setFrozenRows(1); sh.getRange(1,1,1,head.length).setFontWeight('bold').setBackground('#fff3c4'); }
  return sh;
}
function doPost(e){
  var lock = LockService.getScriptLock(); lock.waitLock(15000);
  try{
    var d = JSON.parse(e.postData.contents), ss = SpreadsheetApp.getActive();
    var now = Utilities.formatDate(new Date(), 'Asia/Kolkata', 'yyyy-MM-dd HH:mm:ss');
    var ev = String(d.event||''), vid = String(d.vid||'');
    sheet_(ss,'Events',EV_HEAD).appendRow([now,vid,ev,d.std||'',d.score||0,d.tier||'',d.ad||'',d.campaign||'',d.source||'',d.visits||1,d.device||'',d.detail||'']);
    var vs = sheet_(ss,'Visitors',VI_HEAD);
    var f = vs.getRange(2,1,Math.max(vs.getLastRow()-1,1),1).createTextFinder(vid).matchEntireCell(true).findNext();
    var isWA = ev.indexOf('whatsapp')===0, isCall = ev.indexOf('call')===0, isVisit = ev.indexOf('site_visit')===0;
    if(!f){
      vs.appendRow([vid,now,now,d.visits||1,d.ad||'',d.score||0,d.tier||'',isWA?1:0,isCall?1:0,isVisit?1:0,ev,d.property||'']);
    } else {
      var r = f.getRow(), row = vs.getRange(r,1,1,VI_HEAD.length).getValues()[0];
      row[2]=now; row[3]=Math.max(+row[3]||1,+d.visits||1); row[4]=row[4]||d.ad||''; row[5]=Math.max(+row[5]||0,+d.score||0);
      row[6]=row[5]>=70?'HOT':row[5]>=35?'WARM':'COLD';
      row[7]=(+row[7]||0)+(isWA?1:0); row[8]=(+row[8]||0)+(isCall?1:0); row[9]=(+row[9]||0)+(isVisit?1:0); row[10]=ev;
      vs.getRange(r,1,1,VI_HEAD.length).setValues([row]);
    }
  } finally { lock.releaseLock(); }
  return ContentService.createTextOutput('ok');
}
function doGet(){ return ContentService.createTextOutput('Chokha Ghar tracking chal raha hai ✅'); }

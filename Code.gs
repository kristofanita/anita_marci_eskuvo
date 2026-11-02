// A Google Táblázatból megnyitott Apps Script projektbe illeszd.
const SHEET_NAME = 'RSVP';
function reply_(ok, message, requestId) {
  const payload = JSON.stringify({type:'wedding-rsvp-result', ok:ok, message:message, requestId:requestId});
  const html = '<!doctype html><html><body><script>var result=' + payload + ';var target=window;for(var i=0;i<10;i++){if(target===target.parent)break;target=target.parent;target.postMessage(result,"*");}' + '</' + 'script></body></html>';
  return HtmlService.createHtmlOutput(html)
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}
// Read-only receipt lookup: no names, email addresses or allergies are returned.
function doGet(e) {
  const p=e&&e.parameter?e.parameter:{};
  const requestId=String(p.request_id||'');
  const callback=String(p.callback||'');
  const validId=/^[a-zA-Z0-9_-]{20,100}$/.test(requestId);
  let state='pending';
  if(p.action!=='status'||!validId) state='invalid';
  else {
    try {
      const spreadsheet=SpreadsheetApp.getActiveSpreadsheet();
      const sheet=spreadsheet.getSheetByName(SHEET_NAME);
      if(sheet&&sheet.getLastRow()>0&&sheet.getLastColumn()>=16) {
        const match=sheet.getRange(1,15,sheet.getLastRow(),1)
          .createTextFinder(requestId).matchEntireCell(true).findNext();
        if(match&&sheet.getRange(match.getRow(),16).getValue()==='saved') {
          const summary=sheet.getRange(match.getRow(),1,1,16).getValues()[0];
          if(summary[3]==='nem') state='saved';
          else if(summary[3]==='igen') {
            const people=spreadsheet.getSheetByName('Resztvevok');
            const expected=Number(summary[4])+Number(summary[5]);
            const matches=people&&people.getLastRow()>0
              ?people.getRange(1,2,people.getLastRow(),1).createTextFinder(requestId).matchEntireCell(true).findAll():[];
            if(expected>0&&matches.length===expected) state='saved';
          }
        }
      }
    } catch(err) { console.error(err); state='unavailable'; }
  }
  const result=JSON.stringify({type:'wedding-rsvp-status',requestId:requestId,state:state});
  if(/^rsvpCheck_[a-zA-Z0-9_]+$/.test(callback)) {
    return ContentService.createTextOutput(callback+'('+result+');')
      .setMimeType(ContentService.MimeType.JAVASCRIPT);
  }
  return ContentService.createTextOutput(result).setMimeType(ContentService.MimeType.JSON);
}
function doPost(e) {
  const p = e && e.parameter ? e.parameter : {};
  const requestId = String(p.request_id || '').slice(0, 100);
  const lock=LockService.getScriptLock();
  try {
    lock.waitLock(10000);
    if (!/^[a-zA-Z0-9_-]{20,100}$/.test(requestId)) throw new Error('Invalid request ID');
    if (p.website) return reply_(false, 'Érvénytelen beküldés.', requestId);
    if (!p.family || !p.email || !['igen','nem'].includes(p.attendance)) throw new Error('Hiányzó adatok');
    const sheet=SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAME);
    if (!sheet) throw new Error('Nincs RSVP nevű munkalap');
    const safe=v=>{const s=String(v??'').slice(0,3000);return /^[=+\-@]/.test(s)?"'"+s:s;};
    let names=[];
    let guests=[];
    if(p.attendance==='igen') {
      try {names=JSON.parse(p.guest_names||'[]');} catch(err) {throw new Error('Hibás névlista');}
      const count=Number(p.adults)+Number(p.children);
      if(!Number.isInteger(count)||count<1||count>60||!Array.isArray(names)||names.length!==count||names.some(n=>typeof n!=='string'||!n.trim())) throw new Error('Hiányos résztvevőlista');
    }
    if(p.attendance==='igen') {
      try {guests=JSON.parse(p.guests||'[]');} catch(err) {throw new Error('Invalid guest details');}
      const adultCount=Number(p.adults), childCount=Number(p.children);
      if(!Number.isInteger(adultCount)||adultCount<1||adultCount>30||
         !Number.isInteger(childCount)||childCount<0||childCount>30||
         !Array.isArray(guests)||guests.length!==names.length) throw new Error('Invalid guest details');
      guests.forEach((guest,i)=>{
        const child=i>=adultCount;
        if(!guest||typeof guest.name!=='string'||guest.name.trim()!==names[i].trim()||
           guest.name.length>150||guest.isChild!==child||
           typeof guest.diet!=='string'||!guest.diet.trim()||guest.diet.length>1000||
           typeof guest.age!=='string'||
           (guest.age!==''&&(!child||!/^\d+$/.test(guest.age)||Number(guest.age)>17))) throw new Error('Invalid guest details');
      });
    }
    const timestamp=new Date();
    const submissionId=requestId||Utilities.getUuid();
    let peopleSheet;
    if(guests.length) {
      const spreadsheet=SpreadsheetApp.getActiveSpreadsheet();
      peopleSheet=spreadsheet.getSheetByName('Resztvevok')||spreadsheet.insertSheet('Resztvevok');
      if(peopleSheet.getLastRow()===0) peopleSheet.appendRow([
        'Beküldés ideje', 'Jelentkezés azonosítója', 'Család / társaság',
        'E-mail', 'Teljes név', 'Felnőtt / gyermek', 'Életkor (év)',
        'Ételallergia / speciális étrend'
      ]);
    }
    // Az első 10 oszlop az eredeti táblázat szerkezetével kompatibilis.
    sheet.appendRow([timestamp,safe(p.family),safe(p.email),safe(p.attendance),safe(p.attendance==='igen'?p.adults:0),safe(p.attendance==='igen'?p.children:0),safe(p.ages),safe(p.diet),safe(p.accommodation),safe(p.message),safe(names.join(' | ')),safe(p.travel),safe(p.can_drive),safe(p.extra_seats),safe(submissionId),'']);
    const summaryRow=sheet.getLastRow();
    if(guests.length) {
      const rows=guests.map(guest=>[
        timestamp,safe(submissionId),safe(p.family),safe(p.email),safe(guest.name.trim()),
        guest.isChild?'Gyermek':'Felnőtt',guest.isChild&&guest.age!==''?Number(guest.age):'',safe(guest.diet.trim())
      ]);
      peopleSheet.getRange(peopleSheet.getLastRow()+1,1,rows.length,rows[0].length).setValues(rows);
    }
    // Commit the receipt only after both summary and participant writes complete.
    SpreadsheetApp.flush();
    sheet.getRange(summaryRow,16).setValue('saved');
    SpreadsheetApp.flush();
    return reply_(true, 'A jelentkezéseteket sikeresen rögzítettük.', requestId);
  } catch (err) {
    console.error(err);
    return reply_(false, 'Nem sikerült menteni a jelentkezést. Kérjük, próbáljátok újra.', requestId);
  } finally {
    if (lock.hasLock()) lock.releaseLock();
  }
}

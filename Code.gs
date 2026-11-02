// A Google Táblázatból megnyitott Apps Script projektbe illeszd.
const SHEET_NAME = 'RSVP';
function reply_(ok, message, requestId) {
  const payload = JSON.stringify({type:'wedding-rsvp-result', ok:ok, message:message, requestId:requestId});
  const html = '<!doctype html><html><body><script>var result=' + payload + ';var target=window;for(var i=0;i<10;i++){if(target===target.parent)break;target=target.parent;target.postMessage(result,"*");}' + '</' + 'script></body></html>';
  return HtmlService.createHtmlOutput(html)
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}
function doPost(e) {
  const p = e && e.parameter ? e.parameter : {};
  const requestId = String(p.request_id || '').slice(0, 100);
  const lock=LockService.getScriptLock();
  try {
    lock.waitLock(10000);
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
           typeof guest.diet!=='string'||guest.diet.length>1000||
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
    sheet.appendRow([timestamp,safe(p.family),safe(p.email),safe(p.attendance),safe(p.attendance==='igen'?p.adults:0),safe(p.attendance==='igen'?p.children:0),safe(p.ages),safe(p.diet),safe(p.accommodation),safe(p.message),safe(names.join(' | ')),safe(p.travel),safe(p.can_drive),safe(p.extra_seats)]);
    if(guests.length) {
      const rows=guests.map(guest=>[
        timestamp,safe(submissionId),safe(p.family),safe(p.email),safe(guest.name.trim()),
        guest.isChild?'Gyermek':'Felnőtt',guest.isChild&&guest.age!==''?Number(guest.age):'',safe(guest.diet.trim())
      ]);
      peopleSheet.getRange(peopleSheet.getLastRow()+1,1,rows.length,rows[0].length).setValues(rows);
    }
    return reply_(true, 'A jelentkezéseteket sikeresen rögzítettük.', requestId);
  } catch (err) {
    console.error(err);
    return reply_(false, 'Nem sikerült menteni a jelentkezést. Kérjük, próbáljátok újra.', requestId);
  } finally {
    if (lock.hasLock()) lock.releaseLock();
  }
}

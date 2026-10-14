// Create a NEW Apps Script bound to a NEW spreadsheet for this trip.
const TRIP_ID = 'tohoku-2027-0419';
function doGet(e) { return handle(e.parameter || {}, false); }
function doPost(e) { try { return handle(JSON.parse(e.postData.contents), true); } catch(err) { return output({ok:false,error:'資料格式錯誤'}); } }
function output(value) { return ContentService.createTextOutput(JSON.stringify(value)).setMimeType(ContentService.MimeType.JSON); }
function handle(body,write) {
  const secret=PropertiesService.getScriptProperties().getProperty('SYNC_TOKEN');
  if(body.tripId!==TRIP_ID) return output({ok:false,error:'旅程識別不符'});
  if(!secret || body.token!==secret) return output({ok:false,error:'同步碼錯誤或尚未設定 SYNC_TOKEN'});
  const lock=LockService.getScriptLock();
  try {
    lock.waitLock(15000);
    const book=SpreadsheetApp.getActiveSpreadsheet();
    const sheet=book.getSheetByName('Tohoku2027_Data')||book.insertSheet('Tohoku2027_Data');
    if(write) {
      if(!Array.isArray(body.shopping)||!Array.isArray(body.memos)) throw Error('清單格式錯誤');
      // Chunk JSON so image data does not exceed a single cell's text limit.
      const json=JSON.stringify({shopping:body.shopping,memos:body.memos});
      if(json.length>4500000) throw Error('同步資料過大，請減少圖片');
      const chunks=[];for(let i=0;i<json.length;i+=40000)chunks.push([json.slice(i,i+40000)]);
      // Prefix chunks so user text cannot be interpreted as spreadsheet formulas.
      sheet.clearContents();sheet.getRange(1,1,chunks.length,1).setValues(chunks.map(c=>['JSON:'+c[0]]));
    }
    const raw=sheet.getLastRow()?sheet.getRange(1,1,sheet.getLastRow(),1).getValues().map(r=>String(r[0]).replace(/^JSON:/,'')).join(''):'';
    const state=raw?JSON.parse(raw):{shopping:[],memos:[]};
    return output({ok:true,tripId:TRIP_ID,...state});
  } catch(err) { return output({ok:false,error:err.message}); }
  finally { if(lock.hasLock()) lock.releaseLock(); }
}

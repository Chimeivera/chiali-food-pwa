/**
 * 115年度主管成長營｜全新手機版獨立後端
 * 獨立資料庫：不讀取、不修改任何舊系統的試算表。
 * 部署為網頁應用程式：執行身分=我，存取權=任何人。
 * GitHub 靜態網站透過 JSONP 呼叫 doGet，避免 Apps Script CORS 問題。
 * 重要：公開匿名網頁不提供真實身分驗證，姓名選擇僅供識別。
 */
const SPREADSHEET_ID='1kCojgvR_8Y4CszOMVz9spXc1uB4OPuP3KovQrrmTrsU';
const DATA_SHEET='餐點登記';
const LOG_SHEET='餐點異動紀錄';
const HEADERS=['登記時間','批次編號','製作人員','餐點名稱','類別','數量','單位','可供應人數','預估費用','備註','餐點ID','狀態','最後更新時間'];
const NAMES=['曉鳳','雅鈴','淑英','佩君','幸修','秀玲','佳容','麗娥','欣蓁','美惠'];
const CATEGORIES=['主食','點心','飲品','水果','其他'];
const UNITS=['份','盤','盒','杯','個','鍋','其他'];
function doGet(e){
  const p=(e&&e.parameter)||{};
  const callback=String(p.callback||'');
  if(!/^cb_[A-Za-z0-9_]{1,70}$/.test(callback)){
    return ContentService.createTextOutput('請由正式餐點登記網頁進入。').setMimeType(ContentService.MimeType.TEXT);
  }
  let result;
  try{
    if(p.action==='ping')result={ok:true,service:'新手機版餐點登記',version:1};
    else if(p.action==='list')result=listMeals();
    else if(p.action==='submit')result=submitMeals(JSON.parse(String(p.payload||'{}')));
    else if(p.action==='change')result=changeMeal(JSON.parse(String(p.payload||'{}')));
    else result={ok:false,error:'不支援的操作'};
  }catch(ex){result={ok:false,error:String(ex.message||ex)}}
  return ContentService.createTextOutput(callback+'('+JSON.stringify(result).replace(/\u2028/g,'\\u2028').replace(/\u2029/g,'\\u2029')+');')
    .setMimeType(ContentService.MimeType.JAVASCRIPT);
}
function sheet_(){
 const ss=SpreadsheetApp.openById(SPREADSHEET_ID),s=ss.getSheetByName(DATA_SHEET);
 if(!s)throw Error('找不到新試算表的餐點登記工作表');
 const h=s.getRange(1,1,1,HEADERS.length).getDisplayValues()[0];
 if(h.some((v,i)=>v!==HEADERS[i]))throw Error('新試算表欄位標題不符，已停止寫入');
 return s;
}
function log_(){
 const s=SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName(LOG_SHEET);
 if(!s)throw Error('找不到餐點異動紀錄工作表');
 return s;
}
function safe_(s){s=String(s??'');return /^[=+\-@]/.test(s)?"'"+s:s}
function name_(v){const s=String(v||'').trim();if(!NAMES.includes(s))throw Error('請選擇有效姓名');return s}
function item_(x){
 if(!x||typeof x!=='object')throw Error('餐點格式不正確');
 const dish=String(x.dish||'').trim(),category=String(x.category||''),unit=String(x.unit||''),note=String(x.note||'').trim();
 const quantity=Number(x.quantity),serves=Number(x.serves),cost=Number(x.cost);
 if(!dish||dish.length>80||!CATEGORIES.includes(category)||!UNITS.includes(unit)||!Number.isInteger(quantity)||quantity<1||quantity>10000||!Number.isInteger(serves)||serves<1||serves>100000||!Number.isFinite(cost)||cost<0||cost>10000000||note.length>300)throw Error('餐點名稱、類別、數量、供應人數或費用格式不正確');
 return {dish,category,quantity,unit,serves,cost,note};
}
function obj_(r){return {id:String(r[10]||''),name:String(r[2]||''),dish:String(r[3]||''),category:String(r[4]||''),quantity:Number(r[5])||0,unit:String(r[6]||''),serves:Number(r[7])||0,cost:Number(r[8])||0,note:String(r[9]||'')}}
function listMeals(){
 try{const s=sheet_(),n=s.getLastRow();const rows=n>1?s.getRange(2,1,n-1,HEADERS.length).getValues():[];
  return {ok:true,rows:rows.filter(r=>r[2]&&String(r[11])!=='已刪除').map(obj_)};
 }catch(ex){return {ok:false,error:String(ex.message||ex)}}
}
function submitMeals(data){
 const lock=LockService.getScriptLock();
 try{
 const name=name_(data.name),items=Array.isArray(data.items)?data.items.map(item_):[];
 if(!items.length||items.length>10)throw Error('每次請登記 1 至 10 道餐點');
 const key=String(data.requestId||'');
 if(!/^[A-Za-z0-9_-]{12,90}$/.test(key))throw Error('登記識別碼無效，請重新整理頁面');
 lock.waitLock(15000);
 const s=sheet_(),last=s.getLastRow();
 if(last>1){
  const prev=s.getRange(2,2,last-1,1).getValues();
  if(prev.some(r=>r[0]===key))return {ok:true,count:items.length,duplicate:true};
 }
 const now=new Date(),rows=items.map(x=>[now,key,safe_(name),safe_(x.dish),x.category,x.quantity,x.unit,x.serves,x.cost,safe_(x.note),Utilities.getUuid(),'有效',now]);
 s.getRange(last+1,1,rows.length,HEADERS.length).setValues(rows);
 SpreadsheetApp.flush();return {ok:true,count:rows.length};
 }catch(ex){return {ok:false,error:String(ex.message||ex)}}finally{if(lock.hasLock())lock.releaseLock()}
}
function changeMeal(data){
 const lock=LockService.getScriptLock();
 try{
 const actor=name_(data.actor),id=String(data.id||''),action=String(data.action||'');
 if(!id||!['edit','delete'].includes(action))throw Error('操作參數不正確');
 const item=action==='edit'?item_(data.item):null;
 lock.waitLock(15000);
 const s=sheet_(),last=s.getLastRow();if(last<2)throw Error('找不到餐點');
 const ids=s.getRange(2,11,last-1,1).getValues(),idx=ids.findIndex(x=>String(x[0])===id);
 if(idx<0)throw Error('找不到此筆餐點');
 const row=idx+2,old=s.getRange(row,1,1,HEADERS.length).getValues()[0];
 if(String(old[11])==='已刪除')throw Error('這筆餐點已刪除');
 const before=obj_(old);
 if(before.name!==actor)throw Error('只能修改自己姓名下的餐點');
 const now=new Date();
 if(action==='edit'){old[3]=safe_(item.dish);old[4]=item.category;old[5]=item.quantity;old[6]=item.unit;old[7]=item.serves;old[8]=item.cost;old[9]=safe_(item.note)}
 else old[11]='已刪除';
 old[12]=now;
 const after=action==='edit'?obj_(old):{...before,status:'已刪除'};
 log_().appendRow([now,action==='edit'?'修改':'刪除',safe_(actor),id,safe_(JSON.stringify(before)),safe_(JSON.stringify(after))]);
 s.getRange(row,1,1,HEADERS.length).setValues([old]);SpreadsheetApp.flush();
 return {ok:true};
 }catch(ex){return {ok:false,error:String(ex.message||ex)}}finally{if(lock.hasLock())lock.releaseLock()}
}

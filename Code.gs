/* Muzammil Creatives – HR & Attendance backend
   Sheet se Extensions > Apps Script me paste karein. Pehle setup() ek baar run karein. */
const TZ='Asia/Karachi',START='09:00',END='18:00',CUT='11:00';
const H={
 Users:['email','name','phone','designation','salary','hash','role','active','joined','pic'],
 Attendance:['date','email','name','cin','cout','status','note'],
 Requests:['id','email','name','type','date','time','reason','status','created','decided','by']};
const TYPES=['Leave','Short Leave','Late Sitting','CheckOut Missing','CheckIn Missing'];

const sh=n=>SpreadsheetApp.getActive().getSheetByName(n);
function rows(n){const v=sh(n).getDataRange().getValues(),h=v.shift();
 return v.map((r,i)=>{const o={_r:i+2};h.forEach((k,j)=>o[k]=String(r[j]));return o})}
const put=(n,o)=>sh(n).appendRow(H[n].map(k=>o[k]==null?'':String(o[k])));
const upd=(n,r,o)=>Object.keys(o).forEach(k=>sh(n).getRange(r,H[n].indexOf(k)+1).setValue(String(o[k])));
const hash=p=>Utilities.base64Encode(Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256,'mc|'+p));
const out=o=>ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON);

/* ---- time helpers (Pakistan time) ---- */
const fd=(d,f)=>Utilities.formatDate(d,TZ,f),today=()=>fd(new Date(),'yyyy-MM-dd'),hm=()=>fd(new Date(),'HH:mm');
const wk=s=>{const d=new Date(s+'T12:00:00Z').getUTCDay();return d>0&&d<6}; // Mon-Fri
const addD=(s,n)=>{const d=new Date(s+'T12:00:00Z');d.setUTCDate(d.getUTCDate()+n);return d.toISOString().slice(0,10)};
const wdays=(a,b)=>{let n=0;for(let d=a;d<b;d=addD(d,1))if(wk(addD(d,1)))n++;return n};

/* ---- one-time setup ---- */
function setup(){
 Object.keys(H).forEach(n=>{const s=sh(n)||SpreadsheetApp.getActive().insertSheet(n);
  if(!s.getLastRow()){s.getRange(1,1,1,H[n].length).setValues([H[n]]).setFontWeight('bold');s.setFrozenRows(1)}
  s.getRange(1,1,s.getMaxRows(),H[n].length).setNumberFormat('@')});
 if(!rows('Users').some(u=>u.role=='HR'))put('Users',{email:'almuzammildigitalexperts@gmail.com',name:'Muzammil Ahmed',
  phone:'03270062558',designation:'HR Manager',salary:0,hash:hash('62558'),role:'HR',active:1,joined:today(),pic:''});
 ScriptApp.getProjectTriggers().forEach(t=>ScriptApp.deleteTrigger(t));
 ScriptApp.newTrigger('closeDays').timeBased().atHour(0).nearMinute(5).everyDays(1).inTimezone(TZ).create();
}

/* ---- auth token (signed, 12h) ---- */
function sec(){const p=PropertiesService.getScriptProperties();let s=p.getProperty('S');if(!s){s=Utilities.getUuid();p.setProperty('S',s)}return s}
const sign=x=>Utilities.base64EncodeWebSafe(Utilities.computeHmacSha256Signature(x,sec()));
const mk=e=>{const x=e+'|'+(Date.now()+12*36e5);return Utilities.base64EncodeWebSafe(x)+'.'+sign(x)};
function auth(t){try{const [a,s]=String(t).split('.'),x=Utilities.newBlob(Utilities.base64DecodeWebSafe(a)).getDataAsString();
 if(sign(x)!==s)return;const [e,x2]=x.split('|');if(Date.now()>+x2)return;
 return rows('Users').find(u=>u.email==e&&u.active=='1')}catch(_){}}
const pub=u=>({email:u.email,name:u.name,phone:u.phone,designation:u.designation,pic:u.pic,role:u.role,joined:u.joined,salary:u.salary});

/* ---- daily close: absent / missing checkout + auto-approve ---- */
function closeDays(){
 const t=today(),A=rows('Attendance'),have={},add=[];A.forEach(a=>have[a.date+a.email]=a);
 rows('Users').filter(u=>u.role=='EMP'&&u.active=='1'&&u.joined).forEach(u=>{
  for(let d=u.joined;d<t;d=addD(d,1)){if(!wk(d))continue;const a=have[d+u.email];
   if(!a)add.push(H.Attendance.map(k=>({date:d,email:u.email,name:u.name,status:'Absent',note:'No check-in'}[k]||'')));
   else if(a.status=='In')upd('Attendance',a._r,{status:'CheckOut Missing'})}});
 if(add.length){const s=sh('Attendance');s.getRange(s.getLastRow()+1,1,add.length,H.Attendance.length).setValues(add)}
 rows('Requests').filter(r=>r.status=='Pending'&&wdays(fd(new Date(+r.created),'yyyy-MM-dd'),t)>=5).forEach(r=>decide(r,'Approved','Auto (5 days)'));
}
function decide(r,st,by){upd('Requests',r._r,{status:st,decided:Date.now(),by});if(st=='Approved')apply(r)}
function apply(r){ // approved request -> timesheet
 const a=rows('Attendance').find(x=>x.date==r.date&&x.email==r.email),T=r.type,o={};
 if(T=='Leave'){o.status='Leave';o.note='Approved leave'}
 else{
  if(T=='CheckIn Missing'||(T!='CheckOut Missing'&&!(a&&a.cin)))o.cin=T=='CheckIn Missing'&&r.time||START;
  if(T=='CheckOut Missing'||!(a&&a.cout))o.cout=r.time||END;
  o.note=T+' approved'+(r.time?' @'+r.time:'');
  const i=o.cin||(a&&a.cin)||'',c=(T=='CheckIn Missing'&&!(a&&a.cout))?'':(o.cout||(a&&a.cout)||'');
  if(T=='CheckIn Missing'&&!(a&&a.cout))delete o.cout;
  o.status=i&&c?'Present':i?(r.date<today()?'CheckOut Missing':'In'):'Absent';
 }
 a?upd('Attendance',a._r,o):put('Attendance',Object.assign({date:r.date,email:r.email,name:r.name},o));
}

/* ---- payslip (month = yyyy-MM) ---- */
function slip(u,m){
 if(!/^\d{4}-\d\d$/.test(m))throw Error('Bad month');
 const first=m+'-01',last=new Date(Date.UTC(+m.slice(0,4),+m.slice(5),0)).toISOString().slice(0,10),rel=addD(last,10),d=today();
 if(last<u.joined)throw Error('You had not joined in this month');
 let W=0,pre=0;for(let x=first;x<=last;x=addD(x,1))if(wk(x)){W++;if(x<u.joined)pre++}
 const a=rows('Attendance').filter(x=>x.email==u.email&&x.date>=first&&x.date<=last),
  ab=a.filter(x=>x.status=='Absent').length,lv=a.filter(x=>x.status=='Leave').length,
  hold=a.filter(x=>x.status=='CheckOut Missing').length+rows('Requests').filter(r=>r.email==u.email&&r.status=='Pending'&&r.date>=first&&r.date<=last).length;
 if(d<rel)return{m,state:'wait',msg:'Payslip for '+m+' will be available on '+rel+'.'};
 if(hold)return{m,state:'hold',msg:hold+' item(s) are unresolved (missing check-out or pending request). The payslip is released once HR approves them or they are resolved.'};
 const sal=+u.salary||0,ded=Math.round(sal/W*(ab+pre));
 return{m,state:'ok',salary:sal,workDays:W,absent:ab,prejoin:pre,leave:lv,deduction:ded,net:sal-ded,absentDates:a.filter(x=>x.status=='Absent').map(x=>x.date)};
}

function saveEmp(e){
 const em=String(e.email||'').trim().toLowerCase();if(!em||!String(e.name||'').trim())throw Error('Name and email are required');
 const x=rows('Users').find(u=>u.email==em),
  o={name:e.name.trim(),phone:e.phone||'',designation:e.designation||'',salary:+e.salary||0,active:e.active=='0'?0:1,pic:e.pic||(x?x.pic:'')};
 if(x){if(x.role=='HR')throw Error('HR account cannot be edited here');if(e.pass)o.hash=hash(e.pass);upd('Users',x._r,o)}
 else{if(!e.pass)throw Error('Password is required for a new ID');put('Users',Object.assign({email:em,role:'EMP',joined:e.joined||today(),hash:hash(e.pass)},o))}
 return{ok:1};
}

/* ---- API ---- */
function doGet(){return out({ok:1,app:'Muzammil Creatives'})}
function doPost(e){
 let q;try{q=JSON.parse(e.postData.contents)}catch(_){return out({error:'Bad request'})}
 const L=LockService.getScriptLock();L.waitLock(25000);
 try{return out(route(q))}catch(err){return out({error:String(err.message||err)})}finally{L.releaseLock()}
}
function route(q){
 if(q.a=='login'){
  const u=rows('Users').find(u=>u.email==String(q.email).trim().toLowerCase()&&u.active=='1');
  if(!u||u.hash!=hash(q.pass))throw Error('Wrong email or password');
  return{token:mk(u.email),me:pub(u)};
 }
 const u=auth(q.t);if(!u)throw Error('AUTH');
 const c=CacheService.getScriptCache();if(!c.get('cd')){closeDays();c.put('cd','1',60)}
 const hr=u.role=='HR',d=today(),t=hm(),A=()=>rows('Attendance'),need=()=>{if(!hr)throw Error('HR only')};
 switch(q.a){
  case 'me':{
   const a=A().filter(x=>x.email==u.email).sort((x,y)=>x.date<y.date?1:-1),td=a.find(x=>x.date==d);
   return{me:pub(u),ms:Date.now(),date:d,att:a.slice(0,45),canIn:wk(d)&&!td&&t<=CUT,canOut:!!(td&&td.cin&&!td.cout),
    msg:!wk(d)?'Weekend: office is off':td?(td.status=='Leave'?'Approved leave today':td.cin?(td.cout?'Day complete':'Checked in at '+td.cin):''):
     t>CUT?'Check-in closed after 11:00 AM. Today is marked absent.':'You have not checked in yet'};}
  case 'in':{
   if(!wk(d))throw Error('Weekend: office is off');if(t>CUT)throw Error('Check-in closed after 11:00 AM');
   if(A().some(x=>x.date==d&&x.email==u.email))throw Error('Already marked for today');
   put('Attendance',{date:d,email:u.email,name:u.name,cin:t,status:'In'});return{ok:1};}
  case 'out':{
   const x=A().find(x=>x.date==d&&x.email==u.email);if(!x||!x.cin)throw Error('Check in first');if(x.cout)throw Error('Already checked out');
   upd('Attendance',x._r,{cout:t,status:'Present',note:t<END?'Early checkout':''});return{ok:1};}
  case 'reqs':return{list:rows('Requests').filter(r=>hr||r.email==u.email).reverse()};
  case 'req':{
   if(!TYPES.includes(q.type)||!/^\d{4}-\d\d-\d\d$/.test(q.date)||!String(q.reason||'').trim())throw Error('Please fill type, date and reason');
   if(q.date>d&&q.type!='Leave'&&q.type!='Short Leave')throw Error('Future dates are allowed only for leave');
   put('Requests',{id:Utilities.getUuid().slice(0,8),email:u.email,name:u.name,type:q.type,date:q.date,time:q.time||'',reason:q.reason,status:'Pending',created:Date.now()});return{ok:1};}
  case 'slip':return slip(u,q.m);
  case 'live':{need();const a=A().filter(x=>x.date==d);
   return{date:d,list:rows('Users').filter(x=>x.role=='EMP'&&x.active=='1').map(e=>{const x=a.find(y=>y.email==e.email)||{};
    return{name:e.name,designation:e.designation,pic:e.pic,cin:x.cin||'',cout:x.cout||'',status:x.status||(!wk(d)?'Weekend':t>CUT?'Absent':'Not yet')}})};}
  case 'rep':{need();return{list:A().filter(x=>(!q.from||x.date>=q.from)&&(!q.to||x.date<=q.to)&&(!q.email||x.email==q.email)).sort((x,y)=>x.date<y.date?1:-1)};}
  case 'dec':{need();const r=rows('Requests').find(x=>x.id==q.id&&x.status=='Pending');if(!r)throw Error('Request is not pending');
   decide(r,q.st=='Approved'?'Approved':'Rejected',u.name);return{ok:1};}
  case 'emps':{need();return{list:rows('Users').filter(x=>x.role=='EMP').map(x=>Object.assign(pub(x),{active:x.active}))};}
  case 'save':{need();return saveEmp(q.e);}
 }
 throw Error('Unknown action');
}

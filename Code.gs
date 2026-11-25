const SHEET_NAME = 'Результати ELECTRO LAB';
const PROP_KEY = 'ELECTRO_LAB_SPREADSHEET_ID';
const HEADERS = ['Дата й час','ID спроби','Прізвище та ім’я','Клас','Бал','Максимум','Оцінка словами','Q1: полюси','Q2: умова магнітної дії','Q3: сила струму','Q4: кількість витків','Q5: осердя','Q6: фінальна конфігурація','Спроби за питаннями','Результати 0,5 А / 1,5 А','Порівняння витків','Порівняння осердя','Фінальна конфігурація','Бали за рівнями','Усі відповіді (JSON)','Усі дані (JSON)','Текст звіту'];
function getSheet_(){
  const props=PropertiesService.getScriptProperties();
  let id=props.getProperty(PROP_KEY), ss;
  if(id){try{ss=SpreadsheetApp.openById(id)}catch(e){ss=null}}
  if(!ss){ss=SpreadsheetApp.create('ELECTRO LAB — результати учнів');props.setProperty(PROP_KEY,ss.getId())}
  let sh=ss.getSheetByName(SHEET_NAME);if(!sh)sh=ss.insertSheet(SHEET_NAME);
  if(sh.getLastRow()===0){sh.appendRow(HEADERS);sh.setFrozenRows(1);sh.getRange(1,1,1,HEADERS.length).setFontWeight('bold');sh.autoResizeColumns(1,HEADERS.length)}
  return {spreadsheet:ss,sheet:sh};
}
function doGet(){const x=getSheet_();return ContentService.createTextOutput(JSON.stringify({ok:true,message:'ELECTRO LAB endpoint is ready',spreadsheetUrl:x.spreadsheet.getUrl()})).setMimeType(ContentService.MimeType.JSON)}
function doPost(e){
  try{
    const raw=e&&e.postData&&e.postData.contents;if(!raw)throw new Error('Empty request');
    const p=JSON.parse(raw);if(!p.student||!p.student.name||!p.student.class||!p.attemptId)throw new Error('Missing student or attempt data');
    const lock=LockService.getScriptLock();lock.waitLock(10000);
    try{
      const x=getSheet_(),sh=x.sheet;
      const last=sh.getLastRow();if(last>1){const ids=sh.getRange(2,2,last-1,1).getValues().flat();if(ids.includes(p.attemptId))return json_({ok:true,duplicate:true})}
      const a=p.answers||{},d=p.data||{},exp=p.experiments||{};
      const qLabels={q1:['Змінити напрямок струму','Збільшити кількість витків','Замінити залізне осердя'],q2:['Струм проходить через котушку','Нагрівання','Дотик магніту'],q3:['Зменшується','Не змінюється','Збільшується'],q4:['Посилиться','Послабиться','Зникне'],q5:['Залізо підсилює поле','Дерево краще проводить струм','Збільшує опір повітря'],q6:['0,5 А · 40 · дерево','1,5 А · 80 · залізо','1,5 А · 40 · дерево']};
      const answerText=id=>a[id]===undefined?'':(qLabels[id][a[id]]||String(a[id]));
      const config=exp.finalConfiguration||{};
      const fmt=v=>v===undefined||v===null?'':JSON.stringify(v);
      sh.appendRow([p.submittedAt||new Date().toISOString(),p.attemptId,p.student.name,p.student.class,p.score||0,p.maxScore||12,p.grade||'',answerText('q1'),answerText('q2'),answerText('q3'),answerText('q4'),answerText('q5'),answerText('q6'),fmt(p.attempts),fmt(exp.current), '40 витків: 5; 80 витків: 9 (умовні значення)','Залізо: 10; дерево: 2 (умовні значення)',config.i!==undefined?config.i+' A; '+config.n+' витків; '+config.core+'; дія '+config.power:'',fmt(p.awards),fmt(a),fmt(d),p.report||'']);
      return json_({ok:true,spreadsheetUrl:x.spreadsheet.getUrl()});
    }finally{lock.releaseLock()}
  }catch(err){return json_({ok:false,error:String(err)})}
}
function json_(obj){return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON)}

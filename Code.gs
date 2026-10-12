/**
 * ระบบจัดการคะแนนและติดตามงานนักเรียน ภาคเรียน 2/2569
 * Google Sheets + Apps Script
 * ไฟล์ที่ต้องมี: Code.gs, SmartPaste.html, RosterLink.html
 */
const APP = {
  TERM: '2569-2', MENU: 'ระบบจัดการคะแนน',
  SHEETS: { CONFIG: 'Config', TEMPLATE: 'Template', ASSIGNMENTS: 'Assignments', SUBMISSIONS: 'Submissions', ROSTER_MAP: 'RosterRoomMap' },
  CLASS_HEADERS: ['เลขที่','เลขประจำตัว','ชื่อ-นามสกุล','งานก่อนกลางภาค','ชิ้นงานก่อนกลางภาค','ก่อนกลางภาค','สอบกลางภาค','งานหลังกลางภาค','ชิ้นงานหลังกลางภาค','หลังกลางภาค','สอบปลายภาค','คะแนนรวมสุทธิ','เกรด','คุณลักษณะ 1','คุณลักษณะ 2','คุณลักษณะ 3','คุณลักษณะ 4','คุณลักษณะ 5','คุณลักษณะ 6','คุณลักษณะ 7','คุณลักษณะ 8','สรุปคุณลักษณะฯ','อ่านคิดวิเคราะห์ 1','อ่านคิดวิเคราะห์ 2','อ่านคิดวิเคราะห์ 3','อ่านคิดวิเคราะห์ 4','อ่านคิดวิเคราะห์ 5','สรุปการอ่านฯ','Smart Alert','DemoFlag','สถานะเกรด (ร/มส)'],
  ASSIGNMENT_HEADERS: ['AssignmentID','TermID','CourseCode','Room','Title','Category','MaxScore','DueAt','IsDemo','CreatedAt'],
  SUBMISSION_HEADERS: ['StudentID','AssignmentID','CourseCode','Room','Status','SubmittedAt','Score','MaxScore','IsDemo'],
  ROSTER_HEADERS: ['Room','Year','GradeID','RoomID','SourceURL','UpdatedAt'],
  COURSES: [
    ['ส32103','สังคมศึกษา 4','5/4,5/5,5/6,5/12,5/13,5/17',20,10,20,10,10,30,''],
    ['ส33265','เหตุการณ์ปัจจุบัน 2','6/17',20,10,20,10,10,30,''],
    ['ส23262','มหาวชิราวุธศึกษา 2','3/13,3/14,3/15,3/16',15,15,20,15,15,20,'']
  ]
};

function onOpen() {
  SpreadsheetApp.getUi().createMenu(APP.MENU)
    .addItem('1. ตั้งค่าระบบ / สร้างชีต','setupSystem').addSeparator()
    .addItem('2. นำเข้ารายชื่อนักเรียน (Smart Paste)','showSmartPaste')
    .addItem('3. สร้างลิงก์รายชื่อนักเรียน Academic+','showRosterLinkBuilder')
    .addItem('4. สร้างข้อมูลจำลอง','generateDemoData')
    .addItem('5. ล้างข้อมูลจำลอง','clearDemoData').addSeparator()
    .addItem('6. ตรวจสอบข้อมูลห้องเรียน','validateClassrooms').addToUi();
}

function setupSystem() {
  const ss = SpreadsheetApp.getActive();
  const oldConfig = ss.getSheetByName(APP.SHEETS.CONFIG);
  const oldIds = {};
  let oldTemplate = '';
  if (oldConfig) {
    oldTemplate = String(oldConfig.getRange('B8').getDisplayValue() || '').trim();
    if (oldConfig.getLastRow() >= 3) oldConfig.getRange(3,4,oldConfig.getLastRow()-2,10).getValues().forEach(r => { if (r[0] && r[9]) oldIds[String(r[0]).trim()] = String(r[9]).trim(); });
  }
  const config = getOrCreateSheet_(ss, APP.SHEETS.CONFIG);
  config.getRange('A1:B8').setValues([
    ['Setting','Value'],['TermID',APP.TERM],['DemoMode',false],['QualityPassLevel',2],['SchoolYear',2569],['Semester',2],['LastUpdated',new Date()],['RosterURLTemplate',oldTemplate]
  ]);
  config.getRange('D1').setValue('COURSE CONFIGURATION');
  config.getRange('D2:M2').setValues([['CourseCode','Subject','Rooms','PreWorkMax','PreProjectMax','MidMax','PostWorkMax','PostProjectMax','FinalMax','ClassroomCourseID']]);
  const courseRows = APP.COURSES.map(r => { const a = r.slice(); if (oldIds[a[0]]) a[9] = oldIds[a[0]]; return a; });
  config.getRange(3,4,courseRows.length,10).setValues(courseRows);
  config.setFrozenRows(2); config.getRange('A1:B1').setFontWeight('bold').setBackground('#173A72').setFontColor('#fff');
  config.getRange('D2:M2').setFontWeight('bold').setBackground('#173A72').setFontColor('#fff').setWrap(true);
  config.setColumnWidth(1,170); config.setColumnWidth(2,220); config.setColumnWidths(4,10,145); config.getRange('M3:M').setNumberFormat('@');

  const template = getOrCreateSheet_(ss, APP.SHEETS.TEMPLATE);
  template.clear(); template.getRange(1,1,1,APP.CLASS_HEADERS.length).setValues([APP.CLASS_HEADERS]); styleClassSheet_(template);
  resetHeaderIfNeeded_(getOrCreateSheet_(ss,APP.SHEETS.ASSIGNMENTS),APP.ASSIGNMENT_HEADERS);
  resetHeaderIfNeeded_(getOrCreateSheet_(ss,APP.SHEETS.SUBMISSIONS),APP.SUBMISSION_HEADERS);
  resetHeaderIfNeeded_(getOrCreateSheet_(ss,APP.SHEETS.ROSTER_MAP),APP.ROSTER_HEADERS);
  [APP.SHEETS.ASSIGNMENTS,APP.SHEETS.SUBMISSIONS,APP.SHEETS.ROSTER_MAP].forEach(n => { const sh=ss.getSheetByName(n); sh.setFrozenRows(1); const w=sh.getLastColumn(); if(w) sh.getRange(1,1,1,w).setFontWeight('bold').setBackground('#173A72').setFontColor('#fff').setWrap(true); });
  SpreadsheetApp.flush();
  SpreadsheetApp.getUi().alert('ตั้งค่าระบบเรียบร้อย','สร้าง/ตรวจสอบ Config, Template, Assignments, Submissions และ RosterRoomMap แล้ว\nหมายเหตุ: Template จะถูกล้างและสร้างหัวตารางใหม่ ส่วนข้อมูลในชีตอื่นจะไม่ถูกล้าง',SpreadsheetApp.getUi().ButtonSet.OK);
}

function showSmartPaste() { SpreadsheetApp.getUi().showModalDialog(HtmlService.createHtmlOutputFromFile('SmartPaste').setWidth(680).setHeight(610),'นำเข้ารายชื่อนักเรียน (Smart Paste)'); }
function showRosterLinkBuilder() { SpreadsheetApp.getUi().showModalDialog(HtmlService.createHtmlOutputFromFile('RosterLink').setWidth(760).setHeight(760),'สร้างลิงก์รายชื่อนักเรียน Academic+'); }

function getRosterLinkBuilderData() {
  const courses = getAllCourses_().map(c => ({code:c.code,subject:c.subject,rooms:c.rooms.split(',').map(s=>s.trim()).filter(Boolean)}));
  const mapSheet = SpreadsheetApp.getActive().getSheetByName(APP.SHEETS.ROSTER_MAP);
  const mappings = {};
  if (mapSheet && mapSheet.getLastRow() > 1) mapSheet.getRange(2,1,mapSheet.getLastRow()-1,APP.ROSTER_HEADERS.length).getValues().forEach(r => {
    const room=normalizeRoom_(r[0]); if(room) mappings[room]={room:room,year:String(r[1]||''),gradeId:String(r[2]||''),roomId:String(r[3]||''),sourceUrl:String(r[4]||'')};
  });
  const cfg=SpreadsheetApp.getActive().getSheetByName(APP.SHEETS.CONFIG);
  return {courses:courses,mappings:mappings,template:cfg?String(cfg.getRange('B8').getValue()||''):'',defaultSampleRoom:'',defaultSampleUrl:''};
}

function createRosterLink(payload) {
  payload=payload||{}; const room=normalizeRoom_(payload.targetRoom); if(!room) throw new Error('กรุณาระบุห้อง เช่น 6/17');
  const course=getCourse_(payload.courseCode); if(!course) throw new Error('ไม่พบรายวิชา');
  if(!course.rooms.split(',').map(s=>s.trim()).includes(room)) throw new Error('ห้อง '+room+' ไม่อยู่ในรายการห้องของวิชา '+course.code+'\nห้องที่ใช้ได้: '+course.rooms);
  const map=getRosterMapping_(room); if(!map) throw new Error('ยังไม่มีรหัส Academic+ ของห้อง '+room+' กรุณาขยายส่วน “ห้องใหม่ / บันทึก URL” แล้วบันทึก URL จริงของห้องนี้ก่อน');
  const url=buildAcademicUrl_(map);
  return {success:true,url:url,room:'ม.'+room,course:course.code+' '+course.subject,mapping:map};
}

function saveRosterRoomMapping(payload) {
  payload=payload||{}; const room=normalizeRoom_(payload.room); if(!room) throw new Error('กรุณาระบุห้องในรูปแบบ 6/17');
  const url=String(payload.sourceUrl||'').trim(); if(!/^https:\/\/academic\.mvsk\.ac\.th\//i.test(url)) throw new Error('กรุณาใช้ URL จริงจาก academic.mvsk.ac.th');
  const year=getQueryParam_(url,'year'), gradeId=getQueryParam_(url,'id_classyear'), roomId=getQueryParam_(url,'id_rooms');
  if(!year||!gradeId||!roomId) throw new Error('URL ต้องมีพารามิเตอร์ year, id_classyear และ id_rooms');
  const mapSheet=getOrCreateSheet_(SpreadsheetApp.getActive(),APP.SHEETS.ROSTER_MAP);
  resetHeaderIfNeeded_(mapSheet,APP.ROSTER_HEADERS);
  const last=mapSheet.getLastRow(); let targetRow=0;
  if(last>1){const vals=mapSheet.getRange(2,1,last-1,1).getDisplayValues(); for(let i=0;i<vals.length;i++) if(normalizeRoom_(vals[i][0])===room){targetRow=i+2;break;}}
  if(!targetRow) targetRow=mapSheet.getLastRow()+1;
  mapSheet.getRange(targetRow,1,1,APP.ROSTER_HEADERS.length).setValues([[room,year,gradeId,roomId,url,new Date()]]);
  mapSheet.getRange(targetRow,2,1,3).setNumberFormat('@'); mapSheet.getRange(1,1,1,APP.ROSTER_HEADERS.length).setFontWeight('bold').setBackground('#173A72').setFontColor('#fff');
  const mapping={room:room,year:year,gradeId:gradeId,roomId:roomId,sourceUrl:url};
  return {success:true,room:'ม.'+room,mapping:mapping,url:buildAcademicUrl_(mapping)};
}

function buildAcademicUrl_(mapping) {
  const base='https://academic.mvsk.ac.th/pretest/print-pdf';
  return base+'?year='+encodeURIComponent(mapping.year)+'&id_classyear='+encodeURIComponent(mapping.gradeId)+'&id_rooms='+encodeURIComponent(mapping.roomId);
}
function getRosterMapping_(room) {
  const sh=SpreadsheetApp.getActive().getSheetByName(APP.SHEETS.ROSTER_MAP); if(!sh||sh.getLastRow()<2)return null;
  const rows=sh.getRange(2,1,sh.getLastRow()-1,APP.ROSTER_HEADERS.length).getValues();
  for(let i=0;i<rows.length;i++) if(normalizeRoom_(rows[i][0])===room) return {room:room,year:String(rows[i][1]||''),gradeId:String(rows[i][2]||''),roomId:String(rows[i][3]||''),sourceUrl:String(rows[i][4]||'')};
  return null;
}
function getQueryParam_(url,key){const m=String(url).match(new RegExp('[?&]'+key+'=([^&#]*)','i'));return m?decodeURIComponent(m[1].replace(/\+/g,' ')).trim():'';}
function normalizeRoom_(value){return String(value||'').replace(/[\u200B-\u200D\uFEFF]/g,'').replace(/^\s*ม\s*\.\s*/i,'').replace(/\s+/g,'').replace(/-/g,'/').trim().replace(/^(\d{1,2})\/(\d{1,2})$/,(_,a,b)=>Number(a)+'/'+Number(b));}

function importStudentList(rawText,subjectCode) {
  if(!String(rawText||'').trim()) throw new Error('กรุณาวางข้อความรายชื่อนักเรียนก่อน');
  const course=getCourse_(subjectCode); if(!course) throw new Error('ไม่พบรหัสวิชาใน Config กรุณารันตั้งค่าระบบก่อน');
  const text=cleanInputText_(rawText); const roomMatch=text.match(/ม\s*\.\s*(\d{1,2})\s*\/\s*(\d{1,2})/i);
  if(!roomMatch) throw new Error('ไม่พบชื่อห้อง เช่น ม.5/17 ในข้อความที่วาง');
  const room=Number(roomMatch[1])+'/'+Number(roomMatch[2]);
  if(!course.rooms.split(',').map(s=>s.trim()).includes(room)) throw new Error('ห้อง ม.'+room+' ไม่อยู่ในรายการห้องของวิชา '+course.code+'\nห้องที่ตั้งไว้: '+course.rooms);
  const students=parseStudents_(text); if(!students.length) throw new Error('แยกรายชื่อไม่สำเร็จ ต้องมีเลขที่ + รหัสนักเรียน 5 หลัก + คำนำหน้าชื่อ + ชื่อ-นามสกุล');
  const ids=new Set(),nums=new Set(),dups=[]; students.forEach(s=>{if(ids.has(s.studentId))dups.push('รหัส '+s.studentId);if(nums.has(String(s.number)))dups.push('เลขที่ '+s.number);ids.add(s.studentId);nums.add(String(s.number));});
  if(dups.length) throw new Error('พบข้อมูลซ้ำในข้อความ: '+Array.from(new Set(dups)).join(', '));
  const ss=SpreadsheetApp.getActive(),name=makeClassSheetName_(course.code,room); let sh=ss.getSheetByName(name);
  if(!sh){const t=ss.getSheetByName(APP.SHEETS.TEMPLATE);if(!t)throw new Error('ไม่พบ Template กรุณาตั้งค่าระบบก่อน');sh=t.copyTo(ss).setName(name);sh.getRange(1,1,1,APP.CLASS_HEADERS.length).setValues([APP.CLASS_HEADERS]);styleClassSheet_(sh);}
  const width=APP.CLASS_HEADERS.length,last=sh.getLastRow(),oldRows=new Map(),roster=new Map();
  if(last>1) sh.getRange(2,1,last-1,width).getValues().forEach(r=>{const id=String(r[1]||'').trim().padStart(5,'0');if(!id.trim())return;oldRows.set(id,r);roster.set(id,{number:Number(r[0])||999,studentId:id,fullName:String(r[2]||'').trim()});});
  students.forEach(s=>roster.set(s.studentId,s));
  const all=Array.from(roster.values()).filter(s=>s.studentId&&s.fullName).sort((a,b)=>Number(a.number)-Number(b.number));
  const rows=all.map(s=>{const id=String(s.studentId).trim().padStart(5,'0'),r=Array(width).fill('');r[0]=Number(s.number);r[1]=id;r[2]=String(s.fullName||'').replace(/[\u200B-\u200D\uFEFF]/g,'').trim();const old=oldRows.get(id);if(old){[3,4,6,7,8,10].forEach(c=>r[c]=old[c]===undefined?'':old[c]);r[29]=old[29]||'';r[30]=old[30]||'';}return r;});
  if(sh.getMaxRows()<rows.length+1)sh.insertRowsAfter(sh.getMaxRows(),rows.length+1-sh.getMaxRows());
  if(sh.getLastRow()>1)sh.getRange(2,1,sh.getLastRow()-1,width).clearContent();
  if(rows.length)sh.getRange(2,1,rows.length,width).setValues(rows);
  applyClassFormulas_(sh,course);applyGradeStatusValidation_(sh);styleClassSheet_(sh);SpreadsheetApp.flush();
  return {success:true,room:'ม.'+room,course:course.code+' '+course.subject,sheetName:name,imported:students.length,totalInSheet:rows.length,message:'นำเข้ารายชื่อสำเร็จ'};
}
function parseStudents_(text){
  const re=/(\d{1,2})\s*(\d{5})\s*(เด็กชาย|เด็กหญิง|นางสาว|นาย)\s*/g, matches=[];let m;
  while((m=re.exec(text))!==null)matches.push({index:m.index,number:Number(m[1]),studentId:m[2],prefix:m[3],nameStart:m.index+m[0].length});
  const out=[];matches.forEach((x,i)=>{const end=i+1<matches.length?matches[i+1].index:text.length;let name=text.slice(x.nameStart,end).replace(/[\u200B-\u200D\uFEFF]/g,'').replace(/\s+/g,' ').trim();name=name.replace(/^(?:ชื่อ\s*-?\s*นามสกุล|ชื่อ-นามสกุล)\s*/i,'').trim();name=name.split(/(?:โรงเรียน|พิมพ์รายงาน|หน้าที่\s*\d|รวมทั้งหมด)/)[0].trim();if(name&&name.length<=100)out.push({number:x.number,studentId:x.studentId,fullName:x.prefix+name});});return out;
}
function cleanInputText_(v){return String(v||'').replace(/[\u200B-\u200D\uFEFF]/g,'').replace(/\u00A0/g,' ').replace(/\r\n?/g,'\n').trim();}

function applyClassFormulas_(sh,course){course=course||getCourseFromClassSheet_(sh.getName());
  sh.getRange('F2').setFormula('=ARRAYFORMULA(IF(A2:A="","",ROUND(D2:D+E2:E,0)))');
  sh.getRange('J2').setFormula('=ARRAYFORMULA(IF(A2:A="","",ROUND(H2:H+I2:I,0)))');
  sh.getRange('L2').setFormula('=ARRAYFORMULA(IF(A2:A="","",ROUND(F2:F,0)+ROUND(G2:G,0)+ROUND(J2:J,0)+ROUND(K2:K,0)))');
  sh.getRange('M2').setFormula('=ARRAYFORMULA(IF(A2:A="","",IF(AE2:AE<>"",AE2:AE,IF(L2:L="","",IF(L2:L>=80,4,IF(L2:L>=75,3.5,IF(L2:L>=70,3,IF(L2:L>=65,2.5,IF(L2:L>=60,2,IF(L2:L>=55,1.5,IF(L2:L>=50,1,0)))))))))))');
  const q='=ARRAYFORMULA(IF(A2:A="","",IF(M2:M="","",IF(((M2:M=4)+(M2:M=3.5)+(M2:M=3)+(M2:M=2.5))>0,3,IF(((M2:M=2)+(M2:M=1.5)+(M2:M=1))>0,2,IF(((M2:M=0)+(M2:M="ร")+(M2:M="มส"))>0,1,"")))))';
  for(let c=14;c<=21;c++)sh.getRange(2,c).setFormula(q);
  sh.getRange('V2').setFormula('=ARRAYFORMULA(IF(A2:A="","",IF(M2:M="","",IF((N2:N>=2)*(O2:O>=2)*(P2:P>=2)*(Q2:Q>=2)*(R2:R>=2)*(S2:S>=2)*(T2:T>=2)*(U2:U>=2),"ผ่าน","ไม่ผ่าน"))))');
  for(let c=23;c<=27;c++)sh.getRange(2,c).setFormula(q);
  sh.getRange('AB2').setFormula('=ARRAYFORMULA(IF(A2:A="","",IF(M2:M="","",IF((W2:W>=2)*(X2:X>=2)*(Y2:Y>=2)*(Z2:Z>=2)*(AA2:AA>=2),"ผ่าน","ไม่ผ่าน"))))');
  const cm=course?course.preWorkMax+course.preProjectMax+course.postWorkMax+course.postProjectMax:50,em=course?course.midMax+course.finalMax:50;
  sh.getRange('AC2').setFormula('=ARRAYFORMULA(IF(A2:A="","",IF(L2:L="","",IF(((MOD(L2:L,10)=4)+(MOD(L2:L,10)=9))>0,"ตรวจสอบคะแนนเฉียดเกรด",IF(((D2:D+E2:E+H2:H+I2:I)>='+Math.ceil(cm*.8)+')*((G2:G+K2:K)<'+Math.floor(em*.5)+'),"ตรวจสอบคะแนนเก็บ/สอบ","")))))');
}
function applyGradeStatusValidation_(sh){const rule=SpreadsheetApp.newDataValidation().requireValueInList(['ร','มส'],true).setAllowInvalid(false).setHelpText('เว้นว่างสำหรับเกรดตัวเลข หรือเลือก ร / มส').build();sh.getRange('AE2:AE').setDataValidation(rule);}

function generateDemoData(){const ui=SpreadsheetApp.getUi();if(ui.alert('สร้างข้อมูลจำลอง','สร้างคะแนนจำลองเฉพาะแถวที่ช่องคะแนน D,E,G,H,I,K ว่างทั้งหมดใช่หรือไม่?',ui.ButtonSet.YES_NO)!==ui.Button.YES)return;
  const ss=SpreadsheetApp.getActive(),as=ss.getSheetByName(APP.SHEETS.ASSIGNMENTS),su=ss.getSheetByName(APP.SHEETS.SUBMISSIONS);if(!as||!su)throw new Error('กรุณารันตั้งค่าระบบก่อน');let filled=0;
  getAllCourses_().forEach(c=>c.rooms.split(',').map(s=>s.trim()).filter(Boolean).forEach(room=>{const sh=ss.getSheetByName(makeClassSheetName_(c.code,room));if(!sh||sh.getLastRow()<2)return;const rows=sh.getRange(2,1,sh.getLastRow()-1,APP.CLASS_HEADERS.length).getValues();const valid=rows.filter(r=>r[1]&&r[2]);if(!valid.length)return;
    const defs=[['งานก่อนกลางภาค','PRE_WORK',c.preWorkMax],['ชิ้นงานก่อนกลางภาค','PRE_PROJECT',c.preProjectMax],['งานหลังกลางภาค','POST_WORK',c.postWorkMax],['ชิ้นงานหลังกลางภาค','POST_PROJECT',c.postProjectMax]];
    const existing=as.getLastRow()>1?as.getRange(2,1,as.getLastRow()-1,APP.ASSIGNMENT_HEADERS.length).getValues():[];const keys=new Set(existing.filter(r=>r[8]===true).map(r=>[r[1],r[2],r[3],r[4]].join('|')));const newA=[];const ids={};
    defs.forEach((d,i)=>{const title='Demo-'+d[0],key=[APP.TERM,c.code,room,title].join('|');let old=existing.find(r=>r[8]===true&&[r[1],r[2],r[3],r[4]].join('|')===key);if(old){ids[d[1]]={id:String(old[0]),max:Number(old[6])||d[2]};}else{const id=Utilities.getUuid();ids[d[1]]={id:id,max:d[2]};newA.push([id,APP.TERM,c.code,room,title,d[1],d[2],new Date(Date.now()+(i+1)*86400000),true,new Date()]);}});if(newA.length)appendRows_(as,newA);
    const subRows=[];const subKeys=new Set((su.getLastRow()>1?su.getRange(2,1,su.getLastRow()-1,APP.SUBMISSION_HEADERS.length).getValues():[]).filter(r=>r[8]===true&&r[2]===c.code&&r[3]===room).map(r=>String(r[0])+'|'+String(r[1])));
    valid.forEach(r=>defs.forEach(d=>{const a=ids[d[1]],key=String(r[1])+'|'+a.id;if(subKeys.has(key))return;const status=Math.random()<.75?'ตรงเวลา':Math.random()<.7?'ล่าช้า':'ขาดส่ง';const score=status==='ขาดส่ง'?0:randomScore_(a.max,status==='ตรงเวลา'?.65:.4,1);subRows.push([String(r[1]),a.id,c.code,room,status,status==='ขาดส่ง'?'':new Date(),score,a.max,true]);subKeys.add(key);}));if(subRows.length)appendRows_(su,subRows);
    const count=rows.length,cols=[3,4,6,7,8,10];const updates=rows.map(r=>{const x=cols.map(i=>r[i]);if(r[1]&&cols.every(i=>r[i]===''||r[i]===null)){r[3]=randomScore_(c.preWorkMax,.6,1);r[4]=randomScore_(c.preProjectMax,.6,1);r[6]=randomScore_(c.midMax,.4,1);r[7]=randomScore_(c.postWorkMax,.6,1);r[8]=randomScore_(c.postProjectMax,.6,1);r[10]=randomScore_(c.finalMax,.4,1);r[29]='DEMO';filled++;}return [r[3],r[4],r[6],r[7],r[8],r[10],r[29]||''];});if(count){sh.getRange(2,4,count,2).setValues(updates.map(r=>r.slice(0,2)));sh.getRange(2,7,count,1).setValues(updates.map(r=>[r[2]]));sh.getRange(2,8,count,2).setValues(updates.map(r=>r.slice(3,5)));sh.getRange(2,11,count,1).setValues(updates.map(r=>[r[5]]));sh.getRange(2,30,count,1).setValues(updates.map(r=>[r[6]]));applyClassFormulas_(sh,c);}
  }));const cfg=ss.getSheetByName(APP.SHEETS.CONFIG);if(cfg)cfg.getRange('B3').setValue(true);SpreadsheetApp.flush();ui.alert('สร้างข้อมูลจำลองเสร็จแล้ว','แถวคะแนนที่เติม: '+filled+'\nการสร้างรายการงาน/ส่งงานจะตรวจสอบรายการเดิมก่อนเพิ่ม',ui.ButtonSet.OK);
}
function clearDemoData(){const ui=SpreadsheetApp.getUi();if(ui.alert('ล้างข้อมูลจำลอง','ลบเฉพาะรายการ IsDemo = TRUE และล้างคะแนนเฉพาะแถว DemoFlag = DEMO ใช่หรือไม่?',ui.ButtonSet.YES_NO)!==ui.Button.YES)return;const ss=SpreadsheetApp.getActive();const da=deleteRowsByDemoFlag_(ss.getSheetByName(APP.SHEETS.ASSIGNMENTS),9),ds=deleteRowsByDemoFlag_(ss.getSheetByName(APP.SHEETS.SUBMISSIONS),9);let cleared=0;
  ss.getSheets().forEach(sh=>{if(!isClassSheet_(sh.getName())||sh.getLastRow()<2)return;const n=sh.getLastRow()-1,v=sh.getRange(2,1,n,APP.CLASS_HEADERS.length).getValues();const scoreCols=[4,5,7,8,9,11];const clearRows=[];v.forEach((r,i)=>{if(r[29]==='DEMO')clearRows.push(i+2)});clearRows.forEach(r=>{scoreCols.forEach(c=>sh.getRange(r,c).clearContent());sh.getRange(r,30).clearContent();});cleared+=clearRows.length;applyClassFormulas_(sh);});const cfg=ss.getSheetByName(APP.SHEETS.CONFIG);if(cfg)cfg.getRange('B3').setValue(false);SpreadsheetApp.flush();ui.alert('ล้างข้อมูลจำลองแล้ว','งานที่ลบ: '+da+'\nรายการส่งงานที่ลบ: '+ds+'\nแถวคะแนนที่ล้าง: '+cleared,ui.ButtonSet.OK);}
function validateClassrooms(){const ss=SpreadsheetApp.getActive(),report=[];ss.getSheets().forEach(sh=>{if(!isClassSheet_(sh.getName())||sh.getLastRow()<2)return;const rows=sh.getRange(2,1,sh.getLastRow()-1,3).getDisplayValues(),ids=new Set();rows.forEach((r,i)=>{if(!r[1]&&!r[2])return;const id=String(r[1]).trim();if(!/^\d{5}$/.test(id))report.push(sh.getName()+' แถว '+(i+2)+': เลขประจำตัวไม่ครบ 5 หลัก');if(id&&ids.has(id))report.push(sh.getName()+': เลขประจำตัวซ้ำ '+id);if(id)ids.add(id);if(!r[2])report.push(sh.getName()+' แถว '+(i+2)+': ไม่มีชื่อ-นามสกุล');});});SpreadsheetApp.getUi().alert('ผลตรวจสอบข้อมูล',report.length?report.slice(0,60).join('\n'):'ไม่พบข้อผิดพลาดเบื้องต้น',SpreadsheetApp.getUi().ButtonSet.OK);}

function getOrCreateSheet_(ss,name){return ss.getSheetByName(name)||ss.insertSheet(name);}
function resetHeaderIfNeeded_(sh,headers){if(sh.getLastRow()===0){sh.getRange(1,1,1,headers.length).setValues([headers]);return;}const current=sh.getRange(1,1,1,headers.length).getDisplayValues()[0];if(current.join('|')!==headers.join('|'))throw new Error('หัวตารางชีต '+sh.getName()+' ไม่ตรงกับระบบ กรุณาสำรองและตรวจสอบก่อนแก้ไข');}
function styleClassSheet_(sh){const w=APP.CLASS_HEADERS.length;sh.setFrozenRows(1);sh.getRange(1,1,1,w).setFontWeight('bold').setBackground('#173A72').setFontColor('#fff').setWrap(true);sh.setColumnWidth(1,60);sh.setColumnWidth(2,110);sh.setColumnWidth(3,220);sh.setColumnWidths(4,w-3,125);sh.getRange('B:B').setNumberFormat('@');sh.getRange('D:AC').setVerticalAlignment('middle');sh.getRange('AD:AE').setHorizontalAlignment('center');}
function getAllCourses_(){const sh=SpreadsheetApp.getActive().getSheetByName(APP.SHEETS.CONFIG);if(!sh||sh.getLastRow()<3)throw new Error('กรุณารัน setupSystem ก่อน');return sh.getRange(3,4,sh.getLastRow()-2,10).getValues().filter(r=>String(r[0]||'').trim()).map(r=>({code:String(r[0]).trim(),subject:String(r[1]||'').trim(),rooms:String(r[2]||'').replace(/ม\./g,'').trim(),preWorkMax:Number(r[3])||0,preProjectMax:Number(r[4])||0,midMax:Number(r[5])||0,postWorkMax:Number(r[6])||0,postProjectMax:Number(r[7])||0,finalMax:Number(r[8])||0,classroomCourseId:String(r[9]||'').trim()}));}
function getCourse_(code){return getAllCourses_().find(c=>c.code===String(code||'').trim());}
function getCourseFromClassSheet_(name){return getCourse_(String(name).split('_')[0]);}
function makeClassSheetName_(code,room){return code+'_ม'+String(room).replace('/','-');}
function isClassSheet_(name){return /^ส\d{5}_ม\d+-\d+$/.test(name);}
function appendRows_(sh,rows){if(!rows||!rows.length)return;const start=Math.max(sh.getLastRow()+1,2),end=start+rows.length-1;if(sh.getMaxRows()<end)sh.insertRowsAfter(sh.getMaxRows(),end-sh.getMaxRows());sh.getRange(start,1,rows.length,rows[0].length).setValues(rows);}
function randomScore_(max,minRatio,maxRatio){if(!max||max<0)return 0;return Math.min(max,Math.round(max*(minRatio+Math.random()*(maxRatio-minRatio))));}
function deleteRowsByDemoFlag_(sh,col){if(!sh||sh.getLastRow()<2)return 0;const vals=sh.getRange(2,col,sh.getLastRow()-1,1).getValues();const indexes=[];vals.forEach((r,i)=>{if(r[0]===true||String(r[0]).toUpperCase()==='TRUE')indexes.push(i+2);});let deleted=0;for(let i=indexes.length-1;i>=0;i--){sh.deleteRow(indexes[i]);deleted++;}return deleted;}

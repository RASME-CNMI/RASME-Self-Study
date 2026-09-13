// Local-only tests. No Google services or outbound requests are used.
const fs = require('node:fs');
const vm = require('node:vm');
const assert = require('node:assert/strict');
const path = require('node:path');
const root = path.join(__dirname, '..');
const source = fs.readFileSync(path.join(root, 'Code.js'), 'utf8');
const html = fs.readFileSync(path.join(root, 'index.html'), 'utf8');
const manifest = JSON.parse(fs.readFileSync(path.join(root, 'appsscript.json'), 'utf8'));
new vm.Script(source);
for (const match of html.matchAll(/<script\b[^>]*>([\s\S]*?)<\/script>/g)) {
  new vm.Script(match[1].replace(/<\?[\s\S]*?\?>/g, 'null'));
}
let passed = 0;
function test(name, run) { run(); passed++; console.log('PASS ' + name); }
function context(overrides = {}) {
  const c = vm.createContext({ ...overrides });
  vm.runInContext(source, c);
  Object.assign(c, overrides);
  return c;
}
test('Admin provisioning is not a public RPC', () => {
  const c = context();
  assert.equal(typeof c.setupAdminAccount, 'undefined');
  assert.equal(typeof c.setupAdminAccount_, 'function');
});
test('Blank sheet rows preserve physical row numbers', () => {
  const c = context();
  const rows = c.rows_({getDataRange: () => ({getValues: () => [['ID'], ['a'], [''], ['b']]})});
  assert.equal(rows[0]._row, 2); assert.equal(rows[1]._row, 4);
});
test('Disabled and missing accounts cannot use authenticated RPCs', () => {
  const c = context({sessionUser_: () => ({Status:'Disabled'})});
  assert.throws(() => c.require_('token'), /เข้าสู่ระบบใหม่/);
  c.sessionUser_ = () => null;
  assert.throws(() => c.require_('token'), /เข้าสู่ระบบใหม่/);
  c.sessionUser_ = () => ({Status:'Active'});
  assert.equal(c.require_('token').Status, 'Active');
});
test('Mutation lock releases on failure without forcing a redundant flush', () => {
  const events=[];
  const c=context({ScriptApp:{},ScriptLock:{},LockService:{getScriptLock:()=>({waitLock:()=>events.push('lock'),releaseLock:()=>events.push('release')})},SpreadsheetApp:{flush:()=>events.push('flush')}});
  assert.throws(()=>c.withMutationLock_(()=>{throw Error('test')}),/test/);
  assert.deepEqual(events,['lock','release']);
});
test('Notification network processing runs outside the mutation lock', () => {
  const locked=source.match(/function processAutomationsLocked_\(\)\{[\s\S]*?\n\}/)?.[0]||'';
  assert(!locked.includes('processNotificationQueue_()'));
  assert(source.includes('processNotificationQueue_();const result=withMutationLock_'));
});
test('Evaluation rejects NaN, fractions and out-of-range scores before writes', () => {
  const c=context({require_:()=>({UserID:'u'}),find_:()=>({UserID:'u',Status:'CheckedOut'}),SpreadsheetApp:{getActiveSpreadsheet:()=>({getSheetByName:()=>({})})}});
  for(const score of ['bad',NaN,0,6,1.5]) assert.throws(()=>c.submitEvaluationLocked_('t','b',score,''),/1–5/);
});
test('Booking rechecks procedure status on server', () => {
  const c=context({require_:()=>({UserID:'u'}),find_:(sheet)=>sheet==='Users'?{UserID:'u',Status:'Active'}:sheet==='Sessions'?{SessionID:'s',ProcedureID:'p',Status:'Open'}:sheet==='Procedures'?{ProcedureID:'p',Status:'ComingSoon'}:undefined,refreshBan_:()=>{},banned_:()=>false,cachedSettings_:()=>({}),SpreadsheetApp:{getActiveSpreadsheet:()=>({getSheetByName:n=>n})},LockService:{getScriptLock:()=>({waitLock(){},releaseLock(){}})},rows_:()=>[]});
  assert.throws(()=>c.createBooking('t','s'),/ยังไม่เปิดให้จอง/);
});
test('Booking success is returned even when notification queueing fails', () => {
  const sheets={Users:'Users',ActiveBookings:'ActiveBookings',Sessions:'Sessions',Procedures:'Procedures'};
  const c=context({require_:()=>({UserID:'u'}),find_:(sheet)=>sheet==='Users'?{UserID:'u',Status:'Active',EvaluationPending:false}:sheet==='Sessions'?{SessionID:'s',ProcedureID:'p',Status:'Open'}:sheet==='Procedures'?{ProcedureID:'p',Status:'Active'}:undefined,refreshBan_:()=>{},banned_:()=>false,cachedSettings_:()=>({}),SpreadsheetApp:{getActiveSpreadsheet:()=>({getSheetByName:name=>sheets[name]||name})},LockService:{getScriptLock:()=>({waitLock(){},releaseLock(){}})},rows_:()=>[],sessionView_:()=>({canBook:true,booked:0,maxCapacity:4}),appendBooking_:()=>{},bookingView_:booking=>({id:booking.BookingID,status:booking.Status}),id_:()=> 'BK-001',notifyBooking_:()=>{throw Error('queue unavailable')}});
  const result=c.createBooking('t','s');
  assert.equal(result.ok,true);
  assert.equal(result.bookingId,'BK-001');
  assert.equal(result.notificationQueued,false);
});
test('Duplicate booking returns the existing transaction and does not notify twice', () => {
  let notifications=0;
  const c=context({require_:()=>({UserID:'u'}),find_:()=>({UserID:'u',Status:'Active',EvaluationPending:false}),refreshBan_:()=>{},banned_:()=>false,cachedSettings_:()=>({}),SpreadsheetApp:{getActiveSpreadsheet:()=>({getSheetByName:name=>name})},LockService:{getScriptLock:()=>({waitLock(){},releaseLock(){}})},rows_:sheet=>sheet==='ActiveBookings'?[{BookingID:'BK-OLD',UserID:'u',SessionID:'s',Status:'Reserved'}]:[],bookingView_:booking=>({id:booking.BookingID,status:booking.Status}),notifyBooking_:()=>{notifications++}});
  const result=c.createBooking('t','s');
  assert.equal(result.ok,true);
  assert.equal(result.alreadyProcessed,true);
  assert.equal(result.bookingId,'BK-OLD');
  assert.equal(notifications,0);
});
test('Dashboard booking lookup supplies a predicate function', () => {
  const c=context({require_:()=>({UserID:'u',Status:'Active'}),reconcileUserNoShows_:()=>{},find_:()=>({UserID:'u',Status:'Active'}),refreshBan_:()=>{},SpreadsheetApp:{getActiveSpreadsheet:()=>({getSheetByName:name=>name})},rowsMatchingColumn_:(sheet,key,predicate)=>{assert.equal(typeof predicate,'function');assert.equal(predicate('u'),true);assert.equal(predicate('other'),false);return[]},cachedSheetRows_:()=>[],cachedSettings_:()=>({}),publicUser_:()=>({id:'u'})});
  assert.deepEqual(JSON.parse(JSON.stringify(c.getMyDashboard('t'))),{user:{id:'u'},bookings:[]});
});
test('Booking UI validates the server contract, recovers saved bookings, and hides predicate errors', () => {
  const line=html.split('\n').find(item=>item.includes('async function book(id)'))||'';
  assert(line.includes('result.ok!==true'));
  assert(line.includes("dashboard.bookings||[]"));
  assert(line.includes('ไม่สามารถตรวจสอบผลการจองได้'));
  assert(html.includes('predicate is not a function'));
});
test('Protected writes share lock wrapper', () => {
  for(const name of ['registerUser','resetPassword','submitEvaluation','submitEvaluationV4','adminCompleteCheckout','adminUpdateBookingStatus','adminQuickCheckIn']) {
    const c=context({withMutationLock_:()=> 'locked'});
    assert.equal(c[name](), 'locked');
  }
});
test('HTML IDs are unique', () => {
  const ids=[...html.matchAll(/\bid="([^"]+)"/g)].map(m=>m[1]);
  assert.equal(ids.length,new Set(ids).size);
});
test('Procedure catalogue cards omit location and use a subtle boundary', () => {
  const renderLine=html.split('\n').find(line=>line.includes('function renderProcedures()'))||'';
  assert(!renderLine.includes('procedureCardPlace_'));
  assert(!renderLine.includes('procedure-guidance'));
  assert(html.includes('#procedureGrid{column-gap:18px;row-gap:18px}.procedure{min-height:230px;border:1px solid #d9dee4'));
  assert(html.includes('box-shadow:0 4px 14px rgba(16,24,40,.05)'));
});
test('Procedure catalogue groups available sessions into current and next Thai months', () => {
  const c=context();
  const procedures=[
    {ProcedureID:'p1',ProcedureName:'หัตถการ ก',Status:'Active',SortOrder:1},
    {ProcedureID:'p2',ProcedureName:'หัตถการ ข',Status:'Active',SortOrder:2}
  ];
  const sessions=[
    {ProcedureID:'p1',SessionDate:'2026-09-20',Status:'Open'},
    {ProcedureID:'p1',SessionDate:'2026-09-20',Status:'Confirmed'},
    {ProcedureID:'p2',SessionDate:'2026-10-05',Status:'Full'},
    {ProcedureID:'p2',SessionDate:'2026-11-05',Status:'Open'},
    {ProcedureID:'p2',SessionDate:'2026-09-10',Status:'Open'},
    {ProcedureID:'p2',SessionDate:'2026-09-25',Status:'Cancelled'}
  ];
  const result=JSON.parse(JSON.stringify(c.scheduleOverviewFromRows_(procedures,sessions,'2026-09-12')));
  assert.equal(result[0].label,'กันยายน 2569');
  assert.equal(result[1].label,'ตุลาคม 2569');
  assert.deepEqual(result[0].items,[{id:'p1',name:'หัตถการ ก',dates:['20 กันยายน 2569']}]);
  assert.deepEqual(result[1].items,[{id:'p2',name:'หัตถการ ข',dates:['5 ตุลาคม 2569']}]);
});
test('Procedure catalogue explains booking lead time and renders both schedule periods', () => {
  assert(html.includes('จองได้ล่วงหน้า 7 วัน ก่อนวันที่ต้องการฝึก'));
  assert(html.includes('รายการหัตถการสำหรับรอบเดือนนี้'));
  assert(html.includes('รายการหัตถการสำหรับรอบเดือนถัดไป'));
  assert(html.includes('id="procedureScheduleOverview"'));
  assert(source.includes('scheduleOverview:data.scheduleOverview'));
  assert(source.includes("cachedSheetRows_('Sessions',60)"));
  assert(source.includes("['Procedures','Sessions'].includes(name)"));
});
test('Procedure detail uses icon facts, official hours and concise labels', () => {
  assert(html.includes('aria-label="สถานที่"'));
  assert(html.includes('aria-label="เวลาทำการ"'));
  assert(html.includes('วันจันทร์ - ศุกร์ : 08.00 - 16.30 น.'));
  assert(html.includes('เสาร์ - อาทิตย์ : ปิดให้บริการ'));
  assert(!html.includes('<b>วันจันทร์ - ศุกร์'));
  assert(html.includes('และไม่เกินจำนวนรองรับสูงสุด'));
  assert(html.includes('ต้อง Check-in หลังจากถึงรอบภายใน ${state.settings.noShowGraceMinutes} นาที</li>'));
  assert(!html.includes('มิฉะนั้นจะถูกบันทึกเป็น No-show'));
  assert(!html.includes('<li>มาถึงก่อนเวลาเริ่มรอบ 10 นาที</li>'));
  assert(html.includes('<h3>กฎ / Information</h3>'));
  assert(html.includes('<li>ผู้ใช้บริการต้องเคยผ่านการเรียนหรือ Workshop ในเนื้อหาที่เกี่ยวข้องกับหัตถการที่ต้องการฝึกมาก่อน</li>${procedureRules.map'));
  assert(html.includes('prerequisiteRulePattern=/(?:RAID\\s*419|ผู้เรียนต้องเคยผ่านการเรียนหัวข้อ'));
  assert(html.includes('!prerequisiteRulePattern.test(x)'));
  assert(html.includes('กด “ขอ Check-in” แล้วแจ้งรหัส Check-in 3 ตัวให้เจ้าหน้าที่ตรวจสอบข้อมูลและยืนยัน'));
  assert(html.includes('<h3>เงื่อนไขการเข้าใช้งาน</h3>'));
  assert(!html.includes('<h3>ข้อมูลก่อนจอง</h3>'));
  assert(!html.includes('<h3>ขั้นตอนการเข้าใช้บริการ</h3>'));
  assert(!html.includes('<h3>กฎการใช้พื้นที่และอุปกรณ์</h3>'));
  assert(html.includes('systemRulePattern='));
  assert(html.includes('onclick="showSlots()">รอบทั้งหมด</button>'));
  assert(!html.includes('ตรวจสอบรอบเวลา</button>'));
});
test('Procedure detail reads only four required datasets, not full history/users', () => {
  const reads=[];
  const c=context({require_:()=>({}),CacheService:{getScriptCache:()=>({get:()=>null,put:()=>{}})},Utilities:{formatDate:()=> '2026-09-05'},SpreadsheetApp:{getActiveSpreadsheet:()=>({getSheetByName:n=>n})},rows_:n=>{reads.push(n);return[];},cachedSheetRows_:n=>{reads.push(n);return n==='Procedures'?[{ProcedureID:'p',ProcedureName:'Test',Status:'Active'}]:[];},cachedSettings_:()=>{reads.push('Settings');return{};}});
  const result=c.getProcedureDetail('t','p');
  assert.equal(result.procedure.id,'p');
  assert.deepEqual(reads,['ActiveBookings','Sessions','Procedures','Settings']);
});
test('Single-record lookup reads the key column and matched row only', () => {
  const reads=[];
  const values=[['ID','Name'],['a','Alpha'],['b','Beta']];
  const sheet={getName:()=>'',getLastColumn:()=>2,getLastRow:()=>3,getRange:(row,col,count,width)=>({getValues:()=>{reads.push([row,col,count,width]);return values.slice(row-1,row-1+count).map(r=>r.slice(col-1,col-1+width));}})};
  const c=context();
  assert.equal(c.find_(sheet,'ID','b').Name,'Beta');
  assert.deepEqual(reads,[[1,1,1,2],[2,1,2,1],[3,1,1,2]]);
});
test('Context indexes preserve booking counts while avoiding repeated full scans', () => {
  const c=context({Utilities:{formatDate:()=> '2026-09-05'}}),ctx={bookings:[{SessionID:'s',Status:'Reserved'}],procedures:[{ProcedureID:'p',ProcedureName:'Skill'}],settings:{BOOKING_CLOSE_DAYS:7,MIN_PARTICIPANTS:1,DEFAULT_CAPACITY:4}};
  const session={SessionID:'s',ProcedureID:'p',SessionDate:'2099-09-05',StartTime:'12:00',EndTime:'13:00',Status:'Open'};
  assert.equal(c.sessionView_(session,ctx).booked,1);
  ctx.bookings.push({SessionID:'s',Status:'Confirmed'});
  assert.equal(c.sessionView_(session,ctx).booked,2);
});
test('Cached ISO time values render as clock time instead of year 1899', () => {
  const c=context({Utilities:{formatDate:(value,tz,format)=>format==='HH:mm'?'13:00':'2026-10-20'}});
  assert.equal(c.time_('1899-12-30T06:00:00.000Z'),'13:00');
  assert.equal(c.time_('9:05'),'09:05');
  assert(!c.time_('1899-12-30T06:00:00.000Z').includes('1899'));
});
test('Cached ISO session dates are restored in Bangkok timezone without shifting one day early', () => {
  const c=context({Utilities:{formatDate:(value,timeZone,format)=>{
    assert.equal(timeZone,'Asia/Bangkok');
    return format==='yyyy-MM-dd'?'2026-10-14':'';
  }}});
  assert.equal(c.sheetDate_('2026-10-13T17:00:00.000Z'),'2026-10-14');
  assert.equal(c.sheetDate_('2026-10-14'),'2026-10-14');
});
test('Admin history reads only the configured tail of the master sheet', () => {
  const calls=[],values=Array.from({length:501},(_,i)=>[i?'b'+i:'BookingID',i?'u':'UserID']);
  const sheet={getName:()=> 'Bookings',getLastRow:()=>501,getRange:(row,col,count,width)=>({getValues:()=>{calls.push([row,col,count,width]);return values.slice(row-1,row-1+count).map(r=>r.slice(col-1,col-1+width));}})};
  const c=context();
  const rows=c.tailRows_(sheet,300);
  assert.equal(rows.length,300);
  assert.equal(rows[0].BookingID,'b201');
  assert.deepEqual(calls,[[202,1,300,11]]);
  assert(source.includes('bookingHistory:tailRows_(master,300)'));
});
test('Terminal booking updates master then deletes active row without redundant write', () => {
  const calls=[];
  const c=context({SpreadsheetApp:{getActiveSpreadsheet:()=>({getSheetByName:n=>n})},update_:(n)=>calls.push('update '+n),deleteWhere_:(n,p)=>{assert(p({BookingID:'b'}));calls.push('delete '+n)},invalidateBookingCaches_:()=>calls.push('invalidate')});
  c.updateBooking_('b',{Status:'Cancelled'},true);
  assert.deepEqual(calls,['update Bookings','delete ActiveBookings','invalidate']);
  calls.length=0;c.updateBooking_('b',{Status:'CheckedIn'},false);
  assert.deepEqual(calls,['update Bookings','update ActiveBookings','invalidate']);
});
test('Legacy QR and scanner server entry points are removed', () => {
  const c=context();
  for(const name of ['attendanceByQr','attendanceByQrAuto','getRoomQr','getSessionQr','createAdminScannerHandoff','resumeAdminScannerHandoff']) assert.equal(typeof c[name],'undefined');
});
test('No-show is applied at the grace deadline and reconciled when dashboard opens', () => {
  assert(source.includes('r.start.getTime()+grace*60000'));
  assert(source.includes('reconcileUserNoShows_(auth.UserID)'));
  assert(source.includes("everyMinutes(1)"));
  assert(!source.includes('(grace+buffer)*60000'));
});
test('Check-in keeps server-side role, status and time-window validation', () => {
  assert(source.includes("if(String(admin.Role)!=='Admin')throw Error('ไม่มีสิทธิ์')"));
  assert(source.includes("if(!['Reserved','Confirmed'].includes(String(booking.Status)))"));
  assert(source.includes('if(now<opens||now>closes)'));
});
test('Current check-in flow uses a three-digit booking reference code', () => {
  assert(source.includes('แจ้งรหัส 3 ตัวให้เจ้าหน้าที่ตรวจสอบและยืนยัน'));
  assert(html.includes('รหัสสำหรับ Check-in'));
  assert(html.includes('รหัส 3 ตัวท้ายจากรหัสอ้างอิง'));
  assert(html.includes('กรอกรหัส Check-in'));
  assert(html.includes('ตรวจสอบข้อมูลผู้จอง'));
  assert(html.includes('class="admin-code-form"'));
  assert(html.includes('class="admin-code-actions"'));
  assert(html.includes('type="button" onclick="closeModal()">ย้อนกลับ</button>'));
});
test('Student check-in status endpoint authorizes ownership and returns a stable contract', () => {
  const sheets={ActiveBookings:{name:'ActiveBookings'},Bookings:{name:'Bookings'}};
  const c=context({require_:()=>({UserID:'u1'}),SpreadsheetApp:{getActiveSpreadsheet:()=>({getSheetByName:name=>sheets[name]})},find_:(sheet,key,id)=>sheet.name==='ActiveBookings'?{BookingID:id,UserID:'u1',Status:'CheckedIn',CheckInAt:new Date('2026-09-09T05:05:00Z')}:undefined,Utilities:{formatDate:()=> '09/09/2026 12:05'}});
  const result=c.getMyBookingCheckInStatus('token','BK-001');
  assert.deepEqual(JSON.parse(JSON.stringify(result)),{ok:true,bookingId:'BK-001',status:'CheckedIn',checkInAt:'09/09/2026 12:05',terminal:true});
  c.find_=()=>({BookingID:'BK-002',UserID:'u2',Status:'Reserved'});
  assert.throws(()=>c.getMyBookingCheckInStatus('token','BK-002'),/ไม่พบรายการจองของคุณ/);
});
test('Student check-in polling is lightweight, single-flight and lifecycle bounded', () => {
  assert(html.includes("call('getMyBookingCheckInStatus',state.token,bookingId)"));
  assert(!html.split('\n').find(line=>line.includes('async function pollCheckInStatus_')).includes('getMyDashboard'));
  assert(html.includes('state.checkInPollInFlight'));
  assert(html.includes('generation!==state.checkInPollGeneration'));
  assert(html.includes('Date.now()>=state.checkInPollDeadline'));
  assert(html.includes('if(document.hidden)return'));
  assert(html.includes("document.addEventListener('visibilitychange'"));
  assert(html.includes('clearTimeout(state.checkInPollTimer)'));
  assert(html.includes('function closeModal(){stopCheckInPolling_();'));
  assert(html.includes('function showPage(name){stopCheckInPolling_();'));
});
test('Student check-in modal updates locally on success and handles terminal states', () => {
  const pollLine=html.split('\n').find(line=>line.includes('async function pollCheckInStatus_'))||'';
  assert(pollLine.includes("result.status==='CheckedIn'"));
  assert(pollLine.includes("booking.status='CheckedIn'"));
  assert(pollLine.includes("booking.checkInAt=result.checkInAt||''"));
  assert(html.includes('<h2>Check-in สำเร็จ</h2>'));
  assert(html.includes('เจ้าหน้าที่บันทึกเวลาเข้าใช้บริการเรียบร้อยแล้ว'));
  assert(html.includes('เวลา Check-in:'));
  assert(html.includes('กลับไปยังการจองของฉัน'));
  assert(pollLine.includes('if(result&&result.terminal)'));
});
test('Student check-in polling uses backoff and never performs a check-in mutation', () => {
  const pollLine=html.split('\n').find(line=>line.includes('async function pollCheckInStatus_'))||'';
  assert(pollLine.includes('Math.min(10000'));
  assert(!pollLine.includes('adminQuickCheckIn'));
  assert(!pollLine.includes('update'));
  assert(!source.match(/function getMyBookingCheckInStatus[^\n]+(?:append|update_|deleteWhere_|sendMail_|telegram_)/));
});
test('Duplicate three-digit codes require the admin to choose the correct booking', () => {
  assert(html.includes('matches.length===1'));
  assert(html.includes('matches.map'));
  assert(html.includes('เลือกรายการที่ต้องการ Check-in'));
});
test('Evaluation V2 accepts confidence zero and ten without treating zero as missing', () => {
  const c=context(),a={satisfactionBooking:1,satisfactionEquipment:2,satisfactionEnvironment:3,satisfactionOverall:5,confidenceBefore:0,confidenceAfter:10,feedback:''};
  assert.equal(c.validateEvaluationV2_(a).confidenceBefore,0);
  assert.equal(c.validateEvaluationV2_(a).confidenceAfter,10);
  for(const bad of [undefined,null,'',false,-1,11,1.5,NaN]) assert.throws(()=>c.validateEvaluationV2_({...a,confidenceBefore:bad}));
  for(const bad of [0,6,undefined]) assert.throws(()=>c.validateEvaluationV2_({...a,satisfactionOverall:bad}));
  assert.throws(()=>c.validateEvaluationV2_({...a,feedback:'x'.repeat(2001)}));
});
test('Evaluation header upgrade appends without changing legacy columns', () => {
  const c=context(),base=['EvaluationID','BookingID','UserID','Rating','Feedback','SubmittedAt'];let written;
  let expanded=0;
  const sheet={getLastColumn:()=>6,getMaxColumns:()=>26,insertColumnsAfter:(col,n)=>{assert.equal(col,26);expanded=n},getRange:(r,col,n,w)=>({getValues:()=>[base],setValues:v=>{written={r,col,n,w,v}}})};
  const headers=c.ensureEvaluationHeaders_(sheet);
  assert.deepEqual(Array.from(headers.slice(0,6)),base);assert.equal(written.r,1);assert.equal(written.col,7);
  assert(headers.includes('ConfidenceBefore'));assert(headers.includes('ConfidenceAfter'));assert(headers.includes('ConfidenceMethod'));assert.equal(expanded,headers.length-26);
});
test('Evaluation V3 requires four satisfaction scores and six confidence scores 1-10', () => {
  const c=context(),a={satisfactionStaff:5,satisfactionEquipment:4,satisfactionEnvironment:3,satisfactionOverall:4,confidenceCoreBefore:1,confidenceCoreAfter:10,confidenceSafetyBefore:5,confidenceSafetyAfter:7,confidenceComplicationsBefore:6,confidenceComplicationsAfter:4};
  assert.equal(c.validateEvaluationV3_(a).confidenceCoreBefore,1);
  assert.equal(c.validateEvaluationV3_(a).confidenceCoreAfter,10);
  assert.equal(c.validateEvaluationV3_(a).confidenceComplicationsAfter,4);
  for(const key of Object.keys(a)) {
    for(const bad of [undefined,null,'',false,0,11,1.5,NaN]) assert.throws(()=>c.validateEvaluationV3_({...a,[key]:bad}));
  }
  assert.throws(()=>c.validateEvaluationV3_({...a,satisfactionStaff:6}));
});
test('Evaluation V4 requires six satisfaction scores, before-after confidence and three comments', () => {
  const c=context(),a={satisfactionStaff:5,satisfactionEquipment:4,satisfactionEnvironment:3,satisfactionOverall:4,satisfactionReturn:5,satisfactionRecommend:5,confidenceSelfPracticeBefore:2,confidenceSelfPracticeAfter:4,mostImpressed:'บริการรวดเร็ว',improvementNeeded:'เพิ่มอุปกรณ์',otherSuggestion:'ไม่มี'};
  const valid=c.validateEvaluationV4_(a);
  assert.equal(valid.satisfactionReturn,5);
  assert.equal(valid.confidenceSelfPracticeBefore,2);
  assert.equal(valid.confidenceSelfPracticeAfter,4);
  for(const key of ['satisfactionStaff','satisfactionEquipment','satisfactionEnvironment','satisfactionOverall','satisfactionReturn','satisfactionRecommend','confidenceSelfPracticeBefore','confidenceSelfPracticeAfter']) {
    for(const bad of [undefined,null,'',false,0,6,1.5,NaN]) assert.throws(()=>c.validateEvaluationV4_({...a,[key]:bad}),/1–5/);
  }
  for(const key of ['mostImpressed','improvementNeeded','otherSuggestion']) assert.throws(()=>c.validateEvaluationV4_({...a,[key]:'   '}),/กรุณากรอก/);
});
test('Current evaluation UI uses V4 questions and requires every comment', () => {
  const openLine=html.split('\n').find(line=>line.includes('function openEvaluation(id)'))||'';
  const sendLine=html.split('\n').find(line=>line.includes('async function sendEvaluation(id)'))||'';
  assert(openLine.includes('5. คุณมีโอกาสที่จะกลับมาฝึก Self-study ที่นี่อีกครั้ง'));
  assert(openLine.includes('6. คุณจะแนะนำการบริการฝึก Self-study นี้ให้แก่ผู้อื่น'));
  assert(openLine.includes('confidenceSelfPracticeBefore'));
  assert(openLine.includes('confidenceSelfPracticeAfter'));
  assert(openLine.includes('ก่อนฝึก — '));
  assert(openLine.includes('หลังฝึก — '));
  assert(!openLine.includes('confidenceCore'));
  assert.equal((openLine.match(/textarea[^>]+required/g)||[]).length,3);
  assert(sendLine.includes("call('submitEvaluationV4'"));
  assert(source.includes("'ConfidenceSelfPracticeBefore','ConfidenceSelfPracticeAfter','ConfidenceSelfPracticeChange'"));
});
test('Confidence question has two required fields with ten options each, no default score', () => {
  const c=vm.createContext({esc:s=>s});
  vm.runInContext(html.split('\n').find(l=>l.includes('function confidencePair_(')),c);
  const markup=c.confidencePair_('confidenceCore','Question');
  assert.equal((markup.match(/<select /g)||[]).length,2);
  assert.equal((markup.match(/ required/g)||[]).length,2);
  assert.equal((markup.match(/selected disabled/g)||[]).length,2);
  assert(!markup.includes('value="0"'));
  for(let i=1;i<=10;i++)assert.equal((markup.match(new RegExp('value="'+i+'"','g'))||[]).length,2);
});
test('Public links always use a validated production /exec URL', () => {
  const production='https://script.google.com/macros/s/AKfycbxsR84XffxsZpIGA1ZMuoZNdK-sMaXx92l5mRXY1FiPcOLYdLcf32bVIsHHflleFDjG/exec';
  for(const configured of ['',null,'https://example.com/dev','https://script.google.com/macros/s/id/dev']){
    const c=context({setting_:()=>configured});assert.equal(c.webAppUrl_(),production);
  }
  const valid='https://script.google.com/macros/s/custom_123/exec';
  assert.equal(context({setting_:()=>valid+'?old=1'}).webAppUrl_(),valid);
});
test('Admin login bypasses service terms and opens admin dashboard', () => {
  const calls=[],c={state:{},localStorage:{setItem(){}},isAdmin_:()=>true,openApp:async(a,b)=>calls.push(['openApp',a,b]),acceptServiceTerms_:async()=>calls.push(['terms'])};
  vm.createContext(c);vm.runInContext(html.split('\n').find(l=>l.includes('async function acceptAuth(')),c);
  return c.acceptAuth({token:'t',user:{role:'Admin'}}).then(()=>assert.deepEqual(calls,[['openApp',false,true]]));
});
test('Student login accepts service terms before opening home', () => {
  const calls=[],c={state:{},localStorage:{setItem(){}},isAdmin_:()=>false,openApp:async(a,b)=>calls.push(['openApp',a,b]),acceptServiceTerms_:async()=>calls.push(['terms'])};
  vm.createContext(c);vm.runInContext(html.split('\n').find(l=>l.includes('async function acceptAuth(')),c);
  return c.acceptAuth({token:'t',user:{role:'Student'}}).then(()=>assert.deepEqual(calls,[['terms'],['openApp',true,false]]));
});
test('Transient HTTP 0 failures retry only read-only RPCs', () => {
  assert(html.includes('attempts=safe?3:1'));
  for (const mutation of ['createBooking','cancelMyBooking','checkoutMyBooking','submitEvaluationV3','submitEvaluationV4','adminCompleteCheckout','adminQuickCheckIn']) {
    assert(!html.match(new RegExp("SAFE_RETRY_RPCS=new Set\\([^\\n]+['\"]"+mutation+"['\"]")));
  }
});
test('Three-digit check-in avoids camera, QR generation and request-sheet writes', () => {
  assert(html.includes('function checkInCode_(id)'));
  assert(html.includes("filter(b=>b.canAdminCheckIn&&checkInCode_(b.id)===code)"));
  assert(!html.includes('qrcodejs/1.0.0/qrcode.min.js'));
  assert(!html.includes('html5-qrcode@2.3.8/html5-qrcode.min.js'));
  assert(!source.includes("CheckInRequests:['RequestID'"));
  assert(source.includes('function adminQuickCheckIn(token,bid){return withMutationLock_'));
});
test('Redundant log sheets are not initialized or written', () => {
  for (const name of ['AttendanceLogs','Notifications','LoginAttempts','PenaltyLogs','SystemLogs']) {
    assert(!source.includes(`${name}:[`));
    assert(!source.includes(`getSheetByName('${name}')`));
  }
  assert(source.includes('function logNote_(){return false;}'));
  assert(source.includes('function systemLog_(){return false;}'));
});
test('RPC calls time out and technical network errors are converted to user messages', () => {
  assert(html.includes("rpcOnce_(fn,args,45000)"));
  assert(html.includes('const RPC_INFLIGHT=new Map()'));
  assert(html.includes('RPC_INFLIGHT.has(key)'));
  assert(html.includes('HTTP\\s*0'));
  assert(html.includes('กรุณาตรวจสอบสถานะรายการก่อนทำรายการซ้ำ'));
  assert(html.includes('function userMessage_('));
});
test('Telegram credentials are not embedded in source defaults', () => {
  assert(!/TELEGRAM_BOT_TOKEN','\d+:[A-Za-z0-9_-]{20,}/.test(source));
  assert(source.includes("props.getProperty('TELEGRAM_BOT_TOKEN')"));
  assert(source.includes("props.getProperty('TELEGRAM_CHAT_ID')"));
});
test('Existing booking conflicts explain the required next action', () => {
  assert(source.includes('คุณมีรอบที่จองไว้แล้ว จึงยังไม่สามารถจองรอบหรือหัตถการอื่นได้'));
  assert(source.includes('คุณมีรอบที่กำลังใช้งานอยู่ กรุณากด Check-out ในเมนูการจองของฉัน'));
  assert(source.includes('คุณ Check-out แล้ว แต่ยังไม่ได้ทำแบบประเมิน'));
});
test('Admin status management is role-protected and validated', () => {
  assert(source.includes("if(String(admin.Role)!=='Admin')throw Error('ไม่มีสิทธิ์')"));
  assert(source.includes("const allowed=['Reserved','Confirmed','CheckedIn','CheckedOut','Completed','Cancelled','NoShow','SessionCancelled']"));
  assert(source.includes("const auditNote=note||'ไม่ได้ระบุเหตุผล'"));
  assert(!source.includes("if(note.length<3)throw Error('กรุณาระบุเหตุผล"));
  assert(html.includes('หมายเหตุ <span class="muted">(ไม่บังคับ)</span>'));
  assert(html.includes('ค้นหาชื่อ รหัสนักศึกษา/บุคลากร อีเมล หรือรหัสจอง'));
  assert(html.includes('openAdminStatus'));
});
test('Admin operations view supports safe rapid check-in and incremental refresh', () => {
  assert(source.includes('function adminQuickCheckIn(token,bid){return withMutationLock_'));
  assert(source.includes('canAdminCheckIn:'));
  assert(html.includes('function scheduleAdminRefresh_()'));
  assert(html.includes('function replaceAdminBooking_(booking)'));
  assert(html.includes('กรอกรหัส 3 ตัวจากหน้าจอผู้เรียน'));
  assert(html.includes('.admin-status-cell .badge{padding:0;border:0;border-radius:0;background:transparent!important'));
});
test('User booking and cancellation emails use clear lifecycle messages', () => {
  assert(source.includes("const heading=cancel?'ยกเลิกการจองแล้ว':'ได้รับการจองเรียบร้อยแล้ว'"));
  assert(source.includes("const status=cancel?'ยกเลิกแล้ว':'ได้รับการจองแล้ว'"));
  assert(!source.includes('ขณะรับรายการมีผู้จอง'));
  assert(source.includes('การรับรายการจองนี้ยังไม่ใช่การยืนยันจัดรอบ'));
  assert(source.includes('รอบนี้ได้รับการยืนยันจัดรอบแล้ว'));
  assert(source.includes('function cancellationEmail_('));
  assert(source.includes('จำนวนผู้จองไม่ถึงขั้นต่ำภายในเวลาปิดรับจอง'));
  assert(source.includes("['Cancelled','SessionCancelled'].includes(status)"));
  assert(source.includes("'AdminSessionCancelled':'AdminBookingCancelled'"));
  assert(source.includes('การยกเลิกรอบครั้งนี้ไม่มีผลต่อประวัติ No-show'));
});
test('Authentication copy consistently supports students and personnel', () => {
  assert(html.includes('<label for="loginId">รหัสนักศึกษา/บุคลากร</label>'));
  assert(html.includes('placeholder="กรอกรหัสนักศึกษา/บุคลากร"'));
  assert(html.includes('<label for="regId">รหัสนักศึกษา/บุคลากร</label>'));
  assert(html.includes('รหัสนักศึกษา/บุคลากร หรืออีเมลที่ลงทะเบียนไว้'));
  assert(source.includes('รหัสนักศึกษา/บุคลากรหรือรหัสผ่านไม่ถูกต้อง'));
});
test('Service terms require reading and an accessible unchecked consent checkbox', () => {
  assert(html.includes('<label class="service-terms-consent" for="serviceTermsConsent">'));
  assert(html.includes('<input type="checkbox" id="serviceTermsConsent">'));
  assert(html.includes('ข้าพเจ้าได้อ่านและยอมรับกฎระเบียบและเงื่อนไขการเข้าใช้บริการ Self-Study ของศูนย์ RASME'));
  assert(html.includes('id="acceptServiceTerms" aria-describedby="serviceTermsHint" disabled'));
  assert(html.includes('consent.checked=false'));
  assert(html.includes('button.disabled=!(reachedEnd&&consent.checked)'));
  assert(html.includes('consent.onchange=update'));
  assert(html.includes('.service-terms-consent input:focus-visible'));
});
test('Booking email is plain text, complete and adapts to round confirmation', () => {
  const c=context({setting_:()=>15});
  const base={id:'BK-001',procedureName:'หัตถการทดสอบ',dateDisplay:'09/09/2569',timeDisplay:'09:15–14:00',room:'10-11',floor:'3',status:'Reserved',sessionStatus:'Open',bookedCount:1,minCapacity:2,cancelCutoff:'02/09/2569 23:59 น.'};
  const pending=c.bookingEmail_({FullName:'ผู้ทดสอบ'},base,false).body;
  assert(pending.includes('เรียน คุณ ผู้ทดสอบ'));
  for(const text of ['ข้อมูลการจอง','รหัสการจอง: BK-001','หัตถการ: หัตถการทดสอบ','วันที่: 09/09/2569','เวลา: 09:15–14:00','สถานที่: ห้อง 10-11 · ชั้น 3','สถานะ: ได้รับการจองแล้ว','รายการนี้มีผู้เข้าร่วมทั้งหมด 1 คน (รอบนี้รับขั้นต่ำ 2 คน)','การรับรายการจองนี้ยังไม่ใช่การยืนยันจัดรอบ','ยกเลิกการจองได้ถึง: 02/09/2569 23:59 น.','ติดต่อ: 02 839 5100','อีเมล: simulation.cnmi@gmail.com','<อีเมลนี้ส่งโดยระบบอัตโนมัติ>'])assert(pending.includes(text));
  assert(!/(?:undefined|null|&#x20;)/.test(pending));
  const confirmed=c.bookingEmail_({FullName:'ผู้ทดสอบ'},{...base,status:'Confirmed',sessionStatus:'Confirmed',bookedCount:2},false).body;
  assert(confirmed.includes('รอบนี้ได้รับการยืนยันจัดรอบแล้ว'));
  assert(!confirmed.includes('การรับรายการจองนี้ยังไม่ใช่การยืนยันจัดรอบ'));
});
test('Current user guidance and emails contain no QR or camera instructions', () => {
  const current=[html.split('\n').find(line=>line.includes("el('serviceTermsBody').innerHTML"))||'',html.split('\n').find(line=>line.includes('async function openProcedure'))||'',source.match(/function bookingEmail_\([\s\S]*?\n\}/)?.[0]||'',source.match(/function reminder_\([^\n]+/)?.[0]||''].join('\n');
  assert(!/(?:สแกน QR|แสดง QR|เปิดกล้อง|QR Code)/i.test(current));
  assert(!current.includes('taksinai.nou@mahidol.ac.th'));
  assert(!current.includes('Simulation.cnmi@gmail.com'));
});
test('Session cancellation email is complete plain text with safe fallbacks', () => {
  const c=context();
  const mail=c.cancellationEmail_({FullName:'ผู้ทดสอบ'},{id:'BK-001',procedureName:'หัตถการทดสอบ',dateDisplay:'09/09/2569',timeDisplay:'09:00–10:00',room:'10-11',floor:'3'},'จำนวนผู้จองไม่ถึงขั้นต่ำภายในเวลาปิดรับจอง');
  for(const text of ['เรียน คุณ ผู้ทดสอบ','แจ้งยกเลิกรอบฝึก','รหัสการจอง: BK-001','หัตถการ: หัตถการทดสอบ','วันที่: 09/09/2569','เวลา: 09:00–10:00','สถานที่: ห้อง 10-11 · ชั้น 3','สถานะ: ศูนย์ยกเลิกรอบ','สาเหตุการยกเลิก','จำนวนผู้จองไม่ถึงขั้นต่ำภายในเวลาปิดรับจอง','ไม่มีผลต่อประวัติ No-show','เลือกรอบอื่นและทำรายการจองใหม่','ติดต่อ: 02 839 5100','อีเมล: simulation.cnmi@gmail.com','<อีเมลนี้ส่งโดยระบบอัตโนมัติ>'])assert(mail.body.includes(text));
  assert(!/(?:undefined|null|<br|<div|QR|กล้อง)/i.test(mail.body));
  const fallback=c.cancellationEmail_({}, {id:'BK-002'}, '');
  assert(fallback.body.includes('กรุณาตรวจสอบสถานที่กับเจ้าหน้าที่ศูนย์ RASME'));
  assert(fallback.body.includes('รอบฝึกถูกยกเลิกโดยผู้ดูแลระบบ'));
});
test('Cancelled sessions reconcile every booking and isolate per-booking failures', () => {
  const master={
    B1:{_row:2,BookingID:'B1',UserID:'U1',SessionID:'S1',Status:'Reserved'},
    B2:{_row:3,BookingID:'B2',UserID:'U2',SessionID:'S1',Status:'Confirmed'}
  },deleted=[],notified=[];
  const sheets={ActiveBookings:'ActiveBookings',Bookings:'Bookings',Sessions:'Sessions',Users:'Users'};
  const ctx={bookings:[{...master.B1,_row:5},{...master.B2,_row:6}],sessions:[{SessionID:'S1'}],procedures:[],users:[{UserID:'U1',Email:'u1@example.test'},{UserID:'U2',Email:'u2@example.test'}],settings:{}};
  const c=context({SpreadsheetApp:{getActiveSpreadsheet:()=>({getSheetByName:name=>sheets[name]||name})},find_:(sheet,key,id)=>sheet==='Bookings'?master[id]:undefined,updateRow_:(sheet,row,changes)=>{const item=Object.values(master).find(value=>value._row===row);Object.assign(item,changes);return true;},bookingView_:booking=>({id:booking.BookingID}),notifyCancellation_:(user,details)=>{notified.push(details.id);if(details.id==='B1')throw Error('temporary queue failure');return true;},deleteRowNumbers_:(sheet,rows)=>{deleted.push(...rows);return rows.length;},invalidateBookingCaches_:()=>{}});
  const result=c.cancelSession_('S1',ctx);
  assert.equal(master.B1.Status,'SessionCancelled');
  assert.equal(master.B2.Status,'SessionCancelled');
  assert.deepEqual(notified,['B1','B2']);
  assert.deepEqual(deleted,[6]);
  assert.deepEqual(JSON.parse(JSON.stringify(result)),{sessionId:'S1',processed:1,skipped:0,failed:1,queued:1});
  assert(!JSON.stringify(master).includes('NoShowCount'));
  assert(!JSON.stringify(master).includes('EvaluationPending'));
});
test('Cancelled-session reconciliation retries unfinished notification without duplicate terminal writes', () => {
  const booking={_row:2,BookingID:'B1',UserID:'U1',SessionID:'S1',Status:'SessionCancelled'},deleted=[];
  const ctx={bookings:[{...booking,_row:5}],sessions:[{SessionID:'S1'}],procedures:[],users:[{UserID:'U1',Email:'u1@example.test'}],settings:{}};
  let writes=0,notifications=0;
  const c=context({SpreadsheetApp:{getActiveSpreadsheet:()=>({getSheetByName:name=>name})},find_:(sheet,key,id)=>sheet==='Bookings'?booking:undefined,updateRow_:()=>{writes++;},bookingView_:item=>({id:item.BookingID}),notifyCancellation_:()=>{notifications++;return true;},deleteRowNumbers_:(sheet,rows)=>{deleted.push(...rows);return rows.length;},invalidateBookingCaches_:()=>{}});
  const result=c.cancelSession_('S1',ctx);
  assert.equal(writes,0);
  assert.equal(notifications,1);
  assert.deepEqual(deleted,[5]);
  assert.equal(result.skipped,1);
  assert.equal(result.processed,1);
});
test('Batch row deletion is descending, grouped and invalidates once', () => {
  const calls=[];
  const c=context({invalidateSheetCaches_:()=>calls.push(['invalidate'])});
  const sheet={deleteRows:(start,count)=>calls.push([start,count])};
  assert.equal(c.deleteRowNumbers_(sheet,[3,4,7,7,1]),3);
  assert.deepEqual(calls,[[7,1],[3,2],['invalidate']]);
});
test('Cancellation automation skips no-show on cancelled sessions and resumes reconciliation', () => {
  const locked=source.match(/function processAutomationsLocked_\(\)\{[\s\S]*?\n\}/)?.[0]||'';
  assert(locked.includes("String(s.Status)==='Cancelled'"));
  assert(locked.includes("status==='Cancelled'&&pending"));
  assert(locked.includes('cancelSession_(s.SessionID,ctx)'));
  assert(locked.includes("if(!s||String(s.Status)==='Cancelled')return"));
  const cancelBlock=source.match(/function cancelSession_\([\s\S]*?\n\}/)?.[0]||'';
  assert(!cancelBlock.includes('NoShowCount'));
  assert(!cancelBlock.includes('EvaluationPending'));
});
test('Cancellation notifications are deterministic queued work and retriable outside the lock', () => {
  assert(source.includes("queueId='Q-'+hash_([channel,event,ref].join('|'))"));
  assert(source.includes("['Pending','Retry','Sent'].includes(String(existing.Status))"));
  assert(source.includes("Status:failed?'Failed':'Retry'"));
  assert(source.includes('Math.pow(2,Math.max(0,attempts-1))*30000'));
  assert(source.includes('processNotificationQueue_();const result=withMutationLock_'));
  const cancelBlock=source.match(/function cancelSession_\([\s\S]*?\n\}/)?.[0]||'';
  assert(!cancelBlock.includes('MailApp.sendEmail'));
});
test('Initial HTML response does not read Settings or any spreadsheet', () => {
  const block=source.match(/function doGet\([\s\S]*?\n/)?.[0]||'';
  assert(block.includes('WEB_APP_URL_FALLBACK'));
  assert(!/(?:webAppUrl_|setting_|SS\(|SpreadsheetApp)/.test(block));
});
test('Logged-out boot paints login before starting background bootstrap', () => {
  const boot=html.split('\n').find(line=>line.includes('async function boot()'))||'';
  assert(boot.includes("if(!state.token){showAuth('login');loadPublicBootstrap_('').catch(()=>{});return}"));
  assert(html.includes('function loadPublicBootstrap_(token)'));
  assert(html.includes('function ensurePublicData_()'));
});
test('Booking mutation reads the selected session and procedure, not both full sheets', () => {
  const block=source.match(/function createBooking\([\s\S]*?\n\}/)?.[0]||'';
  assert(block.includes("find_(SS().getSheetByName('Sessions'),'SessionID',sid)"));
  assert(block.includes("find_(SS().getSheetByName('Procedures'),'ProcedureID',session.ProcedureID)"));
  assert(!block.includes("rows_(SS().getSheetByName('Sessions'))"));
  assert(!block.includes("rows_(SS().getSheetByName('Procedures'))"));
});
test('Mutation response view does not scan all active bookings', () => {
  const block=source.match(/function mutationBookingView_\([\s\S]*?\n/)?.[0]||'';
  assert(block.includes('bookings:[booking]'));
  assert(!block.includes("rows_(ss.getSheetByName('ActiveBookings'))"));
  assert(source.includes('booking:mutationBookingView_(updated,session)'));
});
test('Admin mutations update local state without a second dashboard RPC', () => {
  const save=html.split('\n').find(line=>line.includes('async function saveAdminStatus'))||'';
  const checkout=html.split('\n').find(line=>line.includes('async function adminCheckout'))||'';
  assert(save.includes('replaceAdminBooking_'));
  assert(checkout.includes('replaceAdminBooking_'));
  assert(!save.includes('refreshAdminData_'));
  assert(!checkout.includes('refreshAdminData_'));
});
test('Admin cancellation queues notification only after releasing the mutation lock', () => {
  const wrapper=source.match(/function adminUpdateBookingStatus\([\s\S]*?\n/)?.[0]||'';
  const locked=source.match(/function adminUpdateBookingStatusLocked_\([\s\S]*?\n\}/)?.[0]||'';
  assert(wrapper.indexOf('withMutationLock_')<wrapper.indexOf('notifyCancellation_'));
  assert(!locked.includes('notifyCancellation_('));
  assert(locked.includes('result._notification='));
});
test('Optional browser performance diagnostics record RPC duration and response size without arguments', () => {
  assert(html.includes("localStorage.getItem('rasme_perf')==='1'"));
  assert(html.includes("console.info('[RASME perf]'"));
  assert(html.includes('durationMs:Math.round(performance.now()-started)'));
  assert(html.includes('responseBytes:bytes'));
  const report=html.split('\n').find(line=>line.includes('function rpcOnce_'))||'';
  assert(!report.includes('args:args'));
});
test('Telegram retry uses corrected current chat configuration and records delivery', () => {
  const updates=[],payloads=[],queue={_row:2,QueueID:'Q1',Channel:'Telegram',Recipient:'1.2345E+12',Subject:'',Message:'ทดสอบ',Event:'BookingCreated',ReferenceID:'B1',Status:'Retry',Attempts:1,NextAttemptAt:new Date(Date.now()+3600000)};
  const c=context({SpreadsheetApp:{getActiveSpreadsheet:()=>({getSheetByName:()=>({})})},setting_:(key,fallback)=>key==='NOTIFICATION_MAX_RETRIES'?3:fallback,sheetHeaders_:()=>[],rows_:()=>[queue],telegramConfig_:()=>({token:'123456:abcdefghijklmnopqrstuvwxyz',chat:'-1001234567890',tokenValid:true,chatValid:true}),UrlFetchApp:{fetch:(url,options)=>{payloads.push(JSON.parse(options.payload));return{getResponseCode:()=>200,getContentText:()=>'{"ok":true}'}}},MailApp:{},updateRow_:(sheet,row,changes)=>updates.push(changes),logNote_:()=>{},systemLog_:()=>{}});
  c.processNotificationQueue_();
  assert.equal(payloads[0].chat_id,'-1001234567890');
  assert.equal(updates[0].Status,'Sent');
  assert.equal(updates[0].Error,'');
});
test('A previously failed Telegram item receives exactly one recovery attempt', () => {
  const updates=[],queue={_row:2,Channel:'Telegram',Recipient:'old',Subject:'',Message:'ทดสอบ',Event:'BookingCreated',ReferenceID:'B1',Status:'Failed',Attempts:3};
  const c=context({SpreadsheetApp:{getActiveSpreadsheet:()=>({getSheetByName:()=>({})})},setting_:(key,fallback)=>key==='NOTIFICATION_MAX_RETRIES'?3:fallback,sheetHeaders_:()=>[],rows_:()=>[queue],telegramConfig_:()=>({token:'123456:abcdefghijklmnopqrstuvwxyz',chat:'-1001234567890',tokenValid:true,chatValid:true}),UrlFetchApp:{fetch:()=>({getResponseCode:()=>200,getContentText:()=>'{"ok":true}'})},MailApp:{},updateRow_:(sheet,row,changes)=>updates.push(changes),logNote_:()=>{},systemLog_:()=>{}});
  c.processNotificationQueue_();
  assert.equal(updates.length,1);
  assert.equal(updates[0].Status,'Sent');
  queue.Attempts=4;updates.length=0;c.processNotificationQueue_();
  assert.equal(updates.length,0);
});
test('Telegram settings normalize text cells and validate token and chat formats', () => {
  const c=context();
  assert.equal(c.telegramValue_("  '-1001234567890  "),'-1001234567890');
  assert(source.includes("tokenValid:/^\\d+:[A-Za-z0-9_-]{20,}$/"));
  assert(source.includes("รูปแบบ TELEGRAM_CHAT_ID ไม่ถูกต้อง"));
  assert(source.includes("JSON.parse(res.getContentText()||'{}').description"));
});
test('Manifest declares every OAuth scope required by Telegram notification processing', () => {
  const scopes=new Set(manifest.oauthScopes||[]);
  for(const scope of ['https://www.googleapis.com/auth/spreadsheets.currentonly','https://www.googleapis.com/auth/script.external_request','https://www.googleapis.com/auth/script.send_mail','https://www.googleapis.com/auth/script.scriptapp'])assert(scopes.has(scope));
});
test('Thai Apps Script errors keep their guidance without technical prefixes', () => {
  assert(html.includes("replace(/^(?:(?:Error|Exception|TypeError|ReferenceError):\\s*)+"));
  assert(html.includes('hasThai=/[ก-๙]/.test(message)'));
});
setImmediate(()=>console.log(`${passed} local checks passed; live UAT still required.`));

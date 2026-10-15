'use strict';
const STAFF=[
 {id:'quynh',name:'Quỳnh',full:'Nguyễn Thị Quỳnh',role:'KTV trị liệu',col:2,aliases:['quỳnh','nguyễn thị quỳnh']},
 {id:'kien',name:'Kiên',full:'Nguyễn Văn Kiên',role:'KTV trị liệu',col:6,aliases:['kiên','nguyễn văn kiên']},
 {id:'tinh',name:'Tính',full:'Chu Hồng Tính',role:'Massage / Sale',col:10,aliases:['tính','chu hồng tính']},
 {id:'thuong',name:'Thương',full:'Thương',role:'KTV trị liệu',col:14,aliases:['thương','nguyễn thị việt thương']}
];
const KEY='ktd_staff_report_local_v1';
const $=s=>document.getElementById(s);
const vn=(v,max=1)=>new Intl.NumberFormat('vi-VN',{maximumFractionDigits:max}).format(Number(v)||0);
const money=v=>vn(Math.round(v),0)+' đ';
const esc=v=>String(v??'').replace(/[&<>"']/g,c=>({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&#39;'}[c]));
const norm=v=>String(v||'').normalize('NFD').replace(/[\u0300-\u036f]/g,'').replace(/[đĐ]/g,'d').trim().toLowerCase();
let report=null,active='all';
function dateKey(v){
 let d='';
 if(typeof v==='number'&&v>25000&&window.XLSX){
  const p=XLSX.SSF.parse_date_code(v);
  if(p)d=[String(p.d).padStart(2,'0'),String(p.m).padStart(2,'0'),p.y].join('/');
 }else if(typeof v==='string'){
  const m=v.match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})/);
  if(m)d=[m[1].padStart(2,'0'),m[2].padStart(2,'0'),m[3]].join('/');
 }
 return d;
}
function inPeriod(d){
 const p=String(d||'').split('/');
 if(p.length!==3)return false;
 const n=Number(p[2])*10000+Number(p[1])*100+Number(p[0]);
 return n>=20260915&&n<=20261014;
}
function timeValue(v){
 if(typeof v==='number'&&v>=0&&v<1)return v*24;
 if(v instanceof Date)return v.getUTCHours()+v.getUTCMinutes()/60;
 const m=String(v||'').trim().match(/^(\d{1,2}):(\d{2})$/);
 return m?Number(m[1])+Number(m[2])/60:null;
}
function timeLabel(v){
 const h=timeValue(v);
 if(h===null)return String(v||'—').trim()||'—';
 const min=Math.round(h*60);
 return String(Math.floor(min/60)).padStart(2,'0')+':'+String(min%60).padStart(2,'0');
}
function dataSheet(wb,name){
 const sheet=wb.Sheets[name];
 return sheet?XLSX.utils.sheet_to_json(sheet,{header:1,defval:null,raw:true}):null;
}
function readAttendance(wb){
 const rows=dataSheet(wb,'KẾT QUẢ 15.09-14.10');
 if(!rows)throw Error('Không tìm thấy tab KẾT QUẢ 15.09-14.10 trong CHẤM CÔNG KTD.xlsx');
 const base=Number(rows[2]?.[2])||0;
 const lunch=Math.max(0,Number(rows[1]?.[2])||0);
 const out={};
 STAFF.forEach(s=>{
  const title=String(rows[4]?.[s.col]||'');
  const found=title.match(/(\d{4,7})\s*(?:vnđ|vnd|đ)\s*\/\s*giờ/i);
  const rate=found?Number(found[1]):base;
  out[s.id]={...s,rate:rate,days:0,hours:0,after:0,late:0,pay:0,logs:[],cases:[],temporary:0,approved:0,albumin:0};
 });
 rows.slice(6,36).forEach(row=>{
  const date=dateKey(row[0]);
  if(!inPeriod(date))return;
  STAFF.forEach(s=>{
   const o=out[s.id],a=row[s.col],b=row[s.col+1];
   const arrival=timeValue(a),departure=timeValue(b);
   let within=0,extra=0,note='';
   if(arrival!==null&&departure!==null&&departure>arrival){
    within=Math.max(0,Math.min(departure,18)-Math.max(arrival,9)-lunch);
    extra=Math.max(0,departure-Math.max(arrival,18));
    o.days++;o.hours+=within;o.after+=extra;
    if(arrival>9)o.late++;
   }else note=[a,b].filter(v=>typeof v==='string').join(' / ')||'Không có giờ hợp lệ';
   o.logs.push({date:date,arrival:timeLabel(a),departure:timeLabel(b),within:Number(within.toFixed(2)),extra:Number(extra.toFixed(2)),note:note});
  });
 });
 STAFF.forEach(s=>{
  const o=out[s.id];
  o.hours=Number(o.hours.toFixed(2));
  o.after=Number(o.after.toFixed(2));
  o.pay=Math.round(o.hours*o.rate);
 });
 return out;
}
function findPerson(name){
 const a=norm(name);
 const p=STAFF.find(x=>x.aliases.some(n=>norm(n)===a));
 return p?p.id:null;
}
function readWarehouse(wb,people){
 const rows=dataSheet(wb,'13 - CĂN CỨ KTV');
 if(!rows)throw Error('Không tìm thấy tab 13 - CĂN CỨ KTV trong KIỂM KHO.xlsx');
 const seen=new Set();
 let unassigned=0,excluded=0;
 rows.slice(4).forEach(row=>{
  const day=dateKey(row[0]);
  if(!inPeriod(day))return;
  const status=String(row[11]||''),description=String(row[5]||'');
  if(/loại trừ|xuất gộp|điều chuyển/i.test(status+' '+description)){excluded++;return}
  const id=findPerson(row[10]||row[3]);
  if(!id){unassigned++;return}
  const customer=String(row[2]||row[6]||'Không xác định');
  const key=String(row[12]||day+'|'+customer+'|'+id);
  if(seen.has(key))return;
  seen.add(key);
  const approval=String(row[9]||'');
  const approved=/đủ chứng từ|đã duyệt/i.test(approval)||/ca đã duyệt/i.test(status);
  const qty=Math.max(0,Number(row[4])||0);
  const o=people[id];
  o.cases.push({date:day,customer:customer,qty:qty,approved:approved,source:String(row[8]||'13 - CĂN CỨ KTV')});
  if(approved)o.approved++;else o.temporary++;
  o.albumin+=qty;
 });
 STAFF.forEach(s=>people[s.id].cases.sort((a,b)=>b.date.split('/').reverse().join('').localeCompare(a.date.split('/').reverse().join(''))));
 return {unassigned:unassigned,excluded:excluded};
}
function stat(label,value,subtitle){
 return '<div class="stat"><div class="stat-label">'+esc(label)+'</div><div class="stat-number">'+esc(value)+'</div><div class="stat-hint">'+esc(subtitle)+'</div></div>';
}
function setupNavigation(){
 const items=[{id:'all',name:'Tổng quan'}].concat(STAFF);
 $('staff-nav').innerHTML=items.map(p=>'<button class="nav-item '+(active===p.id?'active':'')+'" data-person="'+p.id+'"><span class="nav-circle">'+(p.id==='all'?'◈':esc(p.name[0]))+'</span><strong>'+esc(p.name)+'</strong></button>').join('');
 document.querySelectorAll('[data-person]').forEach(el=>el.addEventListener('click',()=>select(el.dataset.person)));
}
function chart(target,field,label,gold){
 const max=Math.max(1,...STAFF.map(s=>report.people[s.id][field]));
 $(target).innerHTML=STAFF.map(s=>{
  const p=report.people[s.id],v=p[field],width=Math.min(100,Math.max(0,v/max*100));
  return '<div class="bar-row"><span class="bar-name">'+esc(s.name)+'</span><div class="bar-track"><div class="bar-fill '+(gold?'gold':'')+'" style="width:'+width+'%"></div></div><span class="bar-number">'+vn(v)+' '+label+'</span></div>';
 }).join('');
}
function overview(){
 const list=STAFF.map(s=>report.people[s.id]);
 const sum=field=>list.reduce((acc,p)=>acc+p[field],0);
 $('stats').innerHTML=stat('Nhân sự','04','Có dữ liệu chấm công')+stat('Giờ trong ca',vn(sum('hours'))+' h','Khung 09:00–18:00')+stat('Giờ sau 18h',vn(sum('after'))+' h','Chưa duyệt tăng ca')+stat('Ca từ sổ Albumin',vn(sum('temporary')+sum('approved'),0),'Chỉ là một phần hoạt động');
 chart('hours-bars','hours','giờ',false);
 chart('cases-bars','temporary','ca',true);
 $('staff-summary').innerHTML='<table><thead><tr><th>Nhân sự</th><th class="num">Ngày công</th><th class="num">Trong ca</th><th class="num">Sau 18h</th><th class="num">Ca tạm tính</th><th class="num">Ca đã duyệt</th><th class="num">Thu nhập giờ (ước tính)</th></tr></thead><tbody>'+list.map(p=>'<tr><td class="clickable" data-goto="'+p.id+'">'+esc(p.name)+' →</td><td class="num">'+p.days+'</td><td class="num">'+vn(p.hours)+' h</td><td class="num">'+vn(p.after)+' h</td><td class="num">'+p.temporary+'</td><td class="num">'+p.approved+'</td><td class="num">'+money(p.pay)+'</td></tr>').join('')+'</tbody></table>';
 document.querySelectorAll('[data-goto]').forEach(el=>el.addEventListener('click',()=>select(el.dataset.goto)));
}
function detail(id){
 const p=report.people[id];
 const days=p.logs.map(x=>'<tr><td>'+esc(x.date)+'</td><td>'+esc(x.arrival)+'</td><td>'+esc(x.departure)+'</td><td class="num">'+vn(x.within)+'</td><td class="num">'+vn(x.extra)+'</td><td>'+esc(x.note)+'</td></tr>').join('');
 const cases=p.cases.map(x=>'<tr><td>'+esc(x.date)+'</td><td>'+esc(x.customer)+'</td><td class="num">'+vn(x.qty)+'</td><td><span class="tag '+(x.approved?'green':'')+'">'+(x.approved?'Đã duyệt':'Tạm tính')+'</span></td><td class="small">'+esc(x.source)+'</td></tr>').join('')||'<tr><td colspan="5" class="small">Không ghi nhận ca đứng tên nhân sự trong nguồn này. Không có nghĩa là không làm việc.</td></tr>';
 const cells=[['Ngày có công',p.days],['Giờ trong ca',vn(p.hours)+' h'],['Sau 18h',vn(p.after)+' h'],['Ca tạm tính',p.temporary],['Ca đã duyệt',p.approved],['Thu nhập giờ (ước tính)',money(p.pay)]].map(x=>'<div class="minor">'+esc(x[0])+'<strong>'+esc(x[1])+'</strong></div>').join('');
 $('staff-detail').innerHTML='<div class="detail-head"><div><div class="eyebrow">BÁO CÁO CÁ NHÂN</div><h2>'+esc(p.full)+' · '+esc(p.role)+'</h2><p class="panel-sub">15/09–14/10/2026</p></div><span class="tag">Ca chưa duyệt không tính lương</span></div><div class="minor-grid">'+cells+'</div><p class="small" style="margin-top:15px">Đơn giá giờ: '+money(p.rate)+'/giờ. Chưa gồm làm thêm, hoa hồng, thưởng, phụ cấp, khấu trừ. Ca trong kho chỉ dựa trên sổ Albumin.</p><div class="section-title">Chi tiết chấm công</div><div class="table-wrap"><table><thead><tr><th>Ngày</th><th>Đến</th><th>Về</th><th class="num">Trong ca (h)</th><th class="num">Sau 18h (h)</th><th>Ghi chú</th></tr></thead><tbody>'+days+'</tbody></table></div><div class="section-title">Căn cứ ca từ sổ Albumin ('+p.cases.length+' mục)</div><p class="small">Xuất sản phẩm không tự động là ca hợp lệ để tính lương.</p><div class="table-wrap"><table><thead><tr><th>Ngày</th><th>Khách hàng</th><th class="num">Albumin (lọ)</th><th>Trạng thái</th><th>Dòng nguồn</th></tr></thead><tbody>'+cases+'</tbody></table></div>';
 $('staff-detail').classList.remove('hidden');
}
function select(id){
 active=id;
 setupNavigation();
 $('page-title').textContent=id==='all'?'Tổng quan nhân sự':'Báo cáo · '+report?.people[id]?.name;
 $('page-subtitle').textContent=id==='all'?'Tổng hợp 15/09 – 14/10/2026':report?.people[id]?.role+' · 15/09 – 14/10/2026';
 if(!report)return;
 document.querySelectorAll('#dashboard>.stat-grid,#dashboard>.two-panel,#dashboard>.panel:not(#staff-detail)').forEach(el=>el.classList.toggle('hidden',id!=='all'));
 if(id==='all')$('staff-detail').classList.add('hidden');else detail(id);
 window.scrollTo({top:0,behavior:'smooth'});
}
function render(){
 if(!report)return;
 $('dashboard').classList.remove('hidden');
 $('setup').classList.add('hidden');
 $('open-import').textContent='Nhập lại Excel';
 $('source-note').innerHTML='<strong>Nguồn:</strong> CHẤM CÔNG KTD · KẾT QUẢ 15.09-14.10 và KIỂM KHO · 13 - CĂN CỨ KTV. <strong>Ca tạm tính không phải ca được duyệt hoặc được trả lương.</strong> Dữ liệu tổng hợp chỉ lưu trên trình duyệt này.'+(report.extra.unassigned?' Có '+report.extra.unassigned+' dòng kho chưa phân được nhân sự.':'');
 overview();select(active);
}
async function loadExcel(file){
 if(!window.XLSX)throw Error('Chưa tải được thư viện đọc Excel. Kiểm tra kết nối mạng.');
 return XLSX.read(await file.arrayBuffer(),{type:'array',cellDates:false});
}
async function importFiles(){
 const a=$('attendance-file').files[0],b=$('warehouse-file').files[0];
 if(!a||!b){$('import-status').textContent='Vui lòng chọn đủ hai file Excel.';return}
 $('process').disabled=true;$('process').textContent='Đang tổng hợp…';
 try{
  const w1=await loadExcel(a),w2=await loadExcel(b);
  const people=readAttendance(w1),extra=readWarehouse(w2,people);
  report={people:people,extra:extra,updatedAt:new Date().toISOString()};
  try{localStorage.setItem(KEY,JSON.stringify(report))}catch(e){$('import-status').textContent='Không thể lưu trên thiết bị; dữ liệu chỉ có trong phiên này.'}
  render();
 }catch(err){$('import-status').textContent=String(err.message||err)}
 finally{$('process').disabled=false;$('process').textContent='Tạo báo cáo'}
}
function showImport(){$('setup').classList.remove('hidden');$('dashboard').classList.add('hidden');$('import-status').textContent=''}
$('attendance-file').addEventListener('change',e=>$('attendance-name').textContent=e.target.files[0]?.name||'Chọn file chấm công');
$('warehouse-file').addEventListener('change',e=>$('warehouse-name').textContent=e.target.files[0]?.name||'Chọn file kiểm kho');
$('process').addEventListener('click',importFiles);
$('open-import').addEventListener('click',showImport);
$('print').addEventListener('click',()=>window.print());
$('clear-data').addEventListener('click',()=>{
 if(!confirm('Xóa dữ liệu báo cáo đã lưu trên trình duyệt này?'))return;
 localStorage.removeItem(KEY);report=null;active='all';setupNavigation();showImport();
});
try{const saved=localStorage.getItem(KEY);if(saved){report=JSON.parse(saved);render()}}catch(e){localStorage.removeItem(KEY)}
setupNavigation();
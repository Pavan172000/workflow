/* ===========================
 BIS Workflow Management – v2.4 (UI-polish + icon actions)
 - BYYMMNNNN IDs; NRA freeze (workflowName + receivedDateTime + deadline)
 - Actions icon-only; text wrapping + modal; XLSX/CSV export
=========================== */
(function(){ if(!Array.from){ Array.from=function(a,f,t){ var r=[].slice.call(a); return f? r.map(f,t): r; }; }})();
var STORAGE_KEY='bisWorkflowData', EDIT_ID_KEY='bisWorkflowEditId', SEQ_KEY='bisWorkflowSeq';
var RESEARCHERS=['All','New','Aishwarya','Aditya','Arun','Bijesh','Charan','Deepak','David','Hatim','Jayasurya','Megha','Mellissa','Pavan','Pratik','Priya','Subheesha','Sanjana','Shaneza','Udit','Vikas','Vaishnavi','Vivek'];
var RESEARCHERS_NO_ALL=RESEARCHERS.filter(function(n){return n!=='All';});
var DATABASE_SOURCES=['Beauhurst','Bloomberg','Broker','S&P',"Moody's",'Fitch','MSCI','Capital IQ','Orbis','FactSet','Factiva','Dealogic DCM','Dealogic Loans','Dealogic M&A','Dealogic MTN','Dealogic Strategy','Internet','IBISworld'];

function loadData(){ try{var raw=localStorage.getItem(STORAGE_KEY);var arr=raw?JSON.parse(raw):[];return Array.isArray(arr)?arr:[];}catch(e){localStorage.removeItem(STORAGE_KEY);return [];} }
function saveData(a){ localStorage.setItem(STORAGE_KEY, JSON.stringify(a||[])); }

function parseSeqFromId(id,prefix){ if(typeof id!=='string'||!id.startsWith(prefix))return null; var tail=id.slice(prefix.length); return /^\d{4}$/.test(tail)?parseInt(tail,10):null; }
function nextSequence(yyMM){ var m={}; try{m=JSON.parse(localStorage.getItem(SEQ_KEY)||'{}');}catch(e){m={};} var stored=m[yyMM]||0; var max=0;
  loadData().forEach(function(it){ var s=parseSeqFromId(it.id||'','B'+yyMM); if(s&&s>max) max=s; });
  var nxt=Math.max(stored,max)+1; m[yyMM]=nxt; localStorage.setItem(SEQ_KEY, JSON.stringify(m)); return nxt;
}
function generateUniqueId(){ var d=new Date(), yy=(''+d.getFullYear()).slice(-2), mm=(''+(d.getMonth()+1)).padStart(2,'0'); var yyMM=yy+mm; var n=nextSequence(yyMM); return 'B'+yyMM+(''+n).padStart(4,'0'); }

var REQUEST_TYPES={"1":["Biography","Broker Research","Company Information","Datasets","Economic Information","News Run + Article Search","Ratings","ESG Information","IBIS Report + Other Direct Report"],
"2":["Briefing Memo","Company Screening","Corporate Docs","Datasets","Economic Information","ESG information","InfoPack","News Run + Article Search","TravelBook"],
"3":["Industry Research","News Run (Industry)","Deals/Issues","Datasets","Company Profiles","ESG Information","Projects"],
"Alerts":["Alerts","Newsletters","Real Time"],
"NRA":["Meeting","Intranet Mngt","Evaluation","Training / Education","Workflow Mngt"]};

function updateRequestType(){
  var cx=document.getElementById('complexity'); var rt=document.getElementById('requestType'); if(!cx||!rt)return;
  if(!cx.value) cx.value='1';
  var opts=REQUEST_TYPES[cx.value]||[];
  rt.innerHTML='';
  opts.forEach(function(o){ var op=document.createElement('option'); op.value=o; op.textContent=o; rt.appendChild(op); });
  if(rt.dataset.prefillValue){
    var v=rt.dataset.prefillValue;
    var match=Array.prototype.some.call(rt.options,function(x){return x.value===v;});
    if(match) rt.value=v;
    delete rt.dataset.prefillValue;
  }
  applyWorkflowNameFreeze();
}

function applyWorkflowNameFreeze(){
  var cx=document.getElementById('complexity');
  var wf=document.getElementById('workflowName');
  var rdt=document.getElementById('receivedDateTime');
  var ddl=document.getElementById('deadline');
  if(!cx) return;
  var isNRA=(cx.value==='NRA');
  function setFrozen(el,freeze){
    if(!el) return;
    if(freeze){ el.readOnly=true; el.classList.add('readonly'); el.removeAttribute('required');
      if(el.tagName==='INPUT'&&el.type==='text'&&!el.placeholder){ el.placeholder='Not required for NRA'; }
    } else { el.readOnly=false; el.classList.remove('readonly');
      if(['workflowName','receivedDateTime','deadline'].includes(el.id)) el.setAttribute('required','');
      if(el.placeholder==='Not required for NRA') el.placeholder='';
    }
  }
  setFrozen(wf,isNRA); setFrozen(rdt,isNRA); setFrozen(ddl,isNRA);
}

var BASE_COLUMNS=['id','workflowName','clientEmail','receivedDateTime','researcher','status'];
var OPTIONAL_COLUMNS=['requestDetails','deadline','complexity','requestType','foa','noOfCompanies','timeStarted','timeEnded','timeSpent'];
var DB_COLUMNS=Array.from({length:10},function(_,i){return 'database'+(i+1)});
var PAGES_COLUMNS=Array.from({length:10},function(_,i){return 'pages'+(i+1)});

function computeVisibleColumns(items){
  var v=BASE_COLUMNS.slice();
  OPTIONAL_COLUMNS.forEach(function(k){ if(items.some(function(x){ return (x[k]||'')!==''; })) v.push(k); });
  for(var i=1;i<=10;i++){
    var hasDB=items.some(function(x){return (x['database'+i]||'')!=='';});
    var hasPG=items.some(function(x){return (x['pages'+i]||'')!=='';});
    if(hasDB) v.push('database'+i);
    if(hasPG) v.push('pages'+i);
  }
  return v;
}
function titleFor(key){
  var m={id:'Request ID',workflowName:'Workflow Name',clientEmail:'Client Email',receivedDateTime:'Received Date/Time',researcher:'Researcher',status:'Status',
    requestDetails:'Request Details',deadline:'Deadline',complexity:'Complexity',requestType:'Request Type',foa:'FOA',noOfCompanies:'No. of Companies',
    timeStarted:'Time Started',timeEnded:'Time Ended',timeSpent:'Time Spent'};
  if(m[key]) return m[key];
  if(key.indexOf('database')===0) return key.replace('database','Database ');
  if(key.indexOf('pages')===0) return key.replace('pages','Pages ');
  return key;
}

function makeTruncatedCell(text,lines){
  var el=document.createElement('div'); el.className=lines>1?'truncate-3':'truncate-1'; el.title=text||''; el.textContent=text||''; return el;
}

function openModal(title,content){
  var m=document.getElementById('modal'),b=document.getElementById('modalBody'),t=document.getElementById('modalTitle');
  if(!m||!b||!t) return; t.textContent=title||'Full Content'; b.textContent=content||''; m.style.display='block'; m.setAttribute('aria-hidden','false');
}
function closeModal(){ var m=document.getElementById('modal'); if(m){ m.style.display='none'; m.setAttribute('aria-hidden','true'); } }

/** Render table with optional actions (supports icon-only buttons) */
function renderTable(sel,items,opts){
  opts=opts||{};
  var actions=opts.actions||[];
  var table=document.querySelector(sel.table),thead=document.querySelector(sel.thead),tbody=document.querySelector(sel.tbody);
  if(!table||!thead||!tbody) return;
  var cols=computeVisibleColumns(items);

  // Head
  thead.innerHTML='';
  var trh=document.createElement('tr');
  cols.forEach(function(c){ var th=document.createElement('th'); th.textContent=titleFor(c); trh.appendChild(th); });
  if(actions.length){
    actions.forEach(function(a,idx){
      var th=document.createElement('th');
      th.className='nowrap actions-col';
      th.textContent = a.header || 'Action';
      trh.appendChild(th);
    });
  }
  thead.appendChild(trh);

  // Body
  tbody.innerHTML='';
  items.forEach(function(item){
    var tr=document.createElement('tr');
    cols.forEach(function(c){
      var td=document.createElement('td'); td.dataset.col=c;
      if(c==='status' && opts.statusAsDropdown){
        var selEl=document.createElement('select');
        ['In progress','Completed','Yet to Assign','On Hold'].forEach(function(v){ var op=document.createElement('option'); op.value=v; op.textContent=v; selEl.appendChild(op); });
        selEl.value=item.status||'In progress';
        selEl.addEventListener('change', function(){ onStatusChange(item.id, selEl.value, opts.onStatusChanged); });
        td.appendChild(selEl);
      } else if (c==='status' && !opts.statusAsDropdown) {
        // Show badge in read-only tables (e.g., Completed)
        var s=(item[c]||'').toLowerCase();
        var badge=document.createElement('span');
        badge.className='status-badge ' + (
          s==='in progress' ? 'status-in-progress' :
          s==='completed'   ? 'status-completed'   :
          s==='yet to assign'? 'status-yet-to-assign' :
                              'status-on-hold'
        );
        badge.textContent=item[c]||'';
        td.appendChild(badge);
      } else if(c==='requestDetails'){
        td.appendChild(makeTruncatedCell(item[c],3));
        var view=document.createElement('button');
        view.type='button'; view.className='linklike table-btn'; view.textContent='View';
        view.addEventListener('click', function(){ openModal('Request Details', item[c]||''); });
        td.appendChild(view);
      } else if(['clientEmail','workflowName','requestType'].indexOf(c)>=0){
        td.appendChild(makeTruncatedCell(item[c],1));
      } else {
        td.textContent=item[c]||''; td.title=item[c]||'';
      }
      tr.appendChild(td);
    });

    // Actions
    if(actions.length){
      actions.forEach(function(a){
        var td=document.createElement('td'); td.className='nowrap actions-cell';
        var btn=document.createElement('button');
        btn.type='button';
        btn.className = a.className ? ('table-btn ' + a.className) : 'table-btn';
        // accessibility labels for icon-only
        var labelText = a.title || a.label || 'Action';
        btn.setAttribute('aria-label', labelText);
        btn.title = labelText;
        // If icon-only, keep text empty; else show label
        if(!a.className || a.className.indexOf('btn-icon')===-1){
          btn.textContent = a.label || 'Action';
        } else {
          btn.textContent = ''; // icon-only
        }
        btn.addEventListener('click', function(){ a.handler(item.id); });
        td.appendChild(btn); tr.appendChild(td);
      });
    }
    tbody.appendChild(tr);
  });
  return cols;
}

function onStatusChange(id,newStatus,rerenderCb){
  var all=loadData(); var idx=all.findIndex(function(x){return x.id===id;}); if(idx<0) return;
  all[idx].status=newStatus;
  if(newStatus==='Completed' && !all[idx].timeSpent && all[idx].timeStarted && all[idx].timeEnded){
    all[idx].timeSpent=computeTimeSpent(all[idx].timeStarted,all[idx].timeEnded);
  }
  saveData(all);
  if(typeof rerenderCb==='function') rerenderCb();
}

function loadDashboard(){
  var data=loadData().filter(function(x){return (x.status||'')!=='Completed';});
  renderTable(
    {table:'#dashboardTable', thead:'#dashboardTable thead', tbody:'#dashboardBody'},
    data,
    {
      statusAsDropdown:true,
      onStatusChanged:function(){ filterDashboard(); },
      actions:[
        { header:'Action', label:'Edit', title:'Edit', className:'btn-icon btn-edit', handler:editItem }
      ]
    }
  );
}
function filterDashboard(){
  var sEl=document.getElementById('statusFilter'), rEl=document.getElementById('researcherFilter');
  var s=sEl?sEl.value:'All', r=rEl?rEl.value:'All';
  var data=loadData().filter(function(x){return (x.status||'')!=='Completed';});
  data=data.filter(function(it){
    var sOk=s==='All'|| it.status===s;
    var rOk=r==='All'|| it.researcher===r;
    return sOk&&rOk;
  });
  renderTable(
    {table:'#dashboardTable', thead:'#dashboardTable thead', tbody:'#dashboardBody'},
    data,
    {
      statusAsDropdown:true,
      onStatusChanged:function(){ filterDashboard(); },
      actions:[
        { header:'Action', label:'Edit', title:'Edit', className:'btn-icon btn-edit', handler:editItem }
      ]
    }
  );
}
function editItem(id){ localStorage.setItem(EDIT_ID_KEY, id); location.href='workflow_form.html'; }

function loadCompleted(){
  var data=loadData().filter(function(x){return x.status==='Completed';});
  renderTable(
    {table:'#completedTable', thead:'#completedTable thead', tbody:'#completedBody'},
    data,
    {
      // show status as read-only badges
      actions:[
        { header:'Action', label:'Delete', title:'Delete', className:'btn-icon btn-delete', handler:deleteItem }
      ]
    }
  );
}
function filterCompleted(){
  var cEl=document.getElementById('complexityFilter'), rEl=document.getElementById('completedResearcherFilter');
  var startRaw=(document.getElementById('dateRangeStart')||{}).value, endRaw=(document.getElementById('dateRangeEnd')||{}).value;
  var c=cEl?cEl.value:'All', r=rEl?rEl.value:'All';
  var start=startRaw?new Date(startRaw):null, end=endRaw?new Date(endRaw):null; if(end) end.setHours(23,59,59,999);
  var data=loadData().filter(function(x){return x.status==='Completed';});
  data=data.filter(function(it){
    var cmpOk=c==='All'|| it.complexity===c;
    var rOk=r==='All'|| it.researcher===r;
    var dt=it.receivedDateTime?new Date(it.receivedDateTime):null;
    var sOk=!start || (dt&&dt>=start);
    var eOk=!end || (dt&&dt<=end);
    return cmpOk&&rOk&&sOk&&eOk;
  });
  renderTable(
    {table:'#completedTable', thead:'#completedTable thead', tbody:'#completedBody'},
    data,
    {
      actions:[
        { header:'Action', label:'Delete', title:'Delete', className:'btn-icon btn-delete', handler:deleteItem }
      ]
    }
  );
}

function deleteItem(id){ if(!confirm('Delete this item?')) return; saveData(loadData().filter(function(x){return x.id!==id;})); filterCompleted(); }

function computeTimeSpent(a,b){
  var s=new Date(a), e=new Date(b);
  if(isNaN(s)||isNaN(e)||e<s) return '';
  var m=Math.round((e-s)/60000);
  if(m<60) return m+' min';
  var h=Math.floor(m/60), mm=m%60; return mm ? (h+'h '+mm+'m') : (h+'h');
}
function calculateAndSetTimeSpent(){
  var ts=document.getElementById('timeStarted'), te=document.getElementById('timeEnded'), sp=document.getElementById('timeSpent');
  if(!ts||!te||!sp) return; var calc=computeTimeSpent(ts.value, te.value); if(calc) sp.value=calc;
}

function ensureDbRows(count){ count=count||1; for(var i=0;i<count;i++) addDbRow({}); }
function addDbRow(prefill){ prefill=prefill||{}; var c=document.getElementById('dbRows'); if(!c) return;
  var idx=c.querySelectorAll('.db-row').length+1; if(idx>10) return;
  var row=document.createElement('div'); row.className='db-row'; row.dataset.index=String(idx);
  var db=document.createElement('select'); db.name='database'+idx; db.id='database'+idx;
  var first=document.createElement('option'); first.value=''; first.textContent='-- Select database --'; db.appendChild(first);
  DATABASE_SOURCES.forEach(function(src){ var op=document.createElement('option'); op.value=src; op.textContent=src; db.appendChild(op); });
  if(prefill['database'+idx]) db.value=prefill['database'+idx];
  var pg=document.createElement('input'); pg.type='number'; pg.name='pages'+idx; pg.id='pages'+idx; pg.placeholder='Pages'; pg.min='0';
  if(prefill['pages'+idx]!==undefined) pg.value=prefill['pages'+idx];
  var rm=document.createElement('button'); rm.type='button'; rm.className='remove table-btn'; rm.textContent='Remove';
  rm.addEventListener('click', function(){ row.remove(); });
  row.appendChild(db); row.appendChild(pg); row.appendChild(rm); c.appendChild(row);
}

function getFormData(form){
  var fd=new FormData(form); var data={}; fd.forEach(function(v,k){ data[k]=v; });
  if(data.noOfCompanies) data.noOfCompanies=String(data.noOfCompanies);
  if(data.timeStarted&&data.timeEnded){ var calc=computeTimeSpent(data.timeStarted,data.timeEnded); if(calc) data.timeSpent=calc; }
  for(var i=1;i<=10;i++){ data['database'+i]=data['database'+i]||''; data['pages'+i]=data['pages'+i]||''; }
  return data;
}
function populateForm(form,item){
  var keys=['id','workflowName','clientEmail','receivedDateTime','researcher','requestDetails','deadline','status','foa','noOfCompanies','complexity','requestType','timeStarted','timeEnded','timeSpent'].concat(DB_COLUMNS,PAGES_COLUMNS);
  keys.forEach(function(k){
    var el=document.getElementById(k); if(!el) return;
    if(k==='requestType'){ el.dataset.prefillValue=item[k]||''; } else { el.value=item[k]||''; }
  });
  var present=Array.from({length:10},function(_,i){return i+1}).filter(function(i){ return item['database'+i]||item['pages'+i]; }).length;
  var need=Math.max(1,present); ensureDbRows(need);
  for(var i=1;i<=need;i++){ var db=document.getElementById('database'+i); var pg=document.getElementById('pages'+i);
    if(db) db.value=item['database'+i]||''; if(pg) pg.value=item['pages'+i]||'';
  }
  var cx=document.getElementById('complexity'); if(cx && !cx.value) cx.value='1'; updateRequestType(); applyWorkflowNameFreeze();
}
function handleSubmit(e){
  e.preventDefault(); var form=e.target; var incoming=getFormData(form); var all=loadData(); var editId=localStorage.getItem(EDIT_ID_KEY);
  incoming.id=(incoming.id && incoming.id.trim()) ? incoming.id : (editId || generateUniqueId());
  var idx=all.findIndex(function(x){ return x.id===incoming.id; });
  if(idx>=0) all[idx]=Object.assign({},all[idx],incoming); else all.push(incoming);
  saveData(all); localStorage.removeItem(EDIT_ID_KEY); alert('Saved successfully.'); location.href='status_dashboard.html';
}

function buildOptions(select,items,opts){ if(!select) return; opts=opts||{}; select.innerHTML='';
  if(opts.placeholder){ var ph=document.createElement('option'); ph.value=''; ph.textContent=opts.placeholder; select.appendChild(ph); }
  items.forEach(function(v){ var op=document.createElement('option'); op.value=v; op.textContent=v; select.appendChild(op); });
  if(typeof opts.defaultValue!=='undefined') select.value=opts.defaultValue;
}

function tableToCSV(table){
  var rows=[];
  function esc(v){ if(v==null) return ''; var s=String(v).replace(/\r?\n/g,'\n'); return /[",\n]/.test(s)? '"'+s.replace(/"/g,'""')+'"' : s; }
  var head=Array.prototype.map.call(table.querySelectorAll('thead tr th'),function(th){ return th.textContent.trim(); });
  rows.push(head.map(esc).join(','));
  Array.prototype.forEach.call(table.querySelectorAll('tbody tr'),function(tr){
    var cells=Array.prototype.map.call(tr.querySelectorAll('td'),function(td){
      var t=td.getAttribute('title'); if(t) return esc(t);
      var trEl=td.querySelector('.truncate-1, .truncate-3'); if(trEl) return esc(trEl.getAttribute('title')||trEl.textContent.trim());
      var badge=td.querySelector('.status-badge'); if(badge) return esc(badge.textContent.trim());
      return esc(td.textContent.trim());
    });
    rows.push(cells.join(','));
  });
  var blob=new Blob(['\ufeff'+rows.join('\n')],{type:'text/csv;charset=utf-8;'});
  var a=document.createElement('a'); a.href=URL.createObjectURL(blob); a.download='completed_tasks.csv'; a.click(); URL.revokeObjectURL(a.href);
}
function exportToExcel(){
  try{
    var table=document.getElementById('completedTable');
    if(window.XLSX){
      var ws=XLSX.utils.table_to_sheet(table);
      var wb=XLSX.utils.book_new(); XLSX.utils.book_append_sheet(wb, ws, 'Completed Tasks'); XLSX.writeFile(wb, 'completed_tasks.xlsx');
    } else { tableToCSV(table); }
  }catch(err){ console.error(err); alert('Export failed. Falling back to CSV.'); tableToCSV(document.getElementById('completedTable')); }
}

document.addEventListener('DOMContentLoaded', function(){
  try{
    document.body.addEventListener('click', function(e){ var x=e.target.closest('[data-action="close-modal"]'); if(x){ e.preventDefault(); closeModal(); }});
    var page=document.body.getAttribute('data-page') || (location.pathname.split('/').pop().replace('.html','')||'status_dashboard');

    if(page==='status_dashboard'){
      buildOptions(document.getElementById('researcherFilter'), RESEARCHERS, {defaultValue:'All'});
      loadDashboard();
      var s=document.getElementById('statusFilter'); if(s) s.addEventListener('change', filterDashboard);
      var r=document.getElementById('researcherFilter'); if(r) r.addEventListener('change', filterDashboard);
    }

    if(page==='workflow_form'){
      var form=document.getElementById('workflowForm'); if(form) form.addEventListener('submit', handleSubmit);
      buildOptions(document.getElementById('researcher'), RESEARCHERS_NO_ALL, {defaultValue:'New'});
      var cx=document.getElementById('complexity');
      if(cx) cx.addEventListener('change', function(){ updateRequestType(); applyWorkflowNameFreeze(); });
      ensureDbRows(1);
      var ts=document.getElementById('timeStarted'), te=document.getElementById('timeEnded');
      if(ts) ts.addEventListener('change', calculateAndSetTimeSpent);
      if(te) te.addEventListener('change', calculateAndSetTimeSpent);
      var editId=localStorage.getItem(EDIT_ID_KEY);
      if(editId){
        var item=loadData().find(function(x){return x.id===editId;});
        if(item){ var idEl=document.getElementById('id'); if(idEl) idEl.value=editId; populateForm(form,item); }
        else { if(cx&& !cx.value) cx.value='1'; updateRequestType(); applyWorkflowNameFreeze(); }
      } else { if(cx&& !cx.value) cx.value='1'; updateRequestType(); applyWorkflowNameFreeze(); }
      var addBtn=document.getElementById('addDbRow'); if(addBtn) addBtn.addEventListener('click', function(){ addDbRow({}); });
    }

    if(page==='completed'){
      buildOptions(document.getElementById('completedResearcherFilter'), RESEARCHERS, {defaultValue:'All'});
      loadCompleted();
      ['complexityFilter','completedResearcherFilter','dateRangeStart','dateRangeEnd'].forEach(function(id){
        var el=document.getElementById(id); if(el) el.addEventListener('change', filterCompleted);
      });
      var btn=document.getElementById('btnExport'); if(btn) btn.addEventListener('click', function(e){ e.preventDefault(); exportToExcel(); });
    }

    var health=document.getElementById('health'); if(health){ health.textContent='✅ Script loaded OK'; health.className='health ok'; }
  }catch(err){ console.error('Init error:', err); alert('Initialization error. Ensure all files are in the same folder and reload.'); }
});

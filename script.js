// ======================================================================
//  CONFIG
// ======================================================================
const PERSON_COLORS = { Praveen:'#3b82f6', Pavan:'#f3c623', Kumar:'#16a34a', 'Kumar & Praveen':'#8b5cf6' };
const PRIORITY_LABELS = {1:'High',2:'Medium',3:'Low',4:'Very Low'};
const PRIORITY_COLORS = ['#f3c623','#3b82f6','#8b5cf6','#6b7280'];
const STATUS_COLORS = {
  'Hot Lead':'#f3c623','Follow-up Needed':'#3b82f6','Quotation Needed':'#3b82f6',
  'Sample Needed':'#8b5cf6','No Requirement':'#6b7280','Already has Vendor':'#8b5cf6',
  'Not Reachable':'rgba(255,255,255,.5)','Not Interested':'#ef4444','ABSENT':'rgba(255,255,255,.2)'
};

function normalizePersonName(name) {
  if (!name) return '';
  let n = name.replace(/\s+/g,' ').trim();
  if (/^kumar\s*r?$/i.test(n) || /^kuma\s*r$/i.test(n)) return 'Kumar';
  if (/^kumar\s*&\s*praveen$/i.test(n)) return 'Kumar & Praveen';
  return n.charAt(0).toUpperCase() + n.slice(1);
}

// ======================================================================
//  FILE HANDLING
// ======================================================================
let akmFile = null, nehaFile = null;
const akmZone = document.getElementById('akm-zone');
const nehaZone = document.getElementById('neha-zone');
const genBtn = document.getElementById('generate-btn');

function setupZone(zone, input, onFile) {
  zone.addEventListener('dragover', e => { e.preventDefault(); zone.classList.add('dragover'); });
  zone.addEventListener('dragleave', () => zone.classList.remove('dragover'));
  zone.addEventListener('drop', e => { e.preventDefault(); zone.classList.remove('dragover'); if(e.dataTransfer.files.length) onFile(e.dataTransfer.files[0]); });
  input.addEventListener('change', e => { if(e.target.files.length) onFile(e.target.files[0]); });
}
function setAkm(f) { akmFile=f; akmZone.classList.add('has-file'); document.getElementById('akm-filename').textContent=f.name; checkReady(); }
function setNeha(f) { nehaFile=f; nehaZone.classList.add('has-file'); document.getElementById('neha-filename').textContent=f.name; checkReady(); }

setupZone(akmZone, document.getElementById('akm-input'), setAkm);
setupZone(nehaZone, document.getElementById('neha-input'), setNeha);
genBtn.addEventListener('click', () => { document.getElementById('loading').classList.remove('hidden'); setTimeout(buildDashboard, 100); });

const clearBtn = document.getElementById('clear-btn');
function removeFile(type) {
  if (type === 'akm') {
    akmFile = null;
    akmZone.classList.remove('has-file');
    document.getElementById('akm-input').value = '';
    document.getElementById('akm-filename').textContent = '';
  } else {
    nehaFile = null;
    nehaZone.classList.remove('has-file');
    document.getElementById('neha-input').value = '';
    document.getElementById('neha-filename').textContent = '';
  }
  checkReady();
}
function clearAllFiles() {
  removeFile('akm');
  removeFile('neha');
  checkReady();
}
function checkReady() {
  const hasFile = !!akmFile;
  genBtn.style.display = hasFile ? 'inline-flex' : 'none';
  genBtn.disabled = !hasFile;
  genBtn.style.alignItems = 'center';
  clearBtn.style.display = hasFile || nehaFile ? 'inline-flex' : 'none';
  clearBtn.style.alignItems = 'center';
}

// ======================================================================
//  EXCEL PARSING
// ======================================================================
function readExcel(file) {
  return new Promise((resolve, reject) => {
    const reader = new FileReader();
    reader.onload = e => {
      try {
        const wb = XLSX.read(e.target.result, {type:'array', cellFormula:true, cellDates:true});
        resolve(wb);
      } catch(err) { reject(err); }
    };
    reader.onerror = e => reject(new Error('Failed to read file'));
    reader.readAsArrayBuffer(file);
  });
}
function sheetToArray(wb, name) {
  const ws = wb.Sheets[name]; if(!ws) return [];
  return XLSX.utils.sheet_to_json(ws, {header:1, defval:'', raw:true});
}
function resolveCellRef(ref, data) {
  if (typeof ref !== 'string' || !ref.startsWith('=')) return ref;
  const r = ref.substring(1).trim();
  const m = r.match(/^([A-Z]+)(\d+)$/); if(!m) return ref;
  let colIdx = 0;
  for (let i=0; i<m[1].length; i++) colIdx = colIdx*26 + (m[1].charCodeAt(i)-64);
  colIdx -= 1;
  const rowIdx = parseInt(m[2]) - 1;
  if (data[rowIdx] && data[rowIdx][colIdx] !== undefined && data[rowIdx][colIdx] !== '') return data[rowIdx][colIdx];
  return ref;
}
function resolveAllRefs(data) {
  for(let r=0; r<data.length; r++) for(let c=0; c<(data[r]||[]).length; c++) data[r][c] = resolveCellRef(data[r][c], data);
  return data;
}

// ======================================================================
//  STATUS CLASSIFICATION
// ======================================================================
function classifyStatus(remark) {
  if (!remark || remark === '') return 'Not Reachable';
  const s = remark.toLowerCase();
  if (s.includes('absent') && s.length < 20) return 'ABSENT';
  if (s.includes('not went to market')) return 'ABSENT';
  if (s.includes('not interested')) return 'Not Interested';
  if (s.includes('not needed') || s.includes('no requirement') || s.includes('presently not needed') || s.includes('no requir')) return 'No Requirement';
  if (s.includes('already vendor') || s.includes('already order given') || s.includes('already family dealer') || s.includes('already given order') || s.includes('already regentra') || s.includes('already purchased')) return 'Already has Vendor';
  if (s.includes('order received') || s.includes('delivered')) return 'Hot Lead';
  if (s.includes('send sample') || s.includes('give sample') || s.includes('sample testing') || s.includes('sample we check') || s.includes('give me some sample')) return 'Sample Needed';
  if (s.includes('send quotation') || s.includes('send me quotation') || s.includes('give me best quotation') || s.includes('based on quotation') || s.includes('based on it we will decide') || s.includes('send quatation') || s.includes('send me quatation') || s.includes('prepare and send quotation') || s.includes('he will give order')) return 'Hot Lead';
  if (s.includes('interested')) return 'Hot Lead';
  if (s.includes('busy') || s.includes('not responding') || s.includes('wrong number') || s.includes('call disconnected') || s.includes('unreachable') || s.includes('not available') || s.includes('not aviliable') || s.includes('security not allowed') || s.includes('photos not allowed') || s.includes('concern person not available') || s.includes('not there') || s.includes('in meeting') || s.includes('out of city') || s.includes('out of station') || s.includes('not in town') || s.includes('sir is out') || s.includes('owner is out') || s.includes('manager is out') || s.includes('head sir is')) return 'Not Reachable';
  if (s.includes('tomorrow') || s.includes('tommrow') || s.includes('tommorw') || s.includes('tommorow') || s.includes('will update') || s.includes('will contact') || s.includes('check & update') || s.includes('check and update') || s.includes('come tomorrow') || s.includes('come tommrow') || s.includes('come next time') || s.includes('he will update') || s.includes('he will contact') || s.includes('will discuss') || s.includes('will check') || s.includes('will call') || s.includes('he will text') || s.includes('he will get back') || s.includes('will whatsapp') || s.includes('next month') || s.includes('month end') || s.includes('follow up')) return 'Follow-up Needed';
  return 'Follow-up Needed';
}

function classifyNehaRemark(remark) {
  if (!remark || remark === '') return 'No Remark';
  const s = remark.toLowerCase();
  if (s.includes('not interested')) return 'Not Interested';
  if (s.includes('no requirement') || s.includes('present no requirement')) return 'No Requirement';
  if (s.includes('busy') || s.includes('not responding') || s.includes('wrong number') || s.includes('call disconnected') || s.includes('not available') || s.includes('concern person')) return 'Not Reachable';
  if (s.includes('order received') || s.includes('delivered')) return 'Interested';
  if (s.includes('send quotation') || s.includes('send me quotation') || s.includes('send quatation')) return 'Quotation Requested';
  if (s.includes('interested')) return 'Interested';
  if (s.includes('he will call') || s.includes('he will text') || s.includes('he will update') || s.includes('he will discuss') || s.includes('he will talk') || s.includes('he will get back')) return 'Follow-up Pending';
  if (s.includes('email') || s.includes('mail')) return 'Email Shared';
  return 'Other';
}

function extractActionItems(remark) {
  if (!remark) return '';
  const s = remark.toLowerCase();
  const items = [];
  if (s.includes('tomorrow') || s.includes('tommrow') || s.includes('tommorw') || s.includes('tommorow')) items.push('Visit tomorrow');
  if (s.includes('send quotation') || s.includes('send me quotation') || s.includes('send quatation') || s.includes('give me best quotation')) items.push('Send quotation');
  if (s.includes('send sample') || s.includes('give sample') || s.includes('sample testing') || s.includes('give me some sample')) items.push('Send samples');
  if (s.includes('come at') || s.includes('come tommrow') || s.includes('come tomorrow') || s.includes('come next time')) items.push('Schedule visit');
  if (s.includes('month end') || s.includes('next month')) items.push('Follow-up later');
  if (s.includes('will update') || s.includes('check & update') || s.includes('check and update')) items.push('Awaiting update');
  if (s.includes('will contact') || s.includes('he will call')) items.push('Awaiting callback');
  return items.join(' | ');
}

function normalizeDate(d) {
  if (!d || d === '') return '';
  if (d instanceof Date && !isNaN(d)) { const y=d.getFullYear(),m=String(d.getMonth()+1).padStart(2,'0'),dd=String(d.getDate()).padStart(2,'0'); return y+'-'+m+'-'+dd; }
  if (typeof d === 'number' && d > 40000 && d < 60000) {
    const date = new Date((d - 25569) * 86400 * 1000);
    return date.getFullYear()+'-'+String(date.getMonth()+1).padStart(2,'0')+'-'+String(date.getDate()).padStart(2,'0');
  }
  const s = String(d).trim();
  if (s.includes('T')) return s.split('T')[0];
  if (s.match(/^\d{4}-\d{2}-\d{2}/)) return s.split(' ')[0];
  const parts = s.match(/^(\d{1,2})[\/\-](\d{1,2})[\/\-](\d{4})$/);
  if (parts) { const y=parts[3],m=parts[2].padStart(2,'0'),dd=parts[1].padStart(2,'0'); return y+'-'+m+'-'+dd; }
  const p2 = s.match(/^(\d{4})[\/\-](\d{1,2})[\/\-](\d{1,2})$/);
  if (p2) { const y=p2[1],m=p2[2].padStart(2,'0'),dd=p2[3].padStart(2,'0'); return y+'-'+m+'-'+dd; }
  return s;
}

// ======================================================================
//  DATA PROCESSING
// ======================================================================
let DASHBOARD_DATA = {};

async function buildDashboard() {
  try {
    console.log('Starting build...');
    const akmWb = await readExcel(akmFile);
    console.log('AKM workbook loaded, sheets:', akmWb.SheetNames);
    let nehaWb = null;
    if (nehaFile) nehaWb = await readExcel(nehaFile);

    let akmRaw = resolveAllRefs(sheetToArray(akmWb, 'DATA'));
    console.log('AKM rows:', akmRaw.length);

    let nehaRaw = [];
    if (nehaWb) {
      const sn = nehaWb.SheetNames.find(n => n.toLowerCase().includes('neha') || n.toLowerCase().includes('remark'));
      nehaRaw = sheetToArray(nehaWb, sn || nehaWb.SheetNames[0]);
    }

    const akmData = [];
    for (let i = 1; i < akmRaw.length; i++) {
      const row = akmRaw[i];
      if (!row || row.length === 0) continue;
      const person = normalizePersonName(row[2]);
      const clientName = (row[3] || '').toString().trim();
      if (!person && !clientName) continue;
      if (person.toUpperCase() === 'ABSENT') { akmData.push({date:row[1]||'',person:'ABSENT',clientName:'',priority:'',dsoRemark:'ABSENT',area:'',km:0,phone:'',email:''}); continue; }
      if (!person && clientName) continue;
      let priority = row[4];
      if (typeof priority === 'string') priority = parseInt(priority) || '';
      akmData.push({ date:row[1]||'', person, clientName, priority, dsoRemark:(row[5]||'').toString().trim(), area:(row[6]||'').toString().trim(), km:parseFloat(row[7])||0, phone:(row[8]||'').toString().trim(), email:(row[9]||'').toString().trim() });
    }

    const nehaData = [];
    if (nehaRaw.length > 1) {
      for (let i = 1; i < nehaRaw.length; i++) {
        const row = nehaRaw[i]; if (!row || row.length === 0) continue;
        const person = normalizePersonName(row[2]);
        const clientName = (row[3] || '').toString().trim();
        if (!person && !clientName) continue;
        nehaData.push({ date:row[1]||'', person, clientName, area:(row[4]||'').toString().trim(), email:(row[5]||'').toString().trim(), phone:(row[6]||'').toString().trim(), nehaRemark:(row[7]||'').toString().trim(), dsoRemark:(row[8]||'').toString().trim() });
      }
    }

    DASHBOARD_DATA = processAllData(akmData, nehaData);
    renderDashboard();
  } catch (err) {
    console.error(err);
    alert('Error processing files: ' + err.message);
    document.getElementById('loading').classList.add('hidden');
  }
}

function processAllData(akmData, nehaData) {
  const d = { akm: akmData, neha: nehaData, persons: {}, summary: {}, followUps: [], nehaAnalysis: {}, allClientNames: [] };

  const personSet = new Set();
  akmData.forEach(r => { if (r.person && r.person !== 'ABSENT') r.person.split(/\s*&\s*/).forEach(p => personSet.add(p.trim())); });
  const persons = [...personSet].filter(p => p);

  persons.forEach(p => {
    const myLeads = akmData.filter(r => { if (!r.person || r.person === 'ABSENT') return false; return r.person.split(/\s*&\s*/).map(x=>x.trim()).includes(p); });
    const priorityBreakdown = {1:0,2:0,3:0,4:0};
    const statusBreakdown = {};
    const areaMap = {};
    const dailyMap = {};
    let totalKm = 0, phones = 0, emails = 0;
    const uniqueClients = new Set();

    myLeads.forEach(r => {
      const pr = parseInt(r.priority);
      if (pr >= 1 && pr <= 4) priorityBreakdown[pr]++;
      const status = classifyStatus(r.dsoRemark);
      statusBreakdown[status] = (statusBreakdown[status] || 0) + 1;
      if (r.area) areaMap[r.area] = (areaMap[r.area] || 0) + 1;
      totalKm += r.km;
      if (r.phone && r.phone.length >= 7) phones++;
      if (r.email && r.email.includes('@')) emails++;
      if (r.clientName) uniqueClients.add(r.clientName.toLowerCase().trim());
      const dateStr = normalizeDate(r.date);
      if (dateStr) dailyMap[dateStr] = (dailyMap[dateStr] || 0) + 1;
    });

    d.persons[p] = { leads:myLeads, priorityBreakdown, statusBreakdown, areaMap, dailyMap, totalKm:Math.round(totalKm), totalLeads:myLeads.length, phones, emails, uniqueVisits: uniqueClients.size };
  });

  const allLeads = akmData.filter(r => r.person && r.person !== 'ABSENT' && r.clientName);
  let totalKm=0, totalPhones=0, totalEmails=0, p1=0, followups=0;
  const allUniqueClients = new Set();
  allLeads.forEach(r => {
    totalKm += r.km;
    if (r.phone && r.phone.length >= 7) totalPhones++;
    if (r.email && r.email.includes('@')) totalEmails++;
    const pr = parseInt(r.priority);
    if (pr === 1) p1++;
    const s = classifyStatus(r.dsoRemark);
    if (s === 'Follow-up Needed' || s === 'Quotation Needed' || s === 'Sample Needed') followups++;
    if (r.clientName) allUniqueClients.add(r.clientName.toLowerCase().trim());
  });
  d.summary = { totalLeads:allLeads.length, totalPersons:persons.length, totalKm:Math.round(totalKm), totalPhones, totalEmails, p1, followups, absentDays:akmData.filter(r=>r.person==='ABSENT').length, uniqueClients:allUniqueClients.size };

  allLeads.forEach(r => {
    const status = classifyStatus(r.dsoRemark);
    if (['Follow-up Needed','Quotation Needed','Sample Needed','Hot Lead'].includes(status)) {
      d.followUps.push({...r, status, actionItems:extractActionItems(r.dsoRemark)});
    }
  });
  d.followUps.sort((a,b) => { const o={'Hot Lead':0,'Quotation Needed':1,'Sample Needed':2,'Follow-up Needed':3}; return (o[a.status]||4)-(o[b.status]||4); });

  d.allClientNames = allLeads.map(r => ({ clientName:r.clientName, person:r.person, date:normalizeDate(r.date), area:r.area, priority:r.priority, dsoRemark:r.dsoRemark, nehaRemark:'', status:classifyStatus(r.dsoRemark) }));
  if (nehaData.length > 0) {
    nehaData.forEach(nr => {
      const existing = d.allClientNames.find(c => c.clientName.toLowerCase().trim() === nr.clientName.toLowerCase().trim());
      if (existing) { existing.nehaRemark = nr.nehaRemark; }
      else { d.allClientNames.push({ clientName:nr.clientName, person:nr.person, date:normalizeDate(nr.date), area:nr.area, priority:'', dsoRemark:nr.dsoRemark||'', nehaRemark:nr.nehaRemark, status:classifyStatus(nr.dsoRemark) }); }
    });

    const categories = {}, dsoResponses = {}, nehaTimeline = {};
    nehaData.forEach(r => { const cat = classifyNehaRemark(r.nehaRemark); categories[cat] = (categories[cat]||0)+1; const dsoCat = classifyStatus(r.dsoRemark); dsoResponses[dsoCat] = (dsoResponses[dsoCat]||0)+1; const ds = normalizeDate(r.date); if(ds) nehaTimeline[ds]=(nehaTimeline[ds]||0)+1; });
    d.nehaAnalysis = { categories, dsoResponses, nehaTimeline, totalRemarks:nehaData.length };
  }

  return d;
}

// ======================================================================
//  CHART DEFAULTS (Sentry palette)
// ======================================================================
const chartFont = { fontFamily:'Manrope,-apple-system,system-ui,sans-serif' };
const gridColor = 'rgba(255,255,255,.08)';
const labelColor = 'rgba(255,255,255,.6)';

function makeChart(el, opts) {
  const base = {
    chart: { ...chartFont, background:'transparent', toolbar:{show:false} },
    theme: { mode:'dark' },
    grid: { borderColor:gridColor, strokeDashArray:3 },
    tooltip: { theme:'dark', style:{...chartFont, fontSize:'12px'} },
    dataLabels: { style:{...chartFont} },
    ...opts
  };
  const c = new ApexCharts(el, base);
  c.render();
  charts.push(c);
  return c;
}

// ======================================================================
//  RENDER
// ======================================================================
const charts = [];
function destroyCharts() { charts.forEach(c => { try{c.destroy()}catch(e){} }); charts.length = 0; }

function renderDashboard() {
  destroyCharts();
  document.getElementById('upload-section').style.display = 'none';
  document.getElementById('dashboard').style.display = 'block';
  document.getElementById('loading').classList.add('hidden');

  const d = DASHBOARD_DATA;
  document.getElementById('dash-subtitle').textContent = d.summary.totalLeads + ' leads across ' + d.summary.totalPersons + ' persons | ' + d.summary.totalKm + ' km traveled | ' + d.summary.uniqueClients + ' unique clients';

  renderSummaryCards(d);
  renderUniqueVisits(d);
  renderPersonTabs(d);
  renderFollowUps(d);
  renderNehaSection(d);
  setupSearch(d);
}

function renderSummaryCards(d) {
  const s = d.summary;
  document.getElementById('summary-cards').innerHTML =
    card('gold','Total Leads',s.totalLeads,'clients contacted') +
    card('','Persons',s.totalPersons,'active reps') +
    card('gold','Total KM',s.totalKm,'distance traveled') +
    card('gold','Priority 1',s.p1,'high priority') +
    card('gold','Unique Clients',s.uniqueClients,'unique visits') +
    card('','Phones',s.totalPhones,'collected') +
    card('','Emails',s.totalEmails,'collected') +
    card('red','Follow-ups',s.followups,'pending actions');
}
function card(colorCls,label,value,sub){ return '<div class="stat-card"><div class="label">'+label+'</div><div class="value '+(colorCls||'')+'">'+value+'</div>'+(sub?'<div class="sub">'+sub+'</div>':'')+'</div>'; }

// ======================================================================
//  UNIQUE VISITS KPIs
// ======================================================================
function renderUniqueVisits(d) {
  const el = document.getElementById('unique-visits');
  const persons = Object.keys(d.persons);
  const bgColors = { Praveen:'rgba(59,130,246,.15)', Pavan:'rgba(243,198,35,.15)', Kumar:'rgba(22,163,74,.15)' };
  const textColors = { Praveen:'#60a5fa', Pavan:'#f3c623', Kumar:'#4ade80' };

  el.innerHTML = persons.map(p => {
    const pd = d.persons[p];
    const initials = p.substring(0,2).toUpperCase();
    return '<div class="uv-card"><div class="uv-avatar" style="background:'+(bgColors[p]||'rgba(255,255,255,.06)')+';color:'+(textColors[p]||'var(--text)')+'">'+initials+'</div><div class="uv-info"><div class="uv-name">'+p+'</div><div class="uv-count" style="color:'+(textColors[p]||'var(--text)')+'">'+pd.uniqueVisits+'</div><div class="uv-sub">unique clients visited</div></div></div>';
  }).join('');
}

// ======================================================================
//  PERSON TABS
// ======================================================================
function renderPersonTabs(d) {
  const tabsEl = document.getElementById('person-tabs');
  const persons = Object.keys(d.persons);
  tabsEl.innerHTML = persons.map((p,i) => '<button class="person-tab '+(i===0?'active':'')+'" onclick="switchPerson(\''+p+'\')">'+p+'<span class="badge">'+d.persons[p].totalLeads+'</span></button>').join('');
  switchPerson(persons[0]);
}

function switchPerson(name) {
  document.querySelectorAll('.person-tab').forEach(t => t.classList.remove('active'));
  document.querySelectorAll('.person-tab').forEach(t => { if(t.textContent.includes(name)) t.classList.add('active'); });
  const contentEl = document.getElementById('person-content');
  const p = DASHBOARD_DATA.persons[name];
  if (!p) { contentEl.innerHTML = '<p style="color:var(--text-muted)">No data</p>'; return; }

  const ids = ['pr','st','da','ar','km'].map(x => x+'-'+Date.now()+'-'+Math.random().toString(36).slice(2,6));

  contentEl.innerHTML =
    '<div style="display:grid;grid-template-columns:repeat(auto-fit,minmax(140px,1fr));gap:12px;margin-bottom:24px">'+
      card('gold','Leads',p.totalLeads,'')+
      card('gold','KM',p.totalKm,'traveled')+
      card('','Phones',p.phones,'')+
      card('','Emails',p.emails,'')+
      card('gold','Unique',p.uniqueVisits,'clients')+
    '</div>'+
    '<div class="charts-row"><div class="chart-card"><h3>Priority Breakdown</h3><div id="'+ids[0]+'"></div></div><div class="chart-card"><h3>Lead Status Analysis</h3><div id="'+ids[1]+'"></div></div></div>'+
    '<div class="charts-row"><div class="chart-card full"><h3>Daily Activity</h3><div id="'+ids[2]+'"></div></div></div>'+
    '<div class="charts-row"><div class="chart-card"><h3>Area Coverage (Leads)</h3><div id="'+ids[3]+'"></div></div><div class="chart-card"><h3>Area Coverage (KM)</h3><div id="'+ids[4]+'"></div></div></div>';

  setTimeout(() => {
    renderPriorityChart(ids[0], p);
    renderStatusChart(ids[1], p);
    renderDailyChart(ids[2], p);
    renderAreaChart(ids[3], p);
    renderAreaKmChart(ids[4], p);
  }, 60);
}

// ======================================================================
//  CHARTS
// ======================================================================
function renderPriorityChart(id, p) {
  const el = document.getElementById(id); if(!el) return;
  const labels=[], series=[], colors=[];
  [1,2,3,4].forEach(pr => { if(p.priorityBreakdown[pr]>0){ labels.push('P'+pr+' '+PRIORITY_LABELS[pr]); series.push(p.priorityBreakdown[pr]); colors.push(PRIORITY_COLORS[pr-1]); } });
  if(!series.length){ el.innerHTML='<p style="color:var(--text-muted);text-align:center;padding:32px">No priority data</p>'; return; }
  makeChart(el, {
    chart: { type:'donut', height:260, ...chartFont, background:'transparent', toolbar:{show:false} },
    series, labels, colors,
    plotOptions: { pie: { donut: { size:'65%', labels:{ show:true, total:{ show:true, color:'#ffffff', label:'Total', style:{...chartFont} } } } } },
    legend: { position:'bottom', labels:{ colors:labelColor, ...chartFont, fontSize:'11px' } },
    dataLabels: { enabled:true, formatter:v=>v.toFixed(0)+'%', style:{...chartFont, fontSize:'11px', fontWeight:600} }
  });
}

function renderStatusChart(id, p) {
  const el = document.getElementById(id); if(!el) return;
  const labels=[], series=[], colors=[];
  Object.keys(STATUS_COLORS).forEach(c => { if(p.statusBreakdown[c]){labels.push(c);series.push(p.statusBreakdown[c]);colors.push(STATUS_COLORS[c]);} });
  if(!series.length){ el.innerHTML='<p style="color:var(--text-muted);text-align:center;padding:32px">No status data</p>'; return; }
  makeChart(el, {
    chart: { type:'bar', height:260, ...chartFont, background:'transparent', toolbar:{show:false} },
    series: [{ name:'Leads', data:series }],
    colors: ['#3b82f6'],
    plotOptions: { bar: { horizontal:true, borderRadius:4, barHeight:'60%' } },
    xaxis: { categories:labels, labels:{style:{colors:labelColor,...chartFont,fontSize:'11px'}} },
    yaxis: { labels:{style:{colors:labelColor,...chartFont,fontSize:'11px'}} },
    grid: { borderColor:gridColor, strokeDashArray:3 },
    dataLabels: { enabled:true, style:{...chartFont,fontSize:'11px'} },
    legend: { show:false }
  });
}

function renderDailyChart(id, p) {
  const el = document.getElementById(id); if(!el) return;
  const dates = Object.keys(p.dailyMap).sort();
  if(!dates.length){ el.innerHTML='<p style="color:var(--text-muted);text-align:center;padding:32px">No daily data</p>'; return; }
  const shortDates = dates.map(d => { const parts=d.split('-'); return parts[2]+'/'+parts[1]; });
  makeChart(el, {
    chart: { type:'area', height:280, ...chartFont, background:'transparent', toolbar:{show:false} },
    series: [{ name:'Leads', data:dates.map(d=>p.dailyMap[d]) }],
    colors: ['#f3c623'],
    stroke: { curve:'smooth', width:2.5 },
    fill: { type:'gradient', gradient:{ shadeIntensity:1, opacityFrom:0.25, opacityTo:0.02, stops:[0,100] } },
    xaxis: { categories:shortDates, labels:{style:{colors:labelColor,...chartFont,fontSize:'11px'}} },
    yaxis: { labels:{style:{colors:labelColor,...chartFont,fontSize:'11px'}}, min:0, tickAmount:5 },
    grid: { borderColor:gridColor, strokeDashArray:3 },
    markers: { size:4, colors:['#f3c623'], strokeColors:'#000000', strokeWidth:2 },
    dataLabels: { enabled:false },
    legend: { show:false }
  });
}

function renderAreaChart(id, p) {
  const el = document.getElementById(id); if(!el) return;
  const areas = Object.entries(p.areaMap).sort((a,b)=>b[1]-a[1]).slice(0,12);
  if(!areas.length){ el.innerHTML='<p style="color:var(--text-muted);text-align:center;padding:32px">No area data</p>'; return; }
  makeChart(el, {
    chart: { type:'bar', height:Math.max(200,areas.length*32+60), ...chartFont, background:'transparent', toolbar:{show:false} },
    series: [{ name:'Leads', data:areas.map(a=>a[1]) }],
    colors: ['#8b5cf6'],
    plotOptions: { bar: { horizontal:true, borderRadius:3, barHeight:'55%' } },
    xaxis: { categories:areas.map(a=>a[0]), labels:{style:{colors:labelColor,...chartFont,fontSize:'10px'}} },
    yaxis: { labels:{style:{colors:labelColor,...chartFont,fontSize:'10px'}} },
    grid: { borderColor:gridColor, strokeDashArray:3 },
    dataLabels: { enabled:true, style:{...chartFont,fontSize:'10px'} },
    legend: { show:false }
  });
}

function renderAreaKmChart(id, p) {
  const el = document.getElementById(id); if(!el) return;
  const areaKm={}; p.leads.forEach(r=>{if(r.area&&r.km)areaKm[r.area]=(areaKm[r.area]||0)+r.km;});
  const areas = Object.entries(areaKm).sort((a,b)=>b[1]-a[1]).slice(0,12);
  if(!areas.length){ el.innerHTML='<p style="color:var(--text-muted);text-align:center;padding:32px">No KM data</p>'; return; }
  makeChart(el, {
    chart: { type:'bar', height:Math.max(200,areas.length*32+60), ...chartFont, background:'transparent', toolbar:{show:false} },
    series: [{ name:'KM', data:areas.map(a=>Math.round(a[1])) }],
    colors: ['#3b82f6'],
    plotOptions: { bar: { horizontal:true, borderRadius:3, barHeight:'55%' } },
    xaxis: { categories:areas.map(a=>a[0]), labels:{style:{colors:labelColor,...chartFont,fontSize:'10px'}} },
    yaxis: { labels:{style:{colors:labelColor,...chartFont,fontSize:'10px'}} },
    grid: { borderColor:gridColor, strokeDashArray:3 },
    dataLabels: { enabled:true, style:{...chartFont,fontSize:'10px'} },
    legend: { show:false }
  });
}

// ======================================================================
//  SEARCH
// ======================================================================
function setupSearch(d) {
  const input = document.getElementById('search-input');
  const results = document.getElementById('search-results');
  const allClients = d.allClientNames || [];

  input.addEventListener('input', () => {
    const q = input.value.trim().toLowerCase();
    if (q.length < 2) { results.classList.remove('active'); results.innerHTML=''; return; }

    const matches = allClients.filter(c =>
      (c.clientName||'').toLowerCase().includes(q) ||
      (c.dsoRemark||'').toLowerCase().includes(q) ||
      (c.nehaRemark||'').toLowerCase().includes(q) ||
      (c.area||'').toLowerCase().includes(q) ||
      (c.person||'').toLowerCase().includes(q)
    );

    if (!matches.length) { results.innerHTML='<div class="search-no-results">No results found for "'+input.value.trim()+'"</div>'; results.classList.add('active'); return; }

    results.innerHTML = matches.slice(0,20).map(m => {
      const sc = m.status==='Hot Lead'?'badge-hot':m.status==='Follow-up Needed'?'badge-followup':m.status==='Not Interested'?'badge-not-interested':m.status==='No Requirement'?'badge-no-req':m.status==='Already has Vendor'?'badge-vendor':'badge-unreachable';
      return '<div class="search-result-item"><div class="sr-client">'+m.clientName+'</div><div class="sr-meta"><span>'+m.person+'</span><span>'+(m.area||'')+'</span><span>'+(m.date||'')+'</span><span class="badge-status '+sc+'">'+m.status+'</span>'+(m.priority?'<span class="badge-status badge-priority badge-p'+m.priority+'">P'+m.priority+'</span>':'')+'</div>'+(m.dsoRemark?'<div class="sr-remark">DSO: '+m.dsoRemark+'</div>':'')+(m.nehaRemark?'<div class="sr-remark-neha">Neha: '+m.nehaRemark+'</div>':'')+'</div>';
    }).join('');
    results.classList.add('active');
  });

  document.addEventListener('click', e => { if (!e.target.closest('.search-bar')) results.classList.remove('active'); });
  input.addEventListener('focus', () => { if (input.value.trim().length >= 2) input.dispatchEvent(new Event('input')); });
}

// ======================================================================
//  FOLLOW-UPS
// ======================================================================
function renderFollowUps(d) {
  const el = document.getElementById('followup-section');
  if (!d.followUps.length) { el.innerHTML='<p style="color:var(--text-muted);padding:16px">No follow-ups needed!</p>'; return; }

  let html = '<div class="followup-grid">';
  d.followUps.slice(0,30).forEach(f => {
    const dateStr = normalizeDate(f.date);
    const shortDate = dateStr ? dateStr.split('-').slice(1).join('/') : '';
    const sc = f.status==='Hot Lead'?'badge-hot':'badge-followup';
    const pc = f.priority?'badge-p'+f.priority:'';
    html += '<div class="followup-card"><div class="client">'+(f.clientName||'Unknown')+'</div><div class="meta"><span>'+f.person+'</span><span>'+shortDate+'</span><span>'+(f.area||'')+'</span>'+(f.priority?'<span class="badge-status badge-priority '+pc+'">P'+f.priority+'</span>':'')+'<span class="badge-status '+sc+'">'+f.status+'</span></div>'+(f.actionItems?'<div class="action">'+f.actionItems+'</div>':'')+'<div class="remark">'+f.dsoRemark+'</div></div>';
  });
  html += '</div>';
  if (d.followUps.length > 30) html += '<p style="color:var(--text-muted);margin-top:12px;font-size:12px">Showing 30 of '+d.followUps.length+' follow-ups</p>';
  el.innerHTML = html;
}

// ======================================================================
//  NEHA SECTION
// ======================================================================
function renderNehaSection(d) {
  if (!d.nehaAnalysis || !d.nehaAnalysis.categories || !Object.keys(d.nehaAnalysis.categories).length) return;
  const section = document.getElementById('neha-section');
  const el = document.getElementById('neha-content');
  section.style.display = 'block';

  const catId = 'nc-'+Date.now(), dsoId = 'nd-'+Date.now(), tlId = 'nt-'+Date.now();

  el.innerHTML =
    '<div class="summary-grid" style="grid-template-columns:repeat(auto-fit,minmax(150px,1fr));margin-bottom:24px">'+
      card('','Total Remarks',d.nehaAnalysis.totalRemarks,"from Neha's file")+
      card('gold','Interested',d.nehaAnalysis.categories['Interested']||0,'')+
      card('','Quotation Asked',d.nehaAnalysis.categories['Quotation Requested']||0,'')+
      card('','Follow-up Pending',d.nehaAnalysis.categories['Follow-up Pending']||0,'')+
      card('red','Not Interested',d.nehaAnalysis.categories['Not Interested']||0,'')+
    '</div>'+
    '<div class="charts-row"><div class="chart-card"><h3>Neha\'s Remark Categories</h3><div id="'+catId+'"></div></div><div class="chart-card"><h3>DSO Response Status</h3><div id="'+dsoId+'"></div></div></div>'+
    '<div class="charts-row"><div class="chart-card full"><h3>Daily Remarks Volume</h3><div id="'+tlId+'"></div></div></div>';

  setTimeout(() => {
    const catEntries = Object.entries(d.nehaAnalysis.categories);
    const catColors = ['#f3c623','#3b82f6','#8b5cf6','#16a34a','#6b7280','#06b6d4','#f59e0b','#ec4899'];
    makeChart(document.getElementById(catId), {
      chart: { type:'pie', height:280, ...chartFont, background:'transparent', toolbar:{show:false} },
      series: catEntries.map(e=>e[1]), labels: catEntries.map(e=>e[0]), colors: catColors.slice(0,catEntries.length),
      legend: { position:'bottom', labels:{colors:labelColor,...chartFont,fontSize:'11px'} },
      dataLabels: { enabled:true, formatter:v=>v.toFixed(0)+'%', style:{...chartFont,fontSize:'11px'} }
    });

    const dsoEntries = Object.entries(d.nehaAnalysis.dsoResponses).sort((a,b)=>b[1]-a[1]);
    makeChart(document.getElementById(dsoId), {
      chart: { type:'bar', height:280, ...chartFont, background:'transparent', toolbar:{show:false} },
      series: [{ name:'Leads', data:dsoEntries.map(e=>e[1]) }],
      colors: ['#3b82f6'],
      plotOptions: { bar: { horizontal:true, borderRadius:4, barHeight:'55%' } },
      xaxis: { categories:dsoEntries.map(e=>e[0]), labels:{style:{colors:labelColor,...chartFont,fontSize:'10px'}} },
      yaxis: { labels:{style:{colors:labelColor,...chartFont,fontSize:'10px'}} },
      grid: { borderColor:gridColor, strokeDashArray:3 },
      dataLabels: { enabled:true, style:{...chartFont,fontSize:'10px'} },
      legend: { show:false }
    });

    const dates = Object.keys(d.nehaAnalysis.nehaTimeline).sort();
    if (dates.length) {
      const shortDates = dates.map(dd=>{const p=dd.split('-');return p[2]+'/'+p[1];});
      makeChart(document.getElementById(tlId), {
        chart: { type:'bar', height:220, ...chartFont, background:'transparent', toolbar:{show:false} },
        series: [{ name:'Remarks', data:dates.map(dd=>d.nehaAnalysis.nehaTimeline[dd]) }],
        colors: ['#8b5cf6'],
        plotOptions: { bar: { borderRadius:4, columnWidth:'50%' } },
        xaxis: { categories:shortDates, labels:{style:{colors:labelColor,...chartFont,fontSize:'11px'}} },
        yaxis: { labels:{style:{colors:labelColor,...chartFont,fontSize:'11px'}}, min:0 },
        grid: { borderColor:gridColor, strokeDashArray:3 },
        dataLabels: { enabled:false },
        legend: { show:false }
      });
    }
  }, 100);
}

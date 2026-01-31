// Telecom Relatórios - Excel filter + report (SheetJS)
// Padrão: abas por dia (nome como 29042025 etc), tabela com cabeçalho "PERÍODO" e linhas de ocorrência "OCORRÊNCIAS DO DIA : ..."

const els = {
  fileInput: document.getElementById('fileInput'),
  exportPdfBtn: document.getElementById('exportPdfBtn'),
  exportXlsxBtn: document.getElementById('exportXlsxBtn'),
  daySelect: document.getElementById('daySelect'),
  techSelect: document.getElementById('techSelect'),
  motivoSelect: document.getElementById('motivoSelect'),
  periodoSelect: document.getElementById('periodoSelect'),
  searchInput: document.getElementById('searchInput'),
  servicesTable: document.getElementById('servicesTable'),
  reportTable: document.getElementById('reportTable'),
  occList: document.getElementById('occList'),
  hint: document.getElementById('hint'),
  kpis: document.getElementById('kpis'),
};

let WB = null;
let CACHE = {}; // {sheetName: {services:[], occs:[], indexByClientKey:{...}}}

const normalize = (v) => {
  if (v === null || v === undefined) return '';
  return String(v)
    .normalize('NFD').replace(/\p{Diacritic}/gu, '')
    .toLowerCase()
    .replace(/\s+/g, ' ')
    .trim();
};

const safeStr = (v) => (v === null || v === undefined) ? '' : String(v).trim();

// Converte 'ddmmaaaa' -> 'dd/mm/aaaa' quando fizer sentido
function formatBRDate(v){
  const s = safeStr(v);
  if(/^\d{8}$/.test(s)){
    const dd = s.slice(0,2), mm = s.slice(2,4), yyyy = s.slice(4,8);
    return `${dd}/${mm}/${yyyy}`;
  }
  const m = s.match(/^(\d{2})[\-\.](\d{2})[\-\.](\d{4})$/);
  if(m) return `${m[1]}/${m[2]}/${m[3]}`;
  return s;
}

function setDisabled(disabled){
  els.daySelect.disabled = disabled;
  els.techSelect.disabled = disabled;
  els.motivoSelect.disabled = disabled;
  els.periodoSelect.disabled = disabled;
  els.searchInput.disabled = disabled;
  els.exportPdfBtn.disabled = disabled;
  els.exportXlsxBtn.disabled = disabled;
}

function isDateSheet(name){
  return /^\s*\d{8}\s*$/.test(name);
}

function sheetToMatrix(ws){
  return XLSX.utils.sheet_to_json(ws, { header: 1, raw: false, defval: '' });
}

function findHeaderRow(matrix){
  for(let r=0; r<matrix.length; r++){
    for(let c=0; c<Math.min(matrix[r].length, 10); c++){
      if (normalize(matrix[r][c]) === 'periodo') return r;
    }
  }
  return -1;
}

function parseSheet(sheetName){
  if(CACHE[sheetName]) return CACHE[sheetName];

  const ws = WB.Sheets[sheetName];
  const matrix = sheetToMatrix(ws);
  const headerRow = findHeaderRow(matrix);

  if(headerRow === -1){
    CACHE[sheetName] = { services: [], occs: [], indexByClientKey: {} };
    return CACHE[sheetName];
  }

  const headers = matrix[headerRow].map(h => normalize(h));
  const col = (name) => headers.indexOf(normalize(name));

  const idx = {
    periodo: col('PERÍODO'),
    confirmacoes: col('CONFIRMAÇÕES'),
    motivo: col('MOTIVO'),
    tecnico: col('TECNICO'),
    nome: col('NOME'),
    endereco: col('ENDEREÇO'),
    telefone: col('TELEFONE'),
    cpf: col('CPF'),
    rg: col('RG'),
    dtnasc: col('DT.NASC'),
    plano: col('PLANO'),
    vencimento: col('VENCIMENTO'),
    taxa: col('TAXA R$'),
    pagto: col('PAGTO'),
    boleto: col('BOLETO'),
    login: col('LOGIN/SENHA'),
    atendente: col('ATENDENTE'),
    obs: col('OBSERVAÇÃO'),
  };

  const services = [];
  const occs = [];

  for(let r=headerRow+1; r<matrix.length; r++){
    const a = safeStr(matrix[r][0]);
    const aNorm = normalize(a);

    if(aNorm.startsWith('ocorrencias do dia')) break;

    const rowIsEmpty = matrix[r].every(v => normalize(v) === '');
    if(rowIsEmpty) continue;

    const nome = safeStr(matrix[r][idx.nome]);
    const tecnico = safeStr(matrix[r][idx.tecnico]);
    if(!nome && !tecnico) continue;

    services.push({
      sheet: sheetName,
      periodo: safeStr(matrix[r][idx.periodo]),
      confirmacoes: safeStr(matrix[r][idx.confirmacoes]),
      motivo: safeStr(matrix[r][idx.motivo]),
      tecnico,
      nome,
      endereco: safeStr(matrix[r][idx.endereco]),
      telefone: safeStr(matrix[r][idx.telefone]),
      cpf: safeStr(matrix[r][idx.cpf]),
      rg: safeStr(matrix[r][idx.rg]),
      dtnasc: safeStr(matrix[r][idx.dtnasc]),
      plano: safeStr(matrix[r][idx.plano]),
      vencimento: safeStr(matrix[r][idx.vencimento]),
      taxa: safeStr(matrix[r][idx.taxa]),
      pagto: safeStr(matrix[r][idx.pagto]),
      boleto: safeStr(matrix[r][idx.boleto]),
      login: safeStr(matrix[r][idx.login]),
      atendente: safeStr(matrix[r][idx.atendente]),
      obs: safeStr(matrix[r][idx.obs]),
    });
  }

  for(let r=0; r<matrix.length; r++){
    const a = safeStr(matrix[r][0]);
    const aNorm = normalize(a);
    if(!aNorm.startsWith('ocorrencias do dia')) continue;

    const text = a.split(':').slice(1).join(':').trim();
    if(!text) continue;

    let client = text;
    const cutTokens = [' pediu ', '----', ' - ', ' reagend', ' reagenda', ' cliente ', ' nao ', ' não '];
    for(const t of cutTokens){
      const pos = normalize(client).indexOf(normalize(t));
      if(pos > 0){ client = client.slice(0, pos).trim(); break; }
    }
    const words = client.split(/\s+/).filter(Boolean);
    if(words.length > 7) client = words.slice(0, 7).join(' ');

    occs.push({
      sheet: sheetName,
      clientGuess: client,
      clientKey: normalize(client),
      text,
      raw: a
    });
  }

  const indexByClientKey = {};
  for(const o of occs){
    if(!indexByClientKey[o.clientKey]) indexByClientKey[o.clientKey] = o;
  }

  CACHE[sheetName] = { services, occs, indexByClientKey };
  return CACHE[sheetName];
}

function buildSelect(select, items, placeholder='Todos'){
  // items pode ser array de strings OU array de {value,label}
  select.innerHTML = '';
  const opt0 = document.createElement('option');
  opt0.value = '';
  opt0.textContent = placeholder;
  select.appendChild(opt0);

  for(const it of items){
    const opt = document.createElement('option');
    if(typeof it === 'object' && it){
      opt.value = it.value ?? '';
      opt.textContent = it.label ?? (it.value ?? '');
    } else {
      opt.value = it;
      opt.textContent = it;
    }
    select.appendChild(opt);
  }
}

function uniqSorted(arr){
  return Array.from(new Set(arr.filter(Boolean).map(v => String(v).trim())))
    .sort((a,b)=>a.localeCompare(b,'pt-BR',{sensitivity:'base'}));
}

function getFilters(){
  return {
    day: els.daySelect.value,
    tech: els.techSelect.value,
    motivo: els.motivoSelect.value,
    periodo: els.periodoSelect.value,
    q: normalize(els.searchInput.value || ''),
  };
}

function matchOccurrenceForService(service, indexByClientKey){
  const key = normalize(service.nome);
  if(indexByClientKey[key]) return indexByClientKey[key];

  const keys = Object.keys(indexByClientKey);
  for(const k of keys){
    if(!k) continue;
    if(key.includes(k) || k.includes(key)) return indexByClientKey[k];
  }
  return null;
}

function computeView(){
  const f = getFilters();
  if(!f.day) return { rows: [], occs: [], report: [], kpi: null };

  const { services, occs, indexByClientKey } = parseSheet(f.day);

  const rows = services
    .map(s => {
      const occ = matchOccurrenceForService(s, indexByClientKey);
      const status = occ ? 'nao_concluido' : 'concluido';
      return { ...s, status, occText: occ ? occ.text : '' };
    })
    .filter(s => !f.tech || s.tecnico === f.tech)
    .filter(s => !f.motivo || normalize(s.motivo) === normalize(f.motivo))
    .filter(s => !f.periodo || normalize(s.periodo) === normalize(f.periodo))
    .filter(s => !f.q || normalize(s.nome + ' ' + s.endereco).includes(f.q));

  const byTech = {};
  for(const r of rows){
    const t = r.tecnico || 'SEM TÉCNICO';
    if(!byTech[t]) byTech[t] = { tecnico: t, total: 0, concluidos: 0, nao_concluidos: 0 };
    byTech[t].total++;
    if(r.status === 'concluido') byTech[t].concluidos++;
    else byTech[t].nao_concluidos++;
  }
  const report = Object.values(byTech).map(x => ({
    ...x,
    perc: x.total ? Math.round((x.concluidos/x.total)*100) : 0
  })).sort((a,b)=>b.total-a.total);

  const total = rows.length;
  const concluidos = rows.filter(r=>r.status==='concluido').length;
  const nao = total - concluidos;
  const perc = total ? Math.round((concluidos/total)*100) : 0;

  const visibleClientKeys = new Set(rows.map(r=>normalize(r.nome)));
  const occsView = occs.filter(o=>{
    if(!o.clientKey) return true;
    for(const k of visibleClientKeys){
      if(k.includes(o.clientKey) || o.clientKey.includes(k)) return true;
    }
    return false;
  });

  return { rows, occs: occsView, report, kpi: { total, concluidos, nao, perc } };
}

function render(){
  const view = computeView();

  const kpiEls = els.kpis.querySelectorAll('.kpiValue');
  if(view.kpi){
    kpiEls[0].textContent = view.kpi.total;
    kpiEls[1].textContent = view.kpi.concluidos;
    kpiEls[2].textContent = view.kpi.nao;
    kpiEls[3].textContent = view.kpi.perc + '%';
  } else {
    kpiEls.forEach(el=>el.textContent='—');
  }

  const head = els.servicesTable.querySelector('thead');
  const body = els.servicesTable.querySelector('tbody');
  head.innerHTML = '<tr>' + [
    'Período','Técnico','Motivo','Cliente','Endereço','Telefone','Status','Motivo (ocorrência)'
  ].map(h=>`<th>${h}</th>`).join('') + '</tr>';

  body.innerHTML = '';
  for(const r of view.rows){
    const pill = r.status === 'concluido'
      ? '<span class="pill good">Concluído</span>'
      : '<span class="pill bad">Não concluído</span>';

    const occCell = r.occText
      ? `<a class="link" href="#" data-occ="${encodeURIComponent(r.occText)}">Abrir</a>`
      : '<span class="pill warn">—</span>';

    const tr = document.createElement('tr');
    tr.innerHTML = `
      <td>${safeStr(r.periodo)}</td>
      <td>${safeStr(r.tecnico)}</td>
      <td>${safeStr(r.motivo)}</td>
      <td>${safeStr(r.nome)}</td>
      <td>${safeStr(r.endereco)}</td>
      <td>${safeStr(r.telefone)}</td>
      <td>${pill}</td>
      <td>${occCell}</td>
    `;
    body.appendChild(tr);
  }

  body.querySelectorAll('a[data-occ]').forEach(a=>{
    a.addEventListener('click', (e)=>{
      e.preventDefault();
      const text = decodeURIComponent(a.getAttribute('data-occ') || '');
      showOccModal(text);
    });
  });

  els.occList.innerHTML = '';
  if(view.occs.length === 0){
    const li = document.createElement('li');
    li.className = 'occItem';
    li.innerHTML = '<div class="occTitle">Sem ocorrências</div><div class="occText">Nada registrado para esse filtro/dia.</div>';
    els.occList.appendChild(li);
  } else {
    for(const o of view.occs){
      const li = document.createElement('li');
      li.className = 'occItem';
      li.innerHTML = `
        <div class="occTitle">${safeStr(o.clientGuess || 'Ocorrência')}</div>
        <div class="occText">${safeStr(o.text)}</div>
      `;
      els.occList.appendChild(li);
    }
  }

  const rHead = els.reportTable.querySelector('thead');
  const rBody = els.reportTable.querySelector('tbody');
  rHead.innerHTML = '<tr>' + ['Técnico','Total','Concluídos','Não concluídos','%'].map(h=>`<th>${h}</th>`).join('') + '</tr>';
  rBody.innerHTML = '';
  for(const x of view.report){
    const tr = document.createElement('tr');
    tr.innerHTML = `
      <td>${safeStr(x.tecnico)}</td>
      <td>${x.total}</td>
      <td>${x.concluidos}</td>
      <td>${x.nao_concluidos}</td>
      <td>${x.perc}%</td>
    `;
    rBody.appendChild(tr);
  }

  els.exportPdfBtn.disabled = !(WB && els.daySelect.value);
  els.exportXlsxBtn.disabled = !(WB && els.daySelect.value);
}

function showOccModal(text){
  let modal = document.getElementById('occModal');
  if(!modal){
    modal = document.createElement('div');
    modal.id = 'occModal';
    modal.style.position='fixed';
    modal.style.inset='0';
    modal.style.background='rgba(0,0,0,.55)';
    modal.style.display='grid';
    modal.style.placeItems='center';
    modal.style.padding='14px';
    modal.style.zIndex='1000';
    modal.innerHTML = `
      <div style="max-width:720px;width:100%; background: rgba(18,26,51,.98); border:1px solid rgba(255,255,255,.10); border-radius:18px; box-shadow: 0 18px 45px rgba(0,0,0,.5); padding:14px;">
        <div style="display:flex;justify-content:space-between;align-items:center;gap:10px;">
          <div style="font-weight:900;">Detalhe da ocorrência</div>
          <button id="occClose" class="btn secondary" style="padding:8px 10px;border-radius:12px;">Fechar</button>
        </div>
        <div id="occText" style="margin-top:10px;color:#9fb0da;line-height:1.45; white-space:pre-wrap;"></div>
      </div>
    `;
    document.body.appendChild(modal);
    modal.addEventListener('click', (e)=>{ if(e.target === modal) closeOccModal(); });
    modal.querySelector('#occClose').addEventListener('click', closeOccModal);
  }
  modal.querySelector('#occText').textContent = text || '';
  modal.style.display='grid';
}
function closeOccModal(){
  const modal = document.getElementById('occModal');
  if(modal) modal.style.display='none';
}

function exportXlsx(){
  const view = computeView();
  if(!view.kpi) return;

  const dayRaw = els.daySelect.value.trim();
  const day = formatBRDate(dayRaw);

  const rows = view.rows.map(r=>({
    Dia: day,
    Periodo: r.periodo,
    Tecnico: r.tecnico,
    Motivo: r.motivo,
    Cliente: r.nome,
    Endereco: r.endereco,
    Telefone: r.telefone,
    Status: (r.status==='concluido'?'Concluído':'Não concluído'),
    Ocorrencia: r.occText
  }));

  const report = view.report.map(x=>({
    Dia: day,
    Tecnico: x.tecnico,
    Total: x.total,
    Concluidos: x.concluidos,
    NaoConcluidos: x.nao_concluidos,
    Percentual: x.perc + '%'
  }));

  const wb = XLSX.utils.book_new();
  const ws1 = XLSX.utils.json_to_sheet(rows);
  const ws2 = XLSX.utils.json_to_sheet(report);

  XLSX.utils.book_append_sheet(wb, ws1, 'SERVICOS_FILTRADOS');
  XLSX.utils.book_append_sheet(wb, ws2, 'RELATORIO_TECNICOS');

  XLSX.writeFile(wb, `relatorio_${dayRaw}.xlsx`);
}

function exportPDF(){
  const view = computeView();
  if(!view.kpi) return;

  const dayRaw = els.daySelect.value.trim();
  const day = formatBRDate(dayRaw);

  const { jsPDF } = window.jspdf;
  const doc = new jsPDF({ unit: 'pt', format: 'a4' });

  doc.setFont('helvetica','bold');
  doc.setFontSize(16);
  doc.text('Telecom Relatórios', 40, 52);

  doc.setFont('helvetica','normal');
  doc.setFontSize(11);
  doc.text(`Relatório do dia: ${day}`, 40, 74);

  const k = view.kpi;
  doc.text(`Total: ${k.total}   Concluídos: ${k.concluidos}   Não concluídos: ${k.nao}   % Conclusão: ${k.perc}%`, 40, 96);

  doc.setFont('helvetica','bold');
  doc.text('Resumo por técnico', 40, 126);

  const reportRows = view.report.map(x => [x.tecnico, String(x.total), String(x.concluidos), String(x.nao_concluidos), String(x.perc) + '%']);
  doc.autoTable({
    startY: 140,
    head: [['Técnico','Total','Concluídos','Não concluídos','%']],
    body: reportRows,
    styles: { font: 'helvetica', fontSize: 10, cellPadding: 6 },
    headStyles: { fillColor: [24, 34, 66] },
    margin: { left: 40, right: 40 }
  });

  const start = doc.lastAutoTable.finalY + 18;
  doc.setFont('helvetica','bold');
  doc.text('Serviços (filtrados)', 40, start);

  const svcRows = view.rows.slice(0, 180).map(r => ([
    safeStr(r.periodo),
    safeStr(r.tecnico),
    safeStr(r.motivo),
    safeStr(r.nome),
    (r.status==='concluido'?'Concluído':'Não concluído')
  ]));

  doc.autoTable({
    startY: start + 14,
    head: [['Período','Técnico','Motivo','Cliente','Status']],
    body: svcRows,
    styles: { font: 'helvetica', fontSize: 9, cellPadding: 5 },
    headStyles: { fillColor: [24, 34, 66] },
    margin: { left: 40, right: 40 },
    didDrawPage: () => {
      const page = doc.getNumberOfPages();
      doc.setFontSize(9);
      doc.setFont('helvetica','normal');
      doc.text(`Página ${page}`, doc.internal.pageSize.getWidth()-80, doc.internal.pageSize.getHeight()-24);
    }
  });

  doc.save(`relatorio_${dayRaw}.pdf`);
}

els.exportPdfBtn.addEventListener('click', exportPDF);
els.exportXlsxBtn.addEventListener('click', exportXlsx);

function wireFilters(){
  ['daySelect','techSelect','motivoSelect','periodoSelect'].forEach(id=>{
    els[id].addEventListener('change', ()=>{
      if(id==='daySelect') refreshFiltersForDay();
      render();
    });
  });
  els.searchInput.addEventListener('input', ()=>render());
}

function refreshFiltersForDay(){
  const day = els.daySelect.value;
  if(!day){ return; }
  const { services } = parseSheet(day);

  buildSelect(els.techSelect, uniqSorted(services.map(s=>s.tecnico)), 'Todos os técnicos');
  buildSelect(els.motivoSelect, uniqSorted(services.map(s=>s.motivo)), 'Todos os motivos');
  buildSelect(els.periodoSelect, uniqSorted(services.map(s=>s.periodo)), 'Todos os períodos');
  els.searchInput.value = '';
}

els.fileInput.addEventListener('change', async (e)=>{
  const file = e.target.files?.[0];
  if(!file) return;

  const buf = await file.arrayBuffer();
  WB = XLSX.read(buf, { type: 'array' });
  CACHE = {};

  const days = WB.SheetNames.filter(isDateSheet).map(s=>s.trim());
  days.sort((a,b)=>a.localeCompare(b));

  if(days.length === 0){
    els.hint.textContent = 'Não encontrei abas de dia (ex: 29042025). Verifique o padrão do arquivo.';
    setDisabled(true);
    return;
  }

  els.hint.textContent = '';
  buildSelect(els.daySelect, days.map(d=>({value:d,label:formatBRDate(d)})), 'Selecione o dia');
  setDisabled(false);

  els.daySelect.value = days[0];
  refreshFiltersForDay();
  render();
});

wireFilters();
setDisabled(true);
render();

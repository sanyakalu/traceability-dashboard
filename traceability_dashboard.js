const TEMPLATE_URL = 'Traceability_matrix_doc_template.docx';
let ALL = [], templateBuffer = null;

// ── Column definitions ────────────────────────────────────────────────────────
// level: 'req'  = rowspan by (Requirement ID + Released in)
//        'vers' = rowspan by Reqmt for Version (within req group)
//        'risk' = rowspan by Risk ID (within vers group)
//        'leaf' = one cell per row, no rowspan
const COLS = [
  // Requirements
  { key: 'Requirement ID',    label: 'Req ID',       group: 'req',   level: 'req',  sticky: true, urlKey: 'Requirement URL' },
  { key: 'Requirement Name',  label: 'Req Name',     group: 'req',   level: 'req',  sticky: true },
  { key: 'Feature Spec',      label: 'Feature Spec', group: 'req',   level: 'req' },
  { key: 'Released in',       label: 'Released in',  group: 'req',   level: 'req' },
  { key: 'Reqmt for Version', label: 'PIC Version',  group: 'req',   level: 'vers' },
  // Test Cases
  { key: 'Test Case ID',      label: 'TC ID',        group: 'tc',    level: 'leaf', urlKey: 'Test Case URL' },
  { key: 'Test Case Name',    label: 'TC Name',      group: 'tc',    level: 'leaf' },
  { key: 'Test Case State',   label: 'TC State',     group: 'tc',    level: 'leaf' },
  { key: 'Relevant Product',  label: 'Product',      group: 'tc',    level: 'leaf' },
  { key: 'Test Case Release', label: 'TC Release',   group: 'tc',    level: 'leaf' },
  { key: 'Test Point ID',     label: 'Point ID',     group: 'tc',    level: 'leaf' },
  { key: 'Test Plan ID',      label: 'Plan ID',      group: 'tc',    level: 'leaf', urlKey: 'Test Plan URL' },
  { key: 'Test Suite Name',   label: 'Suite Name',   group: 'tc',    level: 'leaf' },
  { key: 'Test Suite ID',     label: 'Suite ID',     group: 'tc',    level: 'leaf', urlKey: 'Test Suite URL' },
  { key: 'Test Plan Tags',    label: 'Plan Tags',    group: 'tc',    level: 'leaf' },
  { key: 'Test Result',       label: 'Result',       group: 'tc',    level: 'leaf' },
  { key: 'Test Result Date',  label: 'Result Date',  group: 'tc',    level: 'leaf' },
  // Risks
  { key: 'Risk ID',           label: 'Risk ID',      group: 'risk',  level: 'risk', urlKey: 'Risk URL' },
  { key: 'Risk Name',         label: 'Risk Name',    group: 'risk',  level: 'risk' },
  // STICRs
  { key: 'STICR ID',          label: 'STICR ID',      group: 'sticr', level: 'leaf', urlKey: 'STICR URL' },
  { key: 'STICR Name',        label: 'STICR Name',    group: 'sticr', level: 'leaf' },
  { key: 'STICR State',       label: 'STICR State',   group: 'sticr', level: 'leaf' },
  { key: 'STICR Release',     label: 'STICR Release', group: 'sticr', level: 'leaf' },
];

const GROUP_META = {
  req:   { label: 'Requirements', cls: 'g-req'   },
  tc:    { label: 'Test Cases',   cls: 'g-tc'    },
  risk:  { label: 'Risks',        cls: 'g-risk'  },
  sticr: { label: 'STICRs',       cls: 'g-sticr' },
};

// ── Utilities ─────────────────────────────────────────────────────────────────
function $(id) { return document.getElementById(id); }

function esc(s) {
  return String(s ?? '')
    .replace(/&/g, '&amp;').replace(/</g, '&lt;')
    .replace(/>/g, '&gt;').replace(/"/g, '&quot;');
}

function val(row, key) {
  const v = row[key];
  if (v === undefined || v === null) return '';
  const s = String(v).trim();
  if (s === '' || s.toLowerCase() === 'nan') return '';
  return s.replace(/^(-?\d+)\.0$/, '$1');
}

function cellHtml(row, col) {
  const v = val(row, col.key);
  if (!v) return '';
  if (col.urlKey) {
    const u = val(row, col.urlKey);
    if (u) return `<a class="lnk" href="${esc(u)}" target="_blank">${esc(v)} &#128279;</a>`;
  }
  if (col.key === 'Test Result') {
    const lv = v.toLowerCase();
    if (lv === 'passed')  return `<span class="badge b-passed">Passed</span>`;
    if (lv === 'failed')  return `<span class="badge b-failed">Failed</span>`;
    if (lv === 'blocked') return `<span class="badge b-blocked">Blocked</span>`;
    if (lv === 'none')    return `<span class="badge b-none">None</span>`;
    return `<span class="badge b-other">${esc(v)}</span>`;
  }
  return esc(v);
}

// ── PIC version checklist ─────────────────────────────────────────────────────
const pvSelected = new Set();
let pvVersions = [];

function pvUpdateLabel() {
  const btn = $('pv-btn');
  if      (pvSelected.size === 0) { btn.textContent = 'PIC Vers (all) ▾'; btn.classList.remove('active'); }
  else if (pvSelected.size === 1) { btn.textContent = [...pvSelected][0] + ' ▾'; btn.classList.add('active'); }
  else                            { btn.textContent = pvSelected.size + ' versions ▾'; btn.classList.add('active'); }
}

function pvRebuildPanel() {
  const panel = $('pv-panel');
  panel.innerHTML = '';
  const allLbl = document.createElement('label'); allLbl.className = 'all-opt';
  const allCb  = document.createElement('input'); allCb.type = 'checkbox'; allCb.checked = pvSelected.size === 0;
  allCb.addEventListener('change', () => { pvSelected.clear(); pvRebuildPanel(); pvUpdateLabel(); render(); });
  allLbl.appendChild(allCb); allLbl.appendChild(document.createTextNode(' All versions'));
  panel.appendChild(allLbl);
  for (const v of pvVersions) {
    const lbl = document.createElement('label');
    const cb  = document.createElement('input'); cb.type = 'checkbox'; cb.checked = pvSelected.has(v);
    cb.addEventListener('change', e => {
      e.target.checked ? pvSelected.add(v) : pvSelected.delete(v);
      pvRebuildPanel(); pvUpdateLabel(); render();
    });
    lbl.appendChild(cb); lbl.appendChild(document.createTextNode(' ' + v));
    panel.appendChild(lbl);
  }
}

$('pv-btn').addEventListener('click', e => { e.stopPropagation(); $('pv-panel').classList.toggle('open'); });
$('pv-panel').addEventListener('click', e => e.stopPropagation());
document.addEventListener('click', () => $('pv-panel').classList.remove('open'));

// ── Filtering ─────────────────────────────────────────────────────────────────
function getFilters() {
  return {
    fs:       $('f-fs').value,
    pv:       [...pvSelected],
    pt:       $('f-pt').value.toLowerCase(),
    lr:       $('f-lr').value.toLowerCase(),
    ri:       $('f-ri').value.trim().toLowerCase(),
    tc:       $('f-tc').value.trim().toLowerCase(),
    pi:       $('f-pi').value.trim().toLowerCase(),
    si:       $('f-si').value.trim().toLowerCase(),
    hasRisk:  $('f-has-risk').checked,
    hasSticr: $('f-has-sticr').checked,
  };
}

function rowMatches(row, f) {
  if (f.fs && val(row, 'Feature Spec') !== f.fs) return false;
  if (f.pv.length > 0 && !f.pv.includes(val(row, 'Reqmt for Version'))) return false;
  if (f.pt && !val(row, 'Test Plan Tags').toLowerCase().includes(f.pt)) return false;
  if (f.lr) {
    const lr = val(row, 'Test Result').toLowerCase();
    if (f.lr === 'none') { if (lr !== '' && lr !== 'none') return false; }
    else                 { if (lr !== f.lr) return false; }
  }
  if (f.ri && !val(row, 'Requirement ID').toLowerCase().includes(f.ri))        return false;
  if (f.tc && !val(row, 'Test Case ID').toLowerCase().includes(f.tc))           return false;
  if (f.pi && !String(val(row, 'Test Plan ID')).toLowerCase().includes(f.pi))  return false;
  if (f.si && !String(val(row, 'Test Suite ID')).toLowerCase().includes(f.si)) return false;
  if (f.hasRisk  && !val(row, 'Risk ID'))  return false;
  if (f.hasSticr && !val(row, 'STICR ID')) return false;
  return true;
}

// ── Latest test case dedup ────────────────────────────────────────────────────

// Pick the best row from a cluster: most recent Passed row, else most recent overall.
function bestInCluster(rows) {
  const passing = rows.filter(r => val(r, 'Test Result').toLowerCase() === 'passed');
  const pool    = passing.length > 0 ? passing : rows;
  return pool.reduce((a, b) => {
    const da = val(a, 'Test Result Date') || '0000-00-00';
    const db = val(b, 'Test Result Date') || '0000-00-00';
    return db > da ? b : a;
  });
}

// Within each (Req ID + Released in + Reqmt for Version + Risk ID) group:
//   1. Rows with the same TC ID are treated as the same test case.
//   2. Rows with different TC IDs but identical TC Name are treated as the same test case.
//   3. From each cluster keep the most recent Passed row (fallback: most recent).
function dedupeLatestTC(data) {
  const groups = new Map();
  for (const row of data) {
    const k = [
      val(row, 'Requirement ID'),
      val(row, 'Released in'),
      val(row, 'Reqmt for Version'),
      val(row, 'Risk ID'),
    ].join('|||');
    if (!groups.has(k)) groups.set(k, []);
    groups.get(k).push(row);
  }

  const result = [];
  for (const rows of groups.values()) {
    if (rows.length === 1) { result.push(rows[0]); continue; }

    // Greedy clustering: seed each cluster from the first unassigned row.
    const assigned = new Array(rows.length).fill(false);
    for (let i = 0; i < rows.length; i++) {
      if (assigned[i]) continue;
      const cluster = [rows[i]];
      assigned[i] = true;
      const seedId   = val(rows[i], 'Test Case ID');
      const seedName = val(rows[i], 'Test Case Name');

      for (let j = i + 1; j < rows.length; j++) {
        if (assigned[j]) continue;
        const sameId   = seedId && seedId === val(rows[j], 'Test Case ID');
        const sameName = !sameId && seedName !== '' && seedName === val(rows[j], 'Test Case Name');
        if (sameId || sameName) {
          cluster.push(rows[j]);
          assigned[j] = true;
        }
      }

      result.push(bestInCluster(cluster));
    }
  }

  return result;
}

// ── Grouping ──────────────────────────────────────────────────────────────────
function buildGroups(data) {
  const sorted = [...data].sort((a, b) => {
    for (const k of ['Requirement ID', 'Released in', 'Reqmt for Version', 'Risk ID', 'Test Case ID', 'Test Point ID', 'STICR ID']) {
      const c = String(val(a, k) || '').localeCompare(String(val(b, k) || ''), undefined, { numeric: true });
      if (c !== 0) return c;
    }
    return 0;
  });

  const r1Map = new Map(), r1Order = [];
  for (const row of sorted) {
    const r1k = `${val(row, 'Requirement ID')}|||${val(row, 'Released in')}`;
    const vk  = val(row, 'Reqmt for Version') || '';
    const rk  = val(row, 'Risk ID') || '';
    if (!r1Map.has(r1k)) { r1Map.set(r1k, { vMap: new Map(), vOrder: [] }); r1Order.push(r1k); }
    const r1 = r1Map.get(r1k);
    if (!r1.vMap.has(vk)) { r1.vMap.set(vk, { rkMap: new Map(), rkOrder: [] }); r1.vOrder.push(vk); }
    const v = r1.vMap.get(vk);
    if (!v.rkMap.has(rk)) { v.rkMap.set(rk, []); v.rkOrder.push(rk); }
    v.rkMap.get(rk).push(row);
  }
  return { r1Map, r1Order };
}

// ── Sticky (vertical header rows + horizontal first 2 cols) ───────────────────
function applySticky() {
  const table = document.querySelector('#tbl-host table');
  if (!table) return;

  const hGrp  = table.querySelector('tr.h-group');
  const hCols = table.querySelector('tr.h-cols');
  if (!hCols) return;

  // ── Vertical sticky: both header rows ──
  const grpH = hGrp ? hGrp.getBoundingClientRect().height : 0;

  if (hGrp) {
    hGrp.querySelectorAll('th').forEach(th => {
      th.style.position = 'sticky';
      th.style.top      = '0px';
      th.style.zIndex   = '4';
    });
  }
  hCols.querySelectorAll('th').forEach(th => {
    th.style.position = 'sticky';
    th.style.top      = `${grpH}px`;
    th.style.zIndex   = '3';
  });

  // ── Horizontal sticky: first 2 column-name header cells ──
  const colThs = Array.from(hCols.querySelectorAll('th'));
  const w0     = colThs[0] ? colThs[0].getBoundingClientRect().width : 0;
  const leftOf = [0, w0];

  for (let i = 0; i < Math.min(2, colThs.length); i++) {
    colThs[i].style.left   = `${leftOf[i]}px`;
    colThs[i].style.zIndex = '6'; // above other sticky th cells
  }
  if (colThs[1]) colThs[1].style.boxShadow = '4px 0 8px -2px rgba(56,61,59,.18)';

  // ── Horizontal sticky: g-req group header band ──
  if (hGrp) {
    const gReq = hGrp.querySelector('th.g-req');
    if (gReq) {
      gReq.style.position = 'sticky';
      gReq.style.left     = '0px';
      gReq.style.zIndex   = '6';
    }
  }

  // ── Horizontal sticky: body c-req cells (Req ID col, then Req Name col) ──
  // There are exactly 2 c-req cells per req group (the two sticky cols).
  table.querySelectorAll('td.c-req').forEach((td, idx) => {
    const col = idx % 2;
    td.style.position = 'sticky';
    td.style.left     = `${leftOf[col]}px`;
    td.style.zIndex   = '2';
    if (col === 1) td.style.boxShadow = '4px 0 8px -2px rgba(56,61,59,.10)';
  });
}

// ── Render ────────────────────────────────────────────────────────────────────
function render() {
  const f      = getFilters();
  let   data   = ALL.filter(r => rowMatches(r, f));
  const latest = $('f-latest').checked;

  if (latest) data = dedupeLatestTC(data);

  ['f-fs','f-pt','f-lr'].forEach(id => $(id).classList.toggle('active', !!$(id).value));
  ['f-ri','f-tc','f-pi','f-si'].forEach(id => $(id).classList.toggle('active', $(id).value.trim() !== ''));
  $('f-latest').closest('label').classList.toggle('active', latest);
  ['f-has-risk','f-has-sticr'].forEach(id => $(id).closest('label').classList.toggle('active', $(id).checked));

  const host = $('tbl-host');
  if (!data.length) {
    const msg = ALL.length
      ? 'No rows match the current filters.'
      : 'Load <strong>traceability_matrix_alarm_panel.csv</strong> using the button above.';
    host.innerHTML = `<div class="empty"><div class="empty-icon">&#128203;</div><p>${msg}</p></div>`;
    $('result-count').textContent = '';
    return;
  }

  const sample  = ALL[0] || {};
  const hasProp = k => Object.prototype.hasOwnProperty.call(sample, k);
  const visCols = COLS.filter(c => hasProp(c.key));

  const reqCols   = visCols.filter(c => c.level === 'req');
  const versCols  = visCols.filter(c => c.level === 'vers');
  const tcCols    = visCols.filter(c => c.level === 'leaf' && c.group !== 'sticr');
  const riskCols  = visCols.filter(c => c.level === 'risk');
  const sticrCols = visCols.filter(c => c.level === 'leaf' && c.group === 'sticr');

  $('result-count').textContent = `${data.length.toLocaleString()} rows`;

  const h = ['<table>'];

  // ── thead row 1: group colour bands ──
  h.push('<thead><tr class="h-group">');
  for (const g of ['req', 'tc', 'risk', 'sticr']) {
    const n = visCols.filter(c => c.group === g).length;
    if (n > 0) h.push(`<th class="${GROUP_META[g].cls}" colspan="${n}">${GROUP_META[g].label}</th>`);
  }
  h.push('</tr>');

  // ── thead row 2: column names ──
  h.push('<tr class="h-cols">');
  for (const c of visCols) h.push(`<th>${esc(c.label)}</th>`);
  h.push('</tr></thead>');

  // ── tbody: 3-level rowspan ──
  // Column order per row:
  //   [req cells (rowspan=r1Span, first row of req group only)]
  //   [vers cell (rowspan=vSpan,  first row of vers group only)]
  //   [TC + test-point leaf cells (every row)]
  //   [risk cells (rowspan=rkSpan, first row of risk group only)]
  //   [STICR leaf cells (every row)]
  h.push('<tbody>');
  const { r1Map, r1Order } = buildGroups(data);

  r1Order.forEach((r1k, ri) => {
    const r1       = r1Map.get(r1k);
    const rowClass = ri % 2 === 0 ? 'rg-even' : 'rg-odd';

    const r1Span = r1.vOrder.reduce((s, vk) => {
      const v = r1.vMap.get(vk);
      return s + v.rkOrder.reduce((ss, rk) => ss + v.rkMap.get(rk).length, 0);
    }, 0);

    r1.vOrder.forEach((vk, vi) => {
      const v     = r1.vMap.get(vk);
      const vSpan = v.rkOrder.reduce((s, rk) => s + v.rkMap.get(rk).length, 0);

      v.rkOrder.forEach((rk, rki) => {
        const rows   = v.rkMap.get(rk);
        const rkSpan = rows.length;

        rows.forEach((row, li) => {
          const isFirst = vi === 0 && rki === 0 && li === 0;
          h.push(`<tr class="${rowClass}${isFirst ? ' grp-start' : ''}">`);

          // req-level: first row of (Req ID + Released in) group
          if (vi === 0 && rki === 0 && li === 0) {
            for (const c of reqCols) {
              h.push(`<td${c.sticky ? ' class="c-req"' : ''} rowspan="${r1Span}">${cellHtml(row, c)}</td>`);
            }
          }

          // vers-level: first row of Reqmt for Version group
          if (rki === 0 && li === 0) {
            for (const c of versCols) {
              h.push(`<td rowspan="${vSpan}">${cellHtml(row, c)}</td>`);
            }
          }

          // TC + test-point leaf: every row
          for (const c of tcCols)    h.push(`<td>${cellHtml(row, c)}</td>`);

          // risk-level: first row of Risk ID group
          if (li === 0) {
            for (const c of riskCols) {
              h.push(`<td rowspan="${rkSpan}">${cellHtml(row, c)}</td>`);
            }
          }

          // STICR leaf: every row
          for (const c of sticrCols) h.push(`<td>${cellHtml(row, c)}</td>`);

          h.push('</tr>');
        });
      });
    });
  });

  h.push('</tbody></table>');
  host.innerHTML = h.join('');
  applySticky();
}

// ── Populate filter dropdowns ─────────────────────────────────────────────────
function populateOptions() {
  function fill(id, key) {
    const el = $(id), first = el.options[0].cloneNode(true);
    el.innerHTML = ''; el.appendChild(first);
    [...new Set(ALL.map(r => val(r, key)).filter(Boolean))]
      .sort((a, b) => a.localeCompare(b, undefined, { numeric: true }))
      .forEach(v => { const o = document.createElement('option'); o.value = o.textContent = v; el.appendChild(o); });
  }
  fill('f-fs', 'Feature Spec');
  pvVersions = [...new Set(ALL.map(r => val(r, 'Reqmt for Version')).filter(Boolean))]
    .sort((a, b) => a.localeCompare(b, undefined, { numeric: true }));
  pvRebuildPanel();
}

// ── CSV loading ───────────────────────────────────────────────────────────────
function loadCsv(file) {
  $('file-label').textContent = `Loading ${file.name}…`;
  Papa.parse(file, {
    header: true,
    skipEmptyLines: true,
    complete(res) {
      ALL = res.data;
      $('file-label').textContent = `${file.name}  ·  ${ALL.length.toLocaleString()} rows`;
      populateOptions();
      render();
    },
    error(err) { $('file-label').textContent = `Error: ${err.message}`; },
  });
}

// ── Template management ───────────────────────────────────────────────────────
function bufferToBase64(buf) {
  const b = new Uint8Array(buf); let s = '';
  for (const x of b) s += String.fromCharCode(x);
  return btoa(s);
}
function base64ToBuffer(b64) {
  const s = atob(b64), b = new Uint8Array(s.length);
  for (let i = 0; i < s.length; i++) b[i] = s.charCodeAt(i);
  return b.buffer;
}

function loadTemplate(buffer, save) {
  templateBuffer = buffer;
  $('btn-export').disabled = false;
  $('btn-export').title    = 'Export current view as DOCX';
  $('btn-tmpl').classList.add('loaded');
  $('btn-tmpl').textContent = '&#10003; Template loaded';
  if (save) { try { localStorage.setItem('tm_template_b64', bufferToBase64(buffer)); } catch {} }
}

async function tryAutoLoadTemplate() {
  try {
    const cached = localStorage.getItem('tm_template_b64');
    if (cached) { loadTemplate(base64ToBuffer(cached), false); return; }
  } catch {}
  try {
    const r = await fetch(TEMPLATE_URL);
    if (r.ok) { loadTemplate(await r.arrayBuffer(), true); return; }
  } catch {}
  $('btn-export').disabled = false;
  $('btn-export').title    = 'Click to select the .docx template, then export';
  $('btn-tmpl').textContent = '&#128196; Load Template';
}

// ── DOCX export ───────────────────────────────────────────────────────────────
function escXml(s) {
  return String(s ?? '')
    .replace(/&/g, '&amp;').replace(/</g, '&lt;')
    .replace(/>/g, '&gt;').replace(/"/g, '&quot;');
}
function makeDocxRow(cells) {
  return `<w:tr>${cells.map(t =>
    `<w:tc><w:p><w:r><w:t xml:space="preserve">${escXml(t)}</w:t></w:r></w:p></w:tc>`
  ).join('')}</w:tr>`;
}
function replaceAllXml(xml, s, r) { return xml.split(s).join(r); }

async function exportDocx() {
  if (!ALL.length)     { alert('Load the CSV first.'); return; }
  if (!templateBuffer) { $('tmpl-input').click(); return; }

  const f    = getFilters();
  let   data = ALL.filter(r => rowMatches(r, f));
  if ($('f-latest').checked) data = dedupeLatestTC(data);

  data = [...data].sort((a, b) => {
    const ri = String(val(a, 'Requirement ID')).localeCompare(String(val(b, 'Requirement ID')), undefined, { numeric: true });
    return ri !== 0 ? ri : String(val(a, 'Test Case ID')).localeCompare(String(val(b, 'Test Case ID')), undefined, { numeric: true });
  });
  if (!data.length) { alert('No rows match the current view.'); return; }

  try {
    const zip = await JSZip.loadAsync(templateBuffer);
    if (!zip.files['word/document.xml']) { alert('Invalid .docx — word/document.xml not found.'); return; }

    let xml = await zip.files['word/document.xml'].async('string');
    const vers    = (pvSelected.size > 0 ? [...pvSelected] : pvVersions).join(', ');
    const versXml = escXml(vers);
    xml = replaceAllXml(xml, '{VERS}', versXml);
    for (const [path, file] of Object.entries(zip.files)) {
      if (/^word\/(header|footer)\d*\.xml$/.test(path)) {
        zip.file(path, replaceAllXml(await file.async('string'), '{VERS}', versXml));
      }
    }

    let hi = xml.indexOf('Test Case ID'); if (hi === -1) hi = xml.indexOf('<w:tbl>');
    if (hi === -1) { alert('Could not find the traceability table.'); return; }
    const ti = xml.indexOf('</w:tbl>', hi);
    if (ti === -1) { alert('Malformed template XML.'); return; }

    const newRows = data.map(row => makeDocxRow([
      val(row, 'Test Case ID'),
      val(row, 'Test Case Name'),
      val(row, 'Test Result'),
      val(row, 'Requirement ID'),
      val(row, 'Requirement Name'),
    ])).join('');

    xml = xml.slice(0, ti) + newRows + xml.slice(ti);
    zip.file('word/document.xml', xml);

    const blob = await zip.generateAsync({
      type: 'blob',
      mimeType: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
    });
    const a = document.createElement('a');
    a.href     = URL.createObjectURL(blob);
    a.download = vers ? `Traceability_${vers}.docx` : 'Traceability_export.docx';
    a.style.display = 'none'; document.body.appendChild(a); a.click();
    setTimeout(() => { URL.revokeObjectURL(a.href); a.remove(); }, 2000);

  } catch (e) { alert('Export failed: ' + e.message); console.error(e); }
}

// ── Event wiring ──────────────────────────────────────────────────────────────
$('csv-input').addEventListener('change', e => { if (e.target.files[0]) loadCsv(e.target.files[0]); });
$('btn-tmpl').addEventListener('click', () => $('tmpl-input').click());
$('tmpl-input').addEventListener('change', e => {
  const f = e.target.files[0]; if (f) f.arrayBuffer().then(buf => loadTemplate(buf, true));
});
['f-fs','f-pt','f-lr'].forEach(id => $(id).addEventListener('change', render));
['f-ri','f-tc','f-pi','f-si'].forEach(id => $(id).addEventListener('input', render));
['f-latest','f-has-risk','f-has-sticr'].forEach(id => $(id).addEventListener('change', render));

$('btn-clear').addEventListener('click', () => {
  ['f-fs','f-pt','f-lr'].forEach(id => $(id).value = '');
  ['f-ri','f-tc','f-pi','f-si'].forEach(id => $(id).value = '');
  ['f-latest','f-has-risk','f-has-sticr'].forEach(id => $(id).checked = false);
  pvSelected.clear(); pvRebuildPanel(); pvUpdateLabel(); render();
});
$('btn-export').addEventListener('click', exportDocx);

// ── Startup ───────────────────────────────────────────────────────────────────
tryAutoLoadTemplate();
const CACHE_KEY = 'tm_data', CACHE_TTL = 60 * 60 * 1000;
const cached = JSON.parse(localStorage.getItem(CACHE_KEY) || 'null');
if (cached && Date.now() - cached.ts < CACHE_TTL) {
  ALL = cached.rows;
  $('file-label').textContent = `Snowflake (cached)  ·  ${ALL.length.toLocaleString()} rows`;
  populateOptions();
  render();
} else {
  fetch('data.json')
    .then(r => r.ok ? r.json() : Promise.reject())
    .then(data => {
      try { localStorage.setItem(CACHE_KEY, JSON.stringify({ ts: Date.now(), rows: data })); } catch (_) {}
      ALL = data;
      $('file-label').textContent = `Snowflake  ·  ${ALL.length.toLocaleString()} rows`;
      populateOptions();
      render();
    })
    .catch(() => {
      fetch('traceability_matrix_alarm_panel.csv')
        .then(r => r.ok ? r.blob() : Promise.reject())
        .then(b => loadCsv(new File([b], 'traceability_matrix_alarm_panel.csv', { type: 'text/csv' })))
        .catch(() => {});
    });
}

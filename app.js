// ─────────────────────────────────────────────
// STATE
// ─────────────────────────────────────────────
let demandMap      = {};  // { [key]: { month, items: [{ group, name, qty }] } }
let supplyRows     = [];  // [{ group, name, qty, month }]
let cardOrder      = [];
let cardChecked    = {};
let cardExpanded   = {};
let collapsedMonths = {};
let nameSelections  = {}; // { "DemandKey_idx": selectedName }
let collapsedStockGroups = {};

function poolKey(group, name) {
  return `${group}__${name}`;
}

function getEffectiveName(key, idx) {
  const k = `${key}_${idx}`;
  if (nameSelections.hasOwnProperty(k)) return nameSelections[k];

  const item = demandMap[key].items[idx];
  const names = getNamesForGroup(item.group);
  if (names.length === 1) return names[0];

  return item.name || '';
}

function hasSupplyMatch(group, name) {
  if (!group || !name) return false;
  return supplyRows.some(r => r.group === group && r.name === name);
}

function getNamesForGroup(group) {
  const seen = new Set(), result = [];
  supplyRows.forEach(r => {
    if (r.group === group && !seen.has(r.name)) { seen.add(r.name); result.push(r.name); }
  });
  return result;
}

// ── Month helpers ──────────────────────────────
function getTodayMonth() {
  const d = new Date();
  return d.getFullYear() * 100 + (d.getMonth() + 1);
}

function addMonths(yymm, n) {
  let y = Math.floor(yymm / 100), m = yymm % 100;
  m += n;
  while (m > 12) { m -= 12; y++; }
  while (m < 1)  { m += 12; y--; }
  return y * 100 + m;
}

function getFullMonthList() {
  if (!Object.keys(demandMap).length) return [];
  const today = getTodayMonth();
  let maxM = today;
  Object.values(demandMap).forEach(d => { if (d.month > maxM) maxM = d.month; });
  supplyRows.forEach(r => { if (r.month !== 'STOCK' && r.month > maxM) maxM = r.month; });
  const end = addMonths(maxM, 3);
  const set = new Set();
  // Always include months that actually have demand data (even if before today)
  Object.values(demandMap).forEach(d => set.add(d.month));
  // Fill in the range from today to end
  let cur = today;
  while (cur <= end) { set.add(cur); cur = addMonths(cur, 1); }
  return [...set].sort((a, b) => a - b);
}

// ─────────────────────────────────────────────
// FILE UPLOAD
// ─────────────────────────────────────────────
document.getElementById('file-input').addEventListener('change', e => {
  const file = e.target.files[0];
  if (!file) return;
  const reader = new FileReader();
  reader.onload = ev => {
    try {
      parseWorkbook(XLSX.read(ev.target.result, { type: 'array' }), file.name);
    } catch(err) { showError('讀取 XLSX 失敗：' + err.message); }
  };
  reader.readAsArrayBuffer(file);
  e.target.value = '';
});

function parseWorkbook(wb, filename) {
  clearError();
  const lcNames = wb.SheetNames.map(s => s.toLowerCase());
  const missing = [];
  if (!lcNames.includes('demand')) missing.push('Demand');
  if (!lcNames.includes('supply')) missing.push('Supply');
  if (missing.length) { showError(`缺少 Sheet：${missing.join('、')}`); return; }

  const demandSheet = wb.Sheets[wb.SheetNames[lcNames.indexOf('demand')]];
  const supplySheet = wb.Sheets[wb.SheetNames[lcNames.indexOf('supply')]];
  const notBlank = obj => !Object.values(obj).every(v => String(v).trim() === '');
  const dRaws = XLSX.utils.sheet_to_json(demandSheet, { defval: '' }).filter(notBlank);
  const sRaws = XLSX.utils.sheet_to_json(supplySheet, { defval: '' }).filter(notBlank);

  // Case-insensitive column picker
  const pick = (obj, ...candidates) => {
    const keys = Object.keys(obj);
    for (const c of candidates) {
      const found = keys.find(k => k.trim().toLowerCase() === c.toLowerCase());
      if (found !== undefined) return String(obj[found]).trim();
    }
    return '';
  };

  const newDemandMap = {};
  dRaws.forEach(row => {
    const demand = pick(row, 'Demand');
    const group  = pick(row, 'Group', 'SPEC');
    const name   = pick(row, 'Name');
    const qty    = Number(pick(row, 'Qty'));
    const month  = Number(pick(row, 'Month'));
    if (!demand || !group || isNaN(qty) || isNaN(month)) return;
    if (!newDemandMap[demand]) newDemandMap[demand] = { month, items: [] };
    newDemandMap[demand].items.push({ group, name, qty });
  });

  const newSupplyRows = [];
  sRaws.forEach(row => {
    const group = pick(row, 'Group', 'SPEC');
    const name  = pick(row, 'Name');
    const qty   = Number(pick(row, 'Qty'));
    const mRaw  = pick(row, 'Month').toUpperCase();
    if (!group || isNaN(qty)) return;
    const month = (mRaw === 'STOCK' || mRaw === '') ? 'STOCK' : Number(pick(row, 'Month'));
    newSupplyRows.push({ group, name, qty, month });
  });

  demandMap       = newDemandMap;
  supplyRows      = newSupplyRows;
  nameSelections  = {};
  cardOrder       = Object.keys(newDemandMap).sort((a, b) => newDemandMap[a].month - newDemandMap[b].month);
  cardChecked     = {};
  cardExpanded    = {};
  collapsedMonths = {};
  collapsedStockGroups = {};
  cardOrder.forEach(k => { cardChecked[k] = true; cardExpanded[k] = false; });

  document.getElementById('file-status').textContent = filename;
  document.getElementById('dl-btn').style.display = '';
  renderAll();
}

// ─────────────────────────────────────────────
// CALCULATION ENGINE
// ─────────────────────────────────────────────
function calculate() {
  // Build supply pools keyed by group+name
  const stockPool   = {};
  const incomingPool = [];
  supplyRows.forEach(r => {
    if (r.month === 'STOCK') {
      const key = poolKey(r.group, r.name);
      stockPool[key] = (stockPool[key] || 0) + r.qty;
    } else {
      incomingPool.push({ group: r.group, name: r.name, qty: r.qty, month: r.month, remaining: r.qty });
    }
  });
  incomingPool.sort((a, b) => a.month - b.month);

  const results = {};
  const committedByKey = {}; // supply key → total qty committed by checked demands

  for (const key of cardOrder) {
    const demand      = demandMap[key];
    const demandMonth = demand.month;

    // Per-item availability using current pool state (before this demand deducts)
    const itemResults = demand.items.map((item, idx) => {
      // If no supply exists for this group at all → ignore (not unresolved)
      const names = getNamesForGroup(item.group);
      if (names.length === 0) {
        return { group: item.group, origName: item.name, effectiveName: '', qty: item.qty,
                 available: null, gap: null, itemStatus: 'ignored' };
      }

      const effectiveName = getEffectiveName(key, idx);
      if (effectiveName && !hasSupplyMatch(item.group, effectiveName)) {
        return { group: item.group, origName: item.name, effectiveName: '', qty: item.qty,
                 available: null, gap: null, itemStatus: 'ignored' };
      }
      if (!effectiveName) {
        return { group: item.group, origName: item.name, effectiveName: '', qty: item.qty,
                 available: null, gap: null, itemStatus: 'unresolved' };
      }
      const supplyKey = poolKey(item.group, effectiveName);
      const sAvail = stockPool[supplyKey] || 0;
      const iAvail = incomingPool
        .filter(r => poolKey(r.group, r.name) === supplyKey && r.month <= demandMonth)
        .reduce((s, r) => s + r.remaining, 0);
      const available = sAvail + iAvail;
      const gap = available - item.qty;
      return { group: item.group, origName: item.name, effectiveName, qty: item.qty,
               available, gap, itemStatus: gap >= 0 ? 'ok' : 'short' };
    });

    // Accumulate committed demand for all checked cards
    if (cardChecked[key]) {
      itemResults.forEach(item => {
        if (item.effectiveName) {
          const sk = poolKey(item.group, item.effectiveName);
          committedByKey[sk] = (committedByKey[sk] || 0) + item.qty;
        }
      });
    }

    const hasUnresolved = itemResults.some(i => i.itemStatus === 'unresolved');
    const hasShort      = itemResults.some(i => i.itemStatus === 'short');
    const calcStatus    = hasUnresolved ? 'unresolved' : hasShort ? 'shortage' : 'ok';
    const deductStatus  = cardChecked[key] ? (calcStatus === 'ok' ? 'ok' : 'shortage') : 'frozen';

    results[key] = { status: deductStatus, calcStatus, items: itemResults };

    // Deduct supply only when checked AND fully satisfiable
    if (cardChecked[key] && calcStatus === 'ok') {
      for (const item of itemResults) {
        if (!item.effectiveName) continue; // skip ignored items
        let toDeduct = item.qty;
        const supplyKey = poolKey(item.group, item.effectiveName);
        const fromStock = Math.min(toDeduct, stockPool[supplyKey] || 0);
        stockPool[supplyKey] = (stockPool[supplyKey] || 0) - fromStock;
        toDeduct -= fromStock;
        if (toDeduct > 0) {
          for (const r of incomingPool) {
            if (poolKey(r.group, r.name) !== supplyKey || r.month > demandMonth) continue;
            const take = Math.min(toDeduct, r.remaining);
            r.remaining -= take;
            toDeduct    -= take;
            if (toDeduct <= 0) break;
          }
        }
      }
    }
  }

  return { results, committedByKey };
}

// ─────────────────────────────────────────────
// RENDER
// ─────────────────────────────────────────────
function renderAll() {
  const { results, committedByKey } = calculate();
  renderCards(results);
  renderSummary(results);
  renderStockTable(committedByKey);
}

// Event delegation for name dropdowns inside cards
document.getElementById('card-list').addEventListener('change', e => {
  if (e.target.classList.contains('name-select') && !e.target.disabled) {
    e.stopPropagation();
    nameSelections[`${e.target.dataset.key}_${e.target.dataset.idx}`] = e.target.value;
    renderAll();
  }
});

function esc(s) {
  return String(s)
    .replace(/&/g,'&amp;').replace(/</g,'&lt;').replace(/>/g,'&gt;').replace(/"/g,'&quot;');
}

function renderCards(results) {
  const list = document.getElementById('card-list');
  list.innerHTML = '';

  if (!cardOrder.length) {
    list.innerHTML = '<div class="empty-state"><div>📂</div><div>上傳 XLSX 以載入需求</div></div>';
    document.getElementById('card-count').textContent = '— 筆';
    return;
  }

  const allMonths = getFullMonthList();

  // Build month → keys map
  const monthKeysMap = {};
  allMonths.forEach(m => { monthKeysMap[m] = []; });
  cardOrder.forEach(key => {
    const m = demandMap[key].month;
    if (monthKeysMap[m] !== undefined) {
      monthKeysMap[m].push(key);
    } else {
      // Demand month outside computed range — add it
      monthKeysMap[m] = [key];
      allMonths.push(m);
      allMonths.sort((a, b) => a - b);
    }
  });

  allMonths.forEach(month => {
    const keys = monthKeysMap[month] || [];
    const isCollapsed = !!collapsedMonths[month];
    const okCount    = keys.filter(k => results[k]?.calcStatus === 'ok').length;
    const shortCount = keys.filter(k => results[k] && results[k].calcStatus !== 'ok').length;

    const groupDiv = document.createElement('div');
    groupDiv.className = 'month-group';

    const hdr = document.createElement('div');
    hdr.className = 'month-header' + (isCollapsed ? ' collapsed' : '');
    hdr.innerHTML = `
      <span class="mh-arrow">▾</span>
      <span class="mh-label">${month}</span>
      <span class="mh-count">${keys.length ? keys.length + '筆' : '—'}</span>
      ${okCount    > 0 ? `<span class="mh-ok">✓${okCount}</span>` : ''}
      ${shortCount > 0 ? `<span class="mh-short">✗${shortCount}</span>` : ''}
    `;
    hdr.onclick = () => { collapsedMonths[month] = !collapsedMonths[month]; renderAll(); };
    hdr.ondragover = e => e.preventDefault();
    hdr.ondrop = e => {
      e.preventDefault();
      if (dragSrcKey && demandMap[dragSrcKey].month !== month) {
        moveDemandToMonth(dragSrcKey, month);
        renderAll();
      }
    };

    const cardsDiv = document.createElement('div');
    cardsDiv.className = 'month-cards' + (isCollapsed ? ' collapsed' : '');
    cardsDiv.ondragover = e => {
      e.preventDefault();
      cardsDiv.classList.add('drag-over-month');
    };
    cardsDiv.ondragleave = e => {
      if (!cardsDiv.contains(e.relatedTarget)) cardsDiv.classList.remove('drag-over-month');
    };
    cardsDiv.ondrop = e => {
      e.preventDefault();
      cardsDiv.classList.remove('drag-over-month');
      if (dragSrcKey && demandMap[dragSrcKey].month !== month) {
        moveDemandToMonth(dragSrcKey, month);
        renderAll();
      }
    };

    if (!keys.length) {
      const dz = document.createElement('div');
      dz.className = 'month-drop-zone';
      dz.textContent = '拖曳至此月份';
      cardsDiv.appendChild(dz);
    }

    keys.forEach(key => {
      const result  = results[key];
      const checked = cardChecked[key];
      const expanded = cardExpanded[key];

      const calcClass = `calc-${result.calcStatus === 'ok' ? 'ok' : result.calcStatus === 'unresolved' ? 'unresolved' : 'short'}`;
      const card = document.createElement('div');
      card.className = `demand-card ${calcClass}${!checked ? ' frozen' : ''}${expanded ? ' expanded' : ''}`;

      // Status badge
      let statusHtml;
      if (!checked) {
        statusHtml = `<div class="card-status status-frozen-txt">—</div>`;
      } else if (result.status === 'ok') {
        statusHtml = `<div class="card-status status-ok">OK</div>`;
      } else if (result.calcStatus === 'unresolved') {
        statusHtml = `<div class="card-status status-unresolved">未完成</div>`;
      } else {
        statusHtml = `<div class="card-status status-short">缺料</div>`;
      }

      const unresCount   = result.items.filter(i => i.itemStatus === 'unresolved').length;
      const ignoredCount = result.items.filter(i => i.itemStatus === 'ignored').length;
      const metaText = `${demandMap[key].items.length}項` +
        (unresCount   > 0 ? ` · ${unresCount} 項未選 Name` : '') +
        (ignoredCount > 0 ? ` · ${ignoredCount} 項無供料` : '');

      card.innerHTML = `
        <div class="card-header">
          <input type="checkbox" class="card-checkbox"${checked?' checked':''} data-key="${esc(key)}">
          <div class="card-expand">▶</div>
          <div class="card-priority">#${cardOrder.indexOf(key)+1}</div>
          <div class="card-info">
            <div class="card-name">${esc(key)}</div>
            <div class="card-meta">${metaText}</div>
          </div>
          ${statusHtml}
        </div>
        <div class="card-detail">
          ${buildDetailTable(key, result.items)}
        </div>
      `;

      // Checkbox handler
      card.querySelector('.card-checkbox').addEventListener('change', e => {
        e.stopPropagation();
        cardChecked[key] = e.target.checked;
        renderAll();
      });

      // Expand/collapse on header click (skip checkbox)
      card.querySelector('.card-header').addEventListener('click', e => {
        if (e.target.closest('.card-checkbox')) return;
        cardExpanded[key] = !cardExpanded[key];
        card.classList.toggle('expanded', cardExpanded[key]);
      });

      // Prevent drag from starting when interacting with controls or detail
      card.querySelector('.card-detail').addEventListener('mousedown', e => e.stopPropagation());
      card.querySelector('.card-header').addEventListener('mousedown', e => {
        if (e.target.closest('.card-checkbox')) e.stopPropagation();
      });

      setupDrag(card, key);
      cardsDiv.appendChild(card);
    });

    groupDiv.appendChild(hdr);
    groupDiv.appendChild(cardsDiv);
    list.appendChild(groupDiv);
  });

  document.getElementById('card-count').textContent = `${cardOrder.length} 筆`;
}

function buildDetailTable(key, items) {
  const rows = items.map((item, idx) => {
    const names         = getNamesForGroup(item.group);
    const noSupply      = names.length === 0;
    const effectiveName = getEffectiveName(key, idx);
    const isUnselected  = !effectiveName;
    const isIgnored     = item.itemStatus === 'ignored';

    let selectHtml;
    if (noSupply) {
      selectHtml = `<select class="name-select unselected" disabled><option>— 無對應原料 —</option></select>`;
    } else {
      const options = ['', ...names].map(n =>
        `<option value="${esc(n)}"${n===effectiveName?' selected':''}>${n ? esc(n) : '-- 未選擇 --'}</option>`
      ).join('');
      selectHtml = `<select class="name-select${isUnselected?' unselected':''}" data-key="${esc(key)}" data-idx="${idx}" onmousedown="event.stopPropagation()">${options}</select>`;
    }

    let availTd, gapTd;
    if (isIgnored) {
      availTd = `<td class="r gap-zero">—</td>`;
      gapTd   = `<td class="r gap-zero">—</td>`;
    } else if (item.available === null) {
      availTd = `<td class="r gap-zero">—</td>`;
      gapTd   = `<td class="r gap-zero">—</td>`;
    } else {
      availTd = `<td class="r">${fmt(item.available)}</td>`;
      const gapClass = item.gap > 0 ? 'gap-pos' : item.gap < 0 ? 'gap-neg' : 'gap-zero';
      const gapStr   = item.gap >= 0 ? `+${fmt(item.gap)}` : fmt(item.gap);
      gapTd = `<td class="r ${gapClass}">${gapStr}</td>`;
    }

    return `<tr>
      <td class="group-cell" title="${esc(item.group)}">${esc(item.group)}</td>
      <td>${selectHtml}</td>
      <td class="r">${fmt(item.qty)}</td>
      ${availTd}
      ${gapTd}
    </tr>`;
  }).join('');

  return `<table class="items-table">
    <thead><tr>
      <th>Group</th>
      <th>Name</th>
      <th class="r">需求</th>
      <th class="r">可用</th>
      <th class="r">缺口</th>
    </tr></thead>
    <tbody>${rows}</tbody>
  </table>`;
}

function renderSummary(results) {
  let ok = 0, short = 0, frozen = 0;
  Object.values(results).forEach(r => {
    if (r.status === 'ok') ok++;
    else if (r.status === 'shortage') short++;
    else frozen++;
  });
  document.getElementById('s-ok').textContent    = ok;
  document.getElementById('s-short').textContent = short;
  document.getElementById('s-frozen').textContent = frozen;
  document.getElementById('s-total').textContent  = cardOrder.length;
}

function renderStockTable(committedByKey) {
  const tbody = document.getElementById('stock-tbody');
  tbody.innerHTML = '';
  const groupMap = {};
  const groupOrder = [];
  const seen = new Set();

  supplyRows.forEach(r => {
    if (!groupMap[r.group]) {
      groupMap[r.group] = [];
      groupOrder.push(r.group);
    }
    const key = poolKey(r.group, r.name);
    if (seen.has(key)) return;
    seen.add(key);
    groupMap[r.group].push({ ...r, key });
  });

  groupOrder.forEach(group => {
    const rows = groupMap[group];
    const isCollapsed = !!collapsedStockGroups[group];

    const hdrTr = document.createElement('tr');
    hdrTr.className = 'stock-group-header';
    hdrTr.innerHTML = `<td colspan="6">
      <div class="sg-header${isCollapsed ? ' collapsed' : ''}">
        <span class="sg-arrow">▾</span>
        <span class="sg-label">${esc(group)}</span>
        <span class="sg-count">${rows.length}筆</span>
      </div>
    </td>`;
    hdrTr.onclick = () => { collapsedStockGroups[group] = !collapsedStockGroups[group]; renderAll(); };
    tbody.appendChild(hdrTr);

    rows.forEach(r => {
      const sOrig    = supplyRows.filter(s => poolKey(s.group, s.name) === r.key && s.month === 'STOCK').reduce((a, s) => a + s.qty, 0);
      const iOrig    = supplyRows.filter(s => poolKey(s.group, s.name) === r.key && s.month !== 'STOCK').reduce((a, s) => a + s.qty, 0);
      const committed = committedByKey[r.key] || 0;
      const balance   = sOrig + iOrig - committed;

      const tr = document.createElement('tr');
      tr.className = `stock-item-row${isCollapsed ? ' collapsed' : ''}`;
      tr.innerHTML = `
        <td></td>
        <td>${esc(r.name)}</td>
        <td>${fmt(sOrig)}</td>
        <td>${fmt(iOrig)}</td>
        <td>${fmt(committed)}</td>
        <td class="${numClass(balance)}">${fmt(balance)}</td>
      `;
      tbody.appendChild(tr);
    });
  });
}

// ─────────────────────────────────────────────
// DOWNLOAD — preserves original upload format
// ─────────────────────────────────────────────
function downloadResult() {
  const wb = XLSX.utils.book_new();

  // Demand sheet with effective names filled in
  const demandData = [['Demand', 'Group', 'Name', 'Qty', 'Month']];
  cardOrder.forEach(key => {
    const { month, items } = demandMap[key];
    items.forEach((item, idx) => {
      demandData.push([key, item.group, getEffectiveName(key, idx), item.qty, month]);
    });
  });
  XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet(demandData), 'Demand');

  // Supply sheet — unchanged
  const supplyData = [['Group', 'Name', 'Qty', 'Month']];
  supplyRows.forEach(r => {
    supplyData.push([r.group, r.name, r.qty, r.month === 'STOCK' ? 'Stock' : r.month]);
  });
  XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet(supplyData), 'Supply');

  XLSX.writeFile(wb, 'kit-check-result.xlsx');
}

function downloadTemplate() {
  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet([
    ['Demand', 'Group', 'Name', 'Qty', 'Month'],
    ['Demand1', 'DIMM 64G', '', 800, 202605],
    ['Demand1', 'AMD/9224', '', 100, 202605],
    ['Demand2', 'DIMM 128G', 'SAMSUNG/DDR5 128G/6400', 1200, 202606],
  ]), 'Demand');
  XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet([
    ['Group', 'Name', 'Qty', 'Month'],
    ['DIMM 64G', 'SAMSUNG/DDR5 64G/6400', 5000, 'Stock'],
    ['AMD/9224', 'AMD/9224', 200, 'Stock'],
    ['DIMM 128G', 'SAMSUNG/DDR5 128G/6400', 3000, 202606],
  ]), 'Supply');
  XLSX.writeFile(wb, 'kit-check-template.xlsx');
}

// ─────────────────────────────────────────────
// DRAG & DROP
// ─────────────────────────────────────────────
let dragSrcKey = null;

function moveDemandToMonth(srcKey, targetMonth) {
  demandMap[srcKey].month = targetMonth;
  // Remove from current position
  const fromIdx = cardOrder.indexOf(srcKey);
  cardOrder.splice(fromIdx, 1);
  // Insert at end of targetMonth's section (before any card with later month)
  let insertIdx = cardOrder.length;
  for (let i = 0; i < cardOrder.length; i++) {
    if (demandMap[cardOrder[i]].month > targetMonth) { insertIdx = i; break; }
  }
  cardOrder.splice(insertIdx, 0, srcKey);
}

function setupDrag(card, key) {
  card.draggable = true;
  card.ondragstart = e => {
    if (e.target.closest('.card-detail') || e.target.tagName === 'SELECT' || e.target.tagName === 'INPUT') {
      e.preventDefault(); return;
    }
    dragSrcKey = key;
    card.classList.add('dragging');
  };
  card.ondragend   = () => card.classList.remove('dragging');
  card.ondragover  = e => e.preventDefault();
  card.ondrop      = e => {
    e.preventDefault();
    if (dragSrcKey && dragSrcKey !== key) {
      const srcMonth = demandMap[dragSrcKey].month;
      const tgtMonth = demandMap[key].month;
      if (srcMonth !== tgtMonth) {
        // Cross-month drop: change month, then reorder before target card
        demandMap[dragSrcKey].month = tgtMonth;
      }
      const from = cardOrder.indexOf(dragSrcKey);
      cardOrder.splice(from, 1);
      const to = cardOrder.indexOf(key);
      cardOrder.splice(to, 0, dragSrcKey);
      renderAll();
    }
  };
}

// ─────────────────────────────────────────────
// UTILS
// ─────────────────────────────────────────────
function fmt(n)        { return Number(n).toLocaleString(); }
function numClass(n)   { return n > 0 ? 'num-pos' : n < 0 ? 'num-neg' : 'num-zero'; }
function showError(msg){ const b = document.getElementById('error-bar'); b.textContent = msg; b.style.display = 'block'; }
function clearError()  { document.getElementById('error-bar').style.display = 'none'; }

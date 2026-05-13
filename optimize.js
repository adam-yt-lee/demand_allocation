// ─────────────────────────────────────────────
// AUTO-OPTIMIZE  ── Branch-and-Bound ILP
//
// 問題建模：
//   決策變數  y_d ∈ {0,1}：第 d 筆需求是否被滿足
//             x_{d,i,n} ∈ {0,1}：第 d 筆需求第 i 項選 Name n
//   目標函數  最大化 Σ y_d
//   約束條件  ① 每個 (d,i) 恰好選一個 Name
//             ② 對每個 Name n、每個月份截止 m：
//                Σ_{d: month_d ≤ m, item uses n} qty * y_d ≤ supply[n][m]
//
// 演算法：以 y_d 為分支變數做 Branch-and-Bound；
//         子問題的 Name 指派以回溯搜索確定可行性。
// ─────────────────────────────────────────────
function autoOptimize() {
  if (!cardOrder.length) return;

  const btn = document.getElementById('optimize-btn');
  btn.disabled = true;
  btn.textContent = '運算中…';

  // Let the browser render the disabled state before blocking
  setTimeout(_runOptimize, 30);
}

function _runOptimize() {
  const btn = document.getElementById('optimize-btn');
  try {
    _solveAndApply();
  } finally {
    btn.disabled = false;
    btn.textContent = '⚡ 自動最佳化';
  }
}

function _solveAndApply() {
  // ── Build supply lookup ──────────────────────────────────────────
  const stockPool      = {};   // group+name → qty
  const incomingByKey  = {};   // group+name → [{qty, month}]  sorted asc

  for (const row of supplyRows) {
    if (row.month === 'STOCK') {
      const key = poolKey(row.group, row.name);
      stockPool[key] = (stockPool[key] || 0) + row.qty;
    } else {
      const key = poolKey(row.group, row.name);
      (incomingByKey[key] = incomingByKey[key] || [])
        .push({ qty: row.qty, month: row.month });
    }
  }
  for (const k in incomingByKey)
    incomingByKey[k].sort((a, b) => a.month - b.month);

  // Cumulative supply of `name` available up to `month`
  function supplyAt(group, name, month) {
    const key = poolKey(group, name);
    let t = stockPool[key] || 0;
    for (const inc of (incomingByKey[key] || []))
      if (inc.month <= month) t += inc.qty;
    return t;
  }

  // ── Build demand list（僅限已勾選的 Demand）──────────────────────
  const checkedKeys   = cardOrder.filter(k => cardChecked[k] === true);
  const uncheckedKeys = cardOrder.filter(k => cardChecked[k] !== true);

  if (!checkedKeys.length) {
    showToast('請先勾選至少一筆需求');
    return;
  }

  const demandList = checkedKeys.map(key => {
    const d     = demandMap[key];
    const month = d.month;
    const items = d.items.map((item, idx) => {
      const selKey    = `${key}_${idx}`;
      const fixedName = item.name || nameSelections[selKey] || '';
      const candidates = fixedName
        ? [fixedName]
        : getNamesForGroup(item.group).filter(n => n);
      return { selKey, candidates, qty: item.qty, month, group: item.group, isFixed: !!fixedName };
    });
    return { key, month, items };
  });

  // Sort by month ascending – tighter constraints first helps pruning
  demandList.sort((a, b) => a.month - b.month || a.key.localeCompare(b.key));

  const N = demandList.length;

  // ── Branch-and-Bound state ───────────────────────────────────────
  let bestCount      = 0;
  let bestSatisfied  = new Set();
  let bestAssignments = {};

  const t0      = Date.now();
  const LIMIT   = 10_000; // 10 s wall-clock limit
  let timedOut  = false;

  // consumed[name][month] = qty allocated to satisfied demands whose
  // adopted month === month.
  // consumedUpTo(name, m) = sum over all months mo ≤ m.
  function consumedUpTo(cons, group, name, month) {
    const byM = cons[poolKey(group, name)];
    if (!byM) return 0;
    let s = 0;
    for (const mo in byM) if (+mo <= month) s += byM[mo];
    return s;
  }

  function copyCons(cons) {
    const c = {};
    for (const n in cons) c[n] = { ...cons[n] };
    return c;
  }

  // Try to assign a valid Name to every item of a demand.
  // Uses backtracking within the demand.  Returns {newCons, newAssign}
  // if feasible, or null.
  function tryAssign(items, cons) {
    const nc = copyCons(cons);
    const na = {};

    function bt(i) {
      if (i >= items.length) return true;
      const item = items[i];
      if (!item.candidates.length) return bt(i + 1); // no supply for group — skip (treated as ignored)
      for (const name of item.candidates) {
        const avail = supplyAt(item.group, name, item.month)
                    - consumedUpTo(nc, item.group, name, item.month);
        if (avail >= item.qty) {
          const key = poolKey(item.group, name);
          if (!nc[key]) nc[key] = {};
          const prev = nc[key][item.month] || 0;
          nc[key][item.month] = prev + item.qty;
          if (!item.isFixed) na[item.selKey] = name;
          if (bt(i + 1)) return true;
          nc[key][item.month] = prev;
          if (!item.isFixed) delete na[item.selKey];
        }
      }
      return false;
    }

    return bt(0) ? { newCons: nc, newAssign: na } : null;
  }

  // Mutable set tracking the current branch's satisfied demands
  const curSatisfied = new Set();

  function bnb(idx, count, cons, assign) {
    if (timedOut) return;
    if (Date.now() - t0 > LIMIT) { timedOut = true; return; }

    // Record new best
    if (count > bestCount) {
      bestCount      = count;
      bestSatisfied  = new Set(curSatisfied);
      bestAssignments = { ...assign };
    }

    // Prune: even if all remaining demands are satisfied, can't beat best
    if (idx >= N || count + (N - idx) <= bestCount) return;

    const d = demandList[idx];

    // Branch 1 – include demand d (try to satisfy it)
    const r = tryAssign(d.items, cons);
    if (r) {
      curSatisfied.add(d.key);
      bnb(idx + 1, count + 1, r.newCons, { ...assign, ...r.newAssign });
      curSatisfied.delete(d.key);
    }

    // Branch 2 – exclude demand d
    bnb(idx + 1, count, cons, assign);
  }

  bnb(0, 0, {}, {});

  // ── Apply results ────────────────────────────────────────────────
  // 1. Name selections（只更新勾選的 Demand）
  Object.assign(nameSelections, bestAssignments);

  // 2. 不改變 cardChecked；重排順序：
  //    最優集合（勾選且可滿足）→ 其餘勾選 → 未勾選
  //    確保 calculate() 的順序式扣料能正確滿足最優集合
  const satArr       = [...bestSatisfied]
    .sort((a, b) => demandMap[a].month - demandMap[b].month);
  const unsatChecked = checkedKeys
    .filter(k => !bestSatisfied.has(k))
    .sort((a, b) => demandMap[a].month - demandMap[b].month);
  cardOrder = [...satArr, ...unsatChecked, ...uncheckedKeys];

  renderAll();

  // ── Show result toast ────────────────────────────────────────────
  const label = timedOut ? '最佳化（達時限，非最優解）' : '最佳解';
  showToast(`⚡ ${label}：${bestCount} / ${checkedKeys.length} 筆勾選需求可滿足`);
}

function showToast(msg) {
  const el = document.getElementById('opt-toast');
  el.textContent = msg;
  el.style.opacity = '1';
  clearTimeout(el._t);
  el._t = setTimeout(() => { el.style.opacity = '0'; }, 5000);
}

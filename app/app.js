(function () {
  'use strict';
  var E = SCEngine;
  var state = { params: null, name: null, meta: {}, run: null, runSnapshot: null, worker: null };

  // ---------- formatting ----------
  function usd(x) {
    if (x === Infinity) return 'No usable output';
    if (typeof x !== 'number' || isNaN(x)) return 'Not defined';
    return (x < 0 ? '-$' : '$') + Math.abs(x).toLocaleString('en-US', { maximumFractionDigits: Math.abs(x) < 100 ? 2 : 0 });
  }
  function bound(x) { return x === Infinity ? 'No finite bound' : usd(x); }
  function pct(x, d) { return (typeof x !== 'number' || isNaN(x)) ? 'Not defined' : (x * 100).toFixed(d == null ? 2 : d) + '%'; }
  function num(x) {
    if (typeof x !== 'number' || isNaN(x)) return 'Not defined';
    if (x === Infinity) return 'Infinite';
    if (Math.abs(x) >= 1e5) return x.toExponential(3) + ' (' + x.toLocaleString('en-US', { maximumFractionDigits: 0 }) + ')';
    return x.toLocaleString('en-US', { maximumFractionDigits: 4 });
  }
  function esc(s) { return String(s == null ? '' : s).replace(/[&<>"']/g, function (c) { return { '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[c]; }); }
  function $(id) { return document.getElementById(id); }

  // ---------- tabs ----------
  var tabs = document.querySelectorAll('nav [role=tab]');
  tabs.forEach(function (t) {
    t.addEventListener('click', function () { showView(t.getAttribute('data-view')); });
  });
  function showView(v) {
    tabs.forEach(function (t) { t.setAttribute('aria-selected', String(t.getAttribute('data-view') === v)); });
    document.querySelectorAll('.view').forEach(function (s) { s.hidden = s.id !== 'view-' + v; });
  }

  // ---------- scenarios ----------
  function load(params, name, meta) {
    state.params = JSON.parse(JSON.stringify(params));
    state.name = name; state.meta = meta || {};
    state.run = null; state.runSnapshot = null;
    $('scenarioName').textContent = name;
    $('scenarioBanner').innerHTML = meta && meta.banner ? '<div class="warn">' + esc(meta.banner) + '</div>' : '';
    renderAssumptions(); refresh(); renderSim();
  }
  $('loadIllustrative').onclick = function () { load(SCScenarios.illustrative.build(), SCScenarios.illustrative.name, { banner: 'HYPOTHETICAL demonstration inputs from PRD section 11. Not literature estimates or clinical dose definitions.' }); };
  $('loadTri').onclick = function () { load(SCScenarios.illustrativeTri.build(), SCScenarios.illustrativeTri.name, { banner: 'HYPOTHETICAL. r_g ~ triangular(1, 2, 6) is a subjective illustrative distribution unrelated to any measured study effect.' }); };
  $('loadEmpty').onclick = function () { load(SCScenarios.empty.build(), 'Empty scenario', {}); };
  function wireData(btn, data, name, banner) {
    if (!data) { btn.disabled = true; btn.title = 'Run build_spreadsheet.py to generate scenario_data.js'; return; }
    btn.onclick = function () { load(data.params, name, { banner: banner, evidence: data.paramEvidence, product: data.product }); };
  }
  wireData($('loadLiterature'), self.SCLiterature, 'Blood stem cells: cord-blood CD34+ expansion (literature-informed)',
    'Literature-informed, NOT validated: values come from Stem_Cell_Microeconomics.xlsx. Several space inputs are engineering estimates or placeholders. No study has measured r_g for human cord-blood CD34+ expansion in real microgravity.');
  wireData($('loadIpsc'), self.SCiPSC, 'Thesis: iPSC-derived heart muscle cells (cardiomyocytes), literature-informed',
    'Literature-informed, NOT validated: values come from Stem_Cell_Microeconomics.xlsx (iPSC_CM sheets). One unit = 1 million correct, working heart muscle cells. The space quality advantage (r_q) comes from small lab studies; check its rating below.');
  var KEY_IDS = ['q_E', 'r_q', 'r_g', 'g_E', 's', 'p_E', 'p_S', 'E_other', 'S_facility', 'S_other', 'L', 'M', 'C_return'];
  wireData($('loadViable'), self.SCViable, 'Space viability case: heart muscle cells on an autonomous free-flyer (projection)',
    'FORWARD-LOOKING CASE. Same heart-cell biology and Earth costs as the thesis scenario, but space uses an autonomous free-flyer (no astronaut time), projected launch prices (Starship-era), shared reentry capsules, and the upper half of the simulated-microgravity yield evidence. Every bet is listed under "Bets this case relies on".');
  var BADGE = { projection: '&#9654; Future bet', ok: '&#10003; In range', weak: '~ Guess', optimistic: '&#9650; Optimistic', conservative: '&#9660; Conservative', none: '? No evidence', out: '! Outside range' };

  // ---------- assumptions ----------
  var GROUPS = ['Product', 'Biology', 'Earth costs', 'Space costs', 'Transport'];
  function renderAssumptions() {
    var html = '';
    var ev = state.meta.evidence || {};
    GROUPS.forEach(function (g) {
      html += '<h3>' + g + '</h3><div class="tablewrap"><table><thead><tr><th>Input</th><th>Unit</th><th>Point estimate</th><th>Distribution</th><th>Min</th><th>Mode</th><th>Max</th><th>Evidence status</th><th>How reasonable?</th></tr></thead><tbody>';
      E.PARAM_DEFS.filter(function (d) { return d.group === g; }).forEach(function (d) {
        var p = state.params[d.id], dist = p.dist || { type: 'fixed' };
        var dis = dist.type === 'fixed' ? ' disabled' : '';
        var e = ev[d.id] || { status: state.meta.evidence ? 'not recorded' : 'speculative scenario (illustrative)' };
        html += '<tr data-id="' + d.id + '"><td><label for="v_' + d.id + '">' + esc(lab(d.id)) + '</label><br><small>' + esc(d.plain) + '</small></td><td>' + esc(d.unit) + '</td>' +
          '<td><input id="v_' + d.id + '" type="number" step="any" data-f="value" value="' + p.value + '"></td>' +
          '<td><select data-f="type" aria-label="Distribution for ' + esc(lab(d.id)) + '">' +
          ['fixed', 'triangular', 'uniform'].map(function (t) { return '<option' + (t === dist.type ? ' selected' : '') + '>' + t + '</option>'; }).join('') + '</select></td>' +
          '<td><input type="number" step="any" data-f="min" aria-label="Min" value="' + (dist.min != null ? dist.min : '') + '"' + dis + '></td>' +
          '<td><input type="number" step="any" data-f="mode" aria-label="Mode" value="' + (dist.mode != null ? dist.mode : '') + '"' + (dist.type === 'triangular' ? '' : ' disabled') + '></td>' +
          '<td><input type="number" step="any" data-f="max" aria-label="Max" value="' + (dist.max != null ? dist.max : '') + '"' + dis + '></td>' +
          '<td>' + esc(e.status) + (e.sources ? '<br><small>' + esc(e.sources) + '</small>' : '') + '</td>' +
          '<td>' + badge(rate(d.id)) + (e.low != null ? '<br><small>evidence: ' + num(e.low) + ' to ' + num(e.high) + '</small>' : '') + '</td></tr>';
      });
      html += '</tbody></table></div>';
    });
    $('paramTables').innerHTML = html;
  }
  $('paramTables').addEventListener('change', function (ev) {
    var el = ev.target, tr = el.closest('tr'); if (!tr || !state.params) return;
    var p = state.params[tr.getAttribute('data-id')], f = el.getAttribute('data-f');
    if (f === 'type') {
      var t = el.value;
      if (t === 'fixed') p.dist = { type: 'fixed' };
      else if (t === 'triangular') p.dist = { type: 'triangular', min: p.dist.min != null ? p.dist.min : p.value, mode: p.value, max: p.dist.max != null ? p.dist.max : p.value };
      else p.dist = { type: 'uniform', min: p.dist.min != null ? p.dist.min : p.value, max: p.dist.max != null ? p.dist.max : p.value };
      renderAssumptions();
    } else {
      var v = el.value === '' ? NaN : Number(el.value);
      if (f === 'value') { p.value = v; if (p.dist.type === 'triangular' && tr.querySelector('[data-f=mode]')) { p.dist.mode = v; tr.querySelector('[data-f=mode]').value = el.value; } }
      else { p.dist[f] = v; if (f === 'mode') { p.value = v; tr.querySelector('[data-f=value]').value = el.value; } }
    }
    refresh(); renderSim();
  });

  // ---------- plain-language answer, viability, reasonableness ----------
  function unitName() { return (state.meta.product && state.meta.product.unit) || 'unit (1 million qualified cells)'; }
  // Scenario-specific wording (from the spreadsheet) overrides the generic engine label.
  function lab(id) { var e = (state.meta.evidence || {})[id]; return (e && e.label) || E.DEF_BY_ID[id].label; }
  function sens() { var s = E.sensitivity(state.params); s.rows.forEach(function (x) { x.label = lab(x.id); }); return s; }
  function rate(id) { return E.reasonableness(state.params, id, (state.meta.evidence || {})[id]); }
  function badge(rt) { return '<span class="badge ' + rt.level + '" title="' + esc(rt.text) + '">' + BADGE[rt.level] + '</span>'; }

  function renderAnswer(r) {
    var h = '<div class="answer">';
    if (r.category === 'unavailable') h += '<p class="big">Neither process produces usable output with these inputs.</p>';
    else if (r.category === 'tie') h += '<p class="big">Space and Earth cost the same: ' + usd(r.c_E) + ' per ' + esc(unitName()) + '.</p>';
    else if (r.category === 'space') {
      h += '<p class="big">Space wins with these assumptions: ' + usd(r.c_S) + ' vs ' + usd(r.c_E) + ' on Earth per ' + esc(unitName()) + (isNaN(r.savings) ? '' : ' (' + pct(r.savings, 0) + ' cheaper)') + '.</p>';
      if (r.baselineFailure) h += '<p>Note: Earth produces nothing usable here, so this is an Earth failure rather than a space advantage.</p>';
    } else {
      var x = r.c_S / r.c_E;
      h += '<p class="big">Earth wins with these assumptions: space costs ' + usd(r.c_S) + ' vs ' + usd(r.c_E) + ' on Earth per ' + esc(unitName()) +
        (isFinite(x) ? ' (' + (x >= 1.995 ? x.toFixed(1) + ' times' : pct(x - 1, 0)) + ' more expensive).' : '.') + '</p>';
    }
    h += '<p>Before launch and return costs, making the cells in orbit would cost ' + usd(r.c_S_prod) + ' per unit vs ' + usd(r.c_E_prod) + ' on Earth.</p>';
    var counts = reasonCounts();
    if (counts.total) h += '<p>Assumption check: ' + counts.ok + ' of ' + counts.total + ' key inputs sit inside the evidence range; ' +
      (counts.projection ? counts.projection + ' are future projections; ' : '') +
      (counts.weak + counts.none) + ' are guesses with little or no evidence; ' + counts.optimistic + ' lean optimistic for space; ' + counts.conservative + ' lean conservative.</p>';
    // List the bets: every input whose support is a projection, a guess, or outside the evidence in space's favor.
    var bets = [];
    if (state.meta.evidence) E.PARAM_DEFS.forEach(function (d) {
      var rt = rate(d.id), e = state.meta.evidence[d.id];
      if (rt.level === 'projection' || rt.level === 'optimistic' || rt.level === 'weak' || (e && e.bet)) {
        bets.push('<li>' + badge(rt) + ' <strong>' + esc(lab(d.id)) + '</strong> = ' + num(state.params[d.id].value) + ' ' + esc(d.unit) + (e && e.basis ? '<br><small>' + esc(e.basis) + '</small>' : '') + '</li>');
      }
    });
    if (bets.length) h += '<details' + (r.category === 'space' ? ' open' : '') + '><summary>Bets this case relies on (' + bets.length + ')</summary><ul class="bets">' + bets.join('') + '</ul></details>';
    $('answer').innerHTML = h + '</div>';
  }

  function vrow(item, need, now, ok, note) {
    return '<tr><td>' + item + '</td><td>' + need + '</td><td class="num">' + now + '</td><td>' +
      (ok === null ? '<span class="badge none">n/a</span>' : ok ? '<span class="badge ok">&#10003; Meets</span>' : '<span class="badge optimistic">&#10007; Not met</span>') +
      (note ? '<br><small>' + note + '</small>' : '') + '</td></tr>';
  }
  function renderViability(r) {
    var v = E.pointValues(state.params), rows = '';
    if (r.L_status === 'not_applicable') rows += vrow('Launch price', 'Not applicable (no launched mass)', usd(v.L) + '/kg', null);
    else if (r.L_status === 'undefined') rows += vrow('Launch price', 'Not defined', usd(v.L) + '/kg', null);
    else rows += vrow('Launch price', r.L_max < 0 ? 'Not viable even at $0/kg' : 'at most ' + usd(r.L_max) + '/kg', usd(v.L) + '/kg', r.L_max >= 0 && v.L <= r.L_max);
    if (isNaN(r.budget)) rows += vrow('Return cost per batch', 'Not defined', usd(v.C_return), null);
    else rows += vrow('Return cost per batch', r.C_return_max < 0 ? 'Not viable even if return were free' : 'at most ' + usd(r.C_return_max), usd(v.C_return), r.C_return_max >= 0 && v.C_return <= r.C_return_max + 1e-6);
    if (!isNaN(r.budget)) rows += vrow('Orbital manufacturing cost per batch (all space production lines)', r.C_S_prod_max < 0 ? 'Not viable even if manufacturing were free' : 'at most ' + usd(r.C_S_prod_max), usd(r.C_S_prod), r.C_S_prod_max >= 0 && r.C_S_prod <= r.C_S_prod_max + 1e-6);
    if (!isNaN(r.r_q_min)) {
      var feasible = r.r_q_min <= r.r_q_cap + 1e-12;
      rows += vrow('<strong>Quality ratio r_q</strong> (share of correct cells, space / Earth)', 'at least ' + num(r.r_q_min) + 'x',
        num(v.r_q) + 'x', v.r_q >= r.r_q_min - 1e-12,
        feasible ? 'Physically possible: the most it could ever be is ' + num(r.r_q_cap) + 'x (100% correct cells in space).'
          : '<strong>Impossible from quality alone:</strong> the most it could ever be is ' + num(r.r_q_cap) + 'x, because Earth already gets ' + pct(v.q_E, 0) + ' correct cells and space cannot exceed 100%.');
    }
    if (!isNaN(r.r_g_min)) rows += vrow('Growth ratio r_g (cells grown, space / Earth)', 'at least ' + num(r.r_g_min) + 'x', num(v.r_g) + 'x', v.r_g >= r.r_g_min - 1e-12);
    var best = isFinite(r.c_S_bestq) ? '<p>Best case for the thesis: even if <strong>100% of space-made cells</strong> came out as the correct type, space would cost <strong>' + usd(r.c_S_bestq) + '</strong> per unit vs ' + usd(r.c_E) + ' on Earth.</p>' : '';
    $('viability').innerHTML = '<div class="tablewrap"><table><thead><tr><th>To match Earth, this must be...</th><th>Requirement</th><th>Current</th><th>Status</th></tr></thead><tbody>' + rows + '</tbody></table></div>' + best;
  }

  function renderMap(r) {
    var v = E.pointValues(state.params);
    if (isNaN(r.budget)) { $('viaMap').innerHTML = '<p class="note">Map unavailable: Earth or space has no usable output.</p>'; return; }
    var K = r.budget - r.C_S_prod - v.C_other_del; // launch + return must fit inside K
    var xMax = Math.max(v.L * 2, 80000, v.M > 0 && K > 0 ? (K / v.M) * 1.2 : 0);
    var yMax = Math.max(v.C_return * 2, K > 0 ? K * 1.2 : 0, 100000);
    var W = 640, H = 300, x0 = 80, x1 = W - 20, y0 = 20, y1 = H - 50;
    function X(a) { return x0 + a / xMax * (x1 - x0); }
    function Y(b) { return y1 - b / yMax * (y1 - y0); }
    var svg = '<svg viewBox="0 0 ' + W + ' ' + H + '" width="100%" role="img" aria-label="Viability map of launch price versus return cost">' +
      '<defs><clipPath id="plotclip"><rect x="' + x0 + '" y="' + y0 + '" width="' + (x1 - x0) + '" height="' + (y1 - y0) + '"/></clipPath></defs>' +
      '<rect x="' + x0 + '" y="' + y0 + '" width="' + (x1 - x0) + '" height="' + (y1 - y0) + '" fill="#f4f4f2" stroke="#999"/>';
    if (K > 0) {
      var pts = v.M > 0 ? [[0, 0], [0, K], [K / v.M, 0]] : [[0, 0], [0, K], [xMax, K], [xMax, 0]];
      svg += '<polygon clip-path="url(#plotclip)" points="' + pts.map(function (p) { return X(p[0]) + ',' + Y(p[1]); }).join(' ') + '" fill="#1f5fa8" opacity="0.25" stroke="#1f5fa8" stroke-width="2"/>';
      svg += '<text x="' + (x0 + 8) + '" y="' + (y1 - 8) + '">Space cheaper (shaded)</text>';
    } else {
      svg += '<text x="' + ((x0 + x1) / 2) + '" y="' + (y0 + 20) + '" text-anchor="middle">No viable region: space loses even with free launch and free return</text>';
      svg += '<text x="' + ((x0 + x1) / 2) + '" y="' + (y0 + 38) + '" text-anchor="middle">(over budget by ' + usd(-K) + ' per batch before any transport)</text>';
    }
    var cx = X(Math.min(v.L, xMax)), cy = Y(Math.min(v.C_return, yMax));
    svg += '<circle cx="' + cx + '" cy="' + cy + '" r="6" fill="#111"/><text x="' + (cx + 9) + '" y="' + (cy - 8) + '">current</text>';
    for (var i = 0; i <= 4; i++) {
      svg += '<text x="' + X(xMax * i / 4) + '" y="' + (y1 + 18) + '" text-anchor="middle">' + usd(xMax * i / 4) + '</text>';
      svg += '<text x="' + (x0 - 6) + '" y="' + (Y(yMax * i / 4) + 4) + '" text-anchor="end">' + usd(yMax * i / 4) + '</text>';
    }
    svg += '<text x="' + ((x0 + x1) / 2) + '" y="' + (H - 8) + '" text-anchor="middle">Launch price (USD per kg)</text>';
    svg += '<text transform="translate(14 ' + ((y0 + y1) / 2) + ') rotate(-90)" text-anchor="middle">Return cost (USD per batch)</text></svg>';
    $('viaMap').innerHTML = svg + '<p class="note">Deterministic. Everything except launch price and return cost is held at its current value (' + num(v.M) + ' kg launched per batch). ' +
      (K > 0 ? 'Space wins anywhere in the shaded triangle: launch cost + return cost must stay below ' + usd(K) + ' per batch.' : '') + '</p>';
  }

  function reasonCounts() {
    var c = { ok: 0, weak: 0, optimistic: 0, conservative: 0, none: 0, out: 0, projection: 0, total: 0 };
    if (!state.meta.evidence) return c;
    KEY_IDS.forEach(function (id) { c[rate(id).level]++; c.total++; });
    return c;
  }
  function renderKeyInputs() {
    var ev = state.meta.evidence || {};
    var html = '<div class="tablewrap"><table class="keyin"><thead><tr><th>Assumption</th><th>Your value</th><th>Evidence range</th><th>How reasonable?</th><th>Evidence strength</th></tr></thead><tbody>';
    KEY_IDS.forEach(function (id) {
      var d = E.DEF_BY_ID[id], p = state.params[id], e = ev[id], rt = rate(id);
      html += '<tr data-id="' + id + '"><td><label for="k_' + id + '"><strong>' + esc(lab(d.id)) + '</strong></label><br><small>' + esc(d.plain) + '</small></td>' +
        '<td class="num"><input id="k_' + id + '" class="ki" type="number" step="any" value="' + p.value + '"><br><small>' + esc(d.unit) + '</small></td>' +
        '<td class="num">' + (e && e.low != null ? num(e.low) + ' to ' + num(e.high) : 'none recorded') + '</td>' +
        '<td>' + badge(rt) + '<br><small>' + esc(rt.text) + '</small></td>' +
        '<td>' + esc(e ? E.evidenceStrength(e.status) : 'None recorded') + (e && e.sources ? '<br><small>Sources: ' + esc(e.sources) + '</small>' : '') + '</td></tr>';
    });
    $('keyInputs').innerHTML = html + '</tbody></table></div>';
    var c = reasonCounts();
    $('reasonSummary').innerHTML = c.total ? '<p class="headline">' + c.ok + ' in range &middot; ' + (c.projection ? c.projection + ' future bets &middot; ' : '') + (c.weak + c.none) + ' guesses &middot; ' + c.optimistic + ' optimistic &middot; ' + c.conservative + ' conservative (of ' + c.total + ' key inputs)</p>'
      : '<p class="note">This scenario has no evidence ranges attached, so inputs cannot be rated. Load a literature-informed scenario to see ratings.</p>';
  }
  $('keyInputs').addEventListener('change', function (ev) {
    var el = ev.target; if (!el.classList.contains('ki')) return;
    var id = el.closest('tr').getAttribute('data-id'), p = state.params[id];
    var v = el.value === '' ? NaN : Number(el.value);
    p.value = v;
    if (p.dist.type !== 'fixed' && isFinite(v)) {
      if (v < p.dist.min) p.dist.min = v;   // explicit: uncertainty range widened to include the new value (shown on Assumptions tab)
      if (v > p.dist.max) p.dist.max = v;
      if (p.dist.type === 'triangular') p.dist.mode = v;
    }
    renderAssumptions(); refresh(); renderSim();
  });

  function renderHinges() {
    var s = sens(), ev = state.meta.evidence || {};
    if (!s.rows.length) { $('hinges').innerHTML = '<p class="note">All inputs are fixed, so nothing to rank. Give inputs a range to see what matters most.</p>'; return; }
    $('hinges').innerHTML = '<ol>' + s.rows.slice(0, 3).map(function (row) {
      var rt = rate(row.id), e = ev[row.id];
      return '<li><strong>' + esc(row.label) + '</strong>: moving it across its range changes the space-minus-Earth cost by ' +
        (isFinite(row.swing) ? usd(row.swing) : 'an unbounded amount') + ' per unit. Evidence: ' + esc(e ? E.evidenceStrength(e.status) : 'none recorded') + ' ' + badge(rt) + '</li>';
    }).join('') + '</ol><p class="note">If the top items have weak evidence, the conclusion is fragile. Full ranking: Results tab.</p>';
  }

  // ---------- deterministic results ----------
  function refresh() {
    if (!state.params) return;
    var errs = E.validate(state.params);
    $('validation').innerHTML = errs.length ? '<div class="errbox"><strong>Fix these inputs before results can be computed:</strong><ul>' + errs.map(function (e) { return '<li>' + esc(e) + '</li>'; }).join('') + '</ul></div>' : '';
    if (errs.length) {
      $('detResults').innerHTML = $('validation').innerHTML; $('answer').innerHTML = $('validation').innerHTML;
      ['breakEven', 'sensChart', 'sensTable', 'viability', 'viaMap', 'hinges'].forEach(function (k) { $(k).innerHTML = ''; });
      renderKeyInputs(); return;
    }
    var r = E.compute(E.pointValues(state.params));
    renderAnswer(r); renderViability(r); renderMap(r); renderKeyInputs(); renderHinges();
    var verdict = { space: 'Space is cheaper at the point estimates.', earth: 'Earth is cheaper at the point estimates.', tie: 'Costs are equal within tolerance.', unavailable: 'Comparison unavailable: neither process produces usable output.' }[r.category];
    if (r.baselineFailure) verdict += ' (Earth baseline produces no usable output.)';
    $('detResults').innerHTML =
      '<p class="headline">' + verdict + ' Relative delivered savings: ' + pct(r.savings) + '.</p>' +
      '<div class="tablewrap"><table><thead><tr><th>Metric (per attempted batch unless stated)</th><th>Earth</th><th>Space</th></tr></thead><tbody>' +
      row('Qualified output, successful batch (units of U cells)', num(r.Y_E), num(r.Y_S_pre) + ' pre-return / ' + num(r.Y_S) + ' delivered') +
      row('Expected usable output Q (units)', num(r.Q_E), num(r.Q_S_pre) + ' pre-return / ' + num(r.Q_S) + ' delivered') +
      row('Production cost per batch', usd(r.C_E_prod), usd(r.C_S_prod)) +
      row('Delivered cost per batch', usd(r.C_E_del), usd(r.C_S_del) + ' (launch ' + usd(r.C_launch) + ')') +
      row('<strong>Production cost per unit</strong> (before transport / return losses)', usd(r.c_E_prod), usd(r.c_S_prod)) +
      row('<strong>Delivered cost per unit</strong> (headline)', usd(r.c_E), usd(r.c_S)) +
      '</tbody></table></div><p class="note">Biological expansion ratio r_g = ' + num(state.params.r_g.value) + ' (shown separately; causal attribution requires controlled experiments). Production relative savings: ' + pct(r.prodSavings) + '.</p>';
    renderBreakEven(r); renderSensitivity();
  }
  function row(a, b, c) { return '<tr><td>' + a + '</td><td class="num">' + b + '</td><td class="num">' + c + '</td></tr>'; }

  function renderBreakEven(r) {
    var L;
    if (r.L_status === 'not_applicable') L = 'Launch price not applicable (allocated launched mass M = 0).';
    else if (r.L_status === 'undefined') L = 'Not defined (Earth has no usable output).';
    else if (r.L_status === 'not_competitive') L = usd(r.L_max) + '/kg: Not competitive even with free outbound launch under these assumptions.';
    else L = usd(r.L_max) + '/kg (current: ' + usd(state.params.L.value) + '/kg)';
    $('breakEven').innerHTML = '<div class="tablewrap"><table><tbody>' +
      '<tr><th>Maximum launch price for space to match Earth (L_max)</th><td>' + L + '</td></tr>' +
      '<tr><th>Break-even gross expansion ratio r_g_min (all else fixed)</th><td>' + (isNaN(r.r_g_min) ? 'Not defined' : num(r.r_g_min) + ' (current r_g = ' + num(state.params.r_g.value) + ')') + '</td></tr>' +
      '<tr><th>Required expected-output ratio Q_S/Q_E &gt; C_S_del/C_E_del</th><td>' + (isNaN(r.reqRatio) ? 'Not defined' : num(r.reqRatio) + ' (current Q_S/Q_E = ' + num(r.outputRatio) + ')') + '</td></tr>' +
      '</tbody></table></div><p class="note">Deterministic, using point estimates.</p>';
  }

  // ---------- sensitivity chart (SVG tornado) ----------
  function renderSensitivity() {
    var s = sens();
    if (!s.rows.length) { $('sensChart').innerHTML = '<p>All inputs are fixed; give inputs a range (triangular or uniform) to see sensitivity.</p>'; $('sensTable').innerHTML = ''; return; }
    var rows = s.rows.slice(0, 12), finite = [];
    rows.forEach(function (r) { [r.diffLow, r.diffHigh].forEach(function (x) { if (isFinite(x)) finite.push(x); }); });
    finite.push(s.baseDiff, 0);
    var lo = Math.min.apply(null, finite), hi = Math.max.apply(null, finite);
    if (lo === hi) { lo -= 1; hi += 1; }
    var W = 760, labelW = 290, H = rows.length * 26 + 50, x0 = labelW + 10, x1 = W - 20;
    function X(v) { return x0 + (Math.max(lo, Math.min(hi, v)) - lo) / (hi - lo) * (x1 - x0); }
    var svg = '<svg viewBox="0 0 ' + W + ' ' + H + '" width="100%" role="img" aria-label="Tornado chart of local sensitivity of delivered cost difference">';
    rows.forEach(function (r, i) {
      var y = 20 + i * 26;
      svg += '<text x="' + labelW + '" y="' + (y + 13) + '" text-anchor="end">' + esc(r.label.length > 44 ? r.label.slice(0, 43) + '…' : r.label) + '</text>';
      var a = X(r.diffLow), b = X(r.diffHigh);
      svg += '<rect x="' + Math.min(a, b) + '" y="' + y + '" width="' + Math.max(1, Math.abs(b - a)) + '" height="18" fill="#1f5fa8" opacity="0.8"></rect>';
      svg += '<circle cx="' + a + '" cy="' + (y + 9) + '" r="4" fill="#fff" stroke="#111"><title>low</title></circle>';
      svg += '<rect x="' + (b - 4) + '" y="' + (y + 5) + '" width="8" height="8" fill="#111"><title>high</title></rect>';
    });
    var zx = X(0), bx = X(s.baseDiff);
    svg += '<line x1="' + zx + '" x2="' + zx + '" y1="10" y2="' + (H - 25) + '" stroke="#111" stroke-dasharray="4 3"></line>';
    svg += '<text x="' + zx + '" y="' + (H - 10) + '" text-anchor="middle">0 (break-even)</text>';
    svg += '<line x1="' + bx + '" x2="' + bx + '" y1="10" y2="' + (H - 25) + '" stroke="#c98a00"></line>';
    svg += '<text x="' + x0 + '" y="12">' + usd(lo) + '</text><text x="' + x1 + '" y="12" text-anchor="end">' + usd(hi) + '</text></svg>';
    $('sensChart').innerHTML = svg + '<p class="note">Circle = input at its low value; square = input at its high value. Orange line = point-estimate difference (' + usd(s.baseDiff) + '/unit). Values beyond the axis or infinite are clipped; see table.</p>';
    $('sensTable').innerHTML = '<div class="tablewrap"><table><thead><tr><th>Input</th><th>Low</th><th>High</th><th>c_S - c_E at low (USD/unit)</th><th>c_S - c_E at high (USD/unit)</th></tr></thead><tbody>' +
      s.rows.map(function (r) { return '<tr><td>' + esc(r.label) + '</td><td class="num">' + num(r.low) + '</td><td class="num">' + num(r.high) + '</td><td class="num">' + usd(r.diffLow) + '</td><td class="num">' + usd(r.diffHigh) + '</td></tr>'; }).join('') + '</tbody></table></div>';
  }

  // ---------- simulation ----------
  $('runSim').onclick = function () {
    if (!state.params) { alert('Load or create a scenario first.'); return; }
    var errs = E.validate(state.params);
    if (errs.length) { $('simResults').innerHTML = '<div class="errbox">Simulation not run. ' + errs.map(esc).join(' ') + '</div>'; return; }
    var seed = Number($('seed').value);
    if (!Number.isInteger(seed) || seed < 0) { $('simResults').innerHTML = '<div class="errbox">Seed must be a non-negative integer.</div>'; return; }
    var opts = { draws: Number($('draws').value), seed: seed, timestamp: new Date().toISOString() };
    var snapshot = JSON.stringify(state.params);
    $('runSim').disabled = true; $('cancelSim').disabled = false; $('simProgress').hidden = false; $('simProgress').value = 0;
    var t0 = performance.now();
    var w;
    try { w = new Worker('worker.js'); } catch (e) { w = null; }
    if (!w) { // fallback (e.g. opened from file://)
      var res = E.simulate(JSON.parse(snapshot), opts); finish(res, snapshot, performance.now() - t0); return;
    }
    state.worker = w;
    w.onmessage = function (m) {
      var d = m.data;
      if (d.type === 'progress') $('simProgress').value = d.fraction;
      else if (d.type === 'error') { stopWorker(); $('simResults').innerHTML = '<div class="errbox">' + d.errors.map(esc).join('<br>') + '</div>'; }
      else if (d.type === 'done') { stopWorker(); finish(d.result, snapshot, performance.now() - t0); }
    };
    w.postMessage({ params: JSON.parse(snapshot), opts: opts });
  };
  $('cancelSim').onclick = function () { stopWorker(); $('simResults').innerHTML = '<p>Simulation cancelled. ' + (state.run ? 'Previous results are not shown.' : '') + '</p>'; state.run = null; renderSimStatus(); };
  function stopWorker() {
    if (state.worker) { state.worker.terminate(); state.worker = null; }
    $('runSim').disabled = false; $('cancelSim').disabled = true; $('simProgress').hidden = true;
  }
  function finish(res, snapshot, ms) { res.elapsedMs = ms; state.run = res; state.runSnapshot = snapshot; renderSim(); }

  function renderSimStatus() {
    var stale = state.run && state.runSnapshot !== JSON.stringify(state.params);
    var txt = !state.run ? 'No simulation has been run.' :
      (stale ? 'Results are OUT OF DATE: assumptions changed since the last run. Re-run on the Results tab.' :
        'Space is cheaper in ' + pct(state.run.counts.space / state.run.draws, 1) + ' of modeled scenarios under the selected assumptions (' + state.run.draws.toLocaleString() + ' draws, seed ' + state.run.seed + ').');
    $('simStatusOverview').innerHTML = stale ? '<div class="stale">' + txt + '</div>' : esc(txt);
    return stale;
  }
  function renderSim() {
    var stale = renderSimStatus();
    var r = state.run;
    if (!r) { $('simResults').innerHTML = '<p>No simulation has been run.</p>'; return; }
    var c = r.counts, n = r.draws;
    var html = stale ? '<div class="stale">OUT OF DATE: these results were computed from earlier assumptions. Re-run to update.</div>' : '';
    html += '<p class="headline">Space is cheaper in ' + pct(c.space / n, 1) + ' of modeled scenarios under the selected assumptions.</p>';
    html += '<p class="note">This is a conditional model output, not the probability that space manufacturing will succeed. Stochastic inputs: ' + (r.stochasticIds.length ? r.stochasticIds.join(', ') : 'none (all fixed)') + '. Model ' + esc(r.modelVersion) + ', ' + n.toLocaleString() + ' draws, seed ' + r.seed + ', run ' + esc(r.timestamp) + ', ' + Math.round(r.elapsedMs) + ' ms.</p>';
    html += '<div class="tablewrap"><table><thead><tr><th>Outcome (all ' + n.toLocaleString() + ' draws)</th><th>Draws</th><th>Share</th></tr></thead><tbody>' +
      ['space', 'tie', 'earth', 'unavailable'].map(function (k) {
        var lab = { space: 'Space cheaper', tie: 'Tie (within 1e-9 relative)', earth: 'Earth cheaper', unavailable: 'Comparison unavailable (no output in either)' }[k];
        return '<tr><td>' + lab + '</td><td class="num">' + c[k].toLocaleString() + '</td><td class="num">' + pct(c[k] / n, 2) + '</td></tr>';
      }).join('') + '</tbody></table></div>';
    if (c.baselineFailure) html += '<p class="note">' + c.baselineFailure + ' space-cheaper draws are Earth baseline failures (Earth produced no usable output).</p>';
    html += '<div class="tablewrap"><table><thead><tr><th>Delivered cost per unit (model percentile ranges)</th><th>5th</th><th>Median</th><th>95th</th><th>No-output draws</th></tr></thead><tbody>' +
      '<tr><td>Earth</td><td class="num">' + bound(r.earth.p05) + '</td><td class="num">' + bound(r.earth.p50) + '</td><td class="num">' + bound(r.earth.p95) + '</td><td class="num">' + c.infE + '</td></tr>' +
      '<tr><td>Space</td><td class="num">' + bound(r.space.p05) + '</td><td class="num">' + bound(r.space.p50) + '</td><td class="num">' + bound(r.space.p95) + '</td><td class="num">' + c.infS + '</td></tr></tbody></table></div>';
    html += '<p>Median relative delivered savings: <strong>' + pct(r.savingsMedian) + '</strong> (denominator: ' + r.savingsN.toLocaleString() + ' of ' + n.toLocaleString() + ' draws with defined positive comparison costs).</p>';
    html += costChart(r);
    $('simResults').innerHTML = html;
  }

  // Chart A: delivered cost comparison with 5-95 percentile bands (log scale)
  function costChart(r) {
    var vals = [r.earth.p05, r.earth.p50, r.earth.p95, r.space.p05, r.space.p50, r.space.p95].filter(function (x) { return isFinite(x) && x > 0; });
    if (vals.length < 2) return '<p class="note">Cost chart unavailable: percentiles not finite and positive.</p>';
    var lo = Math.log10(Math.min.apply(null, vals)) - 0.1, hi = Math.log10(Math.max.apply(null, vals)) + 0.1;
    var W = 700, x0 = 90, x1 = W - 30;
    function X(v) { return x0 + (Math.log10(v) - lo) / (hi - lo) * (x1 - x0); }
    var svg = '<svg viewBox="0 0 ' + W + ' 130" width="100%" role="img" aria-label="Delivered cost per unit, median and 5th to 95th model percentile range, log scale">';
    [['Earth', r.earth, 30], ['Space', r.space, 70]].forEach(function (a) {
      var q = a[1], y = a[2];
      svg += '<text x="' + (x0 - 10) + '" y="' + (y + 5) + '" text-anchor="end">' + a[0] + '</text>';
      if (isFinite(q.p05) && isFinite(q.p95) && q.p05 > 0) svg += '<line x1="' + X(q.p05) + '" x2="' + X(q.p95) + '" y1="' + y + '" y2="' + y + '" stroke="#1f5fa8" stroke-width="8" opacity="0.35"></line>';
      if (isFinite(q.p50) && q.p50 > 0) svg += '<circle cx="' + X(q.p50) + '" cy="' + y + '" r="6" fill="#1f5fa8"></circle>';
      if (!isFinite(q.p95)) svg += '<text x="' + x1 + '" y="' + (y - 8) + '" text-anchor="end">95th: no finite bound</text>';
    });
    for (var e = Math.ceil(lo); e <= Math.floor(hi); e++) {
      svg += '<line x1="' + X(Math.pow(10, e)) + '" x2="' + X(Math.pow(10, e)) + '" y1="90" y2="96" stroke="#111"></line><text x="' + X(Math.pow(10, e)) + '" y="112" text-anchor="middle">' + usd(Math.pow(10, e)) + '</text>';
    }
    svg += '<text x="' + x0 + '" y="128">USD per unit (log scale)</text></svg>';
    return '<h3>Chart: simulated delivered cost per unit (probabilistic)</h3>' + svg + '<p class="note">Dot = median; band = 5th to 95th model percentile range. Assumptions: the snapshot of this run.</p>';
  }

  // ---------- evidence ----------
  var PRD_EVIDENCE = [
    { id: 'PRD-1', title: 'Microgravity Reduces the Differentiation and Regenerative Potential of Embryonic Stem Cells', year: 2015, url: 'https://pmc.ncbi.nlm.nih.gov/articles/PMC4652210/', species: 'Mouse', cell: 'Embryonic stem-cell clusters', gravity: 'Actual spaceflight, 15 days', finding: 'Reduced differentiation and retention of stem-cell markers, with subsequent differentiation after return.', direction: 'contextual', limits: 'Supports investigating preservation of characteristics, not human HSC expansion or lower cost.' },
    { id: 'PRD-2', title: 'Spaceflight/microgravity inhibits the proliferation of hematopoietic stem cells by decreasing Kit-Ras/cAMP-CREB pathway networks as evidenced by RNA-Seq assays', year: 2019, url: 'https://pmc.ncbi.nlm.nih.gov/articles/PMC6463920/', species: 'Mouse', cell: 'Hematopoietic stem/progenitor cells', gravity: 'Actual and simulated', finding: 'Reduced proliferation; differentiation ability not changed in the reported comparison.', direction: 'contradictory', limits: 'Counterevidence to a general faster-growth claim. Mouse cells.' },
    { id: 'PRD-3', title: 'NASA Selects Proposals to Enable Manufacturing In Space for Earth', year: 'not extracted', url: 'https://www.nasa.gov/missions/station/nasa-selects-proposals-to-enable-manufacturing-in-space-for-earth/', species: 'not extracted', cell: 'Human HSC (proposed)', gravity: 'Proposed', finding: 'Program description of proposed HSC expansion work.', direction: 'contextual', limits: 'Not a measured commercial yield or cost result.' }
  ];
  function renderEvidence() {
    var list = PRD_EVIDENCE.concat(self.SCLiterature ? SCLiterature.sources : []);
    $('evidenceList').innerHTML = '<p class="note">' + list.length + ' records (inventory only). Fields not available read "not extracted".</p>' + list.map(function (s) {
      return '<div class="evidence"><strong>' + esc(s.id) + ': ' + esc(s.title) + '</strong> (' + esc(s.year || 'not extracted') + ')<br>' +
        '<span class="tag">' + esc(s.direction || 'contextual') + '</span>' + (s.verification ? '<span class="tag">' + esc(s.verification) + '</span>' : '') +
        (s.species ? '<span class="tag">' + esc(s.species) + '</span>' : '') + '<br>' +
        (s.cell ? 'System: ' + esc(s.cell) + (s.gravity ? '; ' + esc(s.gravity) : '') + '<br>' : '') +
        'Finding: ' + esc(s.finding) + '<br>Limitations: ' + esc(s.limits || 'not extracted') +
        (s.params ? '<br>Linked inputs: ' + esc(s.params) : '') +
        '<br><a href="' + esc(s.url) + '" target="_blank" rel="noopener noreferrer">' + esc(s.url) + '</a></div>';
    }).join('');
  }

  $('modelVersion').textContent = E.MODEL_VERSION;
  renderEvidence();
})();

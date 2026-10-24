/*
 * Stem Cell Microgravity Feasibility Calculator - numerical engine.
 * Pure functions only: no DOM, no storage. Loaded by the page, the worker and the tests.
 * Equations follow PRD sections 6, 7 and 10.
 */
(function (root) {
  'use strict';

  var MODEL_VERSION = 'mvp-0.1.0';
  // Relative tolerance used to declare a tie between two per-unit costs (PRD 6: documented tolerance).
  var TIE_TOLERANCE = 1e-9;

  // ---- Parameter catalogue (ids, units, validation kind) ----
  // kind: 'nonneg' (>=0), 'fraction' ([0,1]), 'positive' (>0)
  var PARAM_DEFS = [
    // Biology / output
    { id: 'U', label: 'Cells per reporting unit (U)', unit: 'cells/unit', group: 'Product', kind: 'positive' },
    { id: 'N0', label: 'Viable starting cells per batch (N0_E = N0_S)', unit: 'cells', group: 'Biology', kind: 'nonneg' },
    { id: 'g_E', label: 'Earth gross expansion factor (g_E)', unit: 'fold', group: 'Biology', kind: 'nonneg' },
    { id: 'r_g', label: 'Space/Earth expansion ratio (r_g)', unit: 'ratio', group: 'Biology', kind: 'nonneg' },
    { id: 'q_E', label: 'Earth qualified fraction (q_E)', unit: 'fraction', group: 'Biology', kind: 'fraction' },
    { id: 'r_q', label: 'Space/Earth qualification ratio (r_q)', unit: 'ratio', group: 'Biology', kind: 'nonneg' },
    { id: 's', label: 'Post-return retention of qualified output (s)', unit: 'fraction', group: 'Transport', kind: 'fraction' },
    { id: 'p_E', label: 'Earth batch release probability (p_E)', unit: 'probability', group: 'Biology', kind: 'fraction' },
    { id: 'p_S', label: 'Space batch release probability (p_S)', unit: 'probability', group: 'Biology', kind: 'fraction' },
    // Earth costs per attempted batch
    { id: 'E_cells', label: 'Earth: starting-cell acquisition', unit: 'USD/batch', group: 'Earth costs', kind: 'nonneg' },
    { id: 'E_consumables', label: 'Earth: media, cytokines, consumables', unit: 'USD/batch', group: 'Earth costs', kind: 'nonneg' },
    { id: 'E_labor', label: 'Earth: labor', unit: 'USD/batch', group: 'Earth costs', kind: 'nonneg' },
    { id: 'E_facility', label: 'Earth: GMP facility allocation', unit: 'USD/batch', group: 'Earth costs', kind: 'nonneg' },
    { id: 'E_equipment', label: 'Earth: equipment allocation', unit: 'USD/batch', group: 'Earth costs', kind: 'nonneg' },
    { id: 'E_qc', label: 'Earth: quality testing / release', unit: 'USD/batch', group: 'Earth costs', kind: 'nonneg' },
    { id: 'E_other', label: 'Earth: other production cost', unit: 'USD/batch', group: 'Earth costs', kind: 'nonneg' },
    { id: 'E_delivery', label: 'Earth: delivery cost', unit: 'USD/batch', group: 'Earth costs', kind: 'nonneg' },
    // Space costs per attempted batch
    { id: 'S_cells', label: 'Space: starting-cell acquisition', unit: 'USD/batch', group: 'Space costs', kind: 'nonneg' },
    { id: 'S_consumables', label: 'Space: media, cytokines, consumables', unit: 'USD/batch', group: 'Space costs', kind: 'nonneg' },
    { id: 'S_labor', label: 'Space: operations / crew & ground labor', unit: 'USD/batch', group: 'Space costs', kind: 'nonneg' },
    { id: 'S_facility', label: 'Space: orbital facility allocation', unit: 'USD/batch', group: 'Space costs', kind: 'nonneg' },
    { id: 'S_equipment', label: 'Space: flight hardware allocation', unit: 'USD/batch', group: 'Space costs', kind: 'nonneg' },
    { id: 'S_qc', label: 'Space: quality testing / release', unit: 'USD/batch', group: 'Space costs', kind: 'nonneg' },
    { id: 'S_other', label: 'Space: other production cost', unit: 'USD/batch', group: 'Space costs', kind: 'nonneg' },
    { id: 'L', label: 'Launch price (L)', unit: 'USD/kg', group: 'Transport', kind: 'nonneg' },
    { id: 'M', label: 'Allocated launched mass (M)', unit: 'kg/batch', group: 'Transport', kind: 'nonneg' },
    { id: 'C_return', label: 'Return and recovery cost', unit: 'USD/batch', group: 'Transport', kind: 'nonneg' },
    { id: 'C_other_del', label: 'Other space delivery cost', unit: 'USD/batch', group: 'Transport', kind: 'nonneg' }
  ];
  // One-sentence plain-English meaning of each input (shown in the UI).
  var PLAIN = {
    U: 'How many cells count as one "unit" of product. Costs are reported per unit.',
    N0: 'How many usable cells you start each batch with (same on Earth and in space).',
    g_E: 'How many times the cell number multiplies during one Earth batch (e.g. 20 = 20x more cells at the end).',
    r_g: 'Growth in space compared with Earth. 1 = same growth; 2 = space grows twice as many cells; 0.5 = half as many.',
    q_E: 'Share of the final Earth cells that are the correct, working cell type and pass quality checks.',
    r_q: 'Quality in space compared with Earth: how many times more of the cells come out as the correct type. 1 = no difference; 1.2 = 20% more. Can never push the space share above 100%.',
    s: 'Share of good space-made cells still good after the trip home (re-entry, storage, thawing).',
    p_E: 'Chance an Earth batch passes final release testing (a failed batch produces nothing usable).',
    p_S: 'Chance a space batch succeeds and is released (includes launch delays, hardware faults, contamination).',
    E_cells: 'Cost of the starting cells for one Earth batch.',
    E_consumables: 'Growth media, growth factors and single-use plastics for one Earth batch.',
    E_labor: 'Staff time for one Earth batch.',
    E_facility: 'Share of clean-room (GMP facility) running costs charged to one batch.',
    E_equipment: 'Share of bioreactor and lab equipment purchase cost charged to one batch.',
    E_qc: 'Testing needed to release one Earth batch.',
    E_other: 'Any other Earth manufacturing cost per batch.',
    E_delivery: 'Cost to ship the finished Earth product.',
    S_cells: 'Cost of the starting cells for one space batch.',
    S_consumables: 'Media and consumables flown for one space batch.',
    S_labor: 'Astronaut and ground-team time for one space batch.',
    S_facility: 'Fee for using the orbital lab (power, space, integration) per batch.',
    S_equipment: 'Share of flight-qualified bioreactor cost charged to one batch.',
    S_qc: 'Testing needed to release one space batch.',
    S_other: 'Any other cost for the space batch, including ground processing before launch and after return.',
    L: 'Price to launch 1 kg to the orbital lab.',
    M: 'Kilograms launched per batch (hardware share, media, packaging).',
    C_return: 'Cost to bring one batch back to Earth and recover it.',
    C_other_del: 'Any other cost to deliver the space product.'
  };
  var DEF_BY_ID = {};
  PARAM_DEFS.forEach(function (d) { d.plain = PLAIN[d.id] || ''; DEF_BY_ID[d.id] = d; });

  var EARTH_PROD = ['E_cells', 'E_consumables', 'E_labor', 'E_facility', 'E_equipment', 'E_qc', 'E_other'];
  var SPACE_PROD = ['S_cells', 'S_consumables', 'S_labor', 'S_facility', 'S_equipment', 'S_qc', 'S_other'];

  // ---- Seeded PRNG: mulberry32 (32-bit state, period 2^32). Documented in Methods. ----
  function mulberry32(seed) {
    var a = seed >>> 0;
    return function () {
      a = (a + 0x6D2B79F5) >>> 0;
      var t = a;
      t = Math.imul(t ^ (t >>> 15), t | 1);
      t ^= t + Math.imul(t ^ (t >>> 7), t | 61);
      return ((t ^ (t >>> 14)) >>> 0) / 4294967296;
    };
  }

  // ---- Distributions ----
  function sampleDist(dist, pointValue, u) {
    if (!dist || dist.type === 'fixed') return pointValue;
    if (dist.type === 'uniform') return dist.min + (dist.max - dist.min) * u;
    if (dist.type === 'triangular') {
      var a = dist.min, c = dist.mode, b = dist.max;
      if (b === a) return a;
      var F = (c - a) / (b - a);
      if (u < F) return a + Math.sqrt(u * (b - a) * (c - a));
      return b - Math.sqrt((1 - u) * (b - a) * (b - c));
    }
    throw new Error('Unknown distribution ' + dist.type);
  }

  function distBounds(p) {
    var d = p.dist;
    if (!d || d.type === 'fixed') return { low: p.value, high: p.value };
    return { low: d.min, high: d.max };
  }

  // ---- Validation (PRD 12) ----
  function validate(params) {
    var errors = [];
    PARAM_DEFS.forEach(function (def) {
      var p = params[def.id];
      if (!p) { errors.push(def.label + ': missing.'); return; }
      var vals = [p.value];
      var d = p.dist || { type: 'fixed' };
      if (d.type === 'triangular') vals.push(d.min, d.mode, d.max);
      if (d.type === 'uniform') vals.push(d.min, d.max);
      for (var i = 0; i < vals.length; i++) {
        var v = vals[i];
        if (typeof v !== 'number' || !isFinite(v)) { errors.push(def.label + ': every value must be a finite number.'); return; }
        if (def.kind === 'nonneg' && v < 0) { errors.push(def.label + ': cannot be negative.'); return; }
        if (def.kind === 'fraction' && (v < 0 || v > 1)) { errors.push(def.label + ': must be between 0 and 1.'); return; }
        if (def.kind === 'positive' && v <= 0) { errors.push(def.label + ': must be greater than 0.'); return; }
      }
      if (d.type === 'triangular' && !(d.min <= d.mode && d.mode <= d.max && d.min < d.max)) {
        errors.push(def.label + ': triangular needs min <= mode <= max and min < max.');
      }
      if (d.type === 'uniform' && !(d.min < d.max)) errors.push(def.label + ': uniform needs min < max.');
      if (d.type !== 'fixed' && (p.value < d.min || p.value > d.max)) {
        errors.push(def.label + ': point estimate must lie inside the distribution range.');
      }
    });
    // q_S = q_E * r_q must stay within [0,1] over the whole support (never clipped, never rejection-sampled).
    if (params.q_E && params.r_q && errors.length === 0) {
      var qEmax = distBounds(params.q_E).high, rqMax = distBounds(params.r_q).high, qMax = qEmax * rqMax;
      if (qMax > 1 + 1e-12) {
        errors.push('Impossible combination: Earth gets up to ' + (qEmax * 100).toFixed(0) + '% correct cells, so a quality ratio of ' + rqMax +
          ' would mean ' + (qMax * 100).toFixed(0) + '% of space cells are correct (max(q_E) x max(r_q) = ' + qMax.toFixed(4) + ' exceeds 1). ' +
          'The highest quality ratio allowed is ' + (qEmax > 0 ? (1 / qEmax).toFixed(3) : 'unbounded') +
          '. Lower the quality ratio r_q (value or Max), or lower Earth purity q_E.');
      }
    }
    return errors;
  }

  // ---- Deterministic model for one set of values (PRD 6, 7) ----
  function sum(v, ids) { var t = 0; for (var i = 0; i < ids.length; i++) t += v[ids[i]]; return t; }

  function compute(v) {
    var U = v.U;
    var g_S = v.g_E * v.r_g;              // derived once
    var q_S = v.q_E * v.r_q;              // derived once
    var Y_E = v.N0 * v.g_E * v.q_E / U;
    var Y_S_pre = v.N0 * g_S * q_S / U;
    var Y_S = Y_S_pre * v.s;
    var Q_E = v.p_E * Y_E;
    var Q_S_pre = v.p_S * Y_S_pre;
    var Q_S = v.p_S * Y_S;

    var C_E_prod = sum(v, EARTH_PROD);
    var C_S_prod = sum(v, SPACE_PROD);
    var C_launch = v.L * v.M;
    var C_E_del = C_E_prod + v.E_delivery;
    var C_S_del = C_S_prod + C_launch + v.C_return + v.C_other_del;

    var c_E_prod = perUnit(C_E_prod, Q_E);
    var c_S_prod = perUnit(C_S_prod, Q_S_pre);
    var c_E = perUnit(C_E_del, Q_E);
    var c_S = perUnit(C_S_del, Q_S);

    var cmp = categorize(c_E, c_S, Q_E, Q_S);
    var savings = (Q_E > 0 && Q_S > 0 && c_E > 0) ? 1 - c_S / c_E : NaN;
    var prodSavings = (Q_E > 0 && Q_S_pre > 0 && c_E_prod > 0) ? 1 - c_S_prod / c_E_prod : NaN;

    // Break-even launch price (PRD 7)
    var L_max = NaN, L_status;
    if (v.M === 0) L_status = 'not_applicable';
    else if (!(Q_E > 0) || !isFinite(c_E)) L_status = 'undefined';
    else {
      L_max = (c_E * Q_S - C_S_prod - v.C_return - v.C_other_del) / v.M;
      L_status = L_max < 0 ? 'not_competitive' : 'ok';
    }
    // Break-even gross expansion ratio (closed form valid: no dependency links r_g to costs in this MVP)
    var denom = c_E * v.p_S * v.N0 * v.g_E * q_S * v.s;
    var r_g_min = (isFinite(c_E) && c_E > 0 && denom > 0) ? C_S_del * U / denom : NaN;
    // Required expected-output ratio Q_S/Q_E > C_S_del/C_E_del
    var reqRatio = (C_E_del > 0 && C_S_del > 0) ? C_S_del / C_E_del : NaN;

    // Break-even qualification ratio (team thesis lever), holding r_g and everything else fixed.
    var denomQ = c_E * v.p_S * v.N0 * g_S * v.q_E * v.s;
    var r_q_min = (isFinite(c_E) && c_E > 0 && denomQ > 0) ? C_S_del * U / denomQ : NaN;
    // Physical ceiling: q_S = q_E * r_q cannot exceed 1, so r_q <= 1/q_E.
    var r_q_cap = v.q_E > 0 ? 1 / v.q_E : NaN;
    // Best case for quality: every space cell qualifies (q_S = 1).
    var Q_S_bestq = v.p_S * v.N0 * g_S * v.s / U;
    var c_S_bestq = perUnit(C_S_del, Q_S_bestq);
    // Viability budgets: the most each cost could be (others fixed) for space to match Earth.
    var budget = (isFinite(c_E) && Q_S > 0) ? c_E * Q_S : NaN;       // max total space delivered spend per batch
    var C_S_prod_max = budget - C_launch - v.C_return - v.C_other_del;
    var C_return_max = budget - C_S_prod - C_launch - v.C_other_del;

    return {
      g_S: g_S, q_S: q_S, Y_E: Y_E, Y_S_pre: Y_S_pre, Y_S: Y_S, Q_E: Q_E, Q_S_pre: Q_S_pre, Q_S: Q_S,
      C_E_prod: C_E_prod, C_S_prod: C_S_prod, C_launch: C_launch, C_E_del: C_E_del, C_S_del: C_S_del,
      c_E_prod: c_E_prod, c_S_prod: c_S_prod, c_E: c_E, c_S: c_S,
      category: cmp.category, baselineFailure: cmp.baselineFailure,
      savings: savings, prodSavings: prodSavings,
      L_max: L_max, L_status: L_status, r_g_min: r_g_min, reqRatio: reqRatio,
      r_q_min: r_q_min, r_q_cap: r_q_cap, c_S_bestq: c_S_bestq,
      budget: budget, C_S_prod_max: C_S_prod_max, C_return_max: C_return_max,
      outputRatio: Q_E > 0 ? Q_S / Q_E : NaN
    };
  }

  // Zero expected output with any cost -> no finite cost (Infinity flag). Zero cost with zero output -> also no usable output.
  function perUnit(cost, q) { return q > 0 ? cost / q : Infinity; }

  function categorize(c_E, c_S, Q_E, Q_S) {
    if (!(Q_E > 0) && !(Q_S > 0)) return { category: 'unavailable', baselineFailure: false };
    if (!(Q_S > 0)) return { category: 'earth', baselineFailure: false };
    if (!(Q_E > 0)) return { category: 'space', baselineFailure: true };
    var scale = Math.max(Math.abs(c_E), Math.abs(c_S));
    if (Math.abs(c_S - c_E) <= TIE_TOLERANCE * scale) return { category: 'tie', baselineFailure: false };
    return { category: c_S < c_E ? 'space' : 'earth', baselineFailure: false };
  }

  function pointValues(params) {
    var v = {};
    PARAM_DEFS.forEach(function (d) { v[d.id] = params[d.id].value; });
    return v;
  }

  // ---- Quantiles that respect infinite tails (PRD 10) ----
  // sorted: ascending array that may contain Infinity. Returns Infinity when the quantile touches the infinite part.
  function quantile(sorted, p) {
    var n = sorted.length;
    if (n === 0) return NaN;
    var h = (n - 1) * p, lo = Math.floor(h), hi = Math.ceil(h);
    var a = sorted[lo], b = sorted[hi];
    if (!isFinite(a) || !isFinite(b)) return Infinity;
    return a + (b - a) * (h - lo);
  }

  // ---- Monte Carlo (PRD 6 statistical interpretation, PRD 10) ----
  // Each draw samples long-run parameters and computes expected expenditure / expected qualified output.
  function simulate(params, opts, onProgress) {
    var n = opts.draws, seed = opts.seed >>> 0;
    var rng = mulberry32(seed);
    var stochastic = PARAM_DEFS.filter(function (d) { return params[d.id].dist && params[d.id].dist.type !== 'fixed'; });
    var base = pointValues(params);
    var cE = new Float64Array(n), cS = new Float64Array(n), sav = [];
    var counts = { space: 0, tie: 0, earth: 0, unavailable: 0, baselineFailure: 0, infE: 0, infS: 0 };
    var draws = opts.keepDraws ? [] : null;
    var v = {}, k;
    for (k in base) v[k] = base[k];
    var chunk = 2000;
    for (var i = 0; i < n; i++) {
      // Fixed parameter order (PARAM_DEFS order) guarantees reproducibility for a given seed.
      for (var j = 0; j < stochastic.length; j++) {
        var p = params[stochastic[j].id];
        v[stochastic[j].id] = sampleDist(p.dist, p.value, rng());
      }
      var r = compute(v);
      cE[i] = r.c_E; cS[i] = r.c_S;
      counts[r.category]++;
      if (r.baselineFailure) counts.baselineFailure++;
      if (!isFinite(r.c_E)) counts.infE++;
      if (!isFinite(r.c_S)) counts.infS++;
      if (!isNaN(r.savings)) sav.push(r.savings);
      if (draws) draws.push({ r_g: v.r_g, c_E: r.c_E, c_S: r.c_S, category: r.category });
      if (onProgress && (i + 1) % chunk === 0) onProgress((i + 1) / n);
    }
    var sE = Array.prototype.slice.call(cE).sort(numAsc);
    var sS = Array.prototype.slice.call(cS).sort(numAsc);
    var ratio = [];
    for (i = 0; i < n; i++) if (isFinite(cE[i]) && isFinite(cS[i]) && cE[i] > 0) ratio.push(cS[i] / cE[i]);
    sav.sort(numAsc);
    return {
      modelVersion: MODEL_VERSION, seed: seed, draws: n, timestamp: opts.timestamp || null,
      stochasticIds: stochastic.map(function (d) { return d.id; }),
      counts: counts,
      earth: { p05: quantile(sE, 0.05), p50: quantile(sE, 0.5), p95: quantile(sE, 0.95) },
      space: { p05: quantile(sS, 0.05), p50: quantile(sS, 0.5), p95: quantile(sS, 0.95) },
      savingsMedian: sav.length ? quantile(sav, 0.5) : NaN, savingsN: sav.length,
      ratios: ratio, ratioUnavailable: n - ratio.length,
      drawList: draws
    };
  }
  function numAsc(a, b) { return a - b; }

  // ---- One-at-a-time local sensitivity (PRD 10 chart D) on delivered cost difference c_S - c_E ----
  function sensitivity(params) {
    var base = pointValues(params);
    var b = compute(base);
    var baseDiff = b.c_S - b.c_E;
    var rows = [];
    PARAM_DEFS.forEach(function (d) {
      var bnd = distBounds(params[d.id]);
      if (bnd.low === bnd.high) return;
      var lo = {}, hi = {}, k;
      for (k in base) { lo[k] = base[k]; hi[k] = base[k]; }
      lo[d.id] = bnd.low; hi[d.id] = bnd.high;
      var rl = compute(lo), rh = compute(hi);
      rows.push({ id: d.id, label: d.label, low: bnd.low, high: bnd.high,
        diffLow: rl.c_S - rl.c_E, diffHigh: rh.c_S - rh.c_E,
        swing: Math.abs((rh.c_S - rh.c_E) - (rl.c_S - rl.c_E)) });
    });
    rows.sort(function (x, y) { return (isFinite(y.swing) ? y.swing : 1e308) - (isFinite(x.swing) ? x.swing : 1e308); });
    return { baseDiff: baseDiff, rows: rows };
  }

  // Direction check: +1 if raising this input (others fixed) makes space relatively cheaper, -1 if it helps Earth, 0 if no effect.
  function favorsSpace(params, id) {
    var base = pointValues(params), up = {}, k;
    for (k in base) up[k] = base[k];
    var step = Math.abs(base[id]) * 0.01 || 0.01;
    up[id] = base[id] + step;
    var a = compute(base), b = compute(up);
    var da = a.c_S - a.c_E, db = b.c_S - b.c_E;
    if (!isFinite(da) || !isFinite(db) || Math.abs(db - da) < 1e-12 * Math.max(1, Math.abs(da))) return 0;
    return db < da ? 1 : -1;
  }

  // Reasonableness of a user value against the evidence range recorded for it.
  // evidence: { low, high, status } or undefined. Returns { level, text }.
  var STRENGTH = {
    'measured in target system': 'Strong', 'measured in target system (derived)': 'Strong', 'measured in target system (price)': 'Strong',
    'measured in another system': 'Moderate', 'engineering estimate': 'Weak', 'engineering estimate (derived)': 'Weak',
    'expert judgment': 'Very weak', 'speculative scenario': 'Very weak', 'projection': 'Projection (future bet)'
  };
  function evidenceStrength(status) { return STRENGTH[status] || 'None recorded'; }
  function reasonableness(params, id, evidence) {
    var v = params[id].value;
    if (!evidence || evidence.low == null || evidence.high == null) return { level: 'none', text: 'No evidence range recorded: treat as an assumption' };
    var st = evidenceStrength(evidence.status);
    if (evidence.status === 'projection') {
      return { level: 'projection', text: 'Future projection: depends on technology or prices that do not exist yet' };
    }
    var weak = st === 'Very weak' || st === 'None recorded';
    var tol = 1e-9 * Math.max(1, Math.abs(evidence.high));
    if (v >= evidence.low - tol && v <= evidence.high + tol) {
      return { level: weak ? 'weak' : 'ok', text: weak ? 'Inside the placeholder range, but the range itself is a guess' : 'Inside the range found in the evidence' };
    }
    var dir = favorsSpace(params, id), above = v > evidence.high;
    var helpsSpace = (above && dir > 0) || (!above && dir < 0);
    if (dir === 0) return { level: 'out', text: 'Outside the evidence range (no effect on the result here)' };
    return helpsSpace ? { level: 'optimistic', text: 'Outside the evidence range, in the direction that favors space (optimistic)' }
      : { level: 'conservative', text: 'Outside the evidence range, in the direction that favors Earth (conservative)' };
  }

  var api = {
    MODEL_VERSION: MODEL_VERSION, TIE_TOLERANCE: TIE_TOLERANCE, PARAM_DEFS: PARAM_DEFS, DEF_BY_ID: DEF_BY_ID,
    mulberry32: mulberry32, sampleDist: sampleDist, validate: validate, compute: compute,
    pointValues: pointValues, quantile: quantile, simulate: simulate, sensitivity: sensitivity,
    favorsSpace: favorsSpace, reasonableness: reasonableness, evidenceStrength: evidenceStrength
  };
  root.SCEngine = api;
})(typeof self !== 'undefined' ? self : this);

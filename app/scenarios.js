// Scenario presets. The "illustrative" preset reproduces PRD section 11 exactly (hypothetical numbers).
// The "literature-informed" preset is generated from the spreadsheet by build_spreadsheet.py (scenario_data.js).
(function (root) {
  'use strict';
  function fixed(v) { return { value: v, dist: { type: 'fixed' } }; }

  function illustrative() {
    var p = {};
    SCEngine.PARAM_DEFS.forEach(function (d) { p[d.id] = fixed(0); });
    p.U = fixed(1e6); p.N0 = fixed(1e6); p.g_E = fixed(20); p.r_g = fixed(2);
    p.q_E = fixed(0.5); p.r_q = fixed(1); p.s = fixed(0.8); p.p_E = fixed(1); p.p_S = fixed(1);
    p.E_other = fixed(100000);   // C_E_prod = $100,000 entered as a single line
    p.S_other = fixed(200000);   // C_S_prod = $200,000 entered as a single line
    p.L = fixed(2000); p.M = fixed(50); p.C_return = fixed(50000);
    return p;
  }

  function illustrativeTriangular() {
    var p = illustrative();
    p.r_g = { value: 2, dist: { type: 'triangular', min: 1, mode: 2, max: 6 } };
    return p;
  }

  function empty() {
    var p = {};
    SCEngine.PARAM_DEFS.forEach(function (d) { p[d.id] = fixed(0); });
    p.U = fixed(1e6); p.r_g = fixed(1); p.r_q = fixed(1); p.s = fixed(1); p.p_E = fixed(1); p.p_S = fixed(1);
    return p;
  }

  root.SCScenarios = {
    illustrative: { name: 'Illustrative example (HYPOTHETICAL, PRD section 11)', build: illustrative },
    illustrativeTri: { name: 'Illustrative + subjective triangular r_g(1, 2, 6) (HYPOTHETICAL)', build: illustrativeTriangular },
    empty: { name: 'Empty scenario', build: empty }
  };
})(typeof self !== 'undefined' ? self : this);

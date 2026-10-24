// Runs the Monte Carlo off the UI thread (PRD 10).
importScripts('engine.js');
self.onmessage = function (e) {
  var msg = e.data;
  try {
    var errors = SCEngine.validate(msg.params);
    if (errors.length) { self.postMessage({ type: 'error', errors: errors }); return; }
    var res = SCEngine.simulate(msg.params, msg.opts, function (f) { self.postMessage({ type: 'progress', fraction: f }); });
    self.postMessage({ type: 'done', result: res });
  } catch (err) {
    self.postMessage({ type: 'error', errors: [String(err && err.message || err)] });
  }
};

/**
 * gas-shim.js
 * Meniru google.script.run supaya Index.html lama jalan di Vercel TANPA ubah ratusan baris kode.
 * Di Apps Script asli (google.script sudah ada) file ini tidak melakukan apa-apa.
 */
(function () {
  if (window.google && window.google.script && window.google.script.run) return;

  function getToken() {
    try { return (JSON.parse(sessionStorage.getItem('belacom_user')) || {}).token || ''; }
    catch (e) { return ''; }
  }

  async function call(fn, args, ok, fail) {
    try {
      const res = await fetch('/api/rpc', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json', 'Authorization': 'Bearer ' + getToken() },
        body: JSON.stringify({ fn: fn, args: args })
      });
      if (res.status === 401 && fn !== 'checkUserLogin') {
        sessionStorage.removeItem('belacom_user');
        alert('Sesi berakhir, silakan login ulang.');
        location.reload();
        return;
      }
      const json = await res.json();
      if (json.ok) { if (ok) ok(json.data); }
      else if (fail) fail({ message: json.error || 'Terjadi kesalahan' });
      else console.error('[rpc ' + fn + ']', json.error);
    } catch (e) {
      if (fail) fail({ message: e.message }); else console.error('[rpc ' + fn + ']', e);
    }
  }

  function runner(ok, fail) {
    return new Proxy({}, {
      get: function (_, prop) {
        if (prop === 'withSuccessHandler') return function (f) { return runner(f, fail); };
        if (prop === 'withFailureHandler') return function (f) { return runner(ok, f); };
        return function () { call(String(prop), Array.prototype.slice.call(arguments), ok, fail); };
      }
    });
  }

  window.google = { script: { run: runner() } };
})();

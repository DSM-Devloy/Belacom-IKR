/**
 * /api/rpc  - gateway tunggal Vercel -> Apps Script (Fase 1).
 * Fungsi: (1) login + token bertanda tangan, (2) otorisasi per role, (3) teruskan ke Apps Script
 * dengan secret rahasia sehingga URL Apps Script tidak bisa dipanggil langsung oleh orang lain.
 *
 * ENV (Vercel > Settings > Environment Variables):
 *   GAS_URL     = URL Web App Apps Script (.../exec)
 *   GAS_SECRET  = string acak panjang, sama dengan Script Property API_SECRET
 *   JWT_SECRET  = string acak panjang lain
 */
const crypto = require('crypto');

const ADMIN_ONLY = new Set([
  'getUsersList', 'saveUser', 'deleteUser',
  'getTeamsManagementList', 'saveTeamData', 'deleteTeamRow',
  'getPriceSettings', 'savePriceSetting', 'deletePriceSetting', 'getPriceImpact',
  'getBillingDashboardData', 'updatePaymentStatus'
]);
const ADMIN_LEADER = new Set([
  'saveWorkOrder', 'deleteWorkOrder', 'exportWorkOrdersToExcel'
]);
// Fungsi lain: cukup login. 'getWorkOrders' tetap terbuka untuk semua role karena dipakai
// dashboard/list; penyaringan kolom harga per role bisa ditambahkan di Fase 2.

const b64u = (s) => Buffer.from(s).toString('base64url');

function sign(payload) {
  const h = b64u(JSON.stringify({ alg: 'HS256', typ: 'JWT' }));
  const p = b64u(JSON.stringify(payload));
  const s = crypto.createHmac('sha256', process.env.JWT_SECRET).update(h + '.' + p).digest('base64url');
  return h + '.' + p + '.' + s;
}

function verify(token) {
  const parts = String(token || '').split('.');
  if (parts.length !== 3) return null;
  const expected = crypto.createHmac('sha256', process.env.JWT_SECRET).update(parts[0] + '.' + parts[1]).digest();
  const given = Buffer.from(parts[2], 'base64url');
  if (given.length !== expected.length || !crypto.timingSafeEqual(given, expected)) return null;
  const payload = JSON.parse(Buffer.from(parts[1], 'base64url').toString());
  if (!payload.exp || payload.exp < Date.now() / 1000) return null;
  return payload;
}

async function callGas(fn, args) {
  const r = await fetch(process.env.GAS_URL, {
    method: 'POST',
    headers: { 'Content-Type': 'text/plain;charset=utf-8' },
    body: JSON.stringify({ secret: process.env.GAS_SECRET, fn: fn, args: args }),
    redirect: 'follow'
  });
  return r.json();
}

module.exports = async (req, res) => {
  if (req.method !== 'POST') return res.status(405).json({ ok: false, error: 'Method not allowed' });
  const { fn, args = [] } = req.body || {};
  if (!fn) return res.status(400).json({ ok: false, error: 'fn wajib diisi' });

  try {
    // Login: satu-satunya fungsi tanpa token
    if (fn === 'checkUserLogin') {
      const out = await callGas(fn, args);
      if (!out.ok) return res.status(502).json(out);
      const u = out.data;
      if (u && u.success) {
        u.token = sign({ u: u.username, role: u.role, exp: Math.floor(Date.now() / 1000) + 12 * 3600 });
      }
      return res.status(200).json({ ok: true, data: u });
    }

    const user = verify((req.headers.authorization || '').replace(/^Bearer /, ''));
    if (!user) return res.status(401).json({ ok: false, error: 'Unauthorized' });

    if (ADMIN_ONLY.has(fn) && user.role !== 'Admin')
      return res.status(403).json({ ok: false, error: 'Akses ditolak' });
    if (ADMIN_LEADER.has(fn) && !['Admin', 'Leader'].includes(user.role))
      return res.status(403).json({ ok: false, error: 'Akses ditolak' });

    // User hanya boleh ganti password miliknya sendiri
    if (fn === 'updatePassword' && String(args[0]).toLowerCase() !== String(user.u).toLowerCase())
      return res.status(403).json({ ok: false, error: 'Akses ditolak' });

    // Audit: nama pengubah harga diambil dari token, bukan dari browser
    if (fn === 'savePriceSetting' && args[0]) args[0].updatedBy = user.u;

    const out = await callGas(fn, args);
    return res.status(out.ok ? 200 : 400).json(out);
  } catch (e) {
    return res.status(500).json({ ok: false, error: 'Gateway error: ' + e.message });
  }
};

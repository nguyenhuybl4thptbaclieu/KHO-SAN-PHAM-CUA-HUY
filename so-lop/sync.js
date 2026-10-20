// Đồng bộ với Google Drive thông qua ứng dụng web Google Apps Script của chính bạn.
const CFG = 'so-lop-sync';

export function createSync({ getData, applyRemote, onConflict, onStatus }) {
  let cfg = {};
  try { cfg = JSON.parse(localStorage.getItem(CFG) || '{}') || {}; } catch { cfg = {}; }
  let timer = null, busy = false, again = false, lastPull = 0;
  let status = { state: cfg.url ? 'idle' : 'off', at: cfg.lastAt || '', message: '' };

  const store = () => { try { localStorage.setItem(CFG, JSON.stringify(cfg)); } catch { /* bỏ qua */ } };
  const set = (state, message = '') => { status = { state, at: cfg.lastAt || '', message }; onStatus?.(status); };
  const configured = () => !!(cfg.url && cfg.key);

  async function call(payload) {
    const ctl = new AbortController(), t = setTimeout(() => ctl.abort(), 25000);
    try {
      const res = await fetch(cfg.url, { method: 'POST', body: JSON.stringify({ ...payload, key: cfg.key }), headers: { 'Content-Type': 'text/plain;charset=utf-8' }, signal: ctl.signal, redirect: 'follow' });
      const text = await res.text();
      let j; try { j = JSON.parse(text); } catch { throw Error('Địa chỉ đồng bộ không đúng hoặc chưa cho phép “Bất kỳ ai”.'); }
      if (!j.ok && !j.conflict) throw Error(j.error || 'Không đồng bộ được.');
      return j;
    } finally { clearTimeout(t); }
  }
  function fail(e) {
    if (!navigator.onLine || e.name === 'AbortError' || e instanceof TypeError) set('offline', 'Chưa có mạng · sẽ tự lưu lên Drive khi có mạng');
    else set('error', e.message);
  }
  function done(updatedAt, sentAt) {
    cfg.base = updatedAt;
    if (getData().updatedAt === sentAt) cfg.pending = false;
    cfg.lastAt = new Date().toISOString(); store(); set('ok');
  }

  async function push(force = false) {
    if (!configured()) return;
    if (busy) { again = true; return; }
    busy = true; set('busy');
    const data = getData();
    try {
      const j = await call({ action: 'save', data, base: cfg.base || '', force });
      if (j.conflict) { set('conflict'); onConflict(j.data); }
      else done(j.updatedAt, data.updatedAt);
    } catch (e) { fail(e); }
    finally { busy = false; if (again) { again = false; push(); } }
  }

  async function pull() {
    if (!configured() || busy) return;
    lastPull = Date.now(); busy = true; set('busy');
    let after = null;
    try {
      const j = await call({ action: 'load' });
      const remote = j.data, local = getData();
      if (!remote) { if (local.groups.length || cfg.pending) after = () => push(); else set('ok'); }
      else if ((j.updatedAt || '') === (cfg.base || '')) { if (cfg.pending) after = () => push(); else { cfg.lastAt = new Date().toISOString(); store(); set('ok'); } }
      else if (!cfg.pending || !local.groups.length) { applyRemote(remote); cfg.base = j.updatedAt; cfg.pending = false; cfg.lastAt = new Date().toISOString(); store(); set('ok'); }
      else { set('conflict'); onConflict(remote); }
    } catch (e) { fail(e); }
    finally { busy = false; after?.(); }
  }

  return {
    get status() { return status; },
    get configured() { return configured(); },
    get url() { return cfg.url || ''; },
    markDirty() {
      cfg.pending = true; store();
      if (!configured()) return;
      clearTimeout(timer); timer = setTimeout(() => push(), 1200);
    },
    pull,
    pullIfStale() { if (Date.now() - lastPull > 20000) pull(); },
    flush() { if (cfg.pending && configured()) push(); },
    async connect(url, key) {
      url = String(url).trim(); key = String(key).trim().toUpperCase();
      if (!/^https:\/\/script\.google\.com\/macros\/s\/[\w-]+\/exec$/.test(url)) throw Error('Địa chỉ phải có dạng https://script.google.com/macros/s/…/exec');
      if (!/^[A-Z0-9]{12,40}$/.test(key)) throw Error('Mã kết nối không đúng.');
      const old = cfg; cfg = { url, key, base: '', pending: getData().groups.length > 0 };
      try { await call({ action: 'ping' }); } catch (e) { cfg = old; throw e; }
      store(); await pull();
    },
    disconnect() { cfg = { pending: cfg.pending }; store(); set('off'); },
    useRemote(remote) { applyRemote(remote); cfg.base = remote.updatedAt || ''; cfg.pending = false; cfg.lastAt = new Date().toISOString(); store(); set('ok'); },
    useLocal() { return push(true); },
  };
}

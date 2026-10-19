/*
 * cloud.js — 登入與共用設定（Supabase）。
 *
 * 部門的人共用同一份「記住的設定」：誰最後存，大家打開看到的就是誰的。
 *   資料庫  Supabase 專案（js/config.js 的 SUPABASE_URL / SUPABASE_KEY）
 *   位置    schema ie_manpower 的 settings 表，一筆 = (app, key) → value(jsonb)
 *           app 是這個 PWA 的專案代號（PROJECT_ID），key 是設定名稱（例如 step1、report）
 *   登入    Supabase Auth 的 email + 密碼（帳號由管理者在 Supabase 後台建立，不開放註冊）
 *
 * 各步驟用 IEApp.settings.read(name) / write(name, obj) 讀寫設定：
 *   讀    一律讀這台電腦的本機備份（localStorage），所以是同步的、離線也能用
 *   寫    先存本機，再在背景上傳 Supabase（連續打字時等停下來 1 秒才傳）
 *   開啟  登入後先把 Supabase 上最新的設定抓下來蓋掉本機備份，才載入各步驟
 *   離線  照常運作，改過的設定記為「還沒上傳」，下次連上網路時再傳
 *
 * 不用 supabase-js 套件，直接用 fetch 呼叫 Supabase 的 REST API（少一個要從網路載入的檔案）。
 */
(function (global) {
  'use strict';

  var IEApp = global.IEApp;
  var h = IEApp.h;
  var C = global.APP_CONFIG || {};
  var URL_BASE = String(C.SUPABASE_URL || '').replace(/\/+$/, '');
  var KEY = C.SUPABASE_KEY || '';
  var SCHEMA = C.SUPABASE_SCHEMA || 'ie_manpower';
  var APP = IEApp.projectId;
  var ENABLED = !!(URL_BASE && KEY);

  var SESSION_KEY = IEApp.storeKey('_session');   // 登入資訊（只存這台電腦）
  var SYNC_KEY = IEApp.storeKey('_sync');         // { dirty: { 名稱: true }, meta: { 名稱: { by, at } } }
  var PUSH_DELAY = 1000;
  var TIMEOUT = 10000;

  var session = null;          // { access_token, refresh_token, expires_at(ms), email }
  var refreshing = null;
  var pushTimer = null;
  var pushing = null;
  var status = { state: 'idle', text: '' };
  var els = {};

  /* ---------- 本機儲存 ---------- */
  function readLocal(k) {
    try { return localStorage.getItem(k); } catch (e) { return null; }
  }
  function writeLocal(k, v) {
    try { if (v === null) localStorage.removeItem(k); else localStorage.setItem(k, v); } catch (e) { /* 存不了就算了 */ }
  }
  function readJson(k, fallback) {
    try { var v = JSON.parse(readLocal(k) || 'null'); return v === null ? fallback : v; } catch (e) { return fallback; }
  }
  function syncState() {
    var s = readJson(SYNC_KEY, {}) || {};
    s.dirty = s.dirty || {};
    s.meta = s.meta || {};
    return s;
  }
  function saveSyncState(s) { writeLocal(SYNC_KEY, JSON.stringify(s)); }

  // 這台電腦上這個專案的設定名稱（不含 _ 開頭的內部資料）
  function localNames() {
    var prefix = IEApp.storeKey('');
    var out = [];
    try {
      for (var i = 0; i < localStorage.length; i++) {
        var k = localStorage.key(i);
        if (k && k.indexOf(prefix) === 0 && k.charAt(prefix.length) !== '_') out.push(k.slice(prefix.length));
      }
    } catch (e) { /* 不能用 localStorage */ }
    return out;
  }

  /* ---------- 各步驟用的設定 API ---------- */
  var settings = {
    // 回傳之前存的物件；沒有就回傳 null
    read: function (name) { return readJson(IEApp.storeKey(name), null); },
    // 存下設定：本機馬上存好，Supabase 稍後在背景上傳
    write: function (name, value) {
      var text = JSON.stringify(value);
      if (readLocal(IEApp.storeKey(name)) === text) return;   // 內容沒變（例如打開畫面時重存一次）就不算修改
      writeLocal(IEApp.storeKey(name), text);
      if (!ENABLED) return;
      var s = syncState();
      s.dirty[name] = true;
      saveSyncState(s);
      schedulePush();
    },
    // 這幾個設定最後是誰、什麼時候改的（取最新的一筆）；沒有回傳 null
    lastChange: function (names) {
      var meta = syncState().meta;
      var best = null;
      (names || []).forEach(function (n) {
        var m = meta[n];
        if (m && m.at && (!best || m.at > best.at)) best = m;
      });
      return best;
    }
  };

  /* ---------- 連線 ---------- */
  function request(path, opts) {
    opts = opts || {};
    var ctrl = typeof AbortController !== 'undefined' ? new AbortController() : null;
    var timer = ctrl ? setTimeout(function () { ctrl.abort(); }, TIMEOUT) : null;
    var headers = Object.assign({ apikey: KEY }, opts.headers || {});
    if (opts.json !== undefined) headers['Content-Type'] = 'application/json';
    return fetch(URL_BASE + path, {
      method: opts.method || 'GET',
      headers: headers,
      body: opts.json !== undefined ? JSON.stringify(opts.json) : undefined,
      signal: ctrl ? ctrl.signal : undefined
    }).then(function (res) {
      if (timer) clearTimeout(timer);
      return res.text().then(function (text) {
        var data = null;
        try { data = text ? JSON.parse(text) : null; } catch (e) { data = text; }
        if (!res.ok) {
          var err = new Error((data && (data.msg || data.message || data.error_description || data.error)) || ('HTTP ' + res.status));
          err.status = res.status;
          err.code = data && (data.error_code || data.code || data.error);
          throw err;
        }
        return data;
      });
    }, function (e) {
      if (timer) clearTimeout(timer);
      var err = new Error('連不上 Supabase（網路中斷或被擋）');
      err.network = true;
      err.cause = e;
      throw err;
    });
  }

  /* ---------- 登入 ---------- */
  function saveSession(data, email) {
    session = {
      access_token: data.access_token,
      refresh_token: data.refresh_token,
      expires_at: Date.now() + (data.expires_in || 3600) * 1000,
      email: (data.user && data.user.email) || email || (session && session.email) || ''
    };
    writeLocal(SESSION_KEY, JSON.stringify(session));
    renderUser();
  }
  function clearSession() {
    session = null;
    writeLocal(SESSION_KEY, null);
  }

  function signIn(email, password) {
    return request('/auth/v1/token?grant_type=password', { method: 'POST', json: { email: email, password: password } })
      .then(function (data) { saveSession(data, email); });
  }

  function refresh() {
    if (!session || !session.refresh_token) return Promise.reject(Object.assign(new Error('沒有登入'), { auth: true }));
    if (!refreshing) {
      refreshing = request('/auth/v1/token?grant_type=refresh_token', { method: 'POST', json: { refresh_token: session.refresh_token } })
        .then(function (data) { saveSession(data); })
        .catch(function (e) {
          if (!e.network) e.auth = true;      // 被 Supabase 拒絕（例如帳號被停用）＝要重新登入
          throw e;
        })
        .then(function (v) { refreshing = null; return v; }, function (e) { refreshing = null; throw e; });
    }
    return refreshing;
  }

  // 拿到有效的登入憑證（快過期就先換新的）
  function token() {
    if (!session) return Promise.reject(Object.assign(new Error('沒有登入'), { auth: true }));
    if (session.expires_at - 60000 > Date.now()) return Promise.resolve(session.access_token);
    return refresh().then(function () { return session.access_token; });
  }

  // 帶登入憑證呼叫資料表；憑證失效時換一次再試
  function api(path, opts, retried) {
    return token().then(function (t) {
      var o = Object.assign({}, opts);
      o.headers = Object.assign({ Authorization: 'Bearer ' + t }, opts.headers || {});
      return request(path, o);
    }).catch(function (e) {
      if (e.status === 401 && !retried) {
        session.expires_at = 0;
        return api(path, opts, true);
      }
      if (e.status === 401) e.auth = true;
      throw e;
    });
  }

  /* ---------- 同步 ---------- */
  // 把 Supabase 上這個 PWA 的設定抓下來：沒有「還沒上傳」的就用雲端的蓋掉本機；
  // 本機有、雲端沒有的設定，標成要上傳（第一次使用時把這台電腦原本的設定傳上去）。
  function pull() {
    setStatus('busy', '同步中…');
    return api('/rest/v1/settings?app=eq.' + encodeURIComponent(APP) + '&select=key,value,updated_at,updated_by',
      { headers: { 'Accept-Profile': SCHEMA } })
      .then(function (rows) {
        var s = syncState();
        var remote = {};
        (rows || []).forEach(function (row) {
          remote[row.key] = true;
          if (!s.dirty[row.key]) writeLocal(IEApp.storeKey(row.key), JSON.stringify(row.value));
          s.meta[row.key] = { by: row.updated_by || '', at: row.updated_at || '' };
        });
        localNames().forEach(function (name) { if (!remote[name]) s.dirty[name] = true; });
        saveSyncState(s);
        updateNotes();
        return push();
      })
      .catch(handleError);
  }

  function schedulePush() {
    clearTimeout(pushTimer);
    pushTimer = setTimeout(function () { push().catch(function () {}); }, PUSH_DELAY);
  }

  // 把「還沒上傳」的設定傳到 Supabase
  function push() {
    if (pushing) return pushing.then(push);
    var s = syncState();
    var names = Object.keys(s.dirty);
    if (!names.length) { setStatus('ok', '已同步'); return Promise.resolve(); }
    if (!session) return Promise.resolve();
    var rows = [];
    names.forEach(function (name) {
      var raw = readLocal(IEApp.storeKey(name));
      var value;
      try { value = JSON.parse(raw); } catch (e) { value = null; }
      if (value !== null && value !== undefined) rows.push({ app: APP, key: name, value: value });
    });
    setStatus('busy', '同步中…');
    pushing = api('/rest/v1/settings?on_conflict=app,key', {
      method: 'POST',
      json: rows,
      headers: { 'Content-Profile': SCHEMA, Prefer: 'resolution=merge-duplicates,return=representation' }
    }).then(function (saved) {
      var st = syncState();
      names.forEach(function (name) {
        // 上傳的這段時間又改了的，留著下一輪再傳
        var now = readLocal(IEApp.storeKey(name));
        var sent = rows.filter(function (r) { return r.key === name; })[0];
        if (!sent || JSON.stringify(sent.value) === now) delete st.dirty[name];
      });
      (saved || []).forEach(function (row) { st.meta[row.key] = { by: row.updated_by || '', at: row.updated_at || '' }; });
      saveSyncState(st);
      updateNotes();
      pushing = null;
      if (Object.keys(st.dirty).length) return push();
      setStatus('ok', '已同步');
    }, function (e) {
      pushing = null;
      return handleError(e);
    });
    return pushing;
  }

  function handleError(e) {
    if (e && e.network) setStatus('offline', '離線：設定先存在這台電腦，連上網路後會自動上傳');
    else if (e && e.auth) { setStatus('error', '登入已過期，請重新登入'); showLogin(true); }
    else setStatus('error', '同步失敗：' + (e && e.message ? e.message : e));
  }

  /* ---------- 畫面：上方的登入者與同步狀態 ---------- */
  function setStatus(state, text) {
    status = { state: state, text: text };
    if (els.status) {
      els.status.textContent = text;
      els.status.className = 'cloud-status is-' + state;
    }
  }

  function renderUser() {
    if (!els.user) return;
    els.user.textContent = session && session.email ? shortName(session.email) : '';
    els.user.title = session && session.email ? session.email : '';
    els.logout.hidden = !session;
  }

  function buildTopbar() {
    var bar = document.querySelector('.topbar');
    if (!bar || els.box) return;
    els.status = h('span', { class: 'cloud-status', 'aria-live': 'polite' });
    els.user = h('span', { class: 'cloud-user' });
    els.logout = h('button', { type: 'button', class: 'cloud-logout', text: '登出', hidden: true, onclick: logout });
    els.box = h('span', { class: 'cloud-box' }, els.status, els.user, els.logout);
    var ver = document.getElementById('app-version');
    bar.insertBefore(els.box, ver);
    setStatus(status.state, status.text);
    renderUser();
  }

  function shortName(email) { return String(email || '').split('@')[0]; }

  function formatTime(iso) {
    var d = new Date(iso);
    if (isNaN(d.getTime())) return '';
    function p(n) { return (n < 10 ? '0' : '') + n; }
    return (d.getMonth() + 1) + '/' + d.getDate() + ' ' + p(d.getHours()) + ':' + p(d.getMinutes());
  }

  // 各步驟上方「上次由誰修改」那一行
  function noteFor(names) {
    var el = h('p', { class: 'sync-note' });
    el.setAttribute('data-keys', names.join(','));
    fillNote(el);
    return el;
  }
  function fillNote(el) {
    var names = (el.getAttribute('data-keys') || '').split(',').filter(Boolean);
    var m = settings.lastChange(names);
    var dirty = syncState().dirty;
    var waiting = names.some(function (n) { return dirty[n]; });
    var text = m ? '共用設定：上次由 ' + shortName(m.by) + ' 於 ' + formatTime(m.at) + ' 修改'
      : '共用設定：還沒有人存過';
    if (waiting) text += '（這台電腦有修改還沒上傳）';
    el.textContent = text;
  }
  function updateNotes() {
    Array.prototype.forEach.call(document.querySelectorAll('.sync-note'), fillNote);
  }

  /* ---------- 畫面：登入 ---------- */
  var loginResolve = null;

  function showLogin(expired) {
    setStatus('busy', '請登入');
    if (els.login) return new Promise(function (res) { loginResolve = res; });
    var email = h('input', { id: 'login-email', type: 'email', class: 'in', autocomplete: 'username', required: true,
      value: (session && session.email) || '' });
    var pass = h('input', { id: 'login-pass', type: 'password', class: 'in', autocomplete: 'current-password', required: true });
    var msg = h('p', { class: 'msg is-error', role: 'alert', hidden: true });
    var btn = h('button', { type: 'submit', class: 'btn-primary', text: '登入' });
    var form = h('form', { class: 'login-card', 'aria-labelledby': 'login-title' },
      h('h1', { id: 'login-title', text: (document.querySelector('.app-name') || {}).textContent || '登入' }),
      h('p', { class: 'login-sub', text: expired ? '登入已過期，請重新登入。輸入的設定都還在這台電腦。' : '請登入。部門共用同一份設定，誰最後存，大家看到的就是誰的。' }),
      h('label', { for: 'login-email', text: '帳號（email）' }), email,
      h('label', { for: 'login-pass', text: '密碼' }), pass,
      msg, btn);
    form.addEventListener('submit', function (e) {
      e.preventDefault();
      msg.hidden = true;
      btn.disabled = true;
      btn.textContent = '登入中…';
      signIn(email.value.trim(), pass.value).then(function () {
        hideLogin();
        var r = loginResolve; loginResolve = null;
        if (r) r();
        else pull();                        // 用到一半過期、重新登入後補上傳
      }).catch(function (err) {
        msg.textContent = err.network ? '連不上 Supabase，請確認網路後再試。'
          : (err.code === 'invalid_credentials' || err.status === 400) ? '帳號或密碼不對。'
          : '登入失敗：' + err.message;
        msg.hidden = false;
        btn.disabled = false;
        btn.textContent = '登入';
      });
    });
    els.login = h('div', { class: 'login-overlay', role: 'dialog', 'aria-modal': 'true' }, form);
    document.body.appendChild(els.login);
    setTimeout(function () { (email.value ? pass : email).focus(); }, 0);
    return new Promise(function (res) { loginResolve = res; });
  }

  function hideLogin() {
    if (els.login) { els.login.remove(); els.login = null; }
  }

  function logout() {
    var pending = Object.keys(syncState().dirty).length;
    if (pending && !global.confirm('這台電腦還有 ' + pending + ' 組設定沒上傳，登出後要等下次登入才會上傳。確定要登出嗎？')) return;
    var t = session && session.access_token;
    clearSession();
    var done = function () { global.location.reload(); };
    if (t) request('/auth/v1/logout', { method: 'POST', headers: { Authorization: 'Bearer ' + t } }).then(done, done);
    else done();
  }

  /* ---------- 開啟時 ---------- */
  // 回傳 Promise：登入好、抓完雲端設定（或確定離線）之後才 resolve，接著才載入各步驟
  function start() {
    buildTopbar();
    if (!ENABLED) { setStatus('off', ''); return Promise.resolve(); }
    session = readJson(SESSION_KEY, null);
    renderUser();
    var ready = session ? token().catch(function (e) {
      if (e.network) return 'offline';               // 沒網路：先用這台電腦的設定
      clearSession();
      return showLogin(true);
    }) : showLogin(false);
    return ready.then(function (r) {
      if (r === 'offline') { setStatus('offline', '離線：設定先存在這台電腦，連上網路後會自動上傳'); return; }
      return pull();
    }).then(function () {
      global.addEventListener('online', function () { if (session) pull(); });
    });
  }

  IEApp.settings = settings;
  IEApp.cloud = {
    enabled: ENABLED,
    start: start,
    noteFor: noteFor,
    pull: pull,
    push: push,
    status: function () { return status; }
  };
})(window);

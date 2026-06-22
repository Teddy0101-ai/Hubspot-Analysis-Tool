/* HubSpot Bulk Recipient Export - content script
 * Runs on app.hubspot.com/email/<portal>/manage/* (the marketing email list).
 * Builds the panel, resolves ids, fires the exports, polls for the finished files'
 * download links, then hands those links to the background worker which fetches +
 * zips + downloads (the worker, unlike this content script, can read the file's
 * cross-origin CDN response).
 */
(function () {
  'use strict';
  if (window.__hsBulkExportLoaded) return;
  window.__hsBulkExportLoaded = true;

  var PORTAL_ID = (location.pathname.match(/\/email\/(\d+)\//) || [])[1];
  if (!PORTAL_ID) return;
  var STATIC_APP = 'ContentEmailUI';
  var STATIC_VER = '2.74364';
  var NOTIF_VER = '1.14628';
  var HARVEST_TIMEOUT_MS = 5 * 60 * 1000;

  // ---------------- helpers ----------------

  function csrfToken() {
    var parts = document.cookie.split('; ');
    var exact = parts.find(function (c) { return c.indexOf('csrf.app=') === 0; });
    if (exact) return exact.split('=').slice(1).join('=');
    var any = parts.find(function (c) { return /^csrf/i.test(c); });
    return any ? any.split('=').slice(1).join('=') : '';
  }
  function apiHeaders() {
    return { 'content-type': 'application/json', 'X-HubSpot-CSRF-hubspotapi': csrfToken() };
  }
  function sleep(ms) { return new Promise(function (r) { setTimeout(r, ms); }); }
  function pad2(n) { return n < 10 ? '0' + n : '' + n; }
  function plural(n) { return n === 1 ? '' : 's'; }

  function defaultZipName() {
    var d = new Date();
    return 'hubspot-recipient-lists-' + d.getFullYear() + pad2(d.getMonth() + 1) + pad2(d.getDate()) +
      '-' + pad2(d.getHours()) + pad2(d.getMinutes());
  }
  function makeRunTag() {
    return (Date.now().toString(36) + Math.random().toString(36).slice(2))
      .toUpperCase().replace(/[^A-Z0-9]/g, '').slice(0, 5);
  }

  // ---------------- selection (native row checkboxes) ----------------

  function rowLinkFor(checkbox) {
    var el = checkbox;
    for (var i = 0; i < 15 && el; i++) {
      el = el.parentElement;
      if (!el) break;
      var links = el.querySelectorAll('a[href*="/details/"]');
      if (links.length === 1) return links[0];
      if (links.length > 1) return null;
    }
    return null;
  }
  function getSelectedEmails() {
    var checks = Array.prototype.filter.call(
      document.querySelectorAll('input[type=checkbox]'),
      function (c) { return c.checked; }
    );
    var seen = new Map();
    for (var i = 0; i < checks.length; i++) {
      var link = rowLinkFor(checks[i]);
      if (!link) continue;
      var m = (link.getAttribute('href') || '').match(/details\/(\d+)/);
      if (!m) continue;
      if (!seen.has(m[1])) seen.set(m[1], (link.textContent || '').trim());
    }
    var out = [];
    seen.forEach(function (name, detailId) { out.push({ detailId: detailId, name: name }); });
    return out;
  }

  // ---------------- HubSpot API calls ----------------

  function resolveMicIds(detailIds) {
    var url = '/api/crm-search/search?portalId=' + PORTAL_ID +
      '&clienttimeout=14000&hs_static_app=' + STATIC_APP + '&hs_static_app_version=' + STATIC_VER;
    var body = {
      count: Math.min(detailIds.length, 200) + 5,
      objectTypeId: '0-29', fetchSource: false, fetchVersion: false,
      requestOptions: { includeObjectVersion: false, properties: ['hs_root_mic_id', 'hs_origin_asset_id'] },
      filterGroups: [{ filters: [{ operator: 'IN', property: 'hs_origin_asset_id', values: detailIds.map(Number) }] }]
    };
    return fetch(url, { method: 'POST', credentials: 'include', headers: apiHeaders(), body: JSON.stringify(body) })
      .then(function (r) { if (!r.ok) throw new Error('ID lookup failed (HTTP ' + r.status + ')'); return r.json(); })
      .then(function (j) {
        var map = {};
        (j.results || []).forEach(function (o) {
          var p = o.properties || {};
          function g(k) { var v = p[k]; return (v && typeof v === 'object') ? v.value : v; }
          var did = String(g('hs_origin_asset_id')), mic = g('hs_root_mic_id');
          if (did && mic) map[did] = String(mic);
        });
        return map;
      });
  }

  function fireExport(micId, emailName) {
    var url = '/api/chirp-frontend-app/v1/gateway/com.hubspot.email.stats.rpc.export.EmailExports/exportEventDigests' +
      '?portalId=' + PORTAL_ID + '&clienttimeout=5000&hs_static_app=' + STATIC_APP + '&hs_static_app_version=' + STATIC_VER;
    var body = {
      startTimestamp: 1262304000000, endTimestamp: Date.now() + 86400000,
      micIds: [Number(micId)], includedBatchListIds: [], eventDigestType: 'ALL',
      emailName: emailName, format: 'XLSX', campaignIds: [], campaignGroupIds: [], campaignSendIds: []
    };
    return fetch(url, { method: 'POST', credentials: 'include', headers: apiHeaders(), body: JSON.stringify(body) })
      .then(function (r) { if (!r.ok) throw new Error('Export request failed (HTTP ' + r.status + ')'); return true; });
  }

  // ---------------- harvest finished files (poll v3 notifications) ----------------

  function listUnreadNotifications() {
    var url = '/api/notification-station/general/v3/notifications/list?clienttimeout=14000' +
      '&hs_static_app=notifications&hs_static_app_version=' + NOTIF_VER +
      '&portalId=' + PORTAL_ID + '&locale=en&showUnreadOnly=true';
    return fetch(url, { credentials: 'include', headers: { accept: 'application/json', 'X-HubSpot-CSRF-hubspotapi': csrfToken() } })
      .then(function (r) { if (!r.ok) throw new Error('notification list HTTP ' + r.status); return r.json(); })
      .then(function (j) { return j.notifications || []; });
  }

  function harvestDownloads(pending, onProgress) {
    var results = {};
    var remaining = new Map();
    pending.forEach(function (p) { remaining.set(p.detailId, p); });
    var deadline = Date.now() + HARVEST_TIMEOUT_MS;
    function finish() { var missing = []; remaining.forEach(function (p) { missing.push(p); }); return { results: results, missing: missing }; }
    function loop() {
      return listUnreadNotifications().then(function (notifs) {
        notifs.forEach(function (n) {
          if (!n.ctaProxyUrl) return;
          var blob = JSON.stringify(n);
          var matchKey = null;
          remaining.forEach(function (p, did) { if (!matchKey && blob.indexOf(p.token) !== -1) matchKey = did; });
          if (matchKey) {
            results[matchKey] = n.ctaProxyUrl;
            remaining.delete(matchKey);
            if (onProgress) onProgress(pending.length - remaining.size, pending.length);
          }
        });
        if (!remaining.size || Date.now() > deadline) return finish();
        return sleep(3500).then(loop);
      }).catch(function () {
        if (Date.now() > deadline) return finish();
        return sleep(3500).then(loop);
      });
    }
    return loop();
  }

  // ---------------- hand off to background for fetch + zip + download ----------------

  function zipViaBackground(zipName, items, onMsg) {
    return new Promise(function (resolve, reject) {
      var port;
      try { port = chrome.runtime.connect({ name: 'bulkExport' }); }
      catch (e) { reject(new Error('Extension background not reachable - reload the extension.')); return; }
      var settled = false;
      port.onMessage.addListener(function (m) {
        if (m.type === 'progress') {
          if (m.stage === 'downloading') onMsg('Downloading file ' + (m.index + 1) + '/' + m.total + ': ' + m.name);
          else if (m.stage === 'zipping') onMsg('Zipping ' + m.count + ' file' + plural(m.count) + '...');
        } else if (m.type === 'fileError') {
          onMsg('Warning: "' + m.name + '" failed (' + m.error + ')');
        } else if (m.type === 'done') {
          settled = true; port.disconnect(); resolve(m);
        } else if (m.type === 'error') {
          settled = true; port.disconnect(); reject(new Error(m.error));
        }
      });
      port.onDisconnect.addListener(function () {
        if (!settled) reject(new Error((chrome.runtime.lastError && chrome.runtime.lastError.message) || 'Background disconnected.'));
      });
      port.postMessage({ type: 'start', zipName: zipName, items: items });
    });
  }

  // ---------------- orchestration ----------------

  var busy = false;

  function runExport() {
    if (busy) return;
    var zipName = (els.name.value || '').trim() || defaultZipName();
    var selected = getSelectedEmails();
    if (!selected.length) { setStatus('No rows selected. Tick the checkbox on each email row first.', 'err'); return; }

    setBusy(true);
    var runTag = makeRunTag();
    var unresolved = [];
    setStatus('Resolving ' + selected.length + ' email' + plural(selected.length) + '...');

    resolveMicIds(selected.map(function (s) { return s.detailId; }))
      .then(function (micMap) {
        var pending = [];
        selected.forEach(function (s, i) {
          if (micMap[s.detailId]) pending.push({ detailId: s.detailId, name: s.name, micId: micMap[s.detailId], token: runTag + pad2(i) });
          else unresolved.push(s.name);
        });
        if (!pending.length) throw new Error('Could not resolve email IDs (none matched).');
        setStatus('Requesting ' + pending.length + ' export' + plural(pending.length) + ' from HubSpot...');
        return pending.reduce(function (chain, p) {
          return chain.then(function () { return fireExport(p.micId, p.token + ' ' + p.name); });
        }, Promise.resolve()).then(function () { return pending; });
      })
      .then(function (pending) {
        setStatus('Waiting for HubSpot to generate files... (0/' + pending.length + ')');
        return harvestDownloads(pending, function (done, total) {
          setStatus('Waiting for HubSpot to generate files... (' + done + '/' + total + ')');
        }).then(function (h) { return { pending: pending, h: h }; });
      })
      .then(function (ctx) {
        var items = ctx.pending
          .filter(function (p) { return ctx.h.results[p.detailId]; })
          .map(function (p) { return { name: p.name, ctaUrl: ctx.h.results[p.detailId] }; });
        if (!items.length) throw new Error('No download links appeared in time. HubSpot may still be generating them - try again.');
        setStatus('Downloading & zipping ' + items.length + ' file' + plural(items.length) + '...');
        return zipViaBackground(zipName, items, function (msg) { setStatus(msg); })
          .then(function (res) { return { res: res, missing: ctx.h.missing }; });
      })
      .then(function (final) {
        var summary = 'Done. ' + final.res.downloaded + ' file' + plural(final.res.downloaded) + ' saved as ' + final.res.filename + '.';
        var extras = [];
        if (final.missing && final.missing.length) extras.push(final.missing.length + ' not ready');
        if (unresolved.length) extras.push(unresolved.length + ' unresolved');
        if (extras.length) summary += ' (' + extras.join(', ') + ')';
        setStatus(summary, 'ok');
      })
      .catch(function (e) { setStatus('Error: ' + (e.message || e), 'err'); })
      .then(function () { setBusy(false); });
  }

  // ---------------- UI ----------------

  var els = {};

  function buildUI() {
    var panel = document.createElement('div');
    panel.className = 'hsbe-panel';
    panel.innerHTML =
      '<div class="hsbe-head">' +
        '<span class="hsbe-title">Bulk recipient ZIP export</span>' +
        '<button class="hsbe-min" title="Minimize" type="button">_</button>' +
      '</div>' +
      '<div class="hsbe-body">' +
        '<label class="hsbe-label">ZIP file name</label>' +
        '<div class="hsbe-namerow">' +
          '<input class="hsbe-name" type="text" spellcheck="false" />' +
          '<span class="hsbe-ext">.zip</span>' +
        '</div>' +
        '<div class="hsbe-count"><b>0</b> rows selected</div>' +
        '<button class="hsbe-go" type="button">Download selected as ZIP</button>' +
        '<div class="hsbe-status">Tick the checkboxes on the email rows, name the ZIP, then export.</div>' +
      '</div>';
    document.body.appendChild(panel);

    els.panel = panel;
    els.name = panel.querySelector('.hsbe-name');
    els.count = panel.querySelector('.hsbe-count');
    els.go = panel.querySelector('.hsbe-go');
    els.status = panel.querySelector('.hsbe-status');

    els.name.value = defaultZipName();
    els.go.addEventListener('click', runExport);
    panel.querySelector('.hsbe-min').addEventListener('click', function () { panel.classList.toggle('hsbe-collapsed'); });

    setInterval(function () {
      if (panel.classList.contains('hsbe-collapsed')) return;
      var n = getSelectedEmails().length;
      els.count.innerHTML = '<b>' + n + '</b> row' + plural(n) + ' selected';
    }, 900);
  }

  function setStatus(msg, kind) {
    if (!els.status) return;
    els.status.textContent = msg;
    els.status.className = 'hsbe-status' + (kind ? ' hsbe-' + kind : '');
  }
  function setBusy(b) {
    busy = b;
    if (els.go) { els.go.disabled = b; els.go.textContent = b ? 'Working...' : 'Download selected as ZIP'; }
  }

  if (document.body) buildUI();
  else document.addEventListener('DOMContentLoaded', buildUI);
})();

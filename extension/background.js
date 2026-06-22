/* HubSpot Bulk Recipient Export - background service worker
 *
 * A service worker (unlike a content script) can read cross-origin responses -
 * including redirected ones - from hosts in host_permissions, with no CORS. So it
 * fetches each notification CTA (app.hubspot.com, session cookie sent via
 * credentials:'include' + host access), follows the 302 to the signed CDN
 * (*.hubspotusercontent-na1.net), reads the bytes, zips them, and downloads.
 */

// ---------- tiny ZIP writer (store / no compression) ----------

const CRC_TABLE = (() => {
  const t = new Uint32Array(256);
  for (let n = 0; n < 256; n++) {
    let c = n;
    for (let k = 0; k < 8; k++) c = c & 1 ? 0xedb88320 ^ (c >>> 1) : c >>> 1;
    t[n] = c >>> 0;
  }
  return t;
})();
function crc32(bytes) {
  let c = 0xffffffff;
  for (let i = 0; i < bytes.length; i++) c = CRC_TABLE[(c ^ bytes[i]) & 0xff] ^ (c >>> 8);
  return (c ^ 0xffffffff) >>> 0;
}
function u16(n) { const b = new Uint8Array(2); new DataView(b.buffer).setUint16(0, n & 0xffff, true); return b; }
function u32(n) { const b = new Uint8Array(4); new DataView(b.buffer).setUint32(0, n >>> 0, true); return b; }
function concat(arrays) {
  let len = 0; for (const a of arrays) len += a.length;
  const out = new Uint8Array(len); let p = 0;
  for (const a of arrays) { out.set(a, p); p += a.length; }
  return out;
}
function buildZip(files) {
  const enc = new TextEncoder();
  const chunks = [], central = [];
  let offset = 0;
  for (const f of files) {
    const nameBytes = enc.encode(f.name);
    const data = f.bytes;
    const crc = crc32(data);
    const flags = 0x0800;
    const local = concat([
      u32(0x04034b50), u16(20), u16(flags), u16(0), u16(0), u16(0),
      u32(crc), u32(data.length), u32(data.length),
      u16(nameBytes.length), u16(0), nameBytes,
    ]);
    chunks.push(local, data);
    central.push(concat([
      u32(0x02014b50), u16(20), u16(20), u16(flags), u16(0), u16(0), u16(0),
      u32(crc), u32(data.length), u32(data.length),
      u16(nameBytes.length), u16(0), u16(0), u16(0), u16(0), u32(0),
      u32(offset), nameBytes,
    ]));
    offset += local.length + data.length;
  }
  const centralStart = offset;
  let centralSize = 0;
  for (const c of central) { chunks.push(c); centralSize += c.length; }
  chunks.push(concat([
    u32(0x06054b50), u16(0), u16(0),
    u16(files.length), u16(files.length),
    u32(centralSize), u32(centralStart), u16(0),
  ]));
  return concat(chunks);
}
function bytesToBase64(bytes) {
  let bin = '';
  const CH = 0x8000;
  for (let i = 0; i < bytes.length; i += CH) bin += String.fromCharCode.apply(null, bytes.subarray(i, i + CH));
  return btoa(bin);
}

// ---------- fetch a single file (rich errors for diagnosis) ----------

async function fetchFileBytes(ctaUrl) {
  let res;
  try {
    res = await fetch(ctaUrl, { credentials: 'include', redirect: 'follow' });
  } catch (e) {
    throw new Error('network: ' + String(e && e.message || e));
  }
  const host = (() => { try { return new URL(res.url).host; } catch (e) { return '?'; } })();
  const ct = res.headers.get('content-type') || '';
  const buf = new Uint8Array(await res.arrayBuffer());
  if (!res.ok) throw new Error('HTTP ' + res.status + ' @ ' + host + ' ct=' + ct.slice(0, 30));
  if (buf.length < 4 || buf[0] !== 0x50 || buf[1] !== 0x4b) {
    // not a zip/xlsx - likely an HTML login/error page (cookie not sent) or empty.
    let head = '';
    try { head = new TextDecoder().decode(buf.subarray(0, 60)).replace(/\s+/g, ' '); } catch (e) {}
    throw new Error('no PK @ ' + host + ' (' + buf.length + 'b, ct=' + ct.slice(0, 25) + ', "' + head.slice(0, 40) + '")');
  }
  return buf;
}

function sanitize(name) {
  return String(name).replace(/[\\/:*?"<>|]/g, '_').replace(/[\u0000-\u001f]/g, '')
    .replace(/\s+/g, ' ').trim().slice(0, 120) || 'file';
}
function ensureUnique(name, used) {
  if (!used.has(name)) { used.add(name); return name; }
  const dot = name.lastIndexOf('.');
  const base = dot > 0 ? name.slice(0, dot) : name;
  const ext = dot > 0 ? name.slice(dot) : '';
  let i = 2, candidate;
  do { candidate = base + ' (' + (i++) + ')' + ext; } while (used.has(candidate));
  used.add(candidate);
  return candidate;
}

// ---------- message port ----------

chrome.runtime.onConnect.addListener((port) => {
  if (port.name !== 'bulkExport') return;
  port.onMessage.addListener(async (msg) => {
    if (!msg || msg.type !== 'start') return;
    const { zipName, items } = msg;
    const folderName = sanitize(zipName || 'hubspot-recipient-lists');
    const files = [];
    const used = new Set();
    let firstError = null;
    try {
      for (let i = 0; i < items.length; i++) {
        const it = items[i];
        port.postMessage({ type: 'progress', stage: 'downloading', index: i, total: items.length, name: it.name });
        try {
          const bytes = await fetchFileBytes(it.ctaUrl);
          files.push({ name: folderName + '/' + ensureUnique(sanitize(it.name) + '.xlsx', used), bytes });
        } catch (e) {
          if (!firstError) firstError = String(e.message || e);
          port.postMessage({ type: 'fileError', index: i, name: it.name, error: String(e.message || e) });
        }
      }
      if (!files.length) throw new Error('No files downloaded. First failure: ' + (firstError || 'unknown'));

      port.postMessage({ type: 'progress', stage: 'zipping', count: files.length });
      const zip = buildZip(files);
      const dataUrl = 'data:application/zip;base64,' + bytesToBase64(zip);
      const filename = folderName + '.zip';
      chrome.downloads.download({ url: dataUrl, filename, saveAs: false }, () => {
        if (chrome.runtime.lastError) port.postMessage({ type: 'error', error: chrome.runtime.lastError.message });
        else port.postMessage({ type: 'done', downloaded: files.length, requested: items.length, filename });
      });
    } catch (e) {
      port.postMessage({ type: 'error', error: String(e.message || e) });
    }
  });
});

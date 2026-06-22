"use strict";

const logBox = document.getElementById("logBox");
const busy = document.getElementById("busy");

let DATA_DIR = "";
let STEPS = [];

const DOWNLOAD_KEYS = {
  "1": [["1", "BY Countries"]],
  "2": [["2", "BY Tiers"]],
  "3": [["3", "Main"]],
  "4": [["4a", "BY Publication"], ["4b", "Master Panel"]],
  "5": [["5", "masterlist.xlsx"]],
};

function log(text, cls) {
  const stamp = new Date().toLocaleTimeString();
  if (logBox.textContent.startsWith("Ready.")) logBox.textContent = "";
  const span = document.createElement("span");
  if (cls) span.className = cls;
  span.textContent = `\n[${stamp}]\n${text}\n`;
  logBox.appendChild(span);
  logBox.scrollTop = logBox.scrollHeight;
}

function setBusy(on) {
  busy.classList.toggle("hidden", !on);
  document.querySelectorAll("button").forEach((b) => (b.disabled = on));
}

function esc(s) {
  return (s || "").replace(/[&<>"]/g, (c) => ({ "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;" }[c]));
}

async function loadInfo() {
  const data = await (await fetch("/api/info")).json();
  DATA_DIR = data.data_dir;
  STEPS = data.steps;
  document.getElementById("dataDir").textContent = DATA_DIR;
  renderSteps();
  refreshStatus();
}

function renderSteps() {
  const list = document.getElementById("stepList");
  list.innerHTML = "";
  let lastPhase = null;

  for (const step of STEPS) {
    if (step.phase !== lastPhase) {
      const ph = document.createElement("div");
      ph.className = "phase";
      ph.textContent = step.phase;
      list.appendChild(ph);
      lastPhase = step.phase;
    }

    const openBtn = step.input_folder
      ? `<button class="btn btn-sm" data-open="${esc(step.input_folder)}">📂 Open folder</button>`
      : "";
    const downloads = (DOWNLOAD_KEYS[step.id] || [])
      .map(([k, label]) => `<a class="btn btn-sm" href="/download/${k}">⬇ ${esc(label)}</a>`)
      .join("");
    const staticTag = step.static ? `<span class="static-tag">rarely changes</span>` : "";

    const card = document.createElement("div");
    card.className = "step";
    card.id = `step-${step.id}`;
    card.innerHTML = `
      <div class="step-row">
        <div class="step-num">${step.id}</div>
        <div class="step-main">
          <div class="step-title">${esc(step.name)}${staticTag}</div>
          <div class="step-does">${esc(step.does)}</div>
        </div>
        <div class="step-side">
          <span class="pill" id="badge-${step.id}">—</span>
          <button class="btn btn-run" data-run="${step.id}">Run</button>
        </div>
      </div>
      <details class="step-details">
        <summary>Files &amp; details</summary>
        <div class="details-body">
          <p><b>You provide:</b> ${esc(step.provide)}</p>
          <p><b>Produces:</b> <code>${esc(step.output)}</code></p>
          <div class="row-actions">${openBtn}${downloads}</div>
        </div>
      </details>`;
    list.appendChild(card);
  }
  attachHandlers();
}

function attachHandlers() {
  document.querySelectorAll("[data-run]").forEach((b) => (b.onclick = () => runStep(b.dataset.run)));
  document.querySelectorAll("[data-open]").forEach((b) => (b.onclick = () => openFolder(b.dataset.open)));
}

async function runStep(id) {
  setBusy(true);
  log(`Running step ${id}…`);
  try {
    const data = await (await fetch(`/api/run/${id}`, { method: "POST" })).json();
    log(data.log || "(no output)", data.ok ? "ok" : "err");
    if (!data.ok) log(`Step ${id} did not finish. See the message above.`, "err");
  } catch (e) {
    log(`Network error running step ${id}: ${e}`, "err");
  } finally {
    setBusy(false);
    refreshStatus();
  }
}

async function runAll() {
  for (const step of STEPS) {
    setBusy(true);
    log(`Running step ${step.id} (${step.name})…`);
    let ok = false;
    try {
      const data = await (await fetch(`/api/run/${step.id}`, { method: "POST" })).json();
      log(data.log || "(no output)", data.ok ? "ok" : "err");
      ok = data.ok;
    } catch (e) {
      log(`Network error: ${e}`, "err");
    }
    await refreshStatus();
    if (!ok) {
      log(`Stopped at step ${step.id} — it did not finish successfully.`, "err");
      setBusy(false);
      return;
    }
  }
  setBusy(false);
  log("✓ All steps finished. Download the masterlist from step 5.", "ok");
}

async function openFolder(rel) {
  try {
    await fetch("/api/open-folder", {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({ path: rel || "" }),
    });
  } catch (e) {
    log(`Could not open folder: ${e}`, "err");
  }
}

async function refreshStatus() {
  try {
    const data = await (await fetch("/api/status")).json();
    for (const step of STEPS) {
      const badge = document.getElementById(`badge-${step.id}`);
      const card = document.getElementById(`step-${step.id}`);
      if (!badge) continue;
      const s = data.steps[step.id] || {};
      const ready = isOutputReady(step.id, data.outputs);
      card.classList.toggle("done", ready);
      if (ready) {
        badge.textContent = "✓ done";
        badge.className = "pill ok";
      } else if (typeof s.input_count === "number") {
        badge.textContent = s.input_count > 0 ? `${s.input_count} file(s) ready` : "no files yet";
        badge.className = s.input_count > 0 ? "pill has" : "pill";
      } else {
        badge.textContent = "needs earlier steps";
        badge.className = "pill";
      }
    }
  } catch (e) {
    /* ignore transient status errors */
  }
}

function isOutputReady(stepId, outputs) {
  if (stepId === "4") return outputs["4a"] && outputs["4b"];
  return !!outputs[stepId];
}

async function loadExtension() {
  try {
    const m = await (await fetch("/api/extension")).json();
    if (m.name) document.getElementById("extName").textContent = m.name;
    if (m.description) document.getElementById("extDesc").textContent = m.description;
    if (m.version) document.getElementById("extVer").textContent = "v" + m.version;
  } catch (e) {
    /* keep defaults */
  }
}

document.getElementById("runAllBtn").onclick = runAll;
document.getElementById("clearLog").onclick = () => (logBox.textContent = "Ready.");
document.getElementById("copyDir").onclick = async () => {
  try {
    await navigator.clipboard.writeText(DATA_DIR);
    log("Folder path copied to clipboard.", "ok");
  } catch {
    log(DATA_DIR);
  }
};

loadInfo();
loadExtension();
setInterval(refreshStatus, 4000);

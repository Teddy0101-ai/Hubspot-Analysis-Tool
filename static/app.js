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
  const line = `\n[${stamp}]\n${text}\n`;
  if (logBox.textContent.startsWith("Ready.")) logBox.textContent = "";
  const span = document.createElement("span");
  if (cls) span.className = cls;
  span.textContent = line;
  logBox.appendChild(span);
  logBox.scrollTop = logBox.scrollHeight;
}

function setBusy(on) {
  busy.classList.toggle("hidden", !on);
  document.querySelectorAll("button").forEach((b) => (b.disabled = on));
}

async function loadInfo() {
  const res = await fetch("/api/info");
  const data = await res.json();
  DATA_DIR = data.data_dir;
  STEPS = data.steps;
  document.getElementById("dataDir").textContent = DATA_DIR;
  renderSteps();
  refreshStatus();
}

function renderSteps() {
  const list = document.getElementById("stepList");
  list.innerHTML = "";
  for (const step of STEPS) {
    const card = document.createElement("div");
    card.className = "step";
    card.id = `step-${step.id}`;

    const openBtn = step.input_folder
      ? `<button class="btn btn-ghost" data-open="${step.input_folder.replace(DATA_DIR, "").replace(/^[\\/]+/, "")}">📂 Open input folder</button>`
      : "";

    const downloads = (DOWNLOAD_KEYS[step.id] || [])
      .map(([k, label]) => `<a class="btn btn-ghost" href="/download/${k}">⬇ ${label}</a>`)
      .join(" ");

    card.innerHTML = `
      <div class="step-head">
        <span class="step-title"><span class="step-num">${step.id}</span>${step.title.replace(/^Step \d+ — /, "")}</span>
        <span class="step-actions">
          <span class="badge" id="badge-${step.id}">—</span>
          ${openBtn}
          <button class="btn btn-run" data-run="${step.id}">▶ Run</button>
        </span>
      </div>
      <div class="step-needs">${step.needs}</div>
      <div class="step-output">Output: <b>${step.output}</b></div>
      <div class="step-actions" style="margin-top:10px">${downloads}</div>
    `;
    list.appendChild(card);
  }
  attachHandlers();
}

function attachHandlers() {
  document.querySelectorAll("[data-run]").forEach((btn) => {
    btn.onclick = () => runStep(btn.getAttribute("data-run"));
  });
  document.querySelectorAll("[data-open]").forEach((btn) => {
    btn.onclick = () => openFolder(btn.getAttribute("data-open"));
  });
}

async function runStep(id) {
  setBusy(true);
  log(`Running Step ${id}…`);
  try {
    const res = await fetch(`/api/run/${id}`, { method: "POST" });
    const data = await res.json();
    log(data.log || "(no output)", data.ok ? "ok" : "err");
    if (!data.ok) log(`Step ${id} did not complete. See the message above.`, "err");
  } catch (e) {
    log(`Network error running Step ${id}: ${e}`, "err");
  } finally {
    setBusy(false);
    refreshStatus();
  }
}

async function runAll() {
  for (const step of STEPS) {
    setBusy(true);
    log(`Running Step ${step.id}…`);
    let ok = false;
    try {
      const res = await fetch(`/api/run/${step.id}`, { method: "POST" });
      const data = await res.json();
      log(data.log || "(no output)", data.ok ? "ok" : "err");
      ok = data.ok;
    } catch (e) {
      log(`Network error: ${e}`, "err");
    }
    await refreshStatus();
    if (!ok) {
      log(`Stopped at Step ${step.id} because it did not finish successfully.`, "err");
      setBusy(false);
      return;
    }
  }
  setBusy(false);
  log("All steps finished.", "ok");
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
    const res = await fetch("/api/status");
    const data = await res.json();
    for (const step of STEPS) {
      const badge = document.getElementById(`badge-${step.id}`);
      if (!badge) continue;
      const s = data.steps[step.id] || {};
      const outReady = isOutputReady(step.id, data.outputs);
      if (outReady) {
        badge.textContent = "✓ output ready";
        badge.className = "badge ok";
      } else if (typeof s.input_count === "number") {
        badge.textContent = `${s.input_count} input file(s)`;
        badge.className = s.input_count > 0 ? "badge has" : "badge";
      } else {
        badge.textContent = "needs earlier steps";
        badge.className = "badge";
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

document.getElementById("runAllBtn").onclick = runAll;
document.getElementById("clearLog").onclick = () => (logBox.textContent = "Ready.");
document.getElementById("copyDir").onclick = async () => {
  try {
    await navigator.clipboard.writeText(DATA_DIR);
    log("Data folder path copied to clipboard.", "ok");
  } catch {
    log(DATA_DIR);
  }
};
document.querySelector('[data-open=""]').onclick = () => openFolder("");

async function loadExtension() {
  try {
    const res = await fetch("/api/extension");
    const m = await res.json();
    if (m.name) document.getElementById("extName").textContent = m.name;
    if (m.description) document.getElementById("extDesc").textContent = m.description;
    const ver = document.getElementById("extVer");
    if (m.version) ver.textContent = "v" + m.version;
  } catch (e) {
    /* extension section keeps its default text if this fails */
  }
}

loadInfo();
loadExtension();
setInterval(refreshStatus, 4000);

// MS Download -> Upload dashboard. Every URL is relative, so the page works
// under any subpath behind a reverse proxy. All data from the
// server -- file names, URLs, user names -- goes into the DOM as text only.
"use strict";

const STATUS = {
  ok: ["Done", "ok"],
  "no-uba": ["No UBA alerts", "no-uba"],
  failed: ["Failed", "failed"],
  interrupted: ["Interrupted", "interrupted"],
  running: ["Running", "running"],
};
const POLL_RUNNING_MS = 4000;
const POLL_IDLE_MS = 30000;
// Alerts columns shown first, so where the file went and how fast fit on screen;
// the rest of the workbook's columns follow in their own order.
const ALERT_COLUMNS_FIRST = [
  "Alert ID", "User", "File", "Upload App", "Gap (Hours)", "Uploaded (GMT)",
  "Upload Host Name", "Upload Source Country", "Download App", "Downloaded (GMT)", "UBA alert_name",
];

let state = null;          // last /api/state
let selectedId = null;     // run shown in Results
let detail = null;         // last /api/runs/<id>
let formReady = false;
let pollTimer = null;
const sortBy = { alerts: null, bulk: null };

const $ = (id) => document.getElementById(id);

function el(tag, props = {}, ...children) {
  const node = document.createElement(tag);
  for (const [k, v] of Object.entries(props)) {
    if (k === "class") node.className = v;
    else if (k === "text") node.textContent = v;
    else if (k.startsWith("on")) node.addEventListener(k.slice(2), v);
    else node.setAttribute(k, v);
  }
  for (const c of children) if (c != null) node.append(c);
  return node;
}

async function api(path, options) {
  const resp = await fetch(path, Object.assign({ cache: "no-store" }, options));
  let body = null;
  try { body = await resp.json(); } catch (_) { /* not JSON */ }
  if (!resp.ok) throw new Error((body && body.error) || `HTTP ${resp.status}`);
  return body;
}

// ── formatting ──────────────────────────────────────────────────────────────
function when(iso) {
  if (!iso) return "—";
  const d = new Date(iso);
  return isNaN(d) ? iso : d.toLocaleString(undefined, { dateStyle: "medium", timeStyle: "short" });
}
function duration(s) {
  if (s == null) return "—";
  if (s < 60) return `${Math.round(s)}s`;
  const m = Math.floor(s / 60);
  return m < 60 ? `${m}m ${Math.round(s % 60)}s` : `${Math.floor(m / 60)}h ${m % 60}m`;
}
function range(run) {
  return run.start === run.end ? run.start : `${run.start} → ${run.end}`;
}
function num(n) {
  return n == null ? "—" : Number(n).toLocaleString();
}
function badge(status) {
  const [label, cls] = STATUS[status] || [status, "no-uba"];
  return el("span", { class: `badge ${cls}`, text: label });
}

// ── status row ──────────────────────────────────────────────────────────────
function renderStatus() {
  const cfg = state.config;
  if (cfg.ready) {
    $("cfg-value").className = "value ok-text";
    $("cfg-value").textContent = "Netskope token set";
    $("cfg-sub").textContent = `${cfg.tenant} · auth ${cfg.auth_mode}`;
  } else {
    $("cfg-value").className = "value bad-text";
    $("cfg-value").textContent = "No Netskope token";
    $("cfg-sub").textContent = "Set NETSKOPE_API_TOKEN and NETSKOPE_BASE_URL in "
      + "config.env (see config.env.example), then reload.";
  }

  const sched = state.schedule;
  const lastDaily = state.runs.find((r) => r.trigger === "daily");
  // Scheduling is up to the host (cron or a systemd timer running --run-daily);
  // a systemd timer's next trigger is shown when there is one.
  $("sched-value").className = "value";
  $("sched-value").textContent = lastDaily
    ? `Last: ${range(lastDaily)} · ${(STATUS[lastDaily.status] || [lastDaily.status])[0]}`
    : "No daily run yet";
  $("sched-sub").textContent = sched.next ? `Next: ${sched.next}`
    : "Schedule: python server.py --run-daily (cron or a systemd timer)";

  const running = state.running;
  if (running) {
    $("now-value").className = "value run-text";
    $("now-value").replaceChildren(el("span", { class: "spin", text: "●" }), ` Running ${range(running)}`);
    $("now-sub").textContent = `Started ${when(running.started_at)} by ${running.by}`;
  } else {
    $("now-value").className = "value";
    $("now-value").textContent = "Idle";
    $("now-sub").textContent = `${state.runs.length} run(s) kept (last ${state.defaults.keep_runs})`;
  }
}

// ── form ────────────────────────────────────────────────────────────────────
function setupForm() {
  const d = state.defaults;
  const today = new Date().toISOString().slice(0, 10);
  for (const id of ["f-start", "f-end"]) {
    $(id).value = d.yesterday;
    $(id).max = today;
  }
  $("f-hours").value = d.hours;
  $("f-hours").max = d.max_hours;
  const box = $("f-acts");
  box.replaceChildren(...d.activities.map((a) =>
    el("label", {}, el("input", { type: "checkbox", value: a, checked: "" }), a)));
  formReady = true;
}

function renderFormState() {
  const busy = Boolean(state.running);
  const btn = $("run-btn");
  btn.disabled = busy || !state.config.ready;
  btn.textContent = busy ? "Running…" : "Run";
  const msg = $("form-msg");
  if (!msg.classList.contains("bad")) {
    msg.textContent = busy ? "One run at a time — this page updates when it finishes."
      : !state.config.ready ? "Add the Netskope token first (see Credentials)."
      : `Up to ${state.defaults.max_range_days} days per run.`;
  }
}

async function submitRun(event) {
  event.preventDefault();
  const msg = $("form-msg");
  msg.classList.remove("bad");
  const activities = [...$("f-acts").querySelectorAll("input:checked")].map((i) => i.value);
  if (!activities.length) {
    msg.classList.add("bad");
    msg.textContent = "Pick at least one UBA activity.";
    return;
  }
  const payload = {
    start: $("f-start").value,
    end: $("f-end").value || $("f-start").value,
    hours: Number($("f-hours").value),
    activities,
  };
  $("run-btn").disabled = true;
  try {
    const { run } = await api("api/runs", {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify(payload),
    });
    selectedId = run.id;
    msg.textContent = `Started ${range(run)}.`;
  } catch (err) {
    msg.classList.add("bad");
    msg.textContent = err.message;
  }
  await refresh();
}

// ── tables ──────────────────────────────────────────────────────────────────
function isNum(v) {
  return typeof v === "number" || (typeof v === "string" && v !== "" && !isNaN(v));
}

function fillTable(table, header, rows, key, emptyText) {
  const sort = sortBy[key];
  let shown = rows;
  if (sort) {
    const i = sort.col;
    shown = [...rows].sort((a, b) => {
      const x = a[i], y = b[i];
      const r = isNum(x) && isNum(y) ? Number(x) - Number(y) : String(x).localeCompare(String(y));
      return sort.desc ? -r : r;
    });
  }
  const head = el("tr", {}, ...header.map((h, i) => {
    const mark = sort && sort.col === i ? (sort.desc ? " ▼" : " ▲") : "";
    return el("th", { scope: "col", "aria-sort": sort && sort.col === i ? (sort.desc ? "descending" : "ascending") : "none" },
      el("button", { type: "button", text: h + mark, onclick: () => {
        sortBy[key] = sort && sort.col === i ? { col: i, desc: !sort.desc } : { col: i, desc: false };
        renderTables();
      } }));
  }));
  const body = el("tbody");
  if (!shown.length) {
    body.append(el("tr", {}, el("td", { class: "empty", colspan: String(Math.max(header.length, 1)), text: emptyText })));
  }
  for (const row of shown) {
    body.append(el("tr", {}, ...row.map((v) => el("td", {
      class: isNum(v) ? "num" : "", title: String(v), text: String(v),
    }))));
  }
  table.replaceChildren(el("thead", {}, head), body);
}

function reorder(table, first) {
  const order = [
    ...first.map((h) => table.header.indexOf(h)).filter((i) => i >= 0),
    ...table.header.map((_, i) => i).filter((i) => !first.includes(table.header[i])),
  ];
  return Object.assign({}, table, {
    header: order.map((i) => table.header[i]),
    rows: table.rows.map((r) => order.map((i) => r[i])),
  });
}

function renderTables() {
  if (!detail) return;
  const s = detail.summary || {};
  const alerts = reorder(s.alerts || { header: [], rows: [], total: 0 }, ALERT_COLUMNS_FIRST);
  const q = $("alert-filter").value.trim().toLowerCase();
  const rows = q ? alerts.rows.filter((r) => r.some((v) => String(v).toLowerCase().includes(q))) : alerts.rows;
  fillTable($("alerts-table"), alerts.header, rows, "alerts",
    detail.run.status === "running" ? "Running…" : "No alerts in this run.");
  let note = q ? `${rows.length} of ${alerts.rows.length} shown` : `${num(alerts.total)} alert(s)`;
  if (alerts.total > alerts.rows.length) {
    note += ` — the first ${num(alerts.rows.length)} are listed here; the workbook has all of them`;
  }
  $("alerts-note").textContent = alerts.header.length ? note + "." : "";

  const bulk = s.bulk || { title: "UBA 300+ Files", header: [], rows: [], total: 0 };
  $("bulk-sum").textContent = `${bulk.title} (${num(bulk.total)})`;
  fillTable($("bulk-table"), bulk.header, bulk.rows, "bulk", "No UBA alerts of that size.");
}

// ── results ─────────────────────────────────────────────────────────────────
function kpi(n, label, gold) {
  return el("div", { class: "kpi" }, el("div", { class: gold ? "n gold" : "n", text: num(n) }),
    el("div", { class: "l", text: label }));
}

function renderDetail() {
  const box = $("results");
  if (!detail) { box.hidden = true; return; }
  box.hidden = false;
  const run = detail.run;
  $("res-h").replaceChildren(`Results · ${range(run)} `, badge(run.status));
  $("res-sub").textContent = `${run.trigger === "daily" ? "Daily run" : `Run by ${run.by}`} · started ${when(run.started_at)}`
    + ` · ${run.finished_at ? `took ${duration(run.duration_s)}` : "in progress"}`
    + ` · ${run.hours === 0 ? "any gap" : `≤ ${run.hours} h`} · ${run.activities.join(", ")}`;

  const dl = $("dl-btn");
  dl.hidden = !run.workbook;
  if (run.workbook) dl.href = `api/runs/${encodeURIComponent(run.id)}/download`;

  const err = $("res-error");
  const explain = run.status === "no-uba"
    ? "No UBA alerts with these activities in this range, so no user was in scope. Nothing to correlate."
    : run.error;
  err.hidden = !explain;
  err.textContent = explain || "";
  err.className = run.status === "no-uba" ? "sub" : "error";

  const c = run.counts;
  $("kpis").replaceChildren(...(c ? [
    kpi(c.alerts, "Alerts (outside uploads)", true),
    kpi(c.users, c.users_partial ? "Users (in the rows listed)" : "Users"),
    kpi(c.bulk, `UBA ${state.defaults.bulk_files}+ file alerts`),
    kpi(c.uba_alerts, "UBA alerts in scope"),
    kpi(c.events, "Events pulled"),
  ] : []));

  renderTables();
  $("log").textContent = detail.log_tail || "(empty)";
  $("log-full").href = `api/runs/${encodeURIComponent(run.id)}/log`;
}

async function loadDetail(id) {
  try {
    detail = await api(`api/runs/${encodeURIComponent(id)}`);
  } catch (_) {
    detail = null;
  }
  renderDetail();
}

// ── history ─────────────────────────────────────────────────────────────────
function renderHistory() {
  const runs = state.runs;
  $("hist-sub").textContent = runs.length
    ? `Newest first. The last ${state.defaults.keep_runs} runs are kept; older ones are deleted.`
    : "No runs yet.";
  const head = el("tr", {}, ...["Range", "Status", "Alerts", "Users", "Trigger", "Started", "Took", ""]
    .map((h) => el("th", { scope: "col", class: h === "Alerts" || h === "Users" ? "num" : "", text: h })));
  const body = el("tbody");
  for (const run of runs) {
    const c = run.counts || {};
    body.append(el("tr", { class: run.id === selectedId ? "selected" : "" },
      el("td", {}, el("button", { type: "button", class: "linkish", text: range(run),
        onclick: () => select(run.id) })),
      el("td", {}, badge(run.status)),
      el("td", { class: "num", text: num(c.alerts) }),
      el("td", { class: "num", text: num(c.users) }),
      el("td", { text: run.trigger === "daily" ? "Daily" : run.by }),
      el("td", { text: when(run.started_at) }),
      el("td", { text: duration(run.duration_s) }),
      el("td", {}, run.workbook ? el("a", { href: `api/runs/${encodeURIComponent(run.id)}/download`, text: "Excel" }) : "")));
  }
  if (!runs.length) body.append(el("tr", {}, el("td", { class: "empty", colspan: "8", text: "Nothing has run yet." })));
  $("hist-table").replaceChildren(el("thead", {}, head), body);
}

function select(id) {
  selectedId = id;
  sortBy.alerts = sortBy.bulk = null;
  $("alert-filter").value = "";
  renderHistory();
  loadDetail(id);
}

// ── polling ─────────────────────────────────────────────────────────────────
async function refresh() {
  clearTimeout(pollTimer);
  const wasRunning = state && state.running ? state.running.id : null;
  try {
    state = await api("api/state");
  } catch (err) {
    $("now-value").className = "value bad-text";
    $("now-value").textContent = "Dashboard unreachable";
    $("now-sub").textContent = err.message;
    pollTimer = setTimeout(refresh, POLL_IDLE_MS);
    return;
  }
  if (!formReady) setupForm();
  renderStatus();
  renderFormState();
  if (!selectedId || !state.runs.some((r) => r.id === selectedId)) {
    selectedId = state.runs.length ? state.runs[0].id : null;
  }
  renderHistory();
  const sel = state.runs.find((r) => r.id === selectedId);
  // Reload the shown run when it changed: a new selection, still running, or just finished.
  if (sel && (!detail || detail.run.id !== sel.id || sel.status === "running"
              || detail.run.status !== sel.status || wasRunning === sel.id)) {
    await loadDetail(sel.id);
  } else if (!sel) {
    detail = null;
    renderDetail();
  }
  pollTimer = setTimeout(refresh, state.running ? POLL_RUNNING_MS : POLL_IDLE_MS);
}

$("run-form").addEventListener("submit", submitRun);
$("alert-filter").addEventListener("input", renderTables);
document.addEventListener("visibilitychange", () => { if (!document.hidden) refresh(); });
refresh();

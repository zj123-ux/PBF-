"use strict";

const $ = (id) => document.getElementById(id);

let INIT = null;
let currentRunId = null;
let pollTimer = null;
let uploadedPeriod = null;
let hasUploaded = false;

async function getJSON(url, opts) {
  const res = await fetch(url, opts);
  return res.json();
}

function esc(s) {
  return String(s).replace(/[&<>"']/g, (c) => ({
    "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;", "'": "&#39;",
  }[c]));
}

function setStatus(el, text, cls) {
  el.textContent = text;
  el.className = "status" + (cls ? " " + cls : "");
}

function setRunning(on) {
  document.body.classList.toggle("is-running", !!on);
}

function fmt(v) {
  if (v === null || v === undefined) return "";
  if (typeof v === "number") {
    return v.toLocaleString("zh-CN", { maximumFractionDigits: 4 });
  }
  return String(v);
}

// ---------- 日志（按行着色） ----------

function logClass(line) {
  if (/失败|错误|异常|Traceback|ERROR/i.test(line)) return "lerr";
  if (/警告|WARNING/i.test(line)) return "lwarn";
  if (/完成|成功|已保存|已刷新|✓/.test(line)) return "lok";
  if (/^\s*\[Web\]/.test(line)) return "linfo";
  return "";
}

function appendLog(text) {
  if (!text) return;
  const log = $("log");
  const html = text.split("\n").map((ln) => {
    const cls = logClass(ln);
    const safe = esc(ln);
    return cls ? `<span class="${cls}">${safe}</span>` : safe;
  }).join("\n");
  log.insertAdjacentHTML("beforeend", html);
  log.scrollTop = log.scrollHeight;
}

function clearLog() {
  $("log").innerHTML = "";
}

// ---------- 初始化 ----------

async function init() {
  INIT = await getJSON("/api/init");

  // 账号
  $("username").value = INIT.username || "";
  $("client").value = INIT.client || "333";
  $("lang").value = INIT.lang || "zh";
  if (INIT.password_set) {
    $("password").placeholder = "已保存（留空则不改）";
  }

  // 品牌×品类
  const unitSel = $("unit");
  unitSel.innerHTML = "";
  for (const u of INIT.units) {
    const opt = document.createElement("option");
    opt.value = u.name;
    opt.textContent = u.enabled ? `${u.name}（${u.brand}·${u.category}）`
                                : `${u.name}（未配置）`;
    opt.disabled = !u.enabled;
    unitSel.appendChild(opt);
  }
  // 默认选服务端指定的启用单元（否则取第一个启用的）
  const firstEnabled = INIT.units.find((u) => u.enabled);
  if (INIT.default_unit) unitSel.value = INIT.default_unit;
  else if (firstEnabled) unitSel.value = firstEnabled.name;

  // 版本
  const verSel = $("version");
  verSel.innerHTML = "";
  const empty = document.createElement("option");
  empty.value = "";
  empty.textContent = "（不指定 / 仅校验）";
  verSel.appendChild(empty);
  for (const v of INIT.versions) {
    const opt = document.createElement("option");
    opt.value = v.key;
    opt.textContent = v.display;
    verSel.appendChild(opt);
  }

  // 期间
  setupPeriodSelects();
}

// ---------- 期间选择 ----------

function addMonths(y, mo, delta) {
  const total = y * 12 + (mo - 1) + delta;
  return [Math.floor(total / 12), (total % 12) + 1];
}

function setupPeriodSelects() {
  const dir = (INIT.latest_data_dir || "").replace(/^data_/, "");
  const parts = dir.split("_").filter(Boolean);
  let start = parts[0] || INIT.suggested_period || "";
  let end = parts[parts.length - 1] || start;

  const m = /^(\d{4})\.(\d{2})$/.exec(start);
  const today = new Date();
  let y = m ? Number(m[1]) : today.getFullYear();
  let mo = m ? Number(m[2]) : today.getMonth() + 1;

  // 候选项：最新期间前 6 个月 ~ 后 12 个月（共 19 个月）
  const months = [];
  let [cy, cm] = addMonths(y, mo, -6);
  for (let i = 0; i < 19; i++) {
    months.push(`${cy}.${String(cm).padStart(2, "0")}`);
    [cy, cm] = addMonths(cy, cm, 1);
  }

  const startSel = $("period-start");
  const endSel = $("period-end");
  for (const mon of months) {
    startSel.insertAdjacentHTML("beforeend", `<option value="${mon}">${mon}</option>`);
    endSel.insertAdjacentHTML("beforeend", `<option value="${mon}">${mon}</option>`);
  }
  if (start) startSel.value = start;
  if (end) endSel.value = end;
  syncPeriodEnd();
  startSel.addEventListener("change", syncPeriodEnd);
}

// 结束期间不能早于起始期间
function syncPeriodEnd() {
  const start = $("period-start").value;
  const endSel = $("period-end");
  for (const opt of endSel.options) opt.disabled = opt.value < start;
  if (!endSel.value || endSel.value < start) endSel.value = start;
}

// 生成传给后端的期间描述：单个 2026.12 / 区间 2026.11-2026.12
function periodSpec() {
  const start = $("period-start").value;
  const end = $("period-end").value || start;
  return start === end ? start : `${start}-${end}`;
}

// ---------- 账号 ----------

$("btn-save-cred").addEventListener("click", async () => {
  const password = $("password").value;
  const body = {
    username: $("username").value,
    client: $("client").value,
    lang: $("lang").value,
  };
  // 留空且已保存过 → 不覆盖原密码
  if (password === "" && INIT && INIT.password_set) {
    body.password = null;
  } else {
    body.password = password;
  }
  const r = await getJSON("/api/credentials", {
    method: "POST",
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify(body),
  });
  setStatus($("cred-status"), r.ok ? "已保存 ✓" : ("失败：" + (r.error || "")),
            r.ok ? "ok" : "err");
  if (r.ok) INIT.password_set = true;
});

// ---------- 文件选择 ----------

$("file").addEventListener("change", () => {
  const f = $("file").files[0];
  $("file-name").textContent = f
    ? `${f.name} · ${(f.size / 1024).toFixed(0)} KB` : "";
  $("dropzone").classList.toggle("has-file", !!f);
});

(() => {
  // 拖拽上传：把拖入的文件塞进原生 input，触发 change
  const dz = $("dropzone");
  ["dragenter", "dragover"].forEach((ev) =>
    dz.addEventListener(ev, (e) => {
      e.preventDefault();
      dz.classList.add("dragover");
    }));
  ["dragleave", "drop"].forEach((ev) =>
    dz.addEventListener(ev, (e) => {
      e.preventDefault();
      dz.classList.remove("dragover");
    }));
  dz.addEventListener("drop", (e) => {
    if (e.dataTransfer.files.length) {
      $("file").files = e.dataTransfer.files;
      $("file").dispatchEvent(new Event("change"));
    }
  });
})();

// ---------- 上传 ----------

$("btn-upload").addEventListener("click", async () => {
  const file = $("file").files[0];
  if (!file) {
    $("upload-result").innerHTML = '<div class="note bad">请先选择《滚动预测模板表》</div>';
    return;
  }
  const period = periodSpec();
  const fd = new FormData();
  fd.append("file", file);
  fd.append("period", period);
  fd.append("unit", $("unit").value || "");   // 目录名带业务单元，各单元互不覆盖

  $("btn-upload").disabled = true;
  $("upload-result").innerHTML = '<div class="note">正在导出…</div>';
  let r;
  try {
    r = await getJSON("/api/upload", { method: "POST", body: fd });
  } catch (e) {
    $("upload-result").innerHTML =
      '<div class="note bad">导出失败：网络异常</div>';
    return;
  } finally {
    $("btn-upload").disabled = false;
  }
  if (!r.ok) {
    $("upload-result").innerHTML =
      `<div class="note bad">导出失败：${esc(r.error || "")}</div>`;
    return;
  }
  uploadedPeriod = period;
  hasUploaded = true;
  const lines = Object.entries(r.rows)
    .map(([k, v]) => `<li>${esc(k)}：<b>${esc(v)}</b> 行</li>`)
    .join("");
  let html = '<div class="note ok">' +
    `<div class="note-title">导出成功 · 共 ${esc(r.total)} 行</div>` +
    `<div class="note-dir">→ ${esc(r.data_dir)}/</div>` +
    (lines ? `<ul>${lines}</ul>` : "");
  if (r.warnings && r.warnings.length) {
    html += `<ul>${r.warnings.map((w) => `<li>⚠ ${esc(w)}</li>`).join("")}</ul>`;
  }
  html += "</div>";
  $("upload-result").innerHTML = html;
  $("btn-run").disabled = false;
});

// ---------- 环境自检 ----------

$("btn-env").addEventListener("click", async () => {
  $("btn-env").disabled = true;
  $("env-result").innerHTML = '<div class="note">检测中…</div>';
  const qs = new URLSearchParams({
    unit: $("unit").value || "",
    period: periodSpec(),
  });
  let r;
  try {
    r = await getJSON(`/api/env?${qs}`);
  } catch (e) {
    $("env-result").innerHTML = '<div class="note bad">检测失败：网络异常</div>';
    return;
  } finally {
    $("btn-env").disabled = false;
  }
  const items = r.checks.map((c) => {
    const status = c.status || (c.ok ? "ok" : "error");
    const icon = status === "ok" ? "✓" : (status === "warn" ? "!" : "✗");
    return `<div class="env-item ${status}">` +
      `<span class="env-icon">${icon}</span>` +
      `<span class="env-label">${esc(c.label)}</span>` +
      `<span class="env-detail">${esc(c.detail)}</span></div>`;
  }).join("");
  $("env-result").innerHTML = `<div class="env-grid">${items}</div>`;
});

// ---------- 运行 ----------

function runParams() {
  const period = periodSpec();
  const verKey = $("version").value;
  const verDisplay = INIT.versions.find((v) => v.key === verKey)?.display || "";
  return {
    period,
    unit: $("unit").value,
    version_key: verKey,
    version_display: verDisplay,
    switch_version: $("switch-version").checked && !!verKey,
  };
}

function enterRunning(label) {
  clearLog();
  $("result").innerHTML = '<div class="empty">运行中，结束后显示损益核对结果</div>';
  $("btn-run").disabled = true;
  $("btn-clear").disabled = true;
  $("btn-release").style.display = "none";
  $("btn-cancel").disabled = false;
  setStatus($("run-status"), label, "running");
  setRunning(true);
}

function handleStartError(r, prefix) {
  setStatus($("run-status"), prefix + (r.error || ""), "err");
  if ((r.error || "").includes("槽位")) $("btn-release").style.display = "";
}

$("btn-run").addEventListener("click", async () => {
  const r = await getJSON("/api/run", {
    method: "POST",
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify(runParams()),
  });
  if (!r.ok) {
    handleStartError(r, "启动失败：");
    return;
  }
  currentRunId = r.run_id;
  enterRunning("运行中…");
  pollLog(0);
});

// ---------- 批量清除 ----------

$("btn-clear").addEventListener("click", async () => {
  const r = await getJSON("/api/clear", {
    method: "POST",
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify(runParams()),
  });
  if (!r.ok) {
    handleStartError(r, "清除启动失败：");
    return;
  }
  currentRunId = r.run_id;
  enterRunning("清除运行中…");
  pollLog(0);
});

function pollLog(offset) {
  pollTimer = setTimeout(async () => {
    const r = await getJSON(`/api/run/${currentRunId}/log?offset=${offset}`);
    if (!r.ok) {
      setStatus($("run-status"), "日志读取失败：" + (r.error || ""), "err");
      return;
    }
    appendLog(r.lines || "");
    if (r.done) {
      $("btn-cancel").disabled = true;
      setRunning(false);
      if (r.stuck) {
        // 取消/超时后 Excel 未确认退出：槽位仍被占用，先别放行新流程
        $("btn-release").style.display = "";
        $("btn-run").disabled = true;
        $("btn-clear").disabled = true;
        setStatus($("run-status"),
          "已结束，但 Excel 未退出：槽位被占用，确认无宏运行后可强制释放", "err");
      } else {
        $("btn-release").style.display = "none";
        $("btn-run").disabled = !hasUploaded;
        $("btn-clear").disabled = false;
        setStatus($("run-status"),
          r.exit_code === 0 ? "完成 ✓" : `结束（退出码 ${r.exit_code}）`,
          r.exit_code === 0 ? "ok" : "err");
      }
      loadResult();
      return;
    }
    pollLog(r.next_offset);
  }, 1000);
}

async function loadResult() {
  const r = await getJSON(`/api/run/${currentRunId}/result`);
  if (!r.ok || !r.report) {
    $("result").innerHTML =
      '<div class="empty">（无损益核对结果）</div>';
    return;
  }
  if (r.report.locked) {
    $("result").innerHTML =
      `<div class="note bad">${esc(r.report.locked.message || "表单被锁定，流程已中止")}</div>`;
    return;
  }
  if (!r.report.indicators) {
    $("result").innerHTML =
      '<div class="empty">（无损益核对结果）</div>';
    return;
  }
  const ind = r.report.indicators;
  const names = Object.keys(ind);
  const periods = [...new Set(names.flatMap((n) => Object.keys(ind[n])))];
  let html = '<div class="table-scroll"><table class="result"><thead><tr><th>指标</th>';
  for (const p of periods) html += `<th>${esc(p)}</th>`;
  html += "</tr></thead><tbody>";
  for (const n of names) {
    html += `<tr><td>${esc(n)}</td>`;
    for (const p of periods) {
      const v = ind[n][p];
      const num = typeof v === "number";
      const cls = "num" + (num && v < 0 ? " neg" : (num && v > 0 ? " pos" : ""));
      html += `<td class="${cls}">${esc(fmt(v))}</td>`;
    }
    html += "</tr>";
  }
  html += "</tbody></table></div>";
  $("result").innerHTML = html;
}

$("btn-cancel").addEventListener("click", async () => {
  if (!currentRunId) return;
  const r = await getJSON(`/api/run/${currentRunId}/cancel`, { method: "POST" });
  setStatus($("run-status"), r.note || "已请求取消", "running");
});

$("btn-release").addEventListener("click", async () => {
  const r = await getJSON("/api/run/release", { method: "POST" });
  if (!r.ok) {
    setStatus($("run-status"), "释放失败：" + (r.error || ""), "err");
    return;
  }
  $("btn-release").style.display = "none";
  $("btn-run").disabled = !hasUploaded;
  $("btn-clear").disabled = false;
  setRunning(false);
  setStatus($("run-status"), r.note || "已释放运行槽位", "ok");
});

init();

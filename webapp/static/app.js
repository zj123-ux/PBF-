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

function setStatus(el, text, cls) {
  el.textContent = text;
  el.className = "status" + (cls ? " " + cls : "");
}

function fmt(v) {
  if (v === null || v === undefined) return "";
  if (typeof v === "number") {
    return v.toLocaleString("zh-CN", { maximumFractionDigits: 4 });
  }
  return String(v);
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
  $("period").value = INIT.suggested_period || INIT.latest_data_dir.replace("data_", "");
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

// ---------- 上传 ----------

$("btn-upload").addEventListener("click", async () => {
  const file = $("file").files[0];
  if (!file) {
    $("upload-result").textContent = "请先选择《滚动预测模板表.xlsx》";
    return;
  }
  const period = $("period").value.trim();
  const fd = new FormData();
  fd.append("file", file);
  fd.append("period", period);

  $("upload-result").textContent = "正在导出…";
  const r = await getJSON("/api/upload", { method: "POST", body: fd });
  if (!r.ok) {
    $("upload-result").textContent = "导出失败：" + (r.error || "");
    return;
  }
  uploadedPeriod = period;
  hasUploaded = true;
  const lines = Object.entries(r.rows).map(([k, v]) => `  ${k}: ${v} 行`);
  let text = `导出成功，共 ${r.total} 行 → ${r.data_dir}/\n` + lines.join("\n");
  if (r.warnings && r.warnings.length) {
    text += "\n⚠ " + r.warnings.join("\n⚠ ");
  }
  $("upload-result").textContent = text;
  $("btn-run").disabled = false;
});

// ---------- 环境自检 ----------

$("btn-env").addEventListener("click", async () => {
  $("env-result").textContent = "检测中…";
  const r = await getJSON("/api/env");
  const lines = r.checks.map((c) => `${c.ok ? "✓" : "✗"} ${c.label}: ${c.detail}`);
  $("env-result").textContent = lines.join("\n");
});

// ---------- 运行 ----------

$("btn-run").addEventListener("click", async () => {
  const period = $("period").value.trim();
  const verKey = $("version").value;
  const verDisplay = INIT.versions.find((v) => v.key === verKey)?.display || "";
  const body = {
    period,
    unit: $("unit").value,
    version_key: verKey,
    version_display: verDisplay,
    switch_version: $("switch-version").checked && !!verKey,
  };
  const r = await getJSON("/api/run", {
    method: "POST",
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify(body),
  });
  if (!r.ok) {
    setStatus($("run-status"), "启动失败：" + (r.error || ""), "err");
    if ((r.error || "").includes("槽位")) $("btn-release").style.display = "";
    return;
  }
  currentRunId = r.run_id;
  $("log").textContent = "";
  $("result").innerHTML = "";
  $("btn-run").disabled = true;
  $("btn-clear").disabled = true;
  $("btn-release").style.display = "none";
  $("btn-cancel").disabled = false;
  setStatus($("run-status"), "运行中…", "");
  pollLog(0);
});

// ---------- 批量清除 ----------

$("btn-clear").addEventListener("click", async () => {
  const period = $("period").value.trim();
  const verKey = $("version").value;
  const verDisplay = INIT.versions.find((v) => v.key === verKey)?.display || "";
  const body = {
    period,
    unit: $("unit").value,
    version_key: verKey,
    version_display: verDisplay,
    switch_version: $("switch-version").checked && !!verKey,
  };
  const r = await getJSON("/api/clear", {
    method: "POST",
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify(body),
  });
  if (!r.ok) {
    setStatus($("run-status"), "清除启动失败：" + (r.error || ""), "err");
    if ((r.error || "").includes("槽位")) $("btn-release").style.display = "";
    return;
  }
  currentRunId = r.run_id;
  $("log").textContent = "";
  $("result").innerHTML = "";
  $("btn-run").disabled = true;
  $("btn-clear").disabled = true;
  $("btn-release").style.display = "none";
  $("btn-cancel").disabled = false;
  setStatus($("run-status"), "清除运行中…", "");
  pollLog(0);
});

function pollLog(offset) {
  pollTimer = setTimeout(async () => {
    const r = await getJSON(`/api/run/${currentRunId}/log?offset=${offset}`);
    if (!r.ok) {
      setStatus($("run-status"), "日志读取失败：" + (r.error || ""), "err");
      return;
    }
    if (r.lines) {
      $("log").textContent += r.lines;
      $("log").scrollTop = $("log").scrollHeight;
    }
    if (r.done) {
      $("btn-cancel").disabled = true;
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
  if (!r.ok || !r.report || !r.report.indicators) {
    $("result").textContent = "（无损益核对结果）";
    return;
  }
  const ind = r.report.indicators;
  const names = Object.keys(ind);
  const periods = [...new Set(names.flatMap((n) => Object.keys(ind[n])))];
  let html = '<table class="result"><thead><tr><th>指标</th>';
  for (const p of periods) html += `<th>${p}</th>`;
  html += "</tr></thead><tbody>";
  for (const n of names) {
    html += `<tr><td>${n}</td>`;
    for (const p of periods) html += `<td>${fmt(ind[n][p])}</td>`;
    html += "</tr>";
  }
  html += "</tbody></table>";
  $("result").innerHTML = html;
}

$("btn-cancel").addEventListener("click", async () => {
  if (!currentRunId) return;
  const r = await getJSON(`/api/run/${currentRunId}/cancel`, { method: "POST" });
  setStatus($("run-status"), r.note || "已请求取消", "");
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
  setStatus($("run-status"), r.note || "已释放运行槽位", "ok");
});

init();

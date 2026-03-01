const el = {
  todayDate: document.getElementById("todayDate"),
  preSales: document.getElementById("preSales"),
  afterSales: document.getElementById("afterSales"),
  openSettingsBtn: document.getElementById("openSettingsBtn"),
  closeSettingsBtn: document.getElementById("closeSettingsBtn"),
  settingsMask: document.getElementById("settingsMask"),
  scheduleTableBody: document.getElementById("scheduleTableBody"),
  weeklyTableBody: document.getElementById("weeklyTableBody"),
  excelFile: document.getElementById("excelFile"),
  importExcelBtn: document.getElementById("importExcelBtn"),
  addRowBtn: document.getElementById("addRowBtn"),
  saveScheduleBtn: document.getElementById("saveScheduleBtn"),
  exportExcelBtn: document.getElementById("exportExcelBtn"),
  webhookInput: document.getElementById("webhookInput"),
  notifyTimeInput: document.getElementById("notifyTimeInput"),
  notifyCountInput: document.getElementById("notifyCountInput"),
  mentionIdsInput: document.getElementById("mentionIdsInput"),
  timezoneInput: document.getElementById("timezoneInput"),
  saveSettingsBtn: document.getElementById("saveSettingsBtn"),
  testNotifyBtn: document.getElementById("testNotifyBtn"),
  toast: document.getElementById("toast"),
};

let toastTimer = null;

function showToast(message) {
  if (toastTimer) {
    window.clearTimeout(toastTimer);
  }
  el.toast.textContent = message;
  el.toast.classList.add("show");
  toastTimer = window.setTimeout(() => el.toast.classList.remove("show"), 2200);
}

async function api(path, options = {}) {
  const resp = await fetch(path, options);
  const data = await resp.json().catch(() => ({}));
  if (!resp.ok) {
    throw new Error(data.detail || data.reason || "请求失败");
  }
  return data;
}

function setSettingsOpen(open) {
  document.body.classList.toggle("settings-open", open);
}

function renderToday(today, timezone) {
  const now = new Date().toLocaleDateString("zh-CN", { timeZone: timezone || "Asia/Shanghai" });
  el.todayDate.textContent = `${now} (${timezone || "Asia/Shanghai"})`;

  if (!today) {
    el.preSales.textContent = "未安排";
    el.afterSales.textContent = "未安排";
    return;
  }

  el.preSales.textContent = today.pre_sales || "未安排";
  el.afterSales.textContent = today.after_sales || "未安排";
}

function createRow(row = { date: "", pre_sales: "", after_sales: "" }) {
  const tr = document.createElement("tr");
  tr.innerHTML = `
    <td><input type="date" value="${row.date || ""}" /></td>
    <td><input type="text" value="${row.pre_sales || ""}" /></td>
    <td><input type="text" value="${row.after_sales || ""}" /></td>
    <td><button class="btn ghost delete-row">删除</button></td>
  `;
  tr.querySelector(".delete-row").addEventListener("click", () => tr.remove());
  return tr;
}

function renderSchedule(rows) {
  el.scheduleTableBody.innerHTML = "";
  if (!rows.length) {
    const tr = document.createElement("tr");
    tr.innerHTML = "<td colspan='4'>暂无排班</td>";
    el.scheduleTableBody.appendChild(tr);
    return;
  }
  rows.forEach((row) => el.scheduleTableBody.appendChild(createRow(row)));
}

function renderWeeklyTemplates(templates) {
  el.weeklyTableBody.innerHTML = "";
  if (!templates.length) {
    const tr = document.createElement("tr");
    tr.innerHTML = "<td colspan='3'>暂无按周模板</td>";
    el.weeklyTableBody.appendChild(tr);
    return;
  }

  const weekdayName = {
    1: "周一",
    2: "周二",
    3: "周三",
    4: "周四",
    5: "周五",
    6: "周六",
    7: "周日",
  };

  templates.forEach((tpl) => {
    const tr = document.createElement("tr");
    tr.innerHTML = `
      <td>${weekdayName[Number(tpl.weekday)] || tpl.weekday || "--"}</td>
      <td>${tpl.pre_sales || "--"}</td>
      <td>${tpl.after_sales || "--"}</td>
    `;
    el.weeklyTableBody.appendChild(tr);
  });
}

function collectScheduleRows() {
  const rows = [];
  const trList = Array.from(el.scheduleTableBody.querySelectorAll("tr"));
  trList.forEach((tr) => {
    const inputs = tr.querySelectorAll("input");
    if (inputs.length < 3) {
      return;
    }
    const date = inputs[0].value;
    const pre_sales = inputs[1].value.trim();
    const after_sales = inputs[2].value.trim();
    if (!date && !pre_sales && !after_sales) {
      return;
    }
    rows.push({ date, pre_sales, after_sales });
  });
  return rows;
}

function fillSettings(settings) {
  el.webhookInput.value = settings.webhook_url || "";
  el.notifyTimeInput.value = settings.notify_time || "09:00";
  el.notifyCountInput.value = settings.notify_count || 0;
  el.mentionIdsInput.value = (settings.mention_userids || []).join("\n");
  el.timezoneInput.value = settings.timezone || "Asia/Shanghai";
}

function collectSettings() {
  return {
    webhook_url: el.webhookInput.value.trim(),
    notify_time: el.notifyTimeInput.value || "09:00",
    notify_count: Number(el.notifyCountInput.value || 0),
    mention_userids: el.mentionIdsInput.value
      .split(/[\n,\s]+/)
      .map((v) => v.trim())
      .filter(Boolean),
    timezone: el.timezoneInput.value.trim() || "Asia/Shanghai",
  };
}

async function loadToday() {
  const data = await api("/api/today");
  renderToday(data.today, data.timezone);
}

async function loadSettingsPage() {
  const [scheduleData, settingsData] = await Promise.all([api("/api/schedule"), api("/api/settings")]);
  renderSchedule(scheduleData.rows || []);
  renderWeeklyTemplates(scheduleData.weekly_templates || []);
  fillSettings(settingsData || {});
}

el.openSettingsBtn.addEventListener("click", async () => {
  try {
    await loadSettingsPage();
    setSettingsOpen(true);
  } catch (err) {
    showToast(err.message || "加载设置失败");
  }
});

el.closeSettingsBtn.addEventListener("click", () => setSettingsOpen(false));
el.settingsMask.addEventListener("click", () => setSettingsOpen(false));

el.addRowBtn.addEventListener("click", () => {
  const empty = el.scheduleTableBody.querySelector("td[colspan='4']");
  if (empty) {
    el.scheduleTableBody.innerHTML = "";
  }
  el.scheduleTableBody.appendChild(createRow());
});

el.saveScheduleBtn.addEventListener("click", async () => {
  try {
    const rows = collectScheduleRows();
    await api("/api/schedule", {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({ rows }),
    });
    await loadToday();
    showToast("排班已保存");
  } catch (err) {
    showToast(err.message || "保存排班失败");
  }
});

el.importExcelBtn.addEventListener("click", async () => {
  const file = el.excelFile.files?.[0];
  if (!file) {
    showToast("请先选择 Excel 文件");
    return;
  }

  const form = new FormData();
  form.append("file", file);

  try {
    const result = await api("/api/import-excel", { method: "POST", body: form });
    await loadSettingsPage();
    await loadToday();
    showToast(`导入成功：按日期 ${result.count ?? 0} 条，按周模板 ${result.weekly_count ?? 0} 条`);
  } catch (err) {
    showToast(err.message || "导入失败");
  }
});

el.exportExcelBtn.addEventListener("click", () => {
  window.location.href = "/api/export-excel";
});

el.saveSettingsBtn.addEventListener("click", async () => {
  try {
    await api("/api/settings", {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify(collectSettings()),
    });
    await loadToday();
    showToast("通知配置已保存");
  } catch (err) {
    showToast(err.message || "保存通知配置失败");
  }
});

el.testNotifyBtn.addEventListener("click", async () => {
  try {
    const result = await api("/api/notify/test", { method: "POST" });
    showToast(result.sent ? "测试通知发送成功" : `发送失败：${result.reason || "未知错误"}`);
  } catch (err) {
    showToast(err.message || "测试通知失败");
  }
});

loadToday().catch((err) => {
  showToast(err.message || "加载今日值班失败");
});

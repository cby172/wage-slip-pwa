const APP_VERSION = 2;
const SETTINGS_KEY = "payroll_rules";
const SINGLE_OWNER_ID = "00000000-0000-0000-0000-000000000001";

const defaultRules = {
  defaultBaseSalary: 3600,
  annualRaise: 200,
  salaryCap: 5000,
  warehouseAllowance: 400,
  paidLeaveLimit: 3,
  paidLeaveDays: 1,
  springDoubleDays: 30,
  fragmentThreshold: 10,
  payday: 5,
  springHolidayOverrides: { "2026": "2026-02-14" },
  schemaVersion: APP_VERSION
};

const initialWorkers = [
  {
    name: "赵师",
    hire_date: "2023-02-20",
    base_start_salary: 3600,
    base_salary_override: 4200,
    raise_strategy: "spring_holiday",
    is_warehouse_manager: true,
    active: true,
    work_note: "真实初始资料：入职 2023-02-20；当前底薪 4200；当前 5 个工人已统一改为每年过年放假日加薪；当前仓库管理员。",
    interruptions: []
  },
  {
    name: "范师",
    hire_date: "2019-04-30",
    base_start_salary: 3600,
    base_salary_override: 4200,
    raise_strategy: "spring_holiday",
    is_warehouse_manager: false,
    active: true,
    work_note: "真实初始资料：第一段工作 2019-04 到 2022-07，不精确到日，作为备注处理；2025-02-13 重新入职；当前底薪 4200。",
    interruptions: [{ start: "2022-08-01", end: "2025-02-12", note: "中断约 3 年" }]
  },
  {
    name: "王师",
    hire_date: "2023-11-29",
    base_start_salary: 3600,
    base_salary_override: 4000,
    raise_strategy: "spring_holiday",
    is_warehouse_manager: false,
    active: true,
    work_note: "真实初始资料：第一段工作 2023-11-29 到 2025-03-31；2026-04-01 重新入职；当前底薪 4000。",
    interruptions: [{ start: "2025-04-01", end: "2026-03-31", note: "中断约 1 年" }]
  },
  {
    name: "张师",
    hire_date: "2024-09-08",
    base_start_salary: 3600,
    base_salary_override: 4000,
    raise_strategy: "spring_holiday",
    is_warehouse_manager: false,
    active: true,
    work_note: "真实初始资料：入职 2024-09-08；当前底薪 4000。",
    interruptions: []
  },
  {
    name: "刘师",
    hire_date: "2025-07-11",
    base_start_salary: 3600,
    base_salary_override: 3800,
    raise_strategy: "spring_holiday",
    is_warehouse_manager: false,
    active: true,
    work_note: "真实初始资料：入职 2025-07-11；当前底薪 3800。",
    interruptions: []
  }
];

const sampleAttendance = `6月
赵师 9，/14，20/，22/，23/，27，预支1800元。
范师 15，30，
王师 19，
张师 28，
刘师 13，`;

const workerSlipNotes = {
  "赵师": "2.20满年&过年加薪",
  "范师": "2.13满年&过年加薪",
  "王师": "4.1满年&过年加薪",
  "张师": "9.8满年&过年加薪",
  "刘师": "7.11满年&过年加薪"
};

let supabaseClient = null;
let rules = { ...defaultRules };
let workers = [];
let attendanceMonths = [];
let latestResult = null;
let realtimeChannel = null;

const $ = (id) => document.getElementById(id);

function isConfigured() {
  const config = window.WAGE_APP_CONFIG || {};
  return Boolean(config.supabaseUrl && config.supabaseAnonKey);
}

function initSupabase() {
  if (!isConfigured() || !window.supabase) return null;
  return window.supabase.createClient(window.WAGE_APP_CONFIG.supabaseUrl, window.WAGE_APP_CONFIG.supabaseAnonKey);
}

function setStatus(text, mode = "") {
  const element = $("syncStatus");
  element.textContent = text;
  element.className = `status-pill ${mode}`.trim();
}

function toast(message) {
  const element = $("toast");
  element.textContent = message;
  element.classList.add("show");
  clearTimeout(toast.timer);
  toast.timer = setTimeout(() => element.classList.remove("show"), 1800);
}

function initTabs() {
  document.querySelectorAll(".tab").forEach((button) => {
    button.addEventListener("click", () => {
      document.querySelectorAll(".tab").forEach((item) => item.classList.remove("active"));
      document.querySelectorAll(".panel").forEach((item) => item.classList.remove("active"));
      button.classList.add("active");
      $(button.dataset.panel).classList.add("active");
    });
  });
}

function initMonthInputs() {
  const now = new Date();
  $("yearInput").value = now.getFullYear();
}

function bindEvents() {
  $("reloadTopBtn").addEventListener("click", loadAll);
  $("generateBtn").addEventListener("click", generatePayroll);
  $("slipList").addEventListener("click", copySingleSlip);
  $("workerForm").addEventListener("submit", saveWorker);
  $("resetWorkerBtn").addEventListener("click", resetWorkerForm);
  $("workerList").addEventListener("click", handleWorkerAction);
  $("saveRulesBtn").addEventListener("click", saveRules);
  $("seedBtn").addEventListener("click", seedWorkers);
  $("reloadBtn").addEventListener("click", loadAll);
  $("exportBtn").addEventListener("click", exportBackup);
}

async function boot() {
  initTabs();
  initMonthInputs();
  bindEvents();
  $("attendanceInput").value = sampleAttendance;

  if (!isConfigured()) {
    $("configAlert").hidden = false;
    setStatus("需配置", "warn");
    return;
  }

  supabaseClient = initSupabase();
  if (!supabaseClient) {
    $("configAlert").hidden = false;
    setStatus("未连接", "warn");
    return;
  }

  $("appView").hidden = false;
  setStatus("连接中", "warn");
  await loadAll();
}

async function loadAll() {
  setStatus("同步中", "warn");
  await ensureRules();
  await Promise.all([loadWorkers(), loadHistory()]);
  await ensureDefaultWorkers();
  subscribeRealtime();
  renderRules();
  renderWorkers();
  renderHistory();
  resetPayrollResult();
  setStatus("已同步", "ok");
}

async function ensureRules() {
  const { data, error } = await supabaseClient
    .from("app_settings")
    .select("value")
    .eq("key", SETTINGS_KEY)
    .maybeSingle();
  if (error) return handleSupabaseError(error);
  if (data?.value) {
    rules = { ...defaultRules, ...data.value };
    return;
  }
  rules = { ...defaultRules };
  const { error: insertError } = await supabaseClient
    .from("app_settings")
    .insert({ owner_id: SINGLE_OWNER_ID, key: SETTINGS_KEY, value: rules });
  if (insertError) handleSupabaseError(insertError);
}

async function loadWorkers() {
  const { data, error } = await supabaseClient
    .from("workers")
    .select("*")
    .order("active", { ascending: false })
    .order("name", { ascending: true });
  if (error) return handleSupabaseError(error);
  workers = data || [];
}

async function ensureDefaultWorkers() {
  if (workers.length > 0) return;
  const { error } = await supabaseClient.from("workers").insert(withOwner(initialWorkers));
  if (error) return handleSupabaseError(error);
  await loadWorkers();
  toast("已自动创建真实初始工人资料");
}

async function loadHistory() {
  const { data, error } = await supabaseClient
    .from("attendance_months")
    .select("*")
    .order("year", { ascending: false })
    .order("month", { ascending: false })
    .limit(24);
  if (error) return handleSupabaseError(error);
  attendanceMonths = data || [];
}

function subscribeRealtime() {
  if (realtimeChannel) return;
  realtimeChannel = supabaseClient
    .channel("payroll-live")
    .on("postgres_changes", { event: "*", schema: "public", table: "workers" }, async () => {
      await loadWorkers();
      renderWorkers();
    })
    .on("postgres_changes", { event: "*", schema: "public", table: "app_settings" }, async () => {
      await ensureRules();
      renderRules();
      renderWorkerStats();
    })
    .on("postgres_changes", { event: "*", schema: "public", table: "attendance_months" }, async () => {
      await loadHistory();
      renderHistory();
    })
    .subscribe();
}

function handleSupabaseError(error) {
  console.error(error);
  setStatus("异常", "warn");
  if (/JWT|auth|permission|policy|RLS/i.test(error.message)) {
    toast("没有权限，请运行免登录版 SQL");
  } else if (/fetch|network|paused|timeout|Failed/i.test(error.message)) {
    toast("数据库连接失败，可能项目暂停或网络异常");
  } else {
    toast(error.message || "Supabase 操作失败");
  }
}

function renderRules() {
  $("ruleDefaultBaseSalary").value = rules.defaultBaseSalary;
  $("ruleAnnualRaise").value = rules.annualRaise;
  $("ruleSalaryCap").value = rules.salaryCap;
  $("ruleWarehouseAllowance").value = rules.warehouseAllowance;
  $("rulePaidLeaveLimit").value = rules.paidLeaveLimit;
  $("rulePaidLeaveDays").value = rules.paidLeaveDays;
  $("ruleDoubleDays").value = rules.springDoubleDays;
  $("ruleFragmentThreshold").value = rules.fragmentThreshold;
  $("ruleSpringOverrides").value = formatSpringOverrides(rules.springHolidayOverrides || {});
}

async function saveRules() {
  rules = {
    ...rules,
    defaultBaseSalary: numberValue("ruleDefaultBaseSalary", 3600),
    annualRaise: numberValue("ruleAnnualRaise", 200),
    salaryCap: numberValue("ruleSalaryCap", 5000),
    warehouseAllowance: numberValue("ruleWarehouseAllowance", 400),
    paidLeaveLimit: numberValue("rulePaidLeaveLimit", 3),
    paidLeaveDays: numberValue("rulePaidLeaveDays", 1),
    springDoubleDays: numberValue("ruleDoubleDays", 30),
    fragmentThreshold: numberValue("ruleFragmentThreshold", 10),
    springHolidayOverrides: parseSpringOverrides($("ruleSpringOverrides").value),
    schemaVersion: APP_VERSION
  };
  const { error } = await supabaseClient
    .from("app_settings")
    .upsert({ owner_id: SINGLE_OWNER_ID, key: SETTINGS_KEY, value: rules }, { onConflict: "owner_id,key" });
  if (error) return handleSupabaseError(error);
  toast("规则已保存");
  renderWorkerStats();
}

function renderWorkers() {
  const list = $("workerList");
  list.innerHTML = workers.length ? "" : '<p class="meta">还没有工人。打开网页后系统会自动创建真实初始工人资料，也可以手动新增。</p>';
  workers.forEach((worker) => {
    const salary = calculateWorkerSalary(worker, getSelectedPeriod().end);
    const card = document.createElement("article");
    card.className = "worker-card";
    card.innerHTML = `
      <div class="worker-head">
        <div>
          <strong>${escapeHtml(worker.name)}</strong>
          <p class="meta">${worker.active ? "在职" : "停用"} · 底薪 ${money(salary.baseSalary)} · 津贴 ${money(salary.allowance)} · ${worker.raise_strategy === "spring_holiday" ? "过年放假日加薪" : "入职日满年加薪"}</p>
        </div>
        <div class="worker-actions">
          <button type="button" data-edit="${worker.id}">编辑</button>
          <button type="button" data-toggle="${worker.id}">${worker.active ? "停用" : "启用"}</button>
          <button class="danger" type="button" data-delete="${worker.id}">删除</button>
        </div>
      </div>
      <p class="meta">${escapeHtml(worker.work_note || "无备注")}</p>
    `;
    list.appendChild(card);
  });
  renderWorkerStats();
}

function renderWorkerStats() {
  const target = $("dbWorkerStats");
  if (!target) return;
  target.innerHTML = buildWorkerStatsHtml(workers);
}

function buildWorkerStatsHtml(workerList) {
  if (!workerList.length) return '<p class="meta">还没有工人资料。</p>';
  const asOf = new Date();
  const rows = workerList.map((worker) => {
    const salary = calculateWorkerSalary(worker, asOf);
    const workYears = effectiveWorkYears(worker, asOf);
    const idle = formatIdleTime(worker, asOf);
    return `<article class="worker-stat-row">
      <div class="stat-cell stat-name" data-label="姓名">${escapeHtml(worker.name)}</div>
      <div class="stat-cell stat-age" data-label="工龄"><strong>${formatYears(workYears)}</strong><span>年</span></div>
      <div class="stat-cell" data-label="当前底薪">${money(salary.baseSalary)} 元</div>
      <div class="stat-cell" data-label="津贴/合计">${money(salary.allowance)} / ${money(salary.baseSalary + salary.allowance)} 元</div>
      <div class="stat-cell" data-label="状态"><span class="mini-pill ${worker.active ? "ok" : "gray"}">${worker.active ? "在职" : "停用"}</span></div>
      <div class="stat-cell" data-label="仓管"><span class="mini-pill ${worker.is_warehouse_manager ? "accent" : "gray"}">${worker.is_warehouse_manager ? "是" : "否"}</span></div>
      <div class="stat-cell stat-idle" data-label="闲置时间">${escapeHtml(idle)}</div>
      <div class="stat-cell stat-formula" data-label="底薪计算公式">${escapeHtml(formatSalaryFormula(worker, salary, asOf))}</div>
    </article>`;
  }).join("");
  return `<div class="stats-meta">统计日期：${dateKey(asOf)}；工龄已扣除中断/闲置时间，保留 1 位小数。</div>
    <div class="worker-stat-table" role="table" aria-label="工人统计">
      <div class="worker-stat-head" role="row">
        <span>姓名</span><span>工龄</span><span>当前底薪</span><span>津贴/合计</span><span>状态</span><span>仓管</span><span>闲置时间</span><span>底薪计算公式</span>
      </div>
      ${rows}
    </div>`;
}

function effectiveWorkYears(worker, asOf) {
  return effectiveWorkDays(worker, asOf) / 365.2425;
}

function effectiveWorkDays(worker, asOf) {
  const hire = parseDate(worker.hire_date);
  if (!hire || asOf < hire) return 0;
  let days = Math.max(0, daysBetweenInclusive(hire, asOf));
  (worker.interruptions || []).forEach((item) => {
    const start = parseDate(item.start);
    const end = parseDate(item.end);
    if (!start || !end) return;
    const overlapStart = start > hire ? start : hire;
    const overlapEnd = end < asOf ? end : asOf;
    if (overlapEnd >= overlapStart) days -= daysBetweenInclusive(overlapStart, overlapEnd);
  });
  return Math.max(0, days);
}

function formatIdleTime(worker, asOf) {
  const parts = [];
  let totalDays = 0;
  (worker.interruptions || []).forEach((item) => {
    const start = parseDate(item.start);
    const end = parseDate(item.end);
    if (!start || !end) return;
    const cappedEnd = end < asOf ? end : asOf;
    if (cappedEnd < start) return;
    const days = daysBetweenInclusive(start, cappedEnd);
    totalDays += days;
    parts.push(`${item.start}至${item.end}${item.note ? `（${item.note}）` : ""}`);
  });
  if (!parts.length) return "无";
  return `合计 ${formatYears(totalDays / 365.2425)} 年；${parts.join("；")}`;
}

function formatSalaryFormula(worker, salary, asOf) {
  const cap = Number(rules.salaryCap);
  if (worker.base_salary_override !== null && worker.base_salary_override !== undefined && worker.base_salary_override !== "") {
    return `手动底薪 ${money(worker.base_salary_override)}，封顶 ${money(cap)} 后当前底薪 ${money(salary.baseSalary)}`;
  }
  const salaryDate = worker.raise_strategy === "spring_holiday" ? latestSpringStartBefore(asOf) : asOf;
  const fullYears = effectiveFullYears(worker, salaryDate);
  const startSalary = Number(worker.base_start_salary || rules.defaultBaseSalary);
  const beforeCap = startSalary + fullYears * Number(rules.annualRaise);
  const strategy = worker.raise_strategy === "spring_holiday" ? "过年放假日加薪" : "入职日满年加薪";
  return `${strategy}：${money(startSalary)} + ${fullYears}年 × ${money(rules.annualRaise)} = ${money(beforeCap)}，封顶 ${money(cap)} 后当前底薪 ${money(salary.baseSalary)}`;
}

function formatYears(value) {
  return (Math.round(Number(value || 0) * 10) / 10).toFixed(1);
}

async function saveWorker(event) {
  event.preventDefault();
  const id = $("workerId").value;
  const payload = {
    owner_id: SINGLE_OWNER_ID,
    name: $("workerName").value.trim(),
    hire_date: $("hireDate").value,
    base_start_salary: numberValue("baseStartSalary", rules.defaultBaseSalary),
    base_salary_override: $("baseSalaryOverride").value ? numberValue("baseSalaryOverride", 0) : null,
    raise_strategy: $("raiseStrategy").value,
    is_warehouse_manager: $("warehouseManager").checked,
    active: $("workerActive").checked,
    work_note: $("workNote").value.trim(),
    interruptions: parseInterruptions($("interruptionsInput").value)
  };
  if (!payload.name || !payload.hire_date) return toast("请填写姓名和入职日期");
  const request = id
    ? supabaseClient.from("workers").update(payload).eq("id", id)
    : supabaseClient.from("workers").insert(payload);
  const { error } = await request;
  if (error) return handleSupabaseError(error);
  resetWorkerForm();
  await loadWorkers();
  renderWorkers();
  toast("工人已保存");
}

function handleWorkerAction(event) {
  const editId = event.target.dataset.edit;
  const toggleId = event.target.dataset.toggle;
  const deleteId = event.target.dataset.delete;
  if (editId) return fillWorkerForm(editId);
  if (toggleId) return toggleWorker(toggleId);
  if (deleteId) return deleteWorker(deleteId);
}

function fillWorkerForm(id) {
  const worker = workers.find((item) => item.id === id);
  if (!worker) return;
  $("workerId").value = worker.id;
  $("workerName").value = worker.name;
  $("hireDate").value = worker.hire_date;
  $("baseStartSalary").value = worker.base_start_salary;
  $("baseSalaryOverride").value = worker.base_salary_override ?? "";
  $("raiseStrategy").value = worker.raise_strategy;
  $("warehouseManager").checked = worker.is_warehouse_manager;
  $("workerActive").checked = worker.active;
  $("workNote").value = worker.work_note || "";
  $("interruptionsInput").value = formatInterruptions(worker.interruptions || []);
  document.querySelector('[data-panel="panel-workers"]').click();
  window.scrollTo({ top: 0, behavior: "smooth" });
}

function resetWorkerForm() {
  $("workerForm").reset();
  $("workerId").value = "";
  $("baseStartSalary").value = rules.defaultBaseSalary;
  $("raiseStrategy").value = "anniversary";
  $("workerActive").checked = true;
}

async function toggleWorker(id) {
  const worker = workers.find((item) => item.id === id);
  if (!worker) return;
  const { error } = await supabaseClient.from("workers").update({ active: !worker.active }).eq("id", id);
  if (error) return handleSupabaseError(error);
  await loadWorkers();
  renderWorkers();
}

async function deleteWorker(id) {
  const worker = workers.find((item) => item.id === id);
  if (!worker || !confirm(`确定删除 ${worker.name}？已生成的工资条会保留姓名快照。`)) return;
  const { error } = await supabaseClient.from("workers").delete().eq("id", id);
  if (error) return handleSupabaseError(error);
  await loadWorkers();
  renderWorkers();
}

async function seedWorkers() {
  if (!confirm("补齐真实初始工人资料会新增缺失的工人，不会覆盖已有工人。继续？")) return;
  const existingNames = new Set(workers.map((worker) => worker.name));
  const payload = initialWorkers.filter((worker) => !existingNames.has(worker.name));
  if (!payload.length) return toast("真实初始工人资料已经齐全");
  const { error } = await supabaseClient.from("workers").insert(withOwner(payload));
  if (error) return handleSupabaseError(error);
  await loadWorkers();
  renderWorkers();
  toast("真实初始工人资料已补齐");
}

function resetPayrollResult() {
  latestResult = null;
  $("summaryGrid").innerHTML = "";
  $("slipList").innerHTML = '<p class="meta">粘贴当月请假记录后，点击“生成工资条”。</p>';
}

function generatePayroll(showToast = true) {
  if (!workers.length) {
    latestResult = null;
    $("summaryGrid").innerHTML = "";
    $("slipList").innerHTML = '<p class="meta">先添加工人，再生成工资条。</p>';
    return null;
  }
  const rawInput = $("attendanceInput").value.trim();
  const period = getSelectedPeriod(rawInput, false);
  if (!period) {
    latestResult = null;
    $("summaryGrid").innerHTML = "";
    $("slipList").innerHTML = '<p class="meta">请在记录里写月份，例如第一行写“6月”。</p>';
    toast("请在记录里写月份");
    return null;
  }
  latestResult = buildPayroll(period.year, period.month, rawInput);
  renderPayrollResult(latestResult);
  if (showToast) toast("已生成");
  return latestResult;
}

async function savePayroll() {
  const result = generatePayroll(false);
  if (!result) return;
  setStatus("保存中", "warn");
  const { data: monthData, error: monthError } = await supabaseClient
    .from("attendance_months")
    .upsert({
      owner_id: SINGLE_OWNER_ID,
      year: result.year,
      month: result.month,
      raw_input: $("attendanceInput").value.trim(),
      generated_text: result.allText,
      summary: result.summary
    }, { onConflict: "owner_id,year,month" })
    .select()
    .single();
  if (monthError) return handleSupabaseError(monthError);

  await supabaseClient.from("payslips").delete().eq("attendance_month_id", monthData.id);
  const payload = result.rows.map((row) => ({
    owner_id: SINGLE_OWNER_ID,
    attendance_month_id: monthData.id,
    worker_id: row.worker.id,
    worker_name: row.worker.name,
    year: result.year,
    month: result.month,
    gross_amount: row.gross,
    net_amount: row.net,
    advance_amount: row.record.advance,
    slip_text: row.slipText,
    payload: row
  }));
  const { error: slipError } = await supabaseClient.from("payslips").insert(payload);
  if (slipError) return handleSupabaseError(slipError);
  await loadHistory();
  renderHistory();
  setStatus("已同步", "ok");
  toast("工资条已保存");
}

function buildPayroll(year, month, rawInput) {
  const monthDays = daysInMonth(year, month);
  const segments = getPayrollSegments(year, month);
  const parsed = parseAttendance(rawInput, workers);
  const activeWorkers = workers.filter((worker) => worker.active);
  const activeByName = new Map(activeWorkers.map((worker) => [worker.name, worker]));
  const orderedWorkers = parsed.order
    .map((name) => activeByName.get(name))
    .filter(Boolean);
  const orderedNames = new Set(orderedWorkers.map((worker) => worker.name));
  const remainingWorkers = activeWorkers.filter((worker) => !orderedNames.has(worker.name));
  const rows = [...orderedWorkers, ...remainingWorkers].map((worker) => buildWorkerPayroll(worker, parsed.records.get(worker.name) || emptyRecord(worker.name), year, month, monthDays, segments));
  const totalGross = rows.reduce((sum, row) => sum + row.gross, 0);
  const totalNet = rows.reduce((sum, row) => sum + row.net, 0);
  const totalAdvance = rows.reduce((sum, row) => sum + row.record.advance, 0);
  const header = `${year}年${month}月份工资(${monthDays}天)`;
  const allText = [header, "", ...rows.flatMap((row) => [row.slipText, ""]), `总计：${money(totalGross)}元。`].join("\n");
  return {
    year,
    month,
    monthDays,
    rows,
    allText,
    summary: { count: rows.length, totalGross, totalNet, totalAdvance, segments }
  };
}

function buildWorkerPayroll(worker, record, year, month, monthDays, segments) {
  const salary = calculateWorkerSalary(worker, segments.main.end);
  const monthlySalary = salary.baseSalary + salary.allowance;
  const segmentDays = segments.all.reduce((sum, segment) => sum + segment.days, 0);
  const leaveStats = calculateLeaveDeduction(record, year, month, monthlySalary, monthDays, segments);
  const paidDays = Math.max(0, segmentDays - leaveStats.deductDays);
  const gross = Math.max(0, roundYuan((monthlySalary / monthDays) * paidDays));
  const net = Math.max(0, roundYuan(gross - record.advance));
  const lines = [];
  const note = getWorkerSlipNote(worker);
  lines.push(`@${worker.name}${money(monthlySalary)}(${note})`);
  lines.push(`${month}月份工资`);
  if (salary.allowance > 0) {
    lines.push(`(基本工资${money(salary.baseSalary)}+职位津贴${money(salary.allowance)}=${money(monthlySalary)})`);
  }
  if (segments.all.length > 1) {
    lines.push(`两段合并工资：${segments.all.map((segment) => `${formatMd(segment.start)}至${formatMd(segment.end)}共${segment.days}天`).join("；")}`);
  } else if (segmentDays !== monthDays) {
    lines.push(`本次计薪：${formatMd(segments.main.start)}至${formatMd(segments.main.end)}共${segmentDays}天`);
  }
  lines.push(formatLeaveLine(record, leaveStats));
  if (paidDays !== monthDays || leaveStats.deductDays > 0 || segmentDays !== monthDays) {
    lines.push(`${money(monthlySalary)}/${monthDays}×${formatDays(paidDays)}天=${money(gross)}元`);
  }
  if (record.advance > 0) {
    lines.push(`${money(gross)}-预支${money(record.advance)}=${money(net)}元`);
  }
  lines.push(`实发：${money(net)}元。`);
  return { worker, record, salary, monthlySalary, gross, net, paidDays, leaveStats, segments, slipText: lines.join("\n") };
}

function calculateLeaveDeduction(record, year, month, monthlySalary, monthDays, segments) {
  const daily = monthlySalary / monthDays;
  const spring = getSpringInfo(year);
  const doubleStart = addDays(spring.preHolidayStart, -rules.springDoubleDays);
  const doubleEnd = addDays(spring.preHolidayStart, -1);
  let regularLeave = 0;
  let doubleLeave = 0;
  record.items.forEach((item) => {
    const date = makeDate(year, month, item.day);
    if (date.getMonth() !== month - 1) return;
    if (!dateInSegments(date, segments.all)) return;
    if (date >= doubleStart && date <= doubleEnd) doubleLeave += item.amount;
    else regularLeave += item.amount;
  });
  const regularPaidRest = regularLeave > 0 ? Math.min(regularLeave, Number(rules.paidLeaveDays || 0)) : 0;
  const regularDeductDays = Math.max(regularLeave - regularPaidRest, 0);
  const regularDeduction = daily * regularDeductDays;
  const doubleDeduction = daily * doubleLeave * 2;
  return {
    regularLeave,
    doubleLeave,
    paidRest: regularPaidRest,
    regularDeductDays,
    deductDays: regularDeductDays + doubleLeave * 2,
    totalLeave: regularLeave + doubleLeave,
    totalDeduction: roundYuan(regularDeduction + doubleDeduction),
    doubleStart,
    doubleEnd
  };
}

function getPayrollSegments(year, month) {
  const mainStart = makeDate(year, month, 1);
  const mainEnd = makeDate(year, month, daysInMonth(year, month));
  let main = { start: mainStart, end: mainEnd, days: daysBetweenInclusive(mainStart, mainEnd), label: "整月工资" };
  const all = [];

  const springThisYear = getSpringInfo(year);
  if (sameMonth(springThisYear.preHolidayStart, mainStart)) {
    const days = springThisYear.preHolidayStart.getDate();
    main = { start: mainStart, end: springThisYear.preHolidayStart, days, label: "春节前工资" };
  }

  const previousMonthDate = addMonths(mainStart, -1);
  const springForPrevious = getSpringInfo(previousMonthDate.getFullYear());
  if (sameMonth(springForPrevious.resumeDate, previousMonthDate)) {
    const fragment = {
      start: springForPrevious.resumeDate,
      end: makeDate(previousMonthDate.getFullYear(), previousMonthDate.getMonth() + 1, daysInMonth(previousMonthDate.getFullYear(), previousMonthDate.getMonth() + 1)),
      label: "春节后零散工资"
    };
    fragment.days = daysBetweenInclusive(fragment.start, fragment.end);
    if (fragment.days <= rules.fragmentThreshold) all.push(fragment);
  }

  all.push(main);
  return { main, all };
}

function calculateWorkerSalary(worker, periodEnd) {
  const allowance = worker.is_warehouse_manager ? Number(rules.warehouseAllowance) : 0;
  if (worker.base_salary_override !== null && worker.base_salary_override !== undefined && worker.base_salary_override !== "") {
    return { baseSalary: Math.min(Number(worker.base_salary_override), Number(rules.salaryCap)), allowance, fullYears: null };
  }
  const asOf = worker.raise_strategy === "spring_holiday" ? latestSpringStartBefore(periodEnd) : periodEnd;
  const fullYears = effectiveFullYears(worker, asOf);
  const base = Math.min(Number(worker.base_start_salary || rules.defaultBaseSalary) + fullYears * Number(rules.annualRaise), Number(rules.salaryCap));
  return { baseSalary: base, allowance, fullYears };
}

function getWorkerSlipNote(worker) {
  if (workerSlipNotes[worker.name]) return workerSlipNotes[worker.name];
  if (worker.raise_strategy === "spring_holiday") return "过年放假日加薪";
  const hireDate = parseDate(worker.hire_date);
  return hireDate ? `${formatMd(hireDate)}满年加薪` : "满年加薪";
}

function effectiveFullYears(worker, asOf) {
  const hire = parseDate(worker.hire_date);
  if (!hire || asOf < hire) return 0;
  let days = Math.max(0, daysBetweenInclusive(hire, asOf));
  (worker.interruptions || []).forEach((item) => {
    const start = parseDate(item.start);
    const end = parseDate(item.end);
    if (!start || !end) return;
    const overlapStart = start > hire ? start : hire;
    const overlapEnd = end < asOf ? end : asOf;
    if (overlapEnd >= overlapStart) days -= daysBetweenInclusive(overlapStart, overlapEnd);
  });
  return Math.max(0, Math.floor(days / 365.2425));
}

function latestSpringStartBefore(date) {
  const thisYear = getSpringInfo(date.getFullYear()).preHolidayStart;
  if (thisYear <= date) return thisYear;
  return getSpringInfo(date.getFullYear() - 1).preHolidayStart;
}

function getSpringInfo(gregorianYear) {
  const overrideDate = rules.springHolidayOverrides?.[String(gregorianYear)];
  const lunarDates = [];
  for (let day = new Date(gregorianYear, 0, 1); day <= new Date(gregorianYear, 2, 10); day = addDays(day, 1)) {
    const lunar = getLunarParts(day);
    if (lunar.relatedYear === gregorianYear - 1 && lunar.month === "腊月") lunarDates.push({ date: new Date(day), day: lunar.day });
  }
  const hasLastMonth30 = lunarDates.some((item) => item.day === 30);
  const startDay = hasLastMonth30 ? 28 : 27;
  const preHolidayStart = lunarDates.find((item) => item.day === startDay)?.date;
  let resumeDate = null;
  for (let day = new Date(gregorianYear, 0, 1); day <= new Date(gregorianYear, 2, 10); day = addDays(day, 1)) {
    const lunar = getLunarParts(day);
    if (lunar.relatedYear === gregorianYear && lunar.month === "正月" && lunar.day === 8) {
      resumeDate = new Date(day);
      break;
    }
  }
  return {
    gregorianYear,
    hasLastMonth30,
    preHolidayStart: parseDate(overrideDate) || preHolidayStart || new Date(gregorianYear, 1, 1),
    resumeDate: resumeDate || new Date(gregorianYear, 1, 8)
  };
}

function getLunarParts(date) {
  const formatter = new Intl.DateTimeFormat("zh-u-ca-chinese", { year: "numeric", month: "long", day: "numeric" });
  const parts = formatter.formatToParts(date);
  return {
    relatedYear: Number(parts.find((part) => part.type === "relatedYear")?.value),
    month: parts.find((part) => part.type === "month")?.value,
    day: Number(parts.find((part) => part.type === "day")?.value)
  };
}

function parseAttendance(input, workerList) {
  const records = new Map();
  const order = [];
  const names = workerList.map((worker) => worker.name).sort((a, b) => b.length - a.length);
  input.split(/\r?\n/).forEach((line) => {
    const raw = line.trim();
    if (!raw || /^(?:20\d{2}\s*年\s*)?\d{1,2}\s*月份?$/.test(raw)) return;
    const name = names.find((item) => raw.startsWith(item));
    if (!name) return;
    if (!records.has(name)) order.push(name);
    records.set(name, parseWorkerLine(name, raw.slice(name.length), raw));
  });
  return { records, order };
}

function parseWorkerLine(name, body, raw) {
  const advanceMatch = body.match(/预支\s*([0-9]+(?:\.[0-9]+)?)\s*元?/);
  const advance = advanceMatch ? Number(advanceMatch[1]) : 0;
  const cleaned = body
    .replace(/预支\s*[0-9]+(?:\.[0-9]+)?\s*元?/g, "")
    .replace(/[，,、；;。.\s]+/g, " ")
    .trim();
  const items = [];
  cleaned.split(" ").forEach((token) => {
    if (!token) return;
    let match = token.match(/^\/(\d{1,2})$/);
    if (match) return items.push({ day: Number(match[1]), amount: 0.5, type: "morning" });
    match = token.match(/^(\d{1,2})\/$/);
    if (match) return items.push({ day: Number(match[1]), amount: 0.5, type: "afternoon" });
    match = token.match(/^(\d{1,2})$/);
    if (match) return items.push({ day: Number(match[1]), amount: 1, type: "full" });
  });
  return { name, raw, advance, items };
}

function emptyRecord(name) {
  return { name, raw: "", advance: 0, items: [] };
}

function formatLeaveLine(record, stats) {
  if (!record.items.length) return "请假：无";
  const full = record.items.filter((item) => item.type === "full").map((item) => item.day);
  const half = record.items.filter((item) => item.type === "morning" || item.type === "afternoon").map((item) => item.day);
  const parts = [];
  if (full.length) parts.push(`${full.join("/")}号${full.length > 1 ? "各" : ""}一天`);
  if (half.length) parts.push(`${half.join("/")}号${half.length > 1 ? "各" : ""}半天`);
  if (stats.totalLeave <= stats.paidRest && full.length === 1 && half.length === 0) {
    return `请假：${full[0]}号正常休息一天`;
  }
  const regularSummary = `(休${formatDays(stats.paidRest)}请${formatDays(stats.regularDeductDays)}天)`;
  const summaryPrefix = parts.length > 1 ? "，" : "";
  if (stats.doubleLeave > 0) {
    return `请假：${parts.join("，")}${summaryPrefix}${regularSummary.replace(")", `，双薪请${formatDays(stats.doubleLeave)}天)`)}`;
  }
  return `请假：${parts.join("，")}${summaryPrefix}${regularSummary}`;
}

function renderPayrollResult(result) {
  $("summaryGrid").innerHTML = `
    <div class="summary-item">人数<b>${result.summary.count}</b></div>
    <div class="summary-item">应发合计<b>${money(result.summary.totalGross)}</b></div>
    <div class="summary-item">实发合计<b>${money(result.summary.totalNet)}</b></div>
  `;
  const list = $("slipList");
  list.innerHTML = "";
  result.rows.forEach((row, index) => {
    const card = document.createElement("article");
    card.className = "slip-card";
    card.innerHTML = `
      <div class="slip-head">
        <div>
          <strong>${escapeHtml(row.worker.name)}</strong>
          <p class="meta">实发 ${money(row.net)} 元</p>
        </div>
        <div class="slip-actions">
          <button type="button" data-copy-slip="${index}">复制</button>
        </div>
      </div>
      <pre>${escapeHtml(row.slipText)}</pre>
    `;
    list.appendChild(card);
  });
}

async function copyAllSlips() {
  if (!latestResult) generatePayroll(false);
  if (!latestResult) return;
  await copyText(latestResult.allText);
}

async function copySingleSlip(event) {
  const index = event.target.dataset.copySlip;
  if (index === undefined || !latestResult) return;
  await copyText(latestResult.rows[Number(index)].slipText);
}

async function copyText(text) {
  try {
    await navigator.clipboard.writeText(text);
  } catch {
    const fallback = document.createElement("textarea");
    fallback.value = text;
    fallback.style.position = "fixed";
    fallback.style.left = "-9999px";
    document.body.appendChild(fallback);
    fallback.select();
    document.execCommand("copy");
    fallback.remove();
  }
  toast("已复制");
}

function renderHistory() {
  const list = $("historyList");
  list.innerHTML = attendanceMonths.length ? "" : '<p class="meta">还没有保存过工资条。</p>';
  attendanceMonths.forEach((item) => {
    const card = document.createElement("article");
    card.className = "history-card";
    card.innerHTML = `
      <div class="history-head">
        <div>
          <strong>${item.year}年${item.month}月工资</strong>
          <p class="meta">人数 ${item.summary?.count || 0} · 实发 ${money(item.summary?.totalNet || 0)} 元</p>
        </div>
        <div class="history-actions">
          <button type="button" data-load-history="${item.id}">载入</button>
        </div>
      </div>
    `;
    card.querySelector("[data-load-history]").addEventListener("click", () => {
      $("yearInput").value = item.year;
      $("attendanceInput").value = item.raw_input || "";
      document.querySelector('[data-panel="panel-payroll"]').click();
      generatePayroll();
    });
    list.appendChild(card);
  });
}

function exportBackup() {
  const backup = {
    exportedAt: new Date().toISOString(),
    schemaVersion: APP_VERSION,
    rules,
    workers,
    attendanceMonths
  };
  const blob = new Blob([JSON.stringify(backup, null, 2)], { type: "application/json" });
  const url = URL.createObjectURL(blob);
  const link = document.createElement("a");
  link.href = url;
  link.download = `工资管理备份_${new Date().toISOString().slice(0, 10)}.json`;
  document.body.appendChild(link);
  link.click();
  link.remove();
  URL.revokeObjectURL(url);
}

function getSelectedPeriod(rawInput = $("attendanceInput")?.value || "", allowFallback = true) {
  const parsed = parsePeriodFromInput(rawInput);
  if (parsed) {
    const year = parsed.year || Number($("yearInput").value) || new Date().getFullYear();
    return { year, month: parsed.month, start: makeDate(year, parsed.month, 1), end: makeDate(year, parsed.month, daysInMonth(year, parsed.month)) };
  }
  if (!allowFallback) return null;
  const now = new Date();
  const year = Number($("yearInput").value) || now.getFullYear();
  const month = now.getMonth() + 1;
  return { year, month, start: makeDate(year, month, 1), end: makeDate(year, month, daysInMonth(year, month)) };
}

function parsePeriodFromInput(input) {
  const match = String(input || "").match(/(?:^|\n)\s*(?:(20\d{2})\s*年\s*)?([1-9]|1[0-2])\s*月份?\s*(?:\n|$)/);
  if (!match) return null;
  return { year: match[1] ? Number(match[1]) : null, month: Number(match[2]) };
}

function parseInterruptions(text) {
  return text.split(/\r?\n/).map((line) => {
    const raw = line.trim();
    if (!raw) return null;
    const match = raw.match(/(\d{4}-\d{2}-\d{2})\s*(?:到|-|至)\s*(\d{4}-\d{2}-\d{2})(.*)/);
    if (!match) return { start: "", end: "", note: raw };
    return { start: match[1], end: match[2], note: match[3].trim() };
  }).filter(Boolean);
}

function formatInterruptions(items) {
  return (items || []).map((item) => `${item.start || ""} 到 ${item.end || ""} ${item.note || ""}`.trim()).join("\n");
}

function parseSpringOverrides(text) {
  const result = {};
  text.split(/\r?\n/).forEach((line) => {
    const match = line.trim().match(/^(\d{4})\s*[=：:]\s*(\d{4}-\d{2}-\d{2})$/);
    if (match) result[match[1]] = match[2];
  });
  return result;
}

function formatSpringOverrides(overrides) {
  return Object.entries(overrides || {})
    .sort(([a], [b]) => Number(a) - Number(b))
    .map(([year, date]) => `${year}=${date}`)
    .join("\n");
}

function withOwner(value) {
  if (Array.isArray(value)) return value.map((item) => ({ owner_id: SINGLE_OWNER_ID, ...item }));
  return { owner_id: SINGLE_OWNER_ID, ...value };
}

function numberValue(id, fallback) {
  const value = Number($(id).value);
  return Number.isFinite(value) ? value : fallback;
}

function makeDate(year, month, day) {
  return new Date(year, month - 1, day);
}

function parseDate(value) {
  if (!value) return null;
  const date = new Date(`${value}T00:00:00`);
  return Number.isNaN(date.getTime()) ? null : date;
}

function addDays(date, days) {
  const copy = new Date(date);
  copy.setDate(copy.getDate() + days);
  return copy;
}

function addMonths(date, months) {
  const copy = new Date(date);
  copy.setMonth(copy.getMonth() + months);
  return copy;
}

function daysInMonth(year, month) {
  return new Date(year, month, 0).getDate();
}

function daysBetweenInclusive(start, end) {
  return Math.floor((stripTime(end) - stripTime(start)) / 86400000) + 1;
}

function stripTime(date) {
  return new Date(date.getFullYear(), date.getMonth(), date.getDate());
}

function sameMonth(a, b) {
  return a.getFullYear() === b.getFullYear() && a.getMonth() === b.getMonth();
}

function dateInSegments(date, segments) {
  return segments.some((segment) => date >= segment.start && date <= segment.end);
}

function roundYuan(value) {
  return Math.round(Number(value));
}

function money(value) {
  const rounded = roundYuan(Number(value || 0));
  return String(rounded);
}

function formatDays(value) {
  return Number.isInteger(value) ? String(value) : String(value).replace(/\.0$/, "");
}

function formatMd(date) {
  return `${date.getMonth() + 1}.${date.getDate()}`;
}

function dateKey(date) {
  return `${date.getFullYear()}-${String(date.getMonth() + 1).padStart(2, "0")}-${String(date.getDate()).padStart(2, "0")}`;
}

function escapeHtml(value) {
  return String(value ?? "")
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;")
    .replace(/'/g, "&#039;");
}

boot();

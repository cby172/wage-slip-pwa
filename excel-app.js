const EXCEL_STORAGE_KEY = "wage_excel_workers_v1";
const EXCEL_APP_VERSION = 1;

const defaultRules = {
  defaultBaseSalary: 3600,
  annualRaise: 200,
  salaryCap: 5000,
  warehouseAllowance: 400,
  paidLeaveDays: 1,
  springDoubleDays: 30,
  fragmentThreshold: 10,
  springHolidayOverrides: { "2026": "2026-02-14" },
  schemaVersion: EXCEL_APP_VERSION
};

const sampleAttendance = `6月
赵师 9，/14，20/，22/，23/，27，预支1800元。
范师 15，30，
王师 19，
张师 28，
刘师 13，`;

const DEFAULT_EXCEL_FILE_NAME = "工人资料_2026-07-23版.xlsx";

const defaultWorkers = [
  {
    id: "excel-赵师",
    owner_id: "excel-local",
    name: "赵师",
    hire_date: "2023-02-20",
    base_start_salary: 3600,
    base_salary_override: 4200,
    raise_strategy: "spring_holiday",
    is_warehouse_manager: true,
    active: true,
    slip_note: "2.20满年&过年加薪",
    work_note: "当前仓库管理员",
    interruptions: []
  },
  {
    id: "excel-范师",
    owner_id: "excel-local",
    name: "范师",
    hire_date: "2019-04-30",
    base_start_salary: 3600,
    base_salary_override: 4200,
    raise_strategy: "spring_holiday",
    is_warehouse_manager: false,
    active: true,
    slip_note: "2.13满年&过年加薪",
    work_note: "第一段工作 2019-04 到 2022-07，2025-02-13 重新入职",
    interruptions: [{ start: "2022-08-01", end: "2025-02-12", note: "中断约 3 年" }]
  },
  {
    id: "excel-王师",
    owner_id: "excel-local",
    name: "王师",
    hire_date: "2023-11-29",
    base_start_salary: 3600,
    base_salary_override: 4000,
    raise_strategy: "spring_holiday",
    is_warehouse_manager: false,
    active: true,
    slip_note: "4.1满年&过年加薪",
    work_note: "第一段工作 2023-11-29 到 2025-03-31，2026-04-01 重新入职",
    interruptions: [{ start: "2025-04-01", end: "2026-03-31", note: "中断约 1 年" }]
  },
  {
    id: "excel-张师",
    owner_id: "excel-local",
    name: "张师",
    hire_date: "2024-09-08",
    base_start_salary: 3600,
    base_salary_override: 4000,
    raise_strategy: "spring_holiday",
    is_warehouse_manager: false,
    active: true,
    slip_note: "9.8满年&过年加薪",
    work_note: "",
    interruptions: []
  },
  {
    id: "excel-刘师",
    owner_id: "excel-local",
    name: "刘师",
    hire_date: "2025-07-11",
    base_start_salary: 3600,
    base_salary_override: 3800,
    raise_strategy: "spring_holiday",
    is_warehouse_manager: false,
    active: true,
    slip_note: "7.11满年&过年加薪",
    work_note: "",
    interruptions: []
  }
];

let rules = { ...defaultRules };
let workers = [];
let latestResult = null;

const $ = (id) => document.getElementById(id);

function boot() {
  initExcelTabs();
  $("yearInput").value = new Date().getFullYear();
  $("attendanceInput").value = sampleAttendance;
  $("excelInput").addEventListener("change", handleExcelUpload);
  $("generateBtn").addEventListener("click", generatePayroll);
  $("clearExcelBtn").addEventListener("click", clearSavedExcel);
  $("downloadCurrentBtn").addEventListener("click", downloadCurrentExcel);
  $("slipList").addEventListener("click", copySingleSlip);
  initDropUpload();
  loadSavedExcel();
}

async function handleExcelUpload(event) {
  const file = event.target.files?.[0];
  if (!file) return;
  await importExcelFile(file);
  event.target.value = "";
}

function initExcelTabs() {
  document.querySelectorAll(".excel-tabs .tab").forEach((button) => {
    button.addEventListener("click", () => {
      document.querySelectorAll(".excel-tabs .tab").forEach((item) => item.classList.remove("active"));
      document.querySelectorAll("#excel-panel-payroll, #excel-panel-data").forEach((item) => item.classList.remove("active"));
      button.classList.add("active");
      $(button.dataset.panel).classList.add("active");
    });
  });
}

function initDropUpload() {
  const zone = $("uploadZone");
  if (!zone) return;
  ["dragenter", "dragover"].forEach((eventName) => {
    zone.addEventListener(eventName, (event) => {
      event.preventDefault();
      zone.classList.add("dragging");
    });
  });
  ["dragleave", "drop"].forEach((eventName) => {
    zone.addEventListener(eventName, (event) => {
      event.preventDefault();
      zone.classList.remove("dragging");
    });
  });
  zone.addEventListener("drop", async (event) => {
    const file = event.dataTransfer?.files?.[0];
    if (!file) return;
    await importExcelFile(file);
  });
  zone.addEventListener("click", () => $("excelInput").click());
}

async function importExcelFile(file) {
  if (!/\.(xlsx|xls)$/i.test(file.name)) {
    toast("请上传 Excel 文件");
    return;
  }
  try {
    setStatus("读取中", "warn");
    const buffer = await file.arrayBuffer();
    const workbook = XLSX.read(buffer, { type: "array", cellDates: true });
    const parsed = parseWorkbook(workbook, file.name);
    workers = parsed.workers;
    rules = parsed.rules;
    localStorage.setItem(EXCEL_STORAGE_KEY, JSON.stringify(parsed));
    renderDataStatus(parsed);
    renderWorkers();
    resetPayrollResult();
    setStatus("已保存", "ok");
    toast("Excel 已识别并保存到本机");
  } catch (error) {
    setStatus("异常", "warn");
    toast(error.message || "Excel 识别失败");
  }
}

function parseWorkbook(workbook, fileName) {
  const workerSheet = workbook.Sheets["工人资料"];
  if (!workerSheet) throw new Error("Excel 缺少工作表：工人资料");
  const interruptionSheet = workbook.Sheets["中断记录"];
  const rulesSheet = workbook.Sheets["规则参数"];

  const workerRows = sheetRows(workerSheet);
  const interruptionRows = interruptionSheet ? sheetRows(interruptionSheet) : [];
  const parsedRules = rulesSheet ? parseRules(sheetRows(rulesSheet)) : { ...defaultRules };
  const interruptionsByName = parseInterruptionsSheet(interruptionRows);

  const parsedWorkers = workerRows
    .map((row, index) => parseWorkerRow(row, index + 2, interruptionsByName, parsedRules))
    .filter(Boolean);

  if (!parsedWorkers.length) throw new Error("工人资料表没有识别到工人");
  return {
    schemaVersion: EXCEL_APP_VERSION,
    sourceFile: fileName,
    savedAt: new Date().toISOString(),
    rules: parsedRules,
    workers: parsedWorkers
  };
}

function sheetRows(sheet) {
  return XLSX.utils.sheet_to_json(sheet, { defval: "", raw: false, dateNF: "yyyy-mm-dd" }).map((row) => {
    const normalized = {};
    Object.entries(row).forEach(([key, value]) => {
      normalized[String(key).trim()] = typeof value === "string" ? value.trim() : value;
    });
    return normalized;
  });
}

function parseWorkerRow(row, rowNumber, interruptionsByName, parsedRules) {
  const name = text(row["姓名"]);
  if (!name) return null;
  const hireDate = parseExcelDate(row["入职日期"]);
  if (!hireDate) throw new Error(`工人资料第 ${rowNumber} 行：入职日期不能为空或格式错误`);
  const raiseStrategy = parseRaiseStrategy(row["加薪策略"]);
  return {
    id: `excel-${name}`,
    owner_id: "excel-local",
    name,
    hire_date: hireDate,
    base_start_salary: numberOr(row["起始底薪"], parsedRules.defaultBaseSalary),
    base_salary_override: optionalNumber(row["手动底薪"]),
    raise_strategy: raiseStrategy,
    is_warehouse_manager: parseYesNo(row["是否仓管"]),
    active: parseYesNo(row["是否在职"], true),
    slip_note: text(row["工资条备注"]),
    work_note: text(row["备注"]),
    interruptions: interruptionsByName.get(name) || []
  };
}

function parseInterruptionsSheet(rows) {
  const result = new Map();
  rows.forEach((row, index) => {
    const name = text(row["姓名"]);
    if (!name) return;
    const start = parseExcelDate(row["中断开始"]);
    const end = parseExcelDate(row["中断结束"]);
    if (!start || !end) throw new Error(`中断记录第 ${index + 2} 行：中断开始/结束日期格式错误`);
    if (!result.has(name)) result.set(name, []);
    result.get(name).push({ start, end, note: text(row["备注"]) });
  });
  return result;
}

function parseRules(rows) {
  const next = { ...defaultRules, springHolidayOverrides: { ...defaultRules.springHolidayOverrides } };
  rows.forEach((row) => {
    const key = text(row["参数"]);
    if (!key) return;
    const value = text(row["值"]);
    if (key === "春节放假日期覆盖") {
      next.springHolidayOverrides = parseSpringOverrides(value);
    } else if (key in next) {
      next[key] = Number(value);
    }
  });
  return next;
}

function parseSpringOverrides(value) {
  const result = {};
  String(value || "").split(/[;\n；]+/).forEach((line) => {
    const match = line.trim().match(/^(\d{4})\s*[=：:]\s*(\d{4}-\d{2}-\d{2})$/);
    if (match) result[match[1]] = match[2];
  });
  return Object.keys(result).length ? result : { ...defaultRules.springHolidayOverrides };
}

function formatSpringOverrides(overrides) {
  return Object.entries(overrides || {})
    .sort(([a], [b]) => Number(a) - Number(b))
    .map(([year, value]) => `${year}=${value}`)
    .join("；");
}

function loadSavedExcel() {
  const saved = localStorage.getItem(EXCEL_STORAGE_KEY);
  if (!saved) {
    loadDefaultExcelData();
    localStorage.setItem(EXCEL_STORAGE_KEY, JSON.stringify({
      schemaVersion: EXCEL_APP_VERSION,
      sourceFile: DEFAULT_EXCEL_FILE_NAME,
      savedAt: new Date().toISOString(),
      isBuiltIn: true,
      rules,
      workers
    }));
    renderDataStatus({ sourceFile: DEFAULT_EXCEL_FILE_NAME, savedAt: new Date().toISOString(), isBuiltIn: true });
    renderWorkers();
    resetPayrollResult();
    setStatus("内置数据", "ok");
    return;
  }
  try {
    const parsed = JSON.parse(saved);
    workers = parsed.workers || [];
    rules = { ...defaultRules, ...(parsed.rules || {}) };
    renderDataStatus(parsed);
    renderWorkers();
    resetPayrollResult();
    setStatus("已加载", "ok");
  } catch {
    localStorage.removeItem(EXCEL_STORAGE_KEY);
    loadDefaultExcelData();
    renderDataStatus({ sourceFile: DEFAULT_EXCEL_FILE_NAME, savedAt: new Date().toISOString(), isBuiltIn: true });
    renderWorkers();
    resetPayrollResult();
    setStatus("内置数据", "ok");
  }
}

function loadDefaultExcelData() {
  workers = JSON.parse(JSON.stringify(defaultWorkers));
  rules = { ...defaultRules, springHolidayOverrides: { ...defaultRules.springHolidayOverrides } };
}

function renderDataStatus(data) {
  const savedAt = data.savedAt ? new Date(data.savedAt).toLocaleString("zh-CN") : "未知时间";
  const source = data.isBuiltIn ? "内置真实数据" : "上传保存数据";
  $("dataStatus").textContent = `当前使用：${source}；文件：${data.sourceFile || DEFAULT_EXCEL_FILE_NAME}；工人 ${workers.length} 人；保存时间：${savedAt}`;
}

function renderWorkers() {
  $("workerList").innerHTML = workers.length
    ? workers.map((worker) => {
      const salary = calculateWorkerSalary(worker, getSelectedPeriod(null, true).end);
      return `<article class="worker-card">
        <div class="worker-head">
          <div>
            <strong>${escapeHtml(worker.name)}</strong>
            <p class="meta">${worker.active ? "在职" : "停用"} · 底薪 ${money(salary.baseSalary)} · 津贴 ${money(salary.allowance)} · ${worker.raise_strategy === "spring_holiday" ? "过年放假日加薪" : "入职日满年加薪"}</p>
          </div>
        </div>
        <p class="meta">${escapeHtml(worker.work_note || "无备注")}</p>
      </article>`;
    }).join("")
    : '<p class="meta">还没有工人数据。</p>';
}

function clearSavedExcel() {
  if (!confirm("确定恢复为内置真实工人数据？本机上传保存的数据会被替换。")) return;
  localStorage.removeItem(EXCEL_STORAGE_KEY);
  loadDefaultExcelData();
  const data = {
    schemaVersion: EXCEL_APP_VERSION,
    sourceFile: DEFAULT_EXCEL_FILE_NAME,
    savedAt: new Date().toISOString(),
    isBuiltIn: true,
    rules,
    workers
  };
  localStorage.setItem(EXCEL_STORAGE_KEY, JSON.stringify(data));
  renderDataStatus(data);
  renderWorkers();
  resetPayrollResult();
  setStatus("内置数据", "ok");
  toast("已恢复内置真实数据");
}

function downloadCurrentExcel() {
  if (!workers.length) loadDefaultExcelData();
  const workbook = XLSX.utils.book_new();
  const workerRows = [["姓名", "入职日期", "起始底薪", "手动底薪", "加薪策略", "是否仓管", "是否在职", "工资条备注", "备注"]];
  workers.forEach((worker) => {
    workerRows.push([
      worker.name,
      worker.hire_date,
      worker.base_start_salary,
      worker.base_salary_override ?? "",
      worker.raise_strategy === "spring_holiday" ? "过年放假日加薪" : "入职日满年加薪",
      worker.is_warehouse_manager ? "是" : "否",
      worker.active ? "是" : "否",
      worker.slip_note || "",
      worker.work_note || ""
    ]);
  });
  XLSX.utils.book_append_sheet(workbook, XLSX.utils.aoa_to_sheet(workerRows), "工人资料");

  const interruptionRows = [["姓名", "中断开始", "中断结束", "备注"]];
  workers.forEach((worker) => {
    (worker.interruptions || []).forEach((item) => {
      interruptionRows.push([worker.name, item.start || "", item.end || "", item.note || ""]);
    });
  });
  XLSX.utils.book_append_sheet(workbook, XLSX.utils.aoa_to_sheet(interruptionRows), "中断记录");

  const ruleRows = [["参数", "值", "说明"]];
  ruleRows.push(["defaultBaseSalary", rules.defaultBaseSalary, "默认起始底薪"]);
  ruleRows.push(["annualRaise", rules.annualRaise, "每满一年加薪"]);
  ruleRows.push(["salaryCap", rules.salaryCap, "底薪封顶，不含职位津贴"]);
  ruleRows.push(["warehouseAllowance", rules.warehouseAllowance, "仓库管理员职位津贴"]);
  ruleRows.push(["paidLeaveDays", rules.paidLeaveDays, "普通月份有薪假天数"]);
  ruleRows.push(["springDoubleDays", rules.springDoubleDays, "春节前双薪天数"]);
  ruleRows.push(["fragmentThreshold", rules.fragmentThreshold, "春节前后零散工资并入阈值"]);
  ruleRows.push(["春节放假日期覆盖", formatSpringOverrides(rules.springHolidayOverrides), "如 2026=2026-02-14"]);
  XLSX.utils.book_append_sheet(workbook, XLSX.utils.aoa_to_sheet(ruleRows), "规则参数");

  const instructionRows = [
    ["说明", "内容"],
    ["固定字段", "不要修改工作表名称和第一行字段名，否则网页可能无法识别。"],
    ["日期格式", "日期建议使用 yyyy-mm-dd，例如 2026-04-01。"],
    ["上传保存", "上传 Excel 后会保存为当前浏览器本地最新版。"]
  ];
  XLSX.utils.book_append_sheet(workbook, XLSX.utils.aoa_to_sheet(instructionRows), "填写说明");
  XLSX.writeFile(workbook, `工人资料_${dateKey(new Date())}版.xlsx`);
}

function generatePayroll(showToast = true) {
  if (!workers.length) {
    resetPayrollResult("先上传工人资料 Excel，再生成工资条。");
    return null;
  }
  const rawInput = $("attendanceInput").value.trim();
  const period = getSelectedPeriod(rawInput, false);
  if (!period) {
    resetPayrollResult("请在记录里写月份，例如第一行写“6月”。");
    toast("请在记录里写月份");
    return null;
  }
  latestResult = buildPayroll(period.year, period.month, rawInput);
  renderPayrollResult(latestResult);
  if (showToast) toast("已生成");
  return latestResult;
}

function resetPayrollResult(message = "上传 Excel 并粘贴请假记录后，点击“生成工资条”。") {
  latestResult = null;
  $("summaryGrid").innerHTML = "";
  $("slipList").innerHTML = `<p class="meta">${escapeHtml(message)}</p>`;
}

function buildPayroll(year, month, rawInput) {
  const monthDays = daysInMonth(year, month);
  const segments = getPayrollSegments(year, month);
  const parsed = parseAttendance(rawInput, workers);
  const activeWorkers = workers.filter((worker) => worker.active);
  const activeByName = new Map(activeWorkers.map((worker) => [worker.name, worker]));
  const orderedWorkers = parsed.order.map((name) => activeByName.get(name)).filter(Boolean);
  const orderedNames = new Set(orderedWorkers.map((worker) => worker.name));
  const remainingWorkers = activeWorkers.filter((worker) => !orderedNames.has(worker.name));
  const rows = [...orderedWorkers, ...remainingWorkers].map((worker) => buildWorkerPayroll(worker, parsed.records.get(worker.name) || emptyRecord(worker.name), year, month, monthDays, segments));
  const totalGross = rows.reduce((sum, row) => sum + row.gross, 0);
  const totalNet = rows.reduce((sum, row) => sum + row.net, 0);
  const totalAdvance = rows.reduce((sum, row) => sum + row.record.advance, 0);
  const header = `${year}年${month}月份工资(${monthDays}天)`;
  const allText = [header, "", ...rows.flatMap((row) => [row.slipText, ""]), `总计：${money(totalGross)}元。`].join("\n");
  return { year, month, monthDays, rows, allText, summary: { count: rows.length, totalGross, totalNet, totalAdvance, segments } };
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
  if (salary.allowance > 0) lines.push(`(基本工资${money(salary.baseSalary)}+职位津贴${money(salary.allowance)}=${money(monthlySalary)})`);
  if (segments.all.length > 1) {
    lines.push(`两段合并工资：${segments.all.map((segment) => `${formatMd(segment.start)}至${formatMd(segment.end)}共${segment.days}天`).join("；")}`);
  } else if (segmentDays !== monthDays) {
    lines.push(`本次计薪：${formatMd(segments.main.start)}至${formatMd(segments.main.end)}共${segmentDays}天`);
  }
  lines.push(formatLeaveLine(record, leaveStats));
  if (paidDays !== monthDays || leaveStats.deductDays > 0 || segmentDays !== monthDays) lines.push(`${money(monthlySalary)}/${monthDays}×${formatDays(paidDays)}天=${money(gross)}元`);
  if (record.advance > 0) lines.push(`${money(gross)}-预支${money(record.advance)}=${money(net)}元`);
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
  return { regularLeave, doubleLeave, paidRest: regularPaidRest, regularDeductDays, deductDays: regularDeductDays + doubleLeave * 2, totalLeave: regularLeave + doubleLeave, totalDeduction: roundYuan(regularDeduction + doubleDeduction), doubleStart, doubleEnd };
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
    const fragment = { start: springForPrevious.resumeDate, end: makeDate(previousMonthDate.getFullYear(), previousMonthDate.getMonth() + 1, daysInMonth(previousMonthDate.getFullYear(), previousMonthDate.getMonth() + 1)), label: "春节后零散工资" };
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
  if (worker.slip_note) return worker.slip_note;
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
  return { gregorianYear, hasLastMonth30, preHolidayStart: parseDate(overrideDate) || preHolidayStart || new Date(gregorianYear, 1, 1), resumeDate: resumeDate || new Date(gregorianYear, 1, 8) };
}

function getLunarParts(date) {
  const formatter = new Intl.DateTimeFormat("zh-u-ca-chinese", { year: "numeric", month: "long", day: "numeric" });
  const parts = formatter.formatToParts(date);
  return { relatedYear: Number(parts.find((part) => part.type === "relatedYear")?.value), month: parts.find((part) => part.type === "month")?.value, day: Number(parts.find((part) => part.type === "day")?.value) };
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
  const cleaned = body.replace(/预支\s*[0-9]+(?:\.[0-9]+)?\s*元?/g, "").replace(/[，,、；;。.\s]+/g, " ").trim();
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
  if (stats.totalLeave <= stats.paidRest && full.length === 1 && half.length === 0) return `请假：${full[0]}号正常休息一天`;
  const regularSummary = `(休${formatDays(stats.paidRest)}请${formatDays(stats.regularDeductDays)}天)`;
  const summaryPrefix = parts.length > 1 ? "，" : "";
  if (stats.doubleLeave > 0) return `请假：${parts.join("，")}${summaryPrefix}${regularSummary.replace(")", `，双薪请${formatDays(stats.doubleLeave)}天)`)}`;
  return `请假：${parts.join("，")}${summaryPrefix}${regularSummary}`;
}

function renderPayrollResult(result) {
  $("summaryGrid").innerHTML = `<div class="summary-item">人数<b>${result.summary.count}</b></div><div class="summary-item">应发合计<b>${money(result.summary.totalGross)}</b></div><div class="summary-item">实发合计<b>${money(result.summary.totalNet)}</b></div>`;
  $("slipList").innerHTML = "";
  result.rows.forEach((row, index) => {
    const card = document.createElement("article");
    card.className = "slip-card";
    card.innerHTML = `<div class="slip-head"><div><strong>${escapeHtml(row.worker.name)}</strong><p class="meta">实发 ${money(row.net)} 元</p></div><div class="slip-actions"><button type="button" data-copy-slip="${index}">复制</button></div></div><pre>${escapeHtml(row.slipText)}</pre>`;
    $("slipList").appendChild(card);
  });
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

function text(value) {
  return String(value ?? "").trim();
}

function numberOr(value, fallback) {
  const number = Number(value);
  return Number.isFinite(number) ? number : Number(fallback);
}

function optionalNumber(value) {
  const raw = text(value);
  if (!raw) return null;
  const number = Number(raw);
  return Number.isFinite(number) ? number : null;
}

function parseYesNo(value, defaultValue = false) {
  const raw = text(value);
  if (!raw) return defaultValue;
  return ["是", "对", "true", "1", "在职"].includes(raw.toLowerCase());
}

function parseRaiseStrategy(value) {
  const raw = text(value);
  if (raw === "过年放假日加薪" || raw === "spring_holiday") return "spring_holiday";
  return "anniversary";
}

function parseExcelDate(value) {
  if (!value) return "";
  if (value instanceof Date && !Number.isNaN(value.getTime())) return dateKey(value);
  const raw = text(value).replace(/[./]/g, "-");
  const match = raw.match(/^(\d{4})-(\d{1,2})-(\d{1,2})$/);
  if (!match) return "";
  return `${match[1]}-${match[2].padStart(2, "0")}-${match[3].padStart(2, "0")}`;
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
  return String(roundYuan(Number(value || 0)));
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
  return String(value ?? "").replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;").replace(/"/g, "&quot;").replace(/'/g, "&#039;");
}

function setStatus(text, mode = "") {
  $("syncStatus").textContent = text;
  $("syncStatus").className = `status-pill ${mode}`.trim();
}

function toast(message) {
  const element = $("toast");
  element.textContent = message;
  element.classList.add("show");
  clearTimeout(toast.timer);
  toast.timer = setTimeout(() => element.classList.remove("show"), 1800);
}

boot();
